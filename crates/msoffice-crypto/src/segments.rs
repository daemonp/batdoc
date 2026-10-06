//! The 4096-byte segment layout the agile package is encrypted in.
//!
//! [MS-OFFCRYPTO] §2.3.4.15 encrypts `EncryptedPackage` in 4096-byte segments, each under
//! its own IV derived from the segment's index:
//!
//! ```text
//! iv(i) = H(keyData.saltValue || LE32(i))[..keyData.blockSize]
//! ```
//!
//! # Why this is a module and not a `for` loop
//!
//! Plan design **D4**. The segmentation is the same on both sides of the format — herumi
//! rounds the whole plaintext up to a 16-byte multiple and cuts 4096-byte segments off it
//! (`include/encode.hpp:74-97`, `EncContent`), and the decrypt path cuts the same segments
//! back out of the ciphertext — but the two are easy to write twice and hard to notice
//! disagreeing, because the writer and the reader here are the same crate and would agree
//! with each other while both being wrong.
//!
//! That is not hypothetical. The IV derivation in `decrypt_package` hardcoded SHA-512 and
//! a 16-byte truncation until GH #11, so a file naming any other `keyData/@hashAlgorithm`
//! decrypted under the wrong IVs *after* `integrity::verify` had already reported
//! `Verified` under the right ones. One place to derive an IV is what makes that class of
//! bug a single edit rather than a search.
//!
//! # Padding
//!
//! The final segment is zero-padded up to an AES block boundary. On the encrypt side this
//! is a real choice and it matches herumi, whose `data.resize(RoundUp(size, 16))` pads with
//! `'\0'` before segmenting; the reader recovers the original length from the stream's
//! 8-byte prefix and truncates, so the pad bytes never reach a caller. On the decrypt side
//! a well-formed final segment is already a block multiple — AES-CBC preserves length — so
//! the padding is defensive, and it is what keeps a truncated file an error from the cipher
//! rather than a panic in it.
//!
//! Rounding the whole buffer once, as herumi does, and rounding only the final segment, as
//! this does, are the same operation: 4096 is a multiple of 16, so no interior segment can
//! be short.

use crate::error::Error;
use crate::hash::{derive_iv, HashAlgorithm};

/// [MS-OFFCRYPTO] §2.3.4.15 — the agile package segment length.
///
/// A format constant, not a bound: it is fixed by the spec rather than declared by the
/// file, which is why it lives here beside the code that cuts on it rather than in
/// [`crate::limits`]. Same rule that keeps the 16-byte AES block and the 72-byte
/// `EncryptionVerifier` out of that module.
pub(crate) const SEGMENT_LEN: usize = 4096;

/// AES block size in bytes — the multiple the final segment is padded up to.
const AES_BLOCK_LEN: usize = 16;

/// One segment of the package, with the IV that belongs to it.
pub(crate) struct Segment<'a> {
    /// The segment's bytes as stored — up to [`SEGMENT_LEN`], shorter only for the last.
    pub(crate) bytes: &'a [u8],
    /// `H(salt || LE32(index))[..block_size]`.
    ///
    /// The index itself is not carried: decrypt does not need it, encrypt gets it folded
    /// into this value, and a test recovers it from `enumerate()`. A field kept for a
    /// hypothetical caller is the `allow(dead_code)` argument wearing a different hat.
    ///
    /// Not wrapped: an IV is not key material, and on the encrypt side it is derived from
    /// a salt the file publishes in the clear.
    pub(crate) iv: Vec<u8>,
}

impl Segment<'_> {
    /// The segment's bytes padded up to an AES block multiple with zeros.
    ///
    /// Allocates because both callers hand the result to a cipher that needs an owned
    /// buffer anyway. A segment already on a block boundary — every segment but the last,
    /// and the last one too in any well-formed file — copies and does not grow.
    pub(crate) fn padded(&self) -> Vec<u8> {
        let mut padded = self.bytes.to_vec();
        let rem = padded.len() % AES_BLOCK_LEN;
        if rem != 0 {
            padded.resize(padded.len() + (AES_BLOCK_LEN - rem), 0);
        }
        padded
    }
}

/// The segments of one payload, in order, each carrying its own IV.
///
/// Borrows the payload rather than owning it: neither direction needs to modify the input,
/// and a 1 GiB package is exactly the thing not to copy on the way in.
pub(crate) struct Segments<'a> {
    data: &'a [u8],
    salt: &'a [u8],
    hash: HashAlgorithm,
    block_size: usize,
    next_index: u32,
}

impl<'a> Segments<'a> {
    /// Segment `data` under `<keyData>`'s hash, salt and block size.
    ///
    /// Both arguments that size something are checked here rather than per segment, so a
    /// bad one is reported once, before any hashing, and with a message naming the field.
    ///
    /// # Errors
    ///
    /// [`Error::BadParameters`] if `block_size` exceeds the digest the named
    /// hash produces — the IV would be a slice past the end — or if `data` would need more
    /// than `u32::MAX` segments.
    ///
    /// The segment-count check is what makes the `index.to_le_bytes()` in `next` a fact
    /// rather than a hope: the counter is a `u32` because the format says the IV suffix is
    /// LE32, so a payload with more segments than that would silently wrap and reuse every
    /// IV from zero. It is unreachable through this crate's own callers, since
    /// [`crate::limits::PAYLOAD_CEILING`] caps the payload at 1 GiB and therefore at
    /// 262 144 segments, but the cap and this type are separate concerns and a later
    /// streaming caller may not carry the first one.
    pub(crate) fn new(
        data: &'a [u8],
        hash: HashAlgorithm,
        salt: &'a [u8],
        block_size: usize,
    ) -> Result<Self, Error> {
        if block_size > hash.digest_len() {
            return Err(Error::BadParameters(format!(
                "keyData blockSize {} exceeds the {} digest length {}",
                block_size,
                hash.name(),
                hash.digest_len()
            )));
        }
        if data.len().div_ceil(SEGMENT_LEN) > u32::MAX as usize {
            return Err(Error::BadParameters(format!(
                "a {}-byte payload needs more than u32::MAX {SEGMENT_LEN}-byte segments",
                data.len()
            )));
        }
        Ok(Self {
            data,
            salt,
            hash,
            block_size,
            next_index: 0,
        })
    }

    /// How many segments this will yield.
    pub(crate) fn len(&self) -> usize {
        self.data.len().div_ceil(SEGMENT_LEN)
    }
}

impl<'a> Iterator for Segments<'a> {
    /// `Result`, even though [`Segments::new`] has already checked what can go wrong.
    ///
    /// The constructor's check is for the message; this is so that no future edit to
    /// [`derive_iv`] can turn a violated invariant into a slice panic. Yielding the error
    /// costs a caller one `?` and costs a hostile file its denial of service.
    type Item = Result<Segment<'a>, Error>;

    fn next(&mut self) -> Option<Self::Item> {
        let start = (self.next_index as usize).checked_mul(SEGMENT_LEN)?;
        if start >= self.data.len() {
            return None;
        }
        let end = self.data.len().min(start + SEGMENT_LEN);
        let index = self.next_index;
        self.next_index += 1;

        let iv = match derive_iv(self.hash, self.salt, &index.to_le_bytes(), self.block_size) {
            Ok(iv) => iv,
            Err(e) => return Some(Err(e)),
        };
        Some(Ok(Segment {
            bytes: &self.data[start..end],
            iv,
        }))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.len().saturating_sub(self.next_index as usize);
        (remaining, Some(remaining))
    }
}

#[cfg(test)]
#[path = "segments_tests.rs"]
mod tests;
