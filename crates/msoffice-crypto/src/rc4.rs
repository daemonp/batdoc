#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! RC4 as the binary formats use it: one keystream per block, keyed from the password
//! hash and a block number, over every key length [MS-OFFCRYPTO] admits.
//!
//! Two families share this module and differ only in how the per-block key is derived —
//! [`crate::rc4_cryptoapi`] (SHA-1, §2.3.5.2) and [`crate::rc4_office97`] (MD5,
//! §2.3.6.2). What they share lives here: the cipher over a wrapped key, the block loop
//! ([MS-DOC] §2.2.6.2-3: 512-byte blocks numbered from zero at the start of each stream;
//! [MS-XLS] §2.2.10: 1024-byte blocks likewise), and the verifier check, whose one
//! non-obvious rule — the keystream MUST NOT be reset between the verifier and its hash
//! ([MS-OFFCRYPTO] §2.3.5.6) — is easier to get right once than twice.
//!
//! The cipher is RustCrypto's `rc4` with its `zeroize` feature: the 256-byte key schedule
//! is a function of the key and is wiped on drop. The key itself arrives wrapped and is
//! read inside `with_secret` for exactly as long as the schedule takes to build.
//!
//! Behaviour ported from office-crypto `src/method/rc4.rs` (MIT) and msoffcrypto-tool
//! `msoffcrypto/method/rc4_cryptoapi.py` (MIT); see NOTICE.

use crate::error::Error;
use crate::sensitive::{DerivedKey, VerifierPlaintext};
use rc4::{consts::*, Key, KeyInit, Rc4, StreamCipher};
use secure_gate::{ConstantTimeEq, RevealSecret};

/// Per-block key derivation — the one thing the two RC4 families do differently.
pub(crate) trait BlockKeySchedule {
    /// The key for `block`: [MS-OFFCRYPTO] §2.3.5.2 or §2.3.6.2.
    fn block_key(&self, block: u32) -> DerivedKey;
}

/// An RC4 keystream under one block key.
pub(crate) struct Keystream(Box<dyn StreamCipher>);

impl Keystream {
    /// Build the keystream for `key`.
    ///
    /// The key lengths the formats can produce are 6..=16 bytes — RC4 CryptoAPI's
    /// `KeySize` of 48 to 128 bits in 8-bit steps — plus 16 for its zero-padded 40-bit
    /// key and for Office 97/2000 RC4, whose key is always 128 bits. 5 is admitted so a
    /// 40-bit key that somehow arrives unpadded is still a cipher and not a slice panic.
    /// Anything else is a bug in the schedule that produced it, and is refused rather
    /// than sliced: `GenericArray::from_slice` panics on a length mismatch, and the
    /// length reaches here from the file's `KeySize`.
    pub(crate) fn new(key: &DerivedKey) -> Result<Self, Error> {
        macro_rules! cipher_for {
            ($k:expr, $($n:literal => $size:ty),+ $(,)?) => {
                match $k.len() {
                    $( $n => Some(Box::new(Rc4::<$size>::new(Key::<$size>::from_slice($k)))
                        as Box<dyn StreamCipher>), )+
                    _ => None,
                }
            };
        }
        let cipher = key.with_secret(|k| {
            cipher_for!(k, 5 => U5, 6 => U6, 7 => U7, 8 => U8, 9 => U9, 10 => U10, 11 => U11,
                12 => U12, 13 => U13, 14 => U14, 15 => U15, 16 => U16)
        });
        cipher.map(Self).ok_or_else(|| {
            Error::BadParameters(format!(
                "an RC4 key of {} bytes is outside the 5..=16 the binary formats define",
                key.with_secret(|k| k.len())
            ))
        })
    }

    pub(crate) fn apply(&mut self, data: &mut [u8]) {
        self.0.apply_keystream(data);
    }
}

/// Decrypt `data` in place in `block_size`-byte blocks: block 0 at offset 0, and a fresh
/// keystream under a fresh key at every boundary.
///
/// The rekeying is the format's, not an optimisation — msoffcrypto's comment credits it
/// to wvDecrypt (`rc4_cryptoapi.py:82-87`) and [MS-DOC] §2.2.6.3 / [MS-XLS] §2.2.10
/// state it outright. The block loop is also why a stream must be decrypted from its
/// start even where its first bytes are written in the clear: the keystream position
/// counts them.
pub(crate) fn decrypt_in_blocks(
    schedule: &dyn BlockKeySchedule,
    data: &mut [u8],
    block_size: usize,
) -> Result<(), Error> {
    if block_size == 0 {
        return Err(Error::BadParameters(
            "an RC4 block size of zero is a bug in the caller".to_string(),
        ));
    }
    for (i, chunk) in data.chunks_mut(block_size).enumerate() {
        let block = u32::try_from(i).map_err(|_| {
            Error::BadParameters(
                "more than 2^32 RC4 blocks; the block number is a 32-bit field".to_string(),
            )
        })?;
        Keystream::new(&schedule.block_key(block))?.apply(chunk);
    }
    Ok(())
}

/// Decrypt `data` in place under the single key for `block`, from the start of its
/// keystream — PowerPoint's rule, where the block number is the persist object
/// identifier and the object is one block however long it is ([MS-PPT] §2.3.7).
pub(crate) fn decrypt_with_block(
    schedule: &dyn BlockKeySchedule,
    data: &mut [u8],
    block: u32,
) -> Result<(), Error> {
    Keystream::new(&schedule.block_key(block))?.apply(data);
    Ok(())
}

/// The password check — [MS-OFFCRYPTO] §2.3.5.6, and §2.3.6.4 for the MD5 family.
///
/// Under the block-0 key, decrypt `EncryptedVerifier` and then `EncryptedVerifierHash`
/// **from the same keystream** ("MUST NOT be reset between the two decryption
/// operations"), hash the verifier with `digest` — SHA-1 for CryptoAPI, MD5 for Office
/// 97/2000 — and compare. Both plaintexts are wrapped; the comparison happens inside
/// nested closures and is constant-time, the one idiom every comparison in this crate
/// uses (see the ct-eq record in the secure-gate skill).
pub(crate) fn verify_password(
    schedule: &dyn BlockKeySchedule,
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
    digest: fn(&[u8]) -> Vec<u8>,
) -> Result<(), Error> {
    let mut keystream = Keystream::new(&schedule.block_key(0))?;
    let verifier = VerifierPlaintext::new({
        let mut v = encrypted_verifier.to_vec();
        keystream.apply(&mut v);
        v
    });
    let verifier_hash = VerifierPlaintext::new({
        let mut h = encrypted_verifier_hash.to_vec();
        keystream.apply(&mut h);
        h
    });
    let matches = verifier.with_secret(|v| {
        // Wrapped like the value it is compared against: when the password is right,
        // the two are the same bytes. Both `digest` functions allocate exactly their
        // digest length, so the wrapper takes the `Vec` by move with nothing left behind.
        let computed = VerifierPlaintext::new(digest(v));
        computed.with_secret(|c| verifier_hash.with_secret(|h| c.as_slice().ct_eq(h)))
    });
    if matches {
        Ok(())
    } else {
        Err(Error::WrongPassword)
    }
}

#[cfg(test)]
mod tests {
    // The lint header this file opens with is about the parser, not the tests that
    // drive it: a test that cannot unwrap its own fixture is unreadable, and a panic
    // here is a failing test rather than a reachable one. `classify_tests.rs` says the
    // same thing at file scope for the same reason.
    #![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
    use super::*;

    /// The RC4 test vectors on the cipher's own Wikipedia page, as the `rc4` crate's
    /// README quotes them — the check that the dispatch above hands the right key length
    /// to the right cipher, for every length the formats can produce.
    #[test]
    fn keystream_matches_the_public_rc4_vectors_for_each_key_length() {
        for (key, repeat, plain, expected) in [
            (
                &b"Key"[..],
                3usize,
                &b"Plaintext"[..],
                &[0xBB, 0xF3, 0x16, 0xE8, 0xD9, 0x40, 0xAF, 0x0A, 0xD3][..],
            ),
            (
                &b"Wiki"[..],
                2,
                &b"pedia"[..],
                &[0x10, 0x21, 0xBF, 0x04, 0x20][..],
            ),
            (
                &b"Secret"[..],
                1,
                &b"Attack at dawn"[..],
                &[
                    0x45, 0xA0, 0x1F, 0x64, 0x5F, 0xC3, 0x5B, 0x38, 0x35, 0x52, 0x54, 0x4B, 0x9B,
                    0xF5,
                ][..],
            ),
        ] {
            // "Key" and "Wiki" are below the formats' 5-byte floor, so they are padded
            // by repetition to a length the dispatch accepts: RC4's key schedule cycles
            // the key, so "KeyKeyKey" (9 bytes) and "WikiWiki" (8) are the same schedules
            // as "Key" and "Wiki". "Secret" is 6 and goes in as it is.
            let key: Vec<u8> = key
                .iter()
                .copied()
                .cycle()
                .take(key.len() * repeat)
                .collect();
            let mut data = plain.to_vec();
            Keystream::new(&DerivedKey::new(key))
                .unwrap()
                .apply(&mut data);
            assert_eq!(data, expected);
        }
    }

    /// The dispatch refuses what it cannot key rather than panicking in
    /// `GenericArray::from_slice`.
    #[test]
    fn key_lengths_outside_the_formats_are_refused_not_panicked() {
        for len in [0usize, 1, 4, 17, 32, 256] {
            assert!(
                matches!(
                    Keystream::new(&DerivedKey::new(vec![0u8; len])),
                    Err(Error::BadParameters(_))
                ),
                "{len}-byte key must be refused"
            );
        }
        for len in 5..=16usize {
            assert!(
                Keystream::new(&DerivedKey::new(vec![0u8; len])).is_ok(),
                "{len}"
            );
        }
    }

    /// A schedule whose block key is the block number repeated: enough to show that the
    /// block loop rekeys at every boundary and that the keystream restarts there.
    struct Counting;
    impl BlockKeySchedule for Counting {
        fn block_key(&self, block: u32) -> DerivedKey {
            DerivedKey::new(vec![(block & 0xFF) as u8 + 1; 16])
        }
    }

    #[test]
    fn decrypt_in_blocks_rekeys_at_every_boundary() {
        let mut whole = vec![0u8; 100];
        decrypt_in_blocks(&Counting, &mut whole, 40).unwrap();

        // Each block, decrypted alone under its own key from a fresh keystream, must
        // reproduce the corresponding slice of the whole.
        for (i, range) in [(0u32, 0..40usize), (1, 40..80), (2, 80..100)] {
            let mut part = vec![0u8; range.len()];
            decrypt_with_block(&Counting, &mut part, i).unwrap();
            assert_eq!(&whole[range.clone()], &part[..], "block {i}");
        }
        // And block 1 is not block 0's keystream continued.
        let mut continued = vec![0u8; 80];
        decrypt_with_block(&Counting, &mut continued, 0).unwrap();
        assert_ne!(&whole[40..80], &continued[40..80]);
    }

    /// The verifier and its hash come off ONE keystream. A reader that restarts the
    /// cipher between the two decrypts a garbage hash and reports every password wrong
    /// — the shape of bug the spec's one MUST NOT in §2.3.5.6 exists to name.
    #[test]
    fn verifier_hash_is_decrypted_from_the_continued_keystream() {
        let digest: fn(&[u8]) -> Vec<u8> = |v| v.iter().rev().copied().collect();
        let verifier = [7u8; 16];
        let expected_hash = digest(&verifier);

        // Encrypt both under the block-0 keystream, continued.
        let mut ks = Keystream::new(&Counting.block_key(0)).unwrap();
        let mut ev = verifier;
        ks.apply(&mut ev);
        let mut evh = expected_hash.clone();
        ks.apply(&mut evh);

        assert!(verify_password(&Counting, &ev, &evh, digest).is_ok());

        // The negative control: the same hash encrypted from a RESTARTED keystream is
        // what a resetting reader would accept, and this one must refuse it.
        let mut evh_reset = expected_hash;
        Keystream::new(&Counting.block_key(0))
            .unwrap()
            .apply(&mut evh_reset);
        assert_ne!(evh, evh_reset);
        assert!(matches!(
            verify_password(&Counting, &ev, &evh_reset, digest),
            Err(Error::WrongPassword)
        ));
    }
}
