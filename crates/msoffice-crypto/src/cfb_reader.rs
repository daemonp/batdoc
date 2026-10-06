//! Capped reads of the CFB streams classification and decryption share.
//!
//! [`read_encryption_info`] is the detection path: open the container, take
//! `\EncryptionInfo`, refuse anything over `ENCRYPTION_INFO_READ_CAP`. [`read_cfb_streams`]
//! is the decrypt path, and also takes `\EncryptedPackage` under the payload ceiling.
//! Both map `cfb` failures onto [`Error`] so a hostile directory tree is an
//! error (or, for `classify`, `Unknown`), never a panic.

use crate::error::Error;
#[cfg(feature = "crypto-ops")]
use crate::limits::ENCRYPTED_PACKAGE_READ_CAP;
use crate::limits::ENCRYPTION_INFO_READ_CAP;
use std::io::{Cursor, Read};

/// Raw streams extracted from a CFB container.
#[cfg(feature = "crypto-ops")]
pub(crate) struct OfficeCfbStreams {
    /// Raw bytes of the `\EncryptionInfo` stream (including the 8-byte version header).
    pub(crate) encryption_info: Vec<u8>,
    /// Raw bytes of the `\EncryptedPackage` stream
    /// (8-byte LE u64 plaintext size followed by encrypted data).
    pub(crate) encrypted_package: Vec<u8>,
}

/// Open a CFB container from raw bytes and extract the two streams needed for decryption.
///
/// # Errors
/// [`Error::NotACfbFile`] if the container will not open,
/// [`Error::MissingStream`] if either stream is absent,
/// [`Error::BadParameters`] if either exceeds its cap —
/// [`ENCRYPTION_INFO_READ_CAP`] (1 MiB) or [`ENCRYPTED_PACKAGE_READ_CAP`]
/// ([`crate::limits::PAYLOAD_CEILING`], 1 GiB) — and [`Error::Io`] on a read
/// failure.
///
/// The `Io` arm is not decorative: both streams go through [`read_capped`], whose
/// `read_to_end` propagates an [`std::io::Error`] from a container whose FAT chain does
/// not lead where its directory entry claims. [`read_encryption_info`] documents the same
/// arm for the same helper, and the two must not disagree about a function they share.
#[cfg(feature = "crypto-ops")]
pub(crate) fn read_cfb_streams(data: &[u8]) -> Result<OfficeCfbStreams, Error> {
    read_cfb_streams_capped(data, ENCRYPTION_INFO_READ_CAP, ENCRYPTED_PACKAGE_READ_CAP)
}

/// [`read_cfb_streams`] with both caps supplied, so a test can exercise the refusal at a
/// size a test can afford.
///
/// The payload cap is 1 GiB, and `read_capped` refuses only after reading `cap + 1` bytes
/// — so proving it at its real value would cost the suite a gigabyte of allocation and
/// several seconds per run, to demonstrate an inequality. Parameterising instead proves
/// the two things that can actually regress: that the mechanism refuses and that
/// [`read_cfb_streams`] passes the *payload* cap to the *payload* stream. That the number
/// is the shared one is a `const` assertion beside the constants themselves in
/// `limits.rs`, per GH #10's pattern 5 — which is what this file's own test doc says.
///
/// This is the payload cap GH #10 deferred and GH #6 step 2 landed. Be clear about what it
/// buys, because the deferral note it replaces was right on the point: `cfb` bounds every
/// read by the directory entry's `stream_len` *and* by the real FAT chain, so a forged
/// length fails the read rather than allocating, and there is no amplification to stop.
/// What the cap bounds is **memory** — a decrypt holds ciphertext and plaintext at once —
/// and `ENCRYPTED_PACKAGE_READ_CAP` being an alias of `PAYLOAD_CEILING` rather than a
/// second number is what stops this read and `agile::decrypt_package`'s output buffer
/// disagreeing about what is too big.
#[cfg(feature = "crypto-ops")]
fn read_cfb_streams_capped(
    data: &[u8],
    info_cap: usize,
    package_cap: usize,
) -> Result<OfficeCfbStreams, Error> {
    let cursor = Cursor::new(data);
    let mut cfb = cfb::CompoundFile::open(cursor).map_err(|_| Error::NotACfbFile)?;

    let encryption_info = read_capped(
        cfb.open_stream("/EncryptionInfo")
            .map_err(|_| Error::MissingStream("EncryptionInfo"))?,
        info_cap,
        "EncryptionInfo",
    )?;

    let encrypted_package = read_capped(
        cfb.open_stream("/EncryptedPackage")
            .map_err(|_| Error::MissingStream("EncryptedPackage"))?,
        package_cap,
        "EncryptedPackage",
    )?;

    Ok(OfficeCfbStreams {
        encryption_info,
        encrypted_package,
    })
}

/// Open a CFB container and extract only `\EncryptionInfo`.
///
/// What [`crate::classify()`] needs: the payload is never read, so classifying a 200 MB
/// document costs the container header and one small stream. It lives beside
/// `read_cfb_streams` rather than inside `classify` so both spellings of "open the
/// container and take a stream" share one cap and one set of error mappings.
///
/// # Errors
/// [`Error::NotACfbFile`] if the container will not open,
/// [`Error::MissingStream`] if it holds no `EncryptionInfo`,
/// [`Error::BadParameters`] if that stream is larger than
/// [`ENCRYPTION_INFO_READ_CAP`], and [`Error::Io`] on a read failure.
pub(crate) fn read_encryption_info(data: &[u8]) -> Result<Vec<u8>, Error> {
    let cursor = Cursor::new(data);
    let mut cfb = cfb::CompoundFile::open(cursor).map_err(|_| Error::NotACfbFile)?;
    read_capped(
        cfb.open_stream("/EncryptionInfo")
            .map_err(|_| Error::MissingStream("EncryptionInfo"))?,
        ENCRYPTION_INFO_READ_CAP,
        "EncryptionInfo",
    )
}

/// Read at most `cap` bytes, and report anything longer as a bad parameter.
///
/// `take(cap + 1)` rather than `take(cap)`: the extra byte is what distinguishes "exactly
/// at the ceiling" from "over it" without ever allocating past it. The same idiom
/// `odf-crypto` uses for its `MANIFEST_READ_CAP` reads (`odf-crypto/src/classify.rs`).
///
/// `pub(crate)` for `legacy_container`, which reads a binary document's streams under
/// the same idiom and the same kind of cap.
pub(crate) fn read_capped(
    mut stream: impl Read,
    cap: usize,
    what: &'static str,
) -> Result<Vec<u8>, Error> {
    let mut buf = Vec::new();
    stream.by_ref().take(cap as u64 + 1).read_to_end(&mut buf)?;
    if buf.len() > cap {
        return Err(Error::BadParameters(format!(
            "the {what} stream is larger than {cap} bytes"
        )));
    }
    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    /// A CFB container holding one `/EncryptionInfo` stream of `len` bytes.
    ///
    /// The content is never parsed by anything these tests call — `read_encryption_info`
    /// hands back raw bytes — so zeros are enough and the only variable is the length.
    fn cfb_with_encryption_info(len: usize) -> Vec<u8> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut container = cfb::CompoundFile::create(&mut cursor).unwrap();
            let mut stream = container.create_stream("/EncryptionInfo").unwrap();
            stream.write_all(&vec![0u8; len]).unwrap();
            stream.flush().unwrap();
            container.flush().unwrap();
        }
        cursor.into_inner()
    }

    /// The cap is the one bound compiled into **every** configuration, because
    /// `classify` reaches it: this is the first read the crate performs on a file from
    /// outside, before a single version byte has been looked at. Every other bound in
    /// `limits` shipped with a test and this one did not.
    ///
    /// Both halves matter. The over-cap case proves the refusal exists; the
    /// exactly-at-cap case proves the ceiling is inclusive rather than off by one, which
    /// is the whole reason `read_capped` takes `cap + 1` bytes rather than `cap`.
    #[test]
    fn an_encryption_info_stream_over_the_cap_is_refused_and_the_cap_itself_is_not() {
        let over = cfb_with_encryption_info(ENCRYPTION_INFO_READ_CAP + 1);
        // The length, not the bytes: an accepted over-cap stream is a megabyte, and
        // `Vec<u8>`'s `Debug` would bury the assertion that failed under all of it.
        let got = read_encryption_info(&over).map(|b| b.len());
        assert!(
            matches!(&got, Err(Error::BadParameters(msg))
                if msg.contains("EncryptionInfo") && msg.contains("larger than")),
            "one byte over the cap must be refused, got: {got:?}"
        );

        // The control: the same container one byte shorter is accepted whole, so the
        // refusal above is attributable to the length and not to the synthetic file.
        let at = cfb_with_encryption_info(ENCRYPTION_INFO_READ_CAP);
        let bytes = read_encryption_info(&at).expect("exactly at the cap must be accepted");
        assert_eq!(bytes.len(), ENCRYPTION_INFO_READ_CAP);
    }

    /// A stream nowhere near the cap — the size every real file has; the agile fixture
    /// is 1 441 bytes — comes back verbatim. Without this the test above would still
    /// pass on a `read_capped` that refused everything over a much smaller number.
    #[test]
    fn an_ordinary_encryption_info_stream_is_returned_whole() {
        let data = cfb_with_encryption_info(1441);
        assert_eq!(read_encryption_info(&data).unwrap().len(), 1441);
    }

    /// A container with both streams, of independently chosen sizes.
    #[cfg(feature = "crypto-ops")]
    fn cfb_with_both(info_len: usize, package_len: usize) -> Vec<u8> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut container = cfb::CompoundFile::create(&mut cursor).unwrap();
            for (path, len) in [
                ("/EncryptionInfo", info_len),
                ("/EncryptedPackage", package_len),
            ] {
                let mut stream = container.create_stream(path).unwrap();
                stream.write_all(&vec![0u8; len]).unwrap();
                stream.flush().unwrap();
            }
            container.flush().unwrap();
        }
        cursor.into_inner()
    }

    /// The payload cap refuses an over-size `\EncryptedPackage`, and is applied to the
    /// payload rather than to the header.
    ///
    /// Exercised through `read_cfb_streams_capped` at 4 KiB rather than at the real 1 GiB:
    /// `read_capped` refuses only after reading `cap + 1` bytes, so proving the real figure
    /// would cost the suite a gigabyte of allocation per run to demonstrate an inequality.
    /// What can regress is the mechanism and the wiring, and both are here — the third
    /// case would pass if `read_cfb_streams` handed the payload cap to the header stream,
    /// which is the transposition a reader of that function cannot see.
    ///
    /// The real number is covered by the `const` assertion beside the constants
    /// themselves in `limits.rs`, which is a compile error rather than a test failure.
    #[cfg(feature = "crypto-ops")]
    #[test]
    fn an_encrypted_package_over_the_payload_cap_is_refused() {
        const CAP: usize = 4096;
        let info_cap = ENCRYPTION_INFO_READ_CAP;

        // One byte over: refused, and the message names the payload stream rather than
        // the header, so the cap that fired is identifiable from the error alone.
        let over = cfb_with_both(1441, CAP + 1);
        let got = read_cfb_streams_capped(&over, info_cap, CAP).map(|s| s.encrypted_package.len());
        assert!(
            matches!(&got, Err(Error::BadParameters(msg))
                if msg.contains("EncryptedPackage") && msg.contains("larger than")),
            "one byte over the payload cap must be refused, got: {got:?}"
        );

        // Exactly at the cap: accepted whole. The control that makes the refusal
        // attributable to the length rather than to the synthetic container.
        let at = cfb_with_both(1441, CAP);
        let streams = read_cfb_streams_capped(&at, info_cap, CAP)
            .expect("exactly at the payload cap must be accepted");
        assert_eq!(streams.encrypted_package.len(), CAP);
        assert_eq!(streams.encryption_info.len(), 1441);

        // The caps are not interchangeable: a payload comfortably inside its own cap is
        // still accepted when the *header* cap is far smaller than it. Without this, a
        // `read_cfb_streams` that passed `info_cap` to both streams would pass every
        // assertion above.
        let asymmetric = cfb_with_both(64, CAP);
        let streams = read_cfb_streams_capped(&asymmetric, 128, CAP)
            .expect("the payload cap governs the payload, not the header cap");
        assert_eq!(streams.encrypted_package.len(), CAP);
    }
}
