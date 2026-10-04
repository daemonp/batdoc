//! Bridge to `msoffice-crypto` for password-protected Office documents.
//!
//! Detection uses the crate's infallible `classify`; decryption uses
//! `decrypt_ooxml_with_policy` under the default integrity policy, which
//! verifies the package HMAC. This is why a wrong password is reported as
//! [`BatdocError::IncorrectPassword`] rather than surfacing as garbage bytes.
//!
//! On `wasm32-unknown-unknown` the crypto dependency is not in the graph (it
//! does not compile there), so this module is stubs instead: Office decryption
//! is unavailable, neither predicate ever reports `true`, and an encrypted
//! package falls through to the unrecognised-format error. Only PDF password
//! support applies there. The stubs keep every call site target-agnostic.

use crate::error::{BatdocError, Result};
#[cfg(not(target_arch = "wasm32"))]
use std::panic::{self, AssertUnwindSafe};

/// Ceiling for decrypted Office payloads (mirrors the CLI's 256 MiB input
/// cap) so a small ciphertext cannot expand into an unbounded allocation.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) const MAX_DECRYPTED_BYTES: usize = 256 * 1024 * 1024;

/// True when `data` is an encrypted OOXML package this build can decrypt —
/// the narrow question [`decrypt`] answers.
///
/// The `Document::OoxmlPackage` check is load-bearing: `Classification::is_supported`
/// is family-level and answers "would a decrypt path exist if `legacy-binary`
/// were enabled", so it is also `true` for an encrypted legacy `.doc`/`.xls`/`.ppt`
/// (families `Rc4CryptoApi`/`Rc4`/`XorObfuscation`), which this build does not
/// enable. Those must keep classifying as `Format::Doc`/`Format::Xls` so
/// extraction can report them unsupported instead of asking for a password that
/// can never work; use [`is_encrypted`] to report them.
///
/// `msoffice_crypto::classify` never panics and never fails: unreadable input
/// collapses to `Family::Unknown`, which `is_encrypted` reports as `false`.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn is_encrypted_office(data: &[u8]) -> bool {
    if !msoffice_crypto::is_cfb_office(data) {
        return false;
    }
    let class = msoffice_crypto::classify(data);
    matches!(class.document, msoffice_crypto::Document::OoxmlPackage)
        && class.is_encrypted()
        && class.is_supported()
}

/// True for any encrypted MS-OFFCRYPTO container, including legacy families
/// this build cannot decrypt. `needs_password` uses this; detection uses
/// [`is_encrypted_office`], which is narrower.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn is_encrypted(data: &[u8]) -> bool {
    msoffice_crypto::is_cfb_office(data) && msoffice_crypto::classify(data).is_encrypted()
}

/// Decrypt an encrypted Office package to its plain OOXML bytes.
///
/// # Errors
///
/// [`BatdocError::IncorrectPassword`] for a bad password,
/// [`BatdocError::UnsupportedEncryption`] for an unimplemented algorithm,
/// [`BatdocError::Document`] for corruption or an over-limit payload.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn decrypt(data: &[u8], password: &str) -> Result<Vec<u8>> {
    let outcome = panic::catch_unwind(AssertUnwindSafe(|| {
        msoffice_crypto::decrypt_ooxml_with_policy(
            data,
            password,
            msoffice_crypto::IntegrityPolicy::RequireWhereDefined,
        )
    }));
    match outcome {
        Ok(Ok(decrypted)) => {
            if decrypted.package.len() > MAX_DECRYPTED_BYTES {
                return Err(BatdocError::Document(
                    "decrypted Office package exceeds the 256 MiB limit".into(),
                ));
            }
            Ok(decrypted.package)
        }
        Ok(Err(e)) => Err(map_error(e)),
        Err(_) => Err(BatdocError::Document(
            "Office decryption panicked (malformed document)".into(),
        )),
    }
}

/// Map a `msoffice_crypto::Error` to the batdoc error surface.
///
/// `msoffice_crypto::Error` is `#[non_exhaustive]`, so the wildcard arm is
/// required. Messages are this crate's own; the source message is only
/// interpolated for the generic `Document` case and never contains key
/// material or the password.
#[cfg(not(target_arch = "wasm32"))]
fn map_error(e: msoffice_crypto::Error) -> BatdocError {
    use msoffice_crypto::Error as E;
    match e {
        E::WrongPassword => BatdocError::IncorrectPassword,
        E::UnsupportedAlgorithm { what, name } => {
            BatdocError::UnsupportedEncryption(format!("{what} names {name}"))
        }
        E::UnsupportedEncryptionVersion(major, minor) => {
            BatdocError::UnsupportedEncryption(format!("MS-OFFCRYPTO version {major}.{minor}"))
        }
        other => BatdocError::Document(format!("Office decryption failed: {other}")),
    }
}

/// Always `false` on wasm: the crypto dependency does not compile there, so
/// an encrypted Office package cannot be recognised or opened in this build.
// Not `const`: signature parity with the native variant.
#[allow(clippy::missing_const_for_fn)]
#[cfg(target_arch = "wasm32")]
pub(crate) fn is_encrypted_office(_data: &[u8]) -> bool {
    false
}

/// Always `false` on wasm, for the same reason as [`is_encrypted_office`]:
/// without the crypto dependency an Office container cannot be classified at
/// all, so no password can help.
// Not `const`: signature parity with the native variant.
#[allow(clippy::missing_const_for_fn)]
#[cfg(target_arch = "wasm32")]
pub(crate) fn is_encrypted(_data: &[u8]) -> bool {
    false
}

/// Always [`BatdocError::UnsupportedEncryption`] on wasm: Office decryption
/// is unavailable because the crypto dependency does not compile for
/// `wasm32-unknown-unknown`. Unreachable through [`is_encrypted_office`],
/// which is always `false` here; it exists so call sites stay
/// target-agnostic.
#[cfg(target_arch = "wasm32")]
pub(crate) fn decrypt(_data: &[u8], _password: &str) -> Result<Vec<u8>> {
    Err(BatdocError::UnsupportedEncryption(
        "encrypted Office documents are not supported in this build".into(),
    ))
}

#[cfg(all(test, not(target_arch = "wasm32")))]
mod tests {
    use super::*;

    fn plain_package() -> Vec<u8> {
        use std::io::Write;
        let mut buf = std::io::Cursor::new(Vec::new());
        let mut z = zip::ZipWriter::new(&mut buf);
        z.start_file(
            "word/document.xml",
            zip::write::SimpleFileOptions::default(),
        )
        .unwrap();
        z.write_all(b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>")
            .unwrap();
        z.finish().unwrap();
        buf.into_inner()
    }

    #[test]
    fn encrypted_office_is_detected() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_package(), "pw").unwrap();
        assert!(is_encrypted_office(&enc));
        assert!(!is_encrypted_office(&plain_package()));
        assert!(!is_encrypted_office(b"%PDF-1.4\n"));
    }

    /// A synthetic encrypted legacy `.doc`: a CFB whose `/WordDocument` FIB has
    /// `fEncrypted` set and whose `/1Table` opens with an RC4 `CryptoAPI`
    /// `EncryptionVersionInfo` — the shape Office 16 still writes.
    ///
    /// Hand-built because `msoffice-crypto` only encrypts OOXML packages, and
    /// `classify` needs a real container to read the FIB and the header.
    fn encrypted_legacy_doc() -> Vec<u8> {
        use std::io::Write;

        let mut header = Vec::new();
        header.extend_from_slice(&2u16.to_le_bytes()); // vMajor
        header.extend_from_slice(&2u16.to_le_bytes()); // vMinor
        header.extend_from_slice(&0x04u32.to_le_bytes()); // fCryptoAPI, fAES clear
        header.extend_from_slice(&32u32.to_le_bytes()); // HeaderSize
        header.extend_from_slice(&0u32.to_le_bytes()); // EncryptionHeader.Flags
        header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
        header.extend_from_slice(&0x6801u32.to_le_bytes()); // AlgID: RC4
        header.extend_from_slice(&0u32.to_le_bytes()); // AlgIDHash
        header.extend_from_slice(&128u32.to_le_bytes()); // KeySize

        let mut fib = vec![0u8; 0x44];
        // wIdent: the marker every Word binary stream opens with.
        fib[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        // fEncrypted | fWhichTblStm: encrypted, header in /1Table.
        fib[0x0A..0x0C].copy_from_slice(&0x0300u16.to_le_bytes());
        // lKey: the EncryptionHeader's length at the start of /1Table.
        let l_key = u32::try_from(header.len()).unwrap();
        fib[0x0E..0x12].copy_from_slice(&l_key.to_le_bytes());

        let mut cursor = std::io::Cursor::new(Vec::new());
        let mut cfb = cfb::CompoundFile::create(&mut cursor).unwrap();
        cfb.create_stream("/WordDocument")
            .unwrap()
            .write_all(&fib)
            .unwrap();
        cfb.create_stream("/1Table")
            .unwrap()
            .write_all(&header)
            .unwrap();
        cfb.flush().unwrap();
        drop(cfb);
        cursor.into_inner()
    }

    /// `decrypt` opens OOXML packages only, so a legacy binary's encryption must
    /// not make [`is_encrypted_office`] true: detection would otherwise answer
    /// `PasswordRequired` for a `.doc` this build cannot open, and the caller
    /// would prompt for a password that can never work. [`is_encrypted`] is the
    /// predicate that reports it.
    #[test]
    fn encrypted_legacy_doc_is_not_decryptable_office() {
        let doc = encrypted_legacy_doc();
        assert!(is_encrypted(&doc));
        assert!(!is_encrypted_office(&doc));
    }

    /// The consequence of that narrowness: an encrypted legacy `.doc` still
    /// detects as `Format::Doc`, so extraction (not detection) owns the
    /// unsupported-encryption report and no password is ever asked for.
    #[test]
    fn encrypted_legacy_doc_still_detects_as_doc() {
        let doc = encrypted_legacy_doc();
        assert_eq!(crate::detect_format(&doc).unwrap(), crate::Format::Doc);
    }

    #[test]
    fn decrypt_round_trips() {
        let plain = plain_package();
        let enc = msoffice_crypto::encrypt_ooxml(&plain, "pw").unwrap();
        assert_eq!(decrypt(&enc, "pw").unwrap(), plain);
    }

    #[test]
    fn wrong_password_is_incorrect() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_package(), "pw").unwrap();
        let err = decrypt(&enc, "nope").unwrap_err();
        assert!(matches!(err, BatdocError::IncorrectPassword), "got {err:?}");
    }

    #[test]
    fn password_never_appears_in_error() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_package(), "hunter2").unwrap();
        let err = decrypt(&enc, "hunter2-not").unwrap_err().to_string();
        assert!(!err.contains("hunter2"), "leaked password: {err}");
    }
}
