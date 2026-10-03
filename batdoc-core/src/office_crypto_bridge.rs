//! Bridge to `msoffice-crypto` for password-protected Office documents.
//!
//! Detection uses the crate's infallible `classify`; decryption uses
//! `decrypt_ooxml_with_policy` under the default integrity policy, which
//! verifies the package HMAC. This is why a wrong password is reported as
//! [`BatdocError::IncorrectPassword`] rather than surfacing as garbage bytes.
//!
//! On `wasm32-unknown-unknown` the crypto dependency is not in the graph (it
//! does not compile there), so this module is a stub pair instead: Office
//! decryption is unavailable, `is_encrypted_office` never reports `true`, and
//! an encrypted package falls through to the unrecognised-format error. Only
//! PDF password support applies there. The stubs keep every call site
//! target-agnostic.

// Wired into detection/extraction by the next commits on this branch; removed there.
#![allow(dead_code)]

use crate::error::{BatdocError, Result};
#[cfg(not(target_arch = "wasm32"))]
use std::panic::{self, AssertUnwindSafe};

/// Ceiling for decrypted Office payloads (mirrors the CLI's 256 MiB input
/// cap) so a small ciphertext cannot expand into an unbounded allocation.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) const MAX_DECRYPTED_BYTES: usize = 256 * 1024 * 1024;

/// True when `data` is an MS-OFFCRYPTO container declaring a password.
///
/// `msoffice_crypto::classify` never panics and never fails: unreadable input
/// collapses to `Family::Unknown`, which `is_encrypted` reports as `false`.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn is_encrypted_office(data: &[u8]) -> bool {
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
