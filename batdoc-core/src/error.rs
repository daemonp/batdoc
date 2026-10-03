//! Error types for batdoc.
//!
//! Provides a single [`BatdocError`] enum that replaces the previous
//! `Box<dyn std::error::Error>` usage throughout the codebase.

/// All errors that can occur during document parsing and rendering.
///
/// `#[non_exhaustive]`: new variants may be added in minor releases;
/// downstream matches must include a wildcard arm.
#[non_exhaustive]
#[derive(Debug, thiserror::Error)]
pub enum BatdocError {
    /// I/O error (file read, stream read, OLE2 compound file).
    #[error("{0}")]
    Io(#[from] std::io::Error),

    /// ZIP archive error (from `zip` crate).
    #[error("{0}")]
    Zip(#[from] zip::result::ZipError),

    /// Document-level error (unsupported format, corruption).
    #[error("{0}")]
    Document(String),

    /// Pretty-printing error (bat rendering failure).
    #[error("pretty print: {0}")]
    Render(String),

    /// Document is encrypted and no usable password was supplied.
    #[error("document is password-protected")]
    PasswordRequired,

    /// A password was supplied but did not authenticate.
    #[error("incorrect password")]
    IncorrectPassword,

    /// Encrypted with an algorithm or container this build does not handle.
    #[error("unsupported encryption: {0}")]
    UnsupportedEncryption(String),
}

/// Convenience alias used throughout the crate.
pub type Result<T> = std::result::Result<T, BatdocError>;

#[cfg(test)]
mod tests {
    use super::BatdocError;

    #[test]
    fn encryption_error_display_is_stable() {
        assert_eq!(
            BatdocError::PasswordRequired.to_string(),
            "document is password-protected"
        );
        assert_eq!(
            BatdocError::IncorrectPassword.to_string(),
            "incorrect password"
        );
        assert_eq!(
            BatdocError::UnsupportedEncryption("legacy .doc".into()).to_string(),
            "unsupported encryption: legacy .doc"
        );
    }
}
