//! Error types for batdoc.
//!
//! Provides a single [`BatdocError`] enum that replaces the previous
//! `Box<dyn std::error::Error>` usage throughout the codebase.

/// All errors that can occur during document parsing and rendering.
#[derive(Debug, thiserror::Error)]
pub(crate) enum BatdocError {
    /// I/O error (file read, stream read, OLE2 compound file).
    #[error("{0}")]
    Io(#[from] std::io::Error),

    /// ZIP archive error (from `zip` crate).
    #[error("{0}")]
    Zip(#[from] zip::result::ZipError),

    /// Document-level error (unsupported format, encryption, corruption).
    #[error("{0}")]
    Document(String),

    /// Pretty-printing error (bat rendering failure).
    #[error("pretty print: {0}")]
    Render(String),
}

/// Convenience alias used throughout the crate.
pub(crate) type Result<T> = std::result::Result<T, BatdocError>;
