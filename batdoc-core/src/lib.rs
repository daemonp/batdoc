//! batdoc-core — document text extraction library.
//!
//! Converts DOCX, XLSX, PPTX, DOC, XLS, and PDF files to plain text
//! or Markdown. Format detection is by magic bytes, not file extension.

#![allow(clippy::redundant_pub_crate)]

mod codepage;
mod dateconv;
mod doc;
mod docx;
mod error;
mod heuristic;
mod markup;
mod ocr;
mod pdf;
mod pptx;
mod sheet;
mod xls;
mod xlsx;
mod xml_util;

pub use error::{BatdocError, Result};

use std::io::Cursor;

/// Supported document formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Format {
    /// Legacy OLE2 Word 97+ binary format.
    Doc,
    /// Legacy OLE2 Excel 97+ binary format (BIFF8).
    Xls,
    /// Modern OOXML Word (ZIP-based) format.
    Docx,
    /// Modern OOXML Excel (ZIP-based) format.
    Xlsx,
    /// Modern OOXML `PowerPoint` (ZIP-based) format.
    Pptx,
    /// PDF document.
    Pdf,
}

impl std::fmt::Display for Format {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Doc => f.write_str("DOC"),
            Self::Xls => f.write_str("XLS"),
            Self::Docx => f.write_str("DOCX"),
            Self::Xlsx => f.write_str("XLSX"),
            Self::Pptx => f.write_str("PPTX"),
            Self::Pdf => f.write_str("PDF"),
        }
    }
}

/// Detect document format from the first bytes of the file.
///
/// Uses magic-byte signatures (OLE2, ZIP, PDF header), not file
/// extensions — critical for email attachments where MIME types are
/// often wrong.
///
/// # Errors
///
/// Returns [`BatdocError::Document`] if the magic bytes don't match any
/// supported format, or if the file matches a container format (OLE2/ZIP)
/// but doesn't contain a recognised document type.
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Zip`] if the container
/// cannot be parsed.
pub fn detect_format(data: &[u8]) -> Result<Format> {
    // OLE2 compound file: 0xD0CF11E0A1B11AE1
    if data.len() >= 8 && data[..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        let cursor = Cursor::new(data);
        let cfb = cfb::CompoundFile::open(cursor)?;
        if cfb.exists("/WordDocument") {
            return Ok(Format::Doc);
        }
        if cfb.exists("/Workbook") || cfb.exists("/Book") {
            return Ok(Format::Xls);
        }
        return Err(BatdocError::Document(
            "OLE2 file is not a .doc or .xls document".into(),
        ));
    }

    // PDF: %PDF-
    if data.len() >= 5 && &data[..5] == b"%PDF-" {
        return Ok(Format::Pdf);
    }

    // ZIP-based OOXML: PK\x03\x04
    if data.len() >= 4 && &data[..4] == b"PK\x03\x04" {
        let cursor = Cursor::new(data);
        let archive = zip::ZipArchive::new(cursor)?;
        if archive.index_for_name("word/document.xml").is_some() {
            return Ok(Format::Docx);
        }
        if archive.index_for_name("xl/workbook.xml").is_some() {
            return Ok(Format::Xlsx);
        }
        if archive.index_for_name("ppt/presentation.xml").is_some() {
            return Ok(Format::Pptx);
        }
        return Err(BatdocError::Document(
            "ZIP archive is not a .docx, .xlsx, or .pptx file".into(),
        ));
    }

    Err(BatdocError::Document(
        "not a supported document (unrecognized format)".into(),
    ))
}

/// Extract plain text from a document.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_plain(data: &[u8], format: Format) -> Result<String> {
    match format {
        Format::Doc => doc::extract_plain(data),
        Format::Xls => xls::extract_plain(data),
        Format::Docx => docx::extract_plain(data, false),
        Format::Xlsx => xlsx::extract_plain(data),
        Format::Pptx => pptx::extract_plain(data),
        Format::Pdf => pdf::extract_plain(data, false),
    }
}

/// Extract Markdown from a document.
///
/// When `images` is `true`, embedded images in DOCX/XLSX/PPTX are
/// included as reference-style base64 data URIs. Has no effect on
/// DOC, XLS, or PDF (which have no extractable embedded images).
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_markdown(data: &[u8], format: Format, images: bool) -> Result<String> {
    match format {
        Format::Doc => doc::extract_markdown(data),
        Format::Xls => xls::extract_markdown(data),
        Format::Docx => docx::extract_markdown(data, images, false),
        Format::Xlsx => xlsx::extract_markdown(data, images),
        Format::Pptx => pptx::extract_markdown(data, images),
        Format::Pdf => pdf::extract_markdown(data, false),
    }
}

/// Convenience: detect format and extract plain text in one call.
///
/// # Errors
///
/// Returns any error from [`detect_format`] or [`extract_plain`].
pub fn to_plain(data: &[u8]) -> Result<String> {
    let format = detect_format(data)?;
    extract_plain(data, format)
}

/// Convenience: detect format and extract Markdown in one call.
///
/// # Errors
///
/// Returns any error from [`detect_format`] or [`extract_markdown`].
pub fn to_markdown(data: &[u8], images: bool) -> Result<String> {
    let format = detect_format(data)?;
    extract_markdown(data, format, images)
}
