//! batdoc-core — document text extraction library.
//!
//! Converts DOCX, XLSX, PPTX, DOC, XLS, PDF, and raster images to
//! plain text or Markdown. Format detection is by magic bytes, not
//! file extension. Images (and optionally embedded images and textless
//! PDF pages) are read via OCR.

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
    /// Raster image (PNG/JPEG/GIF/WebP/BMP) — always OCR'd.
    Image,
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
            Self::Image => f.write_str("IMAGE"),
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

    // Raster images (OCR input): PNG, JPEG, GIF, WebP, BMP
    if data.len() >= 4 && data[..4] == [0x89, 0x50, 0x4E, 0x47] {
        return Ok(Format::Image); // PNG
    }
    if data.len() >= 3 && data[..3] == [0xFF, 0xD8, 0xFF] {
        return Ok(Format::Image); // JPEG
    }
    if data.len() >= 6 && &data[..3] == b"GIF" {
        return Ok(Format::Image); // GIF
    }
    if data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WEBP" {
        return Ok(Format::Image); // WebP
    }
    if data.len() >= 2 && &data[..2] == b"BM" {
        return Ok(Format::Image); // BMP
    }

    Err(BatdocError::Document(
        "not a supported document (unrecognized format)".into(),
    ))
}

/// Extraction options.
#[derive(Debug, Clone, Copy, Default)]
pub struct ExtractOptions {
    /// Include embedded images as base64 markdown (markdown mode only).
    pub images: bool,
    /// OCR embedded images and textless PDF pages.
    pub ocr: bool,
}

/// Extract plain text from a document.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_plain(data: &[u8], format: Format) -> Result<String> {
    extract_plain_with(data, format, ExtractOptions::default())
}

/// Extract plain text with explicit options (`--ocr` on the CLI).
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_plain_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String> {
    match format {
        Format::Doc => doc::extract_plain(data),
        Format::Xls => xls::extract_plain(data),
        Format::Docx => docx::extract_plain(data, opts.ocr),
        Format::Xlsx => xlsx::extract_plain(data),
        Format::Pptx => pptx::extract_plain(data, opts.ocr),
        Format::Pdf => pdf::extract_plain(data, opts.ocr),
        Format::Image => ocr::extract_image_plain(data),
    }
}

/// Extract Markdown from a document.
///
/// When `images` is `true`, embedded images in DOCX/XLSX/PPTX are
/// included as reference-style base64 data URIs. Has no effect on
/// DOC, XLS, PDF, or Image.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_markdown(data: &[u8], format: Format, images: bool) -> Result<String> {
    extract_markdown_with(data, format, ExtractOptions { images, ocr: false })
}

/// Extract Markdown with explicit options (`--ocr` on the CLI).
///
/// When `images` is `true`, embedded images in DOCX/XLSX/PPTX are
/// included as reference-style base64 data URIs. Has no effect on
/// DOC, XLS, PDF, or Image.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_markdown_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String> {
    match format {
        Format::Doc => doc::extract_markdown(data),
        Format::Xls => xls::extract_markdown(data),
        Format::Docx => docx::extract_markdown(data, opts.images, opts.ocr),
        Format::Xlsx => xlsx::extract_markdown(data, opts.images),
        Format::Pptx => pptx::extract_markdown(data, opts.images, opts.ocr),
        Format::Pdf => pdf::extract_markdown(data, opts.ocr),
        Format::Image => ocr::extract_image_plain(data),
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detect_format_image_magic_bytes() {
        // PNG
        assert_eq!(
            detect_format(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A]).unwrap(),
            Format::Image
        );
        // JPEG
        assert_eq!(detect_format(&[0xFF, 0xD8, 0xFF, 0xE0]).unwrap(), Format::Image);
        // GIF
        assert_eq!(detect_format(b"GIF89a....").unwrap(), Format::Image);
        // WebP (RIFF....WEBP)
        let mut webp = b"RIFF".to_vec();
        webp.extend_from_slice(&[0, 0, 0, 0]);
        webp.extend_from_slice(b"WEBP");
        assert_eq!(detect_format(&webp).unwrap(), Format::Image);
        // BMP
        assert_eq!(detect_format(b"BM....").unwrap(), Format::Image);
    }

    #[test]
    fn detect_format_still_rejects_text() {
        assert!(detect_format(b"hello world, definitely not a document").is_err());
    }

    #[test]
    fn format_image_displays_as_image() {
        assert_eq!(Format::Image.to_string(), "IMAGE");
    }

    #[test]
    fn extract_image_plain_path_errors_without_text() {
        // Real OCR needs models; the garbage path must not.
        let err = extract_plain_with(b"garbage", Format::Image, ExtractOptions::default())
            .unwrap_err()
            .to_string();
        assert!(err.contains("no text found in image"));
    }
}
