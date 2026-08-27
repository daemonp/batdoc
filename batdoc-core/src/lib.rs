//! batdoc-core — document text extraction library.
//!
//! Converts DOCX, XLSX, PPTX, DOC, XLS, PDF, and raster images to
//! plain text or Markdown. Format detection is by magic bytes, not
//! file extension. Images (and optionally embedded images and textless
//! PDF pages) are read via OCR.

#![allow(clippy::redundant_pub_crate)]

mod arena;
mod codepage;
mod dateconv;
mod doc;
mod docx;
mod error;
mod heuristic;
mod markup;
mod ocr;
mod pdf;
mod pdf_geometry;
mod pdf_layout;
mod pdf_ocr;
mod pdf_text;
mod pptx;
mod sheet;
mod sheets;
mod sink;
#[cfg(all(target_arch = "wasm32", feature = "wasm-bindgen"))]
mod wasm;
mod xls;
mod xlsx;
mod xml_util;

pub use error::{BatdocError, Result};
pub use ocr::models_present;
pub use sheets::{BudgetSheetSink, Sheet, SheetSink};
pub use sink::{BudgetSink, ExtractSink, IoSink};

use std::io::Cursor;

/// Supported document formats.
///
/// `#[non_exhaustive]`: new variants may be added in minor releases;
/// downstream matches must include a wildcard arm.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
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
    if data.len() >= 6 && (&data[..6] == b"GIF87a" || &data[..6] == b"GIF89a") {
        return Ok(Format::Image); // GIF
    }
    if data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WEBP" {
        return Ok(Format::Image); // WebP
    }
    // BMP's 2-byte "BM" signature is weak: any file starting with those
    // bytes is routed to OCR and fails with "no text found in image"
    // rather than "unrecognized format". Accepted trade-off — real-world
    // collisions are rare.
    if data.len() >= 2 && &data[..2] == b"BM" {
        return Ok(Format::Image); // BMP
    }

    Err(BatdocError::Document(
        "not a supported document (unrecognized format)".into(),
    ))
}

/// Extraction options.
#[derive(Debug, Clone, Copy)]
pub struct ExtractOptions {
    /// Include embedded images as base64 markdown (markdown mode only).
    pub images: bool,
    /// OCR embedded images (DOCX/PPTX). Has no effect on `Format::Image` —
    /// image input is always OCR'd.
    pub ocr: bool,
    /// Textless or garbled PDF pages fall back to OCR when `true` (the
    /// default), even if `ocr` is `false`. Set `false` to disable that
    /// fallback (Worker-safe / Vault). Ignored when `ocr` is `true`; no
    /// models are downloaded or required when `false`.
    pub auto_ocr: bool,
    /// Stop writing after this many output bytes. `None` means unlimited.
    pub max_output_bytes: Option<u64>,
}

impl Default for ExtractOptions {
    fn default() -> Self {
        Self {
            images: false,
            ocr: false,
            auto_ocr: true,
            max_output_bytes: None,
        }
    }
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

/// Extract plain text with explicit options.
///
/// `opts.ocr` enables OCR for DOCX/PPTX embedded images; the `images` option
/// is ignored in plain mode. Textless PDF pages are OCR'd as a fallback when
/// `opts.auto_ocr` is enabled (the default), with or without `opts.ocr`.
/// `Format::Image` input is always OCR'd regardless of options and returns
/// plain OCR text.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_plain_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String> {
    match format {
        Format::Doc => doc::extract_plain(data),
        Format::Xls => xls::extract_plain(data),
        Format::Docx => docx::extract_plain(data, opts),
        Format::Xlsx => xlsx::extract_plain(data),
        Format::Pptx => pptx::extract_plain(data, opts),
        Format::Pdf => pdf::extract_plain(data, opts),
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
    extract_markdown_with(
        data,
        format,
        ExtractOptions {
            images,
            ocr: false,
            ..Default::default()
        },
    )
}

/// Extract Markdown with explicit options.
///
/// `opts.images` embeds DOCX/XLSX/PPTX images as base64 markdown;
/// `opts.ocr` OCRs DOCX/PPTX embedded images, rendered as blockquotes.
/// Textless/garbled PDF pages fall back to OCR when `opts.auto_ocr` is
/// enabled (the default). `Format::Image` input is always OCR'd regardless
/// of options and returns plain OCR text (no markdown).
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed.
pub fn extract_markdown_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String> {
    match format {
        Format::Doc => doc::extract_markdown(data),
        Format::Xls => xls::extract_markdown(data),
        Format::Docx => docx::extract_markdown(data, opts),
        Format::Xlsx => xlsx::extract_markdown(data, opts.images),
        Format::Pptx => pptx::extract_markdown(data, opts),
        Format::Pdf => pdf::extract_markdown(data, opts),
        Format::Image => ocr::extract_image_plain(data),
    }
}

/// Extract plain text into a sink.
///
/// When `opts.max_output_bytes` is `Some`, writing stops with
/// [`BatdocError::Document`] once that many bytes would be exceeded.
///
/// # Errors
///
/// Returns any error from [`extract_plain_with`], or
/// [`BatdocError::Document`] if the output budget is exceeded.
pub fn extract_plain_to(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
    sink: &mut impl ExtractSink,
) -> Result<()> {
    match opts.max_output_bytes {
        Some(max) => {
            let mut limited = BudgetSink::new(sink, max);
            write_plain(data, format, opts, &mut limited)
        }
        None => write_plain(data, format, opts, sink),
    }
}

fn write_plain(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
    sink: &mut impl ExtractSink,
) -> Result<()> {
    match format {
        Format::Xlsx => xlsx::extract_plain_to(data, sink),
        Format::Xls => xls::extract_plain_to(data, sink),
        Format::Docx => docx::extract_plain_to(data, opts, sink),
        Format::Pptx => pptx::extract_plain_to(data, opts, sink),
        Format::Doc => doc::extract_plain_to(data, sink),
        Format::Pdf => pdf::extract_plain_to(data, opts, sink),
        _ => {
            let text = extract_plain_with(data, format, opts)?;
            sink.write_str(&text)
        }
    }
}

/// Extract Markdown into a sink.
///
/// When `opts.max_output_bytes` is `Some`, writing stops with
/// [`BatdocError::Document`] once that many bytes would be exceeded.
///
/// # Errors
///
/// Returns any error from [`extract_markdown_with`], or
/// [`BatdocError::Document`] if the output budget is exceeded.
pub fn extract_markdown_to(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
    sink: &mut impl ExtractSink,
) -> Result<()> {
    match opts.max_output_bytes {
        Some(max) => {
            let mut limited = BudgetSink::new(sink, max);
            write_markdown(data, format, opts, &mut limited)
        }
        None => write_markdown(data, format, opts, sink),
    }
}

fn write_markdown(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
    sink: &mut impl ExtractSink,
) -> Result<()> {
    match format {
        Format::Xlsx => xlsx::extract_markdown_to(data, opts.images, sink),
        Format::Xls => xls::extract_markdown_to(data, sink),
        Format::Docx => docx::extract_markdown_to(data, opts, sink),
        Format::Pptx => pptx::extract_markdown_to(data, opts, sink),
        Format::Pdf => pdf::extract_markdown_to(data, opts, sink),
        _ => {
            let text = extract_markdown_with(data, format, opts)?;
            sink.write_str(&text)
        }
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

/// Extract all sheets into a `Vec<Sheet>` (collecting — O(cells) memory).
///
/// Prefer [`extract_sheets_to`] on large workbooks. Only `Format::Xls` and
/// `Format::Xlsx` are supported.
///
/// # Errors
///
/// [`BatdocError::Document`] for non-spreadsheet formats or parse failures.
pub fn extract_sheets(data: &[u8], format: Format) -> Result<Vec<Sheet>> {
    extract_sheets_with(data, format, ExtractOptions::default())
}

/// Like [`extract_sheets`] with options. Only `max_output_bytes` is honored;
/// `images` / `ocr` / `auto_ocr` are ignored.
///
/// # Errors
///
/// See [`extract_sheets_to`].
pub fn extract_sheets_with(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
) -> Result<Vec<Sheet>> {
    let mut sheets = Vec::new();
    extract_sheets_to(data, format, opts, &mut sheets)?;
    Ok(sheets)
}

/// Stream structured tabular data into `sink`.
///
/// When `opts.max_output_bytes` is `Some`, wraps `sink` in
/// [`BudgetSheetSink`] (payload estimate: sheet name bytes + per-cell
/// `len+1`; not wire/JS heap size). Peak process memory still includes
/// the input buffer and shared-string table.
///
/// # Errors
///
/// `"tabular extraction is only supported for XLS and XLSX"` for other
/// formats; budget / column / parse errors otherwise.
pub fn extract_sheets_to(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
    sink: &mut impl SheetSink,
) -> Result<()> {
    match opts.max_output_bytes {
        Some(max) => {
            let mut limited = BudgetSheetSink::new(sink, max);
            write_sheets(data, format, &mut limited)
        }
        None => write_sheets(data, format, sink),
    }
}

fn write_sheets(data: &[u8], format: Format, sink: &mut impl SheetSink) -> Result<()> {
    match format {
        Format::Xlsx => xlsx::extract_sheets_to(data, sink),
        Format::Xls => xls::extract_sheets_to(data, sink),
        _ => Err(BatdocError::Document(
            "tabular extraction is only supported for XLS and XLSX".into(),
        )),
    }
}

/// Detect format and extract sheets (collecting).
///
/// # Errors
///
/// [`detect_format`] or [`extract_sheets`] errors.
pub fn to_sheets(data: &[u8]) -> Result<Vec<Sheet>> {
    let format = detect_format(data)?;
    extract_sheets(data, format)
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
        assert_eq!(
            detect_format(&[0xFF, 0xD8, 0xFF, 0xE0]).unwrap(),
            Format::Image
        );
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
    fn detect_format_requires_full_gif_magic() {
        assert_eq!(detect_format(b"GIF87a....").unwrap(), Format::Image);
        assert_eq!(detect_format(b"GIF89a....").unwrap(), Format::Image);
        assert!(detect_format(b"GIFzzz....").is_err());
    }

    #[test]
    fn extract_image_plain_path_errors_without_text() {
        // Real OCR needs models; the garbage path must not.
        let err = extract_plain_with(b"garbage", Format::Image, ExtractOptions::default())
            .unwrap_err()
            .to_string();
        assert!(err.contains("no text found in image"));
    }

    #[test]
    fn extract_plain_to_equals_extract_plain_on_image_garbage() {
        let data = b"garbage";
        let format = Format::Image;
        let opts = ExtractOptions::default();
        let a = extract_plain_with(data, format, opts)
            .unwrap_err()
            .to_string();
        let mut out = String::new();
        let b = extract_plain_to(data, format, opts, &mut out)
            .unwrap_err()
            .to_string();
        assert_eq!(a, b);
        assert!(out.is_empty());
    }

    #[test]
    fn extract_sheets_rejects_non_spreadsheet() {
        for fmt in [Format::Pdf, Format::Docx] {
            let err = extract_sheets(b"%PDF-1.4", fmt).unwrap_err().to_string();
            assert_eq!(err, "tabular extraction is only supported for XLS and XLSX");
        }
    }

    #[allow(clippy::assert_is_empty)] // intentionally brief verbatim assertion
    #[test]
    fn extract_sheets_routes_xlsx_and_honors_budget() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;
        use zip::ZipWriter;

        let mut z = ZipWriter::new(Cursor::new(Vec::new()));
        for (name, body) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>
  </sheetData>
</worksheet>"#,
            ),
        ] {
            z.start_file(name, SimpleFileOptions::default()).unwrap();
            z.write_all(body.as_bytes()).unwrap();
        }
        let data = z.finish().unwrap().into_inner();

        // Routing: public collecting wrapper reaches the real implementation.
        let sheets = extract_sheets(&data, Format::Xlsx).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "S");
        assert_eq!(sheets[0].rows, vec![vec!["Hello"]]);

        // Budget: name "S" = 1; row ["Hello"] = 5+1 = 6 → total 7 > 3.
        let mut out = Vec::<Sheet>::new();
        let err = extract_sheets_to(
            &data,
            Format::Xlsx,
            ExtractOptions {
                max_output_bytes: Some(3),
                ..Default::default()
            },
            &mut out,
        )
        .unwrap_err()
        .to_string();
        assert_eq!(err, "output exceeded 3 bytes");
        assert_eq!(out.len(), 1);
        assert!(out[0].rows.is_empty());
    }

    #[test]
    fn to_sheets_rejects_unrecognized() {
        let err = to_sheets(b"hello, definitely not a document")
            .unwrap_err()
            .to_string();
        assert_eq!(err, "not a supported document (unrecognized format)");
    }
}
