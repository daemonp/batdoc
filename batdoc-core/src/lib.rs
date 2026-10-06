//! batdoc-core — document text extraction library.
//!
//! Converts DOCX, XLSX, PPTX, DOC, XLS, PDF, and raster images to
//! plain text or Markdown. Format detection is by magic bytes, not
//! file extension. Images (and optionally embedded images and textless
//! PDF pages) are read via OCR.
//!
//! Image OCR, embedded-image OCR, and the textless-PDF OCR fallback are
//! compiled in with the `ocr` feature (on by default); disabling it
//! (`--no-default-features`) drops `ocrs`/`rten`/`image` for lean builds.

#![allow(clippy::redundant_pub_crate)]

mod arena;
mod codepage;
mod csv;
mod dateconv;
mod doc;
// ExtractOptions is a small options bag passed by value through the parse
// tree. It stopped being `Copy` when the PDF strip options (a `Vec`) were
// added; threading a borrow through every recursive helper would be churn
// for no gain, so the lint is allowed module-wide.
#[allow(clippy::needless_pass_by_value)]
mod docx;
mod error;
mod heuristic;
mod markup;
#[cfg(feature = "ocr")]
mod ocr;
mod office_crypto_bridge;
#[allow(clippy::needless_pass_by_value)] // see the note on `mod docx`
mod pdf;
mod pdf_geometry;
mod pdf_layout;
mod pdf_ocr;
#[cfg(feature = "ocr")]
mod pdf_raster;
mod pdf_text;
mod pdf_watermark;
#[allow(clippy::needless_pass_by_value)] // see the note on `mod docx`
mod pptx;
mod sheet;
mod sheets;
mod sink;
#[cfg(all(target_arch = "wasm32", feature = "wasm-bindgen"))]
mod wasm;
mod xls;
mod xlsx;
mod xml_util;

pub use csv::{escape_field, to_csv_row, CsvSink};
pub use error::{BatdocError, Result};
#[cfg(feature = "ocr")]
pub use ocr::models_present;
pub use sheets::{BudgetSheetSink, Sheet, SheetSink};
pub use sink::{BudgetSink, ExtractSink, IoSink};

use std::io::Cursor;

/// Whether OCR inference is compiled in (the `ocr` feature). When `false`,
/// [`ExtractOptions::ocr`] and [`ExtractOptions::auto_ocr`] are treated as
/// `false` (there is no engine), and `Format::Image` input fails.
#[cfg(feature = "ocr")]
pub(crate) const OCR_COMPILED: bool = true;
#[cfg(not(feature = "ocr"))]
pub(crate) const OCR_COMPILED: bool = false;

/// Whether OCR model files are present in the cache.
///
/// With the `ocr` feature disabled there is no OCR engine, so this is
/// always `false`.
#[cfg(not(feature = "ocr"))]
#[must_use]
// Not `const`: signature parity with the feature-on re-export.
#[allow(clippy::missing_const_for_fn)]
pub fn models_present() -> bool {
    false
}

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
    /// Raster image (PNG/JPEG/GIF/WebP/BMP) — always OCR'd (when the
    /// `ocr` feature is enabled).
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
/// Returns [`BatdocError::PasswordRequired`] for an encrypted Office package:
/// the plaintext is not reachable without a password, so no format can be
/// named. Use [`detect_format_with`] to supply one, or [`needs_password`] to
/// probe first.
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
        if office_crypto_bridge::is_encrypted_office(data) {
            return Err(BatdocError::PasswordRequired);
        }
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

/// Detect the format, decrypting an encrypted Office package when a password
/// is supplied.
///
/// Without a password an encrypted Office package yields
/// [`BatdocError::PasswordRequired`]. PDF and unencrypted formats behave
/// exactly like [`detect_format`].
///
/// # Errors
///
/// See [`detect_format`], plus [`BatdocError::PasswordRequired`] /
/// [`BatdocError::IncorrectPassword`] / [`BatdocError::UnsupportedEncryption`].
pub fn detect_format_with(data: &[u8], password: Option<&str>) -> Result<Format> {
    if office_crypto_bridge::is_encrypted_office(data) {
        let password = password.ok_or(BatdocError::PasswordRequired)?;
        let plain = office_crypto_bridge::decrypt(data, password)?;
        return detect_format(&plain);
    }
    detect_format(data)
}

/// Whether `data` is encrypted and will need a password to open.
///
/// Never requires a password. Covers the PDF encryption dictionary, encrypted
/// Office packages, and legacy `.doc`/`.xls`/`.ppt` encryption markers. This
/// parses a PDF's encryption dictionary, so it is not free.
///
/// # Errors
///
/// [`BatdocError::Document`] if a PDF header is present but the document
/// cannot be parsed.
pub fn needs_password(data: &[u8]) -> Result<bool> {
    if data.len() >= 5 && &data[..5] == b"%PDF-" {
        // lopdf parses attacker-controlled dictionaries and can panic on
        // malformed input; contain that as a `Document` error rather than
        // letting it abort the caller.
        let probed = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let doc = lopdf::Document::load_mem(data)
                .map_err(|e| BatdocError::Document(format!("PDF parse failed: {e}")))?;
            if !doc.is_encrypted() {
                return Ok(false);
            }
            // Owner-only locks authenticate with the empty user password.
            // Probe authentication rather than full decryption: an
            // undecryptable object must not make an empty password look
            // wrong when the document itself accepts it.
            Ok(doc.authenticate_password("").is_err())
        }));
        return probed.unwrap_or_else(|_| Err(BatdocError::Document("PDF parse panicked".into())));
    }
    Ok(office_crypto_bridge::is_encrypted(data))
}

/// Extraction options.
#[derive(Clone)]
#[allow(clippy::struct_excessive_bools)] // independent feature switches, not state
pub struct ExtractOptions {
    /// Include embedded images as base64 markdown (markdown mode only).
    pub images: bool,
    /// OCR embedded images (DOCX/PPTX). Has no effect on `Format::Image` —
    /// image input is always OCR'd. No effect when built without the
    /// `ocr` feature.
    pub ocr: bool,
    /// Textless or garbled PDF pages fall back to OCR when `true` (the
    /// default), even if `ocr` is `false`. Set `false` to disable that
    /// fallback (Worker-safe / Vault). Ignored when `ocr` is `true`; no
    /// models are downloaded or required when `false`. Forced off when built
    /// without the `ocr` feature.
    pub auto_ocr: bool,
    /// Stop writing after this many output bytes. `None` means unlimited.
    pub max_output_bytes: Option<u64>,
    /// Needles to strip from PDF output: any *reconstructed text run*
    /// containing one (case-insensitive, ignoring whitespace) is removed
    /// before layout. This operates on glyphs along their true direction, so
    /// it reaches diagonal watermarks that per-line output filters cannot
    /// see. Empty means no stripping. PDF only.
    pub strip_text: Vec<String>,
    /// Remove skewed (non-orthogonal) text runs whose normalized signature
    /// repeats across pages — watermark removal without naming the string.
    /// Needs at least two pages to learn a signature; a single-page document
    /// is left untouched (use `strip_text` there). No-op by default.
    pub strip_watermarks: bool,
    /// Password for an encrypted PDF or Office document.
    ///
    /// `None` means "not supplied": PDF still tries the empty user password
    /// (owner-only locks); an encrypted Office package returns
    /// [`BatdocError::PasswordRequired`] without guessing.
    pub password: Option<String>,
}

impl Default for ExtractOptions {
    fn default() -> Self {
        Self {
            images: false,
            ocr: false,
            auto_ocr: true,
            max_output_bytes: None,
            strip_text: Vec::new(),
            strip_watermarks: false,
            password: None,
        }
    }
}

/// Wrapper that renders an optional password without revealing it.
///
/// Used by the manual `Debug` impl below so that `ExtractOptions` can be
/// printed (for example from a `#[derive(Debug)]` container) without ever
/// putting a plaintext password into the output.
struct RedactedPassword<'a>(&'a Option<String>);

impl std::fmt::Debug for RedactedPassword<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if self.0.is_some() {
            f.write_str("Some(\"<redacted>\")")
        } else {
            f.write_str("None")
        }
    }
}

// Manual impl: `#[derive(Debug)]` would print the plaintext password, which
// must never reach `Debug`, logs, or stdout. Every field is listed explicitly
// to satisfy `clippy::missing_fields_in_debug`.
impl std::fmt::Debug for ExtractOptions {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ExtractOptions")
            .field("images", &self.images)
            .field("ocr", &self.ocr)
            .field("auto_ocr", &self.auto_ocr)
            .field("max_output_bytes", &self.max_output_bytes)
            .field("strip_text", &self.strip_text)
            .field("strip_watermarks", &self.strip_watermarks)
            .field("password", &RedactedPassword(&self.password))
            .finish()
    }
}

/// Extract text from a raster image. Requires the `ocr` feature; without it
/// this always returns [`BatdocError::Document`].
#[cfg(feature = "ocr")]
fn extract_image(data: &[u8]) -> Result<String> {
    ocr::extract_image_plain(data)
}

#[cfg(not(feature = "ocr"))]
fn extract_image(_data: &[u8]) -> Result<String> {
    Err(BatdocError::Document(
        "image input requires the `ocr` feature (not compiled in)".into(),
    ))
}

/// OCR a raw image byte buffer to text. With the `ocr` feature disabled this
/// is a no-op returning `Ok(None)`, so DOCX/PPTX embedded-image OCR is
/// silently skipped rather than erroring.
#[cfg(feature = "ocr")]
pub(crate) fn ocr_image_bytes(data: &[u8]) -> Result<Option<String>> {
    ocr::ocr_image_bytes(data)
}

#[cfg(not(feature = "ocr"))]
// The Result/Option shape mirrors the feature-on variant — call sites
// use `?` on it.
#[allow(clippy::unnecessary_wraps, clippy::missing_const_for_fn)]
pub(crate) fn ocr_image_bytes(_data: &[u8]) -> Result<Option<String>> {
    Ok(None)
}

/// If `data` is an encrypted Office package, decrypt it. `None` when it is
/// not encrypted Office (the caller uses `data` as-is).
fn decrypt_office(data: &[u8], opts: &ExtractOptions) -> Result<Option<Vec<u8>>> {
    if !office_crypto_bridge::is_encrypted_office(data) {
        return Ok(None);
    }
    let password = opts
        .password
        .as_deref()
        .ok_or(BatdocError::PasswordRequired)?;
    Ok(Some(office_crypto_bridge::decrypt(data, password)?))
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
/// An encrypted Office package is decrypted first when `opts.password` is
/// set, so `format` may describe the decrypted package even though `data` is
/// still the ciphertext.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed; for encrypted
/// Office input, [`BatdocError::PasswordRequired`] /
/// [`BatdocError::IncorrectPassword`] / [`BatdocError::UnsupportedEncryption`].
pub fn extract_plain_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String> {
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_plain_with(&plain, detect_format(&plain)?, opts);
    }
    match format {
        Format::Doc => doc::extract_plain(data),
        Format::Xls => xls::extract_plain(data),
        Format::Docx => docx::extract_plain(data, opts),
        Format::Xlsx => xlsx::extract_plain(data),
        Format::Pptx => pptx::extract_plain(data, opts),
        Format::Pdf => pdf::extract_plain(data, opts),
        Format::Image => extract_image(data),
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
/// An encrypted Office package is decrypted first when `opts.password` is
/// set, so `format` may describe the decrypted package even though `data` is
/// still the ciphertext.
///
/// # Errors
///
/// Returns [`BatdocError::Io`] or [`BatdocError::Document`] if the
/// document is malformed, encrypted, or cannot be parsed; for encrypted
/// Office input, [`BatdocError::PasswordRequired`] /
/// [`BatdocError::IncorrectPassword`] / [`BatdocError::UnsupportedEncryption`].
pub fn extract_markdown_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String> {
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_markdown_with(&plain, detect_format(&plain)?, opts);
    }
    match format {
        Format::Doc => doc::extract_markdown(data),
        Format::Xls => xls::extract_markdown(data),
        Format::Docx => docx::extract_markdown(data, opts),
        Format::Xlsx => xlsx::extract_markdown(data, opts.images),
        Format::Pptx => pptx::extract_markdown(data, opts),
        Format::Pdf => pdf::extract_markdown(data, opts),
        Format::Image => extract_image(data),
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
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_plain_to(&plain, detect_format(&plain)?, opts, sink);
    }
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
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_markdown_to(&plain, detect_format(&plain)?, opts, sink);
    }
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

/// Convenience: detect (decrypting encrypted Office) and extract plain text.
///
/// # Errors
///
/// See [`extract_plain_with`]; adds [`BatdocError::PasswordRequired`] /
/// [`BatdocError::IncorrectPassword`] / [`BatdocError::UnsupportedEncryption`].
pub fn to_plain_with(data: &[u8], opts: ExtractOptions) -> Result<String> {
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_plain_with(&plain, detect_format(&plain)?, opts);
    }
    extract_plain_with(data, detect_format(data)?, opts)
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

/// Convenience: detect (decrypting encrypted Office) and extract Markdown.
///
/// # Errors
///
/// See [`extract_markdown_with`]; adds the same encryption errors.
pub fn to_markdown_with(data: &[u8], opts: ExtractOptions) -> Result<String> {
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_markdown_with(&plain, detect_format(&plain)?, opts);
    }
    extract_markdown_with(data, detect_format(data)?, opts)
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
#[allow(clippy::needless_pass_by_value)] // matches the extract_*_to family
pub fn extract_sheets_to(
    data: &[u8],
    format: Format,
    opts: ExtractOptions,
    sink: &mut impl SheetSink,
) -> Result<()> {
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_sheets_to(&plain, detect_format(&plain)?, opts, sink);
    }
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

/// Convenience: detect (decrypting encrypted Office) and extract sheets.
///
/// # Errors
///
/// See [`extract_sheets_with`]; adds the same encryption errors.
pub fn to_sheets_with(data: &[u8], opts: ExtractOptions) -> Result<Vec<Sheet>> {
    if let Some(plain) = decrypt_office(data, &opts)? {
        return extract_sheets_with(&plain, detect_format(&plain)?, opts);
    }
    extract_sheets_with(data, detect_format(data)?, opts)
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

    #[cfg(feature = "ocr")]
    #[test]
    fn extract_image_plain_path_errors_without_text() {
        // Real OCR needs models; the garbage path must not.
        let err = extract_plain_with(b"garbage", Format::Image, ExtractOptions::default())
            .unwrap_err()
            .to_string();
        assert!(err.contains("no text found in image"));
    }

    #[cfg(feature = "ocr")]
    #[test]
    fn extract_plain_to_equals_extract_plain_on_image_garbage() {
        let data = b"garbage";
        let format = Format::Image;
        let opts = ExtractOptions::default();
        let a = extract_plain_with(data, format, opts.clone())
            .unwrap_err()
            .to_string();
        let mut out = String::new();
        let b = extract_plain_to(data, format, opts, &mut out)
            .unwrap_err()
            .to_string();
        assert_eq!(a, b);
        assert!(out.is_empty());
    }

    #[cfg(not(feature = "ocr"))]
    #[test]
    fn extract_image_without_ocr_feature_errors() {
        let err = extract_plain_with(b"\x89PNG\r\n", Format::Image, ExtractOptions::default())
            .unwrap_err()
            .to_string();
        assert!(err.contains("`ocr` feature"), "got: {err}");
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
    fn extract_options_default_password_is_none() {
        assert!(ExtractOptions::default().password.is_none());
    }

    #[test]
    fn extract_options_debug_redacts_password() {
        let opts = ExtractOptions {
            password: Some("hunter2".into()),
            ..Default::default()
        };
        let debug = format!("{opts:?}");
        assert!(!debug.contains("hunter2"), "password leaked: {debug}");
        assert!(debug.contains("password: Some(\"<redacted>\")"), "{debug}");
        assert!(
            format!("{:?}", ExtractOptions::default()).contains("password: None"),
            "unset password should render as None"
        );
    }

    #[test]
    fn to_sheets_rejects_unrecognized() {
        let err = to_sheets(b"hello, definitely not a document")
            .unwrap_err()
            .to_string();
        assert_eq!(err, "not a supported document (unrecognized format)");
    }

    /// A minimal docx ZIP carrying `word/document.xml`, for detection tests.
    fn plain_docx() -> Vec<u8> {
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
    fn encrypted_office_detect_requires_password() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_docx(), "pw").unwrap();
        let err = detect_format(&enc).unwrap_err();
        assert!(matches!(err, BatdocError::PasswordRequired), "got {err:?}");
    }

    #[test]
    fn detect_format_with_password_resolves_inner_format() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_docx(), "pw").unwrap();
        assert_eq!(detect_format_with(&enc, Some("pw")).unwrap(), Format::Docx);
        let err = detect_format_with(&enc, None).unwrap_err();
        assert!(matches!(err, BatdocError::PasswordRequired));
    }

    /// A minimal valid PDF (empty page tree), for `needs_password` tests.
    ///
    /// A header-only stub will not do: the probe parses the trailer, and lopdf
    /// rejects a file with no cross-reference table.
    fn plain_pdf() -> Vec<u8> {
        use lopdf::dictionary;

        let mut doc = lopdf::Document::with_version("1.4");
        let pages_id = doc.new_object_id();
        let catalog_id = doc.add_object(lopdf::dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.objects.insert(
            pages_id,
            lopdf::Object::Dictionary(lopdf::dictionary! {
                "Type" => "Pages",
                "Kids" => lopdf::Object::Array(Vec::new()),
                "Count" => 0,
            }),
        );
        doc.trailer.set("Root", catalog_id);
        let mut out = Vec::new();
        doc.save_to(&mut out).unwrap();
        out
    }

    #[test]
    fn needs_password_matrix() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_docx(), "pw").unwrap();
        assert!(needs_password(&enc).unwrap());
        assert!(!needs_password(&plain_docx()).unwrap());
        assert!(!needs_password(&plain_pdf()).unwrap());
    }

    #[test]
    fn needs_password_errors_on_unparseable_pdf() {
        let err = needs_password(b"%PDF-1.4\n%%EOF\n").unwrap_err();
        assert!(matches!(err, BatdocError::Document(_)), "got {err:?}");
    }

    #[test]
    fn encrypted_docx_extracts_with_password() {
        let docx = {
            use std::io::Write;
            let mut buf = std::io::Cursor::new(Vec::new());
            let mut z = zip::ZipWriter::new(&mut buf);
            z.start_file(
                "word/document.xml",
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
            z.write_all(b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>SecretOfficeText</w:t></w:r></w:p></w:body></w:document>")
                .unwrap();
            z.finish().unwrap();
            buf.into_inner()
        };
        let enc = msoffice_crypto::encrypt_ooxml(&docx, "pw").unwrap();

        let err = to_plain(&enc).unwrap_err();
        assert!(matches!(err, BatdocError::PasswordRequired), "got {err:?}");

        let opts = ExtractOptions {
            password: Some("pw".into()),
            ..Default::default()
        };
        let text = to_plain_with(&enc, opts).unwrap();
        assert!(text.contains("SecretOfficeText"), "got {text:?}");
    }

    #[test]
    fn encrypted_docx_wrong_password_is_incorrect() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_docx(), "pw").unwrap();
        let opts = ExtractOptions {
            password: Some("nope".into()),
            ..Default::default()
        };
        let err = to_plain_with(&enc, opts).unwrap_err();
        assert!(matches!(err, BatdocError::IncorrectPassword), "got {err:?}");
    }

    #[test]
    fn explicit_format_entry_decrypts_encrypted_office() {
        let enc = msoffice_crypto::encrypt_ooxml(&plain_docx(), "pw").unwrap();
        // A caller who pre-detected with the password can still pass the
        // original encrypted bytes to the explicit entry point.
        let format = detect_format_with(&enc, Some("pw")).unwrap();
        let opts = ExtractOptions {
            password: Some("pw".into()),
            ..Default::default()
        };
        assert!(extract_plain_with(&enc, format, opts).is_ok());
    }
}
