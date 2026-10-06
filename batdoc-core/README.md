# batdoc-core

Document text extraction library for Rust. Converts `.doc`, `.docx`, `.xls`,
`.xlsx`, `.pptx`, and `.pdf` files — and raster images via OCR — to plain text
or Markdown.

Format detection is by magic bytes, not file extension — works on raw byte
buffers without filesystem access, which makes it suitable for email
attachments, HTTP uploads, and other contexts where filenames are unreliable.

## Usage

```toml
[dependencies]
batdoc-core = { git = "https://github.com/daemonp/batdoc" }
```

```rust
use batdoc_core::{detect_format, extract_markdown, extract_plain, to_markdown, Format};

// One-shot: detect format and extract in a single call
let data: Vec<u8> = std::fs::read("report.docx").unwrap();
let markdown = batdoc_core::to_markdown(&data, false).unwrap();
let plain = batdoc_core::to_plain(&data).unwrap();

// Two-step: detect first, then extract (useful for logging or branching)
let format = batdoc_core::detect_format(&data).unwrap();
println!("Detected format: {format}");
let text = batdoc_core::extract_plain(&data, format).unwrap();
```

## Public API

```rust
pub enum Format { Doc, Xls, Docx, Xlsx, Pptx, Pdf, Image }

// #[non_exhaustive]: match with a wildcard arm.
pub enum BatdocError {
    Io, Zip, Document, Render,
    PasswordRequired,               // encrypted, no usable password supplied
    IncorrectPassword,              // supplied password did not authenticate
    UnsupportedEncryption(String),  // encryption this build does not handle
}
pub type Result<T> = std::result::Result<T, BatdocError>;

pub struct ExtractOptions {
    pub images: bool,            // embed images as base64 data URIs (markdown mode only)
    pub ocr: bool,               // OCR embedded images (DOCX/PPTX)
    pub auto_ocr: bool,          // textless/garbled PDF fallback (default true)
    pub max_output_bytes: Option<u64>,
    pub password: Option<String>, // password for an encrypted PDF/Office document
}

pub fn detect_format(data: &[u8]) -> Result<Format>;
pub fn detect_format_with(data: &[u8], password: Option<&str>) -> Result<Format>;
pub fn needs_password(data: &[u8]) -> Result<bool>;
pub fn extract_plain(data: &[u8], format: Format) -> Result<String>;
pub fn extract_plain_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String>;
pub fn extract_markdown(data: &[u8], format: Format, images: bool) -> Result<String>;
pub fn extract_markdown_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String>;
pub fn to_plain(data: &[u8]) -> Result<String>;
pub fn to_plain_with(data: &[u8], opts: ExtractOptions) -> Result<String>;
pub fn to_markdown(data: &[u8], images: bool) -> Result<String>;
pub fn to_markdown_with(data: &[u8], opts: ExtractOptions) -> Result<String>;
pub fn to_sheets_with(data: &[u8], opts: ExtractOptions) -> Result<Vec<Sheet>>;

pub struct Sheet { pub name: String, pub rows: Vec<Vec<String>> }
```

`extract_markdown` with `images: true` embeds images from DOCX/XLSX/PPTX as
base64 data URIs. Has no effect on DOC, XLS, PDF, or Image.

Set `ExtractOptions.auto_ocr = false` to disable the automatic
textless/garbled-PDF OCR fallback (no model download or requirement).
`Format::Image` is always OCR'd.

Image OCR, embedded-image OCR, and the PDF fallback are behind the
default-on `ocr` feature; `default-features = false` removes them
together with the `ocrs`/`rten`/`image` dependencies.

```rust
// Raster images are always OCR'd — no options needed for `Format::Image`.
let text = batdoc_core::extract_plain_with(&png, Format::Image, ExtractOptions::default()).unwrap();
```

## Password-protected documents

Encrypted PDFs and Office documents (`.docx`/`.xlsx`/`.pptx`) are decrypted
when `ExtractOptions::password` is set:

```rust
let mut opts = batdoc_core::ExtractOptions::default();
opts.password = Some(std::env::var("BATDOC_PASSWORD")?);
let markdown = batdoc_core::to_markdown_with(&data, opts)?;
```

`to_plain_with`, `to_markdown_with`, and `to_sheets_with` decrypt encrypted
Office input before extracting; the `extract_*_with` functions do the same, so
the `format` argument may describe the decrypted package even though `data` is
still the ciphertext. `password: None` means "not supplied": an encrypted
Office package reports `BatdocError::PasswordRequired` without guessing, while
a PDF still tries the empty user password (owner-only locks). A wrong password
reports `BatdocError::IncorrectPassword`.

`detect_format` cannot name the format of an encrypted Office package and
returns `BatdocError::PasswordRequired`; use
`detect_format_with(data, Some(pw))` to detect through the encryption, or
`needs_password(data)` to probe for encryption without a password. Probing a
PDF parses its encryption dictionary, so it is not free.

Encrypted legacy `.doc`/`.xls` files are detected but not decrypted — they
report `BatdocError::UnsupportedEncryption`. `BatdocError` is
`#[non_exhaustive]`, so matches need a wildcard arm. Passwords never appear in
error messages or `Debug` output (`ExtractOptions`'s `Debug` renders the
password as `<redacted>`). The same Office path links on `wasm32`; this repo
vendors `msoffice-crypto` for that (see `crates/msoffice-crypto/BATDOC-FORK.md`
and [WASM.md](../WASM.md)).

## Supported formats

| Format | Detection | Parser |
| -------- | ----------- | -------- |
| `.doc` | OLE2 magic + `/WordDocument` stream | Binary Word 97+ (BIFF-like) |
| `.xls` | OLE2 magic + `/Workbook` stream | BIFF8 (Excel 97+) |
| `.docx` | ZIP magic + `word/document.xml` | OOXML |
| `.xlsx` | ZIP magic + `xl/workbook.xml` | OOXML |
| `.pptx` | ZIP magic + `ppt/presentation.xml` | OOXML |
| `.pdf` | `%PDF-` header | pdf-extract |
| `.png` `.jpg` `.gif` `.webp` `.bmp` | file magic | OCR (ocrs) |

Raster images (`Format::Image`) are always OCR'd; the first use downloads the
OCR models to `$BATDOC_MODELS_DIR`.

## License

MIT
