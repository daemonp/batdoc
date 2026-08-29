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
pub enum BatdocError { Io, Zip, Document, Render }
pub type Result<T> = std::result::Result<T, BatdocError>;

pub struct ExtractOptions {
    pub images: bool,            // embed images as base64 data URIs (markdown mode only)
    pub ocr: bool,               // OCR embedded images (DOCX/PPTX)
    pub auto_ocr: bool,          // textless/garbled PDF fallback (default true)
    pub max_output_bytes: Option<u64>,
}

pub fn detect_format(data: &[u8]) -> Result<Format>;
pub fn extract_plain(data: &[u8], format: Format) -> Result<String>;
pub fn extract_plain_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String>;
pub fn extract_markdown(data: &[u8], format: Format, images: bool) -> Result<String>;
pub fn extract_markdown_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String>;
pub fn to_plain(data: &[u8]) -> Result<String>;
pub fn to_markdown(data: &[u8], images: bool) -> Result<String>;
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
