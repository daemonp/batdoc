# batdoc-core

Document text extraction library for Rust. Converts `.doc`, `.docx`, `.xls`,
`.xlsx`, `.pptx`, and `.pdf` files to plain text or Markdown.

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
pub enum Format { Doc, Xls, Docx, Xlsx, Pptx, Pdf }
pub enum BatdocError { Io, Zip, Document, Render }
pub type Result<T> = std::result::Result<T, BatdocError>;

pub fn detect_format(data: &[u8]) -> Result<Format>;
pub fn extract_plain(data: &[u8], format: Format) -> Result<String>;
pub fn extract_markdown(data: &[u8], format: Format, images: bool) -> Result<String>;
pub fn to_plain(data: &[u8]) -> Result<String>;
pub fn to_markdown(data: &[u8], images: bool) -> Result<String>;
```

`extract_markdown` with `images: true` embeds images from DOCX/XLSX/PPTX as
base64 data URIs. Has no effect on DOC, XLS, or PDF.

## Supported formats

| Format | Detection | Parser |
|--------|-----------|--------|
| `.doc` | OLE2 magic + `/WordDocument` stream | Binary Word 97+ (BIFF-like) |
| `.xls` | OLE2 magic + `/Workbook` stream | BIFF8 (Excel 97+) |
| `.docx` | ZIP magic + `word/document.xml` | OOXML |
| `.xlsx` | ZIP magic + `xl/workbook.xml` | OOXML |
| `.pptx` | ZIP magic + `ppt/presentation.xml` | OOXML |
| `.pdf` | `%PDF-` header | pdf-extract |

## License

MIT
