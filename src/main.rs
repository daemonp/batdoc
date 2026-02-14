//! `batdoc` — bat for `.doc`, `.docx`, `.xls`, `.xlsx`, `.pptx`, and `.pdf` files.
//!
//! Reads legacy OLE2 `.doc` and `.xls`, modern OOXML `.docx`, `.xlsx`, and
//! `.pptx`, and PDF files and dumps their text to stdout. When stdout is a
//! terminal the output is pretty-printed as syntax-highlighted markdown via
//! `bat`; when piped, plain text is emitted.

#![allow(clippy::redundant_pub_crate)]

mod codepage;
mod dateconv;
mod doc;
mod docx;
mod error;
mod heuristic;
mod markup;
mod pdf;
mod pptx;
mod sheet;
mod xls;
mod xlsx;
mod xml_util;

use error::BatdocError;

use bat::{Input, PrettyPrinter};
use is_terminal::IsTerminal;
use std::io::{self, Read, Write};
use std::process;

const USAGE: &str = "\
batdoc - bat for .doc, .docx, .xls, .xlsx, .pptx, and .pdf files

Usage: batdoc [OPTIONS] [FILE...]
       cat FILE | batdoc [OPTIONS]
       batdoc [OPTIONS] -

Options:
  -p, --plain       Force plain text output (no colors, no decorations)
  -m, --markdown    Output as markdown (default when terminal detected)
  -h, --help        Show this help

When stdout is a terminal, output is pretty-printed as syntax-highlighted
markdown with decorations. When piped, output is plain text.

Multiple files can be specified and will be processed in order.
Use - to read from stdin explicitly.

Supports legacy .doc/.xls (OLE2), modern .docx/.xlsx/.pptx (OOXML), and .pdf.
Format is detected by magic bytes, not file extension.";

/// Maximum input file size (256 MiB). Prevents accidental OOM from
/// huge files or zip bombs.
const MAX_INPUT_SIZE: usize = 256 * 1024 * 1024;

// Magic signatures
const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];
const PDF_MAGIC: [u8; 5] = [0x25, 0x50, 0x44, 0x46, 0x2D]; // %PDF-

/// Output mode selection.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Mode {
    /// Detect automatically: markdown to terminal, plain text when piped.
    Auto,
    /// Force plain text output.
    Plain,
    /// Force markdown output.
    Markdown,
}

/// Detected document format based on magic bytes.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Format {
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

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut mode = Mode::Auto;
    let mut files: Vec<String> = Vec::new();

    for arg in &args {
        match arg.as_str() {
            "-h" | "--help" => {
                println!("{USAGE}");
                return;
            }
            "-p" | "--plain" => mode = Mode::Plain,
            "-m" | "--markdown" => mode = Mode::Markdown,
            "-" => files.push("-".to_string()),
            s if s.starts_with('-') => {
                eprintln!("batdoc: unknown option: {s}");
                eprintln!("{USAGE}");
                process::exit(1);
            }
            _ => files.push(arg.clone()),
        }
    }

    // No files specified → read from stdin
    if files.is_empty() {
        files.push("-".to_string());
    }

    let mut exit_code = 0;
    for (i, path) in files.iter().enumerate() {
        let (buf, filename) = if path == "-" {
            let mut buf = Vec::new();
            if let Err(e) = io::stdin().read_to_end(&mut buf) {
                eprintln!("batdoc: stdin: {e}");
                exit_code = 1;
                continue;
            }
            (buf, "stdin".to_string())
        } else {
            match std::fs::read(path) {
                Ok(b) => (b, path.clone()),
                Err(e) => {
                    eprintln!("batdoc: {path}: {e}");
                    exit_code = 1;
                    continue;
                }
            }
        };

        if buf.len() > MAX_INPUT_SIZE {
            #[allow(clippy::cast_precision_loss)] // only used in error message
            let size_mib = buf.len() as f64 / (1024.0 * 1024.0);
            eprintln!(
                "batdoc: {filename}: too large ({size_mib:.1} MiB, max {} MiB)",
                MAX_INPUT_SIZE / (1024 * 1024),
            );
            exit_code = 1;
            continue;
        }

        let multiple = files.len() > 1;

        if let Err(e) = run(&buf, &filename, mode, multiple && i > 0) {
            eprintln!("batdoc: {filename}: {e}");
            exit_code = 1;
        }
    }

    if exit_code != 0 {
        process::exit(exit_code);
    }
}

/// Detect the document format from magic bytes.
///
/// For OLE2 formats, peeks inside the compound file to distinguish
/// `.doc` (has `WordDocument` stream) from `.xls` (has `Workbook` stream).
/// For ZIP-based formats, peeks inside the archive to distinguish
/// `.docx` (has `word/document.xml`) from `.xlsx` (has `xl/workbook.xml`).
fn detect_format(data: &[u8]) -> error::Result<Format> {
    if data.len() >= 8 && data[..8] == OLE2_MAGIC {
        let cursor = std::io::Cursor::new(data);
        let cfb = cfb::CompoundFile::open(cursor)?;
        if cfb.exists("/WordDocument") {
            Ok(Format::Doc)
        } else if cfb.exists("/Workbook") || cfb.exists("/Book") {
            Ok(Format::Xls)
        } else {
            Err(BatdocError::Document(
                "OLE2 file is not a .doc or .xls document".into(),
            ))
        }
    } else if data.len() >= 5 && data[..5] == PDF_MAGIC {
        Ok(Format::Pdf)
    } else if data.len() >= 4 && data[..4] == ZIP_MAGIC {
        let cursor = std::io::Cursor::new(data);
        let archive = zip::ZipArchive::new(cursor)?;
        if archive.index_for_name("word/document.xml").is_some() {
            Ok(Format::Docx)
        } else if archive.index_for_name("xl/workbook.xml").is_some() {
            Ok(Format::Xlsx)
        } else if archive.index_for_name("ppt/presentation.xml").is_some() {
            Ok(Format::Pptx)
        } else {
            Err(BatdocError::Document(
                "ZIP archive is not a .docx, .xlsx, or .pptx file".into(),
            ))
        }
    } else {
        Err(BatdocError::Document(
            "not a supported document (unrecognized format)".into(),
        ))
    }
}

fn run(data: &[u8], filename: &str, mode: Mode, needs_separator: bool) -> error::Result<()> {
    let format = detect_format(data)?;
    let is_tty = io::stdout().is_terminal();

    if needs_separator && !is_tty {
        io::stdout().write_all(b"\n")?;
    }

    match mode {
        Mode::Plain => {
            let text = extract_plain(data, format)?;
            io::stdout().write_all(text.as_bytes())?;
        }
        Mode::Markdown => {
            let md = extract_markdown(data, format)?;
            if is_tty {
                pretty_print(&md, filename)?;
            } else {
                io::stdout().write_all(md.as_bytes())?;
            }
        }
        Mode::Auto => {
            if is_tty {
                let md = extract_markdown(data, format)?;
                pretty_print(&md, filename)?;
            } else {
                let text = extract_plain(data, format)?;
                io::stdout().write_all(text.as_bytes())?;
            }
        }
    }

    Ok(())
}

fn extract_plain(data: &[u8], format: Format) -> error::Result<String> {
    match format {
        Format::Doc => doc::extract_plain(data),
        Format::Xls => xls::extract_plain(data),
        Format::Docx => docx::extract_plain(data),
        Format::Xlsx => xlsx::extract_plain(data),
        Format::Pptx => pptx::extract_plain(data),
        Format::Pdf => pdf::extract_plain(data),
    }
}

fn extract_markdown(data: &[u8], format: Format) -> error::Result<String> {
    match format {
        Format::Doc => doc::extract_markdown(data),
        Format::Xls => xls::extract_markdown(data),
        Format::Docx => docx::extract_markdown(data),
        Format::Xlsx => xlsx::extract_markdown(data),
        Format::Pptx => pptx::extract_markdown(data),
        Format::Pdf => pdf::extract_markdown(data),
    }
}

fn pretty_print(content: &str, filename: &str) -> error::Result<()> {
    let input = Input::from_bytes(content.as_bytes())
        .name(filename)
        .title(filename);

    let theme = std::env::var("BAT_THEME").unwrap_or_else(|_| "ansi".to_string());

    PrettyPrinter::new()
        .input(input)
        .language("Markdown")
        .theme(&theme)
        .header(true)
        .line_numbers(false)
        .grid(true)
        .colored_output(true)
        .true_color(true)
        .paging_mode(bat::PagingMode::QuitIfOneScreen)
        .print()
        .map_err(|e| BatdocError::Render(e.to_string()))?;

    Ok(())
}
