//! `batdoc` — bat for `.doc`, `.docx`, `.xls`, `.xlsx`, `.pptx`, `.pdf`, and image files.
//!
//! Reads legacy OLE2 `.doc` and `.xls`, modern OOXML `.docx`, `.xlsx`, and
//! `.pptx`, PDF, and raster image files and dumps their text to stdout. Image
//! files are always OCR'd; textless PDF pages are OCR'd automatically as a
//! fallback (a textless PDF is a scan). Embedded images in DOCX/PPTX are OCR'd
//! with `--ocr`. When stdout is a terminal the output is pretty-printed as
//! syntax-highlighted markdown via `bat`; when piped, plain text is emitted.

use batdoc_core::{BatdocError, Format};

use bat::{Input, PrettyPrinter};
use is_terminal::IsTerminal;
use std::io::{self, Read, Write};
use std::process;

const USAGE: &str = "\
batdoc - bat for .doc, .docx, .xls, .xlsx, .pptx, .pdf, and image files

Usage: batdoc [OPTIONS] [FILE...]
       cat FILE | batdoc [OPTIONS]
       batdoc [OPTIONS] -

Options:
  -p, --plain       Force plain text output (no colors, no decorations)
  -m, --markdown    Output as markdown (default when terminal detected)
  -i, --images      Embed images as inline base64 data URIs in markdown
      --ocr         OCR embedded images (docx/pptx); textless PDFs already auto-OCR
  -h, --help        Show this help

When stdout is a terminal, output is pretty-printed as syntax-highlighted
markdown with decorations. When piped, output is plain text.

--images extracts embedded images from .docx, .pptx, and .xlsx files and
includes them as ![](data:image/...;base64,...) in the markdown output.
Most useful when piping to a file (batdoc --images report.docx > out.md).
Ignored in plain text mode and for formats without image support (.doc, .xls, .pdf).

--ocr uses the ocrs engine (models downloaded on first use to
$BATDOC_MODELS_DIR, $XDG_CACHE_HOME/batdoc/models, or ~/.cache/batdoc/models).
For .docx/.pptx, embedded images are OCR'd. PDFs need no flag: any page
without a text layer is OCR'd automatically from its embedded images as a
fallback (a textless PDF is a scan). Image files (.png/.jpg/.gif/
.webp/.bmp) are always OCR'd, with or without --ocr.

Multiple files can be specified and will be processed in order.
Use - to read from stdin explicitly.

Supports legacy .doc/.xls (OLE2), modern .docx/.xlsx/.pptx (OOXML), .pdf,
and raster images. Format is detected by magic bytes, not file extension.";

/// Maximum input file size (256 MiB). Prevents accidental OOM from
/// huge files or zip bombs.
const MAX_INPUT_SIZE: usize = 256 * 1024 * 1024;

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

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut mode = Mode::Auto;
    let mut images = false;
    let mut ocr = false;
    let mut files: Vec<String> = Vec::new();

    for arg in &args {
        match arg.as_str() {
            "-h" | "--help" => {
                println!("{USAGE}");
                return;
            }
            "-p" | "--plain" => mode = Mode::Plain,
            "-m" | "--markdown" => mode = Mode::Markdown,
            "-i" | "--images" => images = true,
            "--ocr" => ocr = true,
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

        if let Err(e) = run(&buf, &filename, mode, images, ocr, multiple && i > 0) {
            eprintln!("batdoc: {filename}: {e}");
            exit_code = 1;
        }
    }

    if exit_code != 0 {
        process::exit(exit_code);
    }
}

fn run(
    data: &[u8],
    filename: &str,
    mode: Mode,
    images: bool,
    ocr: bool,
    needs_separator: bool,
) -> batdoc_core::Result<()> {
    use batdoc_core::ExtractOptions;

    let format = batdoc_core::detect_format(data)?;
    let is_tty = io::stdout().is_terminal();

    if needs_separator && !is_tty {
        io::stdout().write_all(b"\n")?;
    }

    let opts = ExtractOptions {
        images,
        ocr,
        ..Default::default()
    };

    // OCR input (flagged, or image input which is always OCR'd) downloads
    // models on first use; say so once per process, before it happens.
    if (ocr || format == Format::Image) && !batdoc_core::models_present() {
        static NOTICE: std::sync::Once = std::sync::Once::new();
        NOTICE.call_once(|| {
            eprintln!(
                "batdoc: OCR models not cached; downloading on first use \
                 (set BATDOC_MODELS_DIR to override the cache location)"
            );
        });
    }

    match mode {
        Mode::Plain => {
            let text = batdoc_core::extract_plain_with(data, format, opts)?;
            io::stdout().write_all(text.as_bytes())?;
        }
        Mode::Markdown => {
            let md = batdoc_core::extract_markdown_with(data, format, opts)?;
            if is_tty && format != Format::Image {
                pretty_print(&md, filename)?;
            } else {
                io::stdout().write_all(md.as_bytes())?;
            }
        }
        Mode::Auto => {
            if is_tty && format != Format::Image {
                let md = batdoc_core::extract_markdown_with(data, format, opts)?;
                pretty_print(&md, filename)?;
            } else {
                let text = batdoc_core::extract_plain_with(data, format, opts)?;
                io::stdout().write_all(text.as_bytes())?;
            }
        }
    }

    Ok(())
}

fn pretty_print(content: &str, filename: &str) -> batdoc_core::Result<()> {
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
