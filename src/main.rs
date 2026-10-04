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
      --password SECRET
                    Password for encrypted PDF/Office documents. If omitted
                    and stdin is a terminal, you are prompted.
      --strip-text STR
                    Remove rotated PDF text containing STR from markdown
                    output (repeatable, case- and whitespace-insensitive).
      --strip-watermarks
                    Remove diagonal PDF text repeated across pages from
                    markdown output
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
without a text layer is OCR'd automatically — from its embedded images, or
(if it has none) from a rendered bitmap of the page. Image files (.png/.jpg/
.gif/.webp/.bmp) are always OCR'd, with or without --ocr.

--strip-text matches against text runs reconstructed along their true
rotation, so it removes watermarks drawn at an angle (which otherwise
shatter into per-letter noise). Matching ignores case and whitespace:
--strip-text draftcopy also matches \"Draft Copy\". Repeat
the flag to strip several strings.

--strip-watermarks needs no string: it removes diagonal text whose written
form repeats across pages. It requires at least two pages to learn a
signature and is a no-op on a single-page document (use --strip-text
there). Horizontal repeated text is left alone, so headers and footers
survive. Both options apply to PDF markdown output only. Piped output
defaults to plain text, which they do not change, so pass -m when
redirecting to a file:

    batdoc -m --strip-watermarks report.pdf > report.md

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

/// Consume the value that follows a `--flag` argument. Exits with `code`
/// (usage error) when the flag was given without one.
fn flag_value(args: &[String], i: &mut usize, flag: &str, code: i32) -> String {
    *i += 1;
    args.get(*i).cloned().unwrap_or_else(|| {
        eprintln!("batdoc: {flag} requires a value");
        eprintln!("{USAGE}");
        process::exit(code);
    })
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut mode = Mode::Auto;
    let mut images = false;
    let mut ocr = false;
    let mut strip_text: Vec<String> = Vec::new();
    let mut strip_watermarks = false;
    let mut password: Option<String> = None;
    let mut files: Vec<String> = Vec::new();

    let mut i = 0;
    while i < args.len() {
        let arg = &args[i];
        match arg.as_str() {
            "-h" | "--help" => {
                println!("{USAGE}");
                return;
            }
            "-p" | "--plain" => mode = Mode::Plain,
            "-m" | "--markdown" => mode = Mode::Markdown,
            "-i" | "--images" => images = true,
            "--ocr" => ocr = true,
            "--strip-watermarks" => strip_watermarks = true,
            "--password" => password = Some(flag_value(&args, &mut i, "--password", 2)),
            s if s.starts_with("--password=") => {
                password = Some(s["--password=".len()..].to_string());
            }
            "--strip-text" => {
                i += 1;
                if let Some(value) = args.get(i) {
                    strip_text.push(value.clone());
                } else {
                    eprintln!("batdoc: --strip-text requires a value");
                    process::exit(1);
                }
            }
            s if s.starts_with("--strip-text=") => {
                strip_text.push(s["--strip-text=".len()..].to_string());
            }
            "-" => files.push("-".to_string()),
            s if s.starts_with('-') => {
                eprintln!("batdoc: unknown option: {s}");
                eprintln!("{USAGE}");
                process::exit(1);
            }
            _ => files.push(arg.clone()),
        }
        i += 1;
    }

    // No files specified → read from stdin
    if files.is_empty() {
        files.push("-".to_string());
    }

    // Only prompt when the user did not supply --password and stdin is a
    // terminal; with a supplied password a failure is reported immediately.
    let allow_prompt = password.is_none() && io::stdin().is_terminal();

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

        if let Err(e) = run(
            &buf,
            &filename,
            mode,
            images,
            ocr,
            &strip_text,
            strip_watermarks,
            files.len() > 1 && i > 0,
            password.clone(),
            allow_prompt,
        ) {
            eprintln!("batdoc: {filename}: {e}");
            exit_code = 1;
        }
    }

    if exit_code != 0 {
        process::exit(exit_code);
    }
}

#[allow(clippy::too_many_arguments, clippy::fn_params_excessive_bools)] // thin CLI shim; each is a distinct flag
fn run(
    data: &[u8],
    filename: &str,
    mode: Mode,
    images: bool,
    ocr: bool,
    strip_text: &[String],
    strip_watermarks: bool,
    needs_separator: bool,
    password: Option<String>,
    allow_prompt: bool,
) -> batdoc_core::Result<()> {
    use batdoc_core::ExtractOptions;

    let mut opts = ExtractOptions {
        images,
        ocr,
        strip_text: strip_text.to_vec(),
        strip_watermarks,
        password,
        ..Default::default()
    };

    // An encrypted document rejects the first attempt with a typed error.
    // Prompt and retry — twice, so a typo is recoverable — but only when
    // interactive; a supplied `--password` that fails is reported as-is.
    let mut prompts = 0u32;
    loop {
        match run_once(data, filename, mode, &opts, needs_separator) {
            Err(BatdocError::PasswordRequired | BatdocError::IncorrectPassword)
                if allow_prompt && prompts < 2 =>
            {
                prompts += 1;
                opts.password = Some(prompt_password()?);
            }
            other => return other,
        }
    }
}

/// Prompt for a password without echo. Returns [`BatdocError::Io`] if the
/// terminal cannot be read.
fn prompt_password() -> batdoc_core::Result<String> {
    rpassword::prompt_password("Password: ").map_err(BatdocError::Io)
}

fn run_once(
    data: &[u8],
    filename: &str,
    mode: Mode,
    opts: &batdoc_core::ExtractOptions,
    needs_separator: bool,
) -> batdoc_core::Result<()> {
    let format = batdoc_core::detect_format_with(data, opts.password.as_deref())?;
    let is_tty = io::stdout().is_terminal();

    // The strip flags act on the positioned (glyph) pipeline, which only the
    // markdown path uses. Piped output is plain by default, so say so rather
    // than silently doing nothing.
    let plain_output = match mode {
        Mode::Plain => true,
        Mode::Markdown => false,
        Mode::Auto => !is_tty,
    };
    if format == Format::Pdf
        && plain_output
        && (!opts.strip_text.is_empty() || opts.strip_watermarks)
    {
        static NOTICE: std::sync::Once = std::sync::Once::new();
        NOTICE.call_once(|| {
            eprintln!(
                "batdoc: --strip-text/--strip-watermarks affect markdown output only; \
                 plain output is unchanged. Use -m to force markdown."
            );
        });
    }

    if needs_separator && !is_tty {
        io::stdout().write_all(b"\n")?;
    }

    // OCR input (flagged, or image input which is always OCR'd) downloads
    // models on first use; say so once per process, before it happens.
    if (opts.ocr || format == Format::Image) && !batdoc_core::models_present() {
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
            let mut sink = batdoc_core::IoSink(io::stdout());
            batdoc_core::extract_plain_to(data, format, opts.clone(), &mut sink)?;
        }
        Mode::Markdown => {
            if is_tty && format != Format::Image {
                let md = batdoc_core::extract_markdown_with(data, format, opts.clone())?;
                pretty_print(&md, filename)?;
            } else {
                let mut sink = batdoc_core::IoSink(io::stdout());
                batdoc_core::extract_markdown_to(data, format, opts.clone(), &mut sink)?;
            }
        }
        Mode::Auto => {
            if is_tty && format != Format::Image {
                let md = batdoc_core::extract_markdown_with(data, format, opts.clone())?;
                pretty_print(&md, filename)?;
            } else {
                let mut sink = batdoc_core::IoSink(io::stdout());
                batdoc_core::extract_plain_to(data, format, opts.clone(), &mut sink)?;
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
