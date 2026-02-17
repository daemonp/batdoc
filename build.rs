use man::prelude::*;
use std::path::Path;

fn main() {
    let page = Manual::new("batdoc")
        .about("cat(1) for doc, docx, xls, xlsx, pptx, and pdf — renders to markdown with bat")
        .author(Author::new("Damon Petta").email("d@disassemble.net"))
        .flag(
            Flag::new()
                .short("-p")
                .long("--plain")
                .help("Force plain text output (no colors, no decorations)."),
        )
        .flag(
            Flag::new()
                .short("-m")
                .long("--markdown")
                .help("Output as markdown (default when terminal detected)."),
        )
        .flag(Flag::new().short("-i").long("--images").help(
            "Embed images as inline base64 data URIs in markdown output. \
                     Extracts embedded images from .docx, .pptx, and .xlsx files. \
                     Most useful when piping to a file \
                     (batdoc --images report.docx > out.md). \
                     Ignored in plain text mode and for formats without image \
                     support (.doc, .xls, .pdf).",
        ))
        .flag(
            Flag::new()
                .short("-h")
                .long("--help")
                .help("Show help information."),
        )
        .arg(Arg::new("[FILE...]"))
        .custom(
            Section::new("description")
                .paragraph(
                    "batdoc reads Office documents and PDFs and dumps their contents \
                     to the terminal as markdown. It is a spiritual successor to \
                     catdoc(1) — cat had catdoc, bat gets batdoc.",
                )
                .paragraph(
                    "Format is detected by magic bytes (file signature), not file \
                     extension. Supported formats: .doc (OLE2 Word 97+), .docx \
                     (OOXML), .xls (BIFF8 Excel 97+), .xlsx (OOXML), .pptx \
                     (OOXML), and .pdf.",
                )
                .paragraph(
                    "When stdout is a terminal, output is pretty-printed as \
                     syntax-highlighted markdown via bat(1) with paging. When \
                     piped, plain text is emitted.",
                )
                .paragraph(
                    "Multiple files can be specified and will be processed in \
                     order. Use \\fB-\\fR to read from stdin explicitly. Maximum \
                     input size is 256 MiB.",
                ),
        )
        .example(
            Example::new()
                .text("View a Word document in the terminal")
                .command("batdoc report.docx"),
        )
        .example(
            Example::new()
                .text("Extract a spreadsheet as plain-text TSV")
                .command("batdoc --plain data.xlsx > data.tsv"),
        )
        .example(
            Example::new()
                .text("Convert a presentation to markdown with embedded images")
                .command("batdoc --images slides.pptx > slides.md"),
        )
        .example(
            Example::new()
                .text("Read from stdin")
                .command("curl -sL https://example.com/file.docx | batdoc"),
        )
        .custom(
            Section::new("environment")
                .paragraph(
                    "batdoc respects the \\fBNO_COLOR\\fR environment variable. \
                     When set, colored output is suppressed even on a terminal.",
                )
                .paragraph(
                    "The \\fBPAGER\\fR environment variable controls which pager \
                     is used when output is displayed on a terminal.",
                ),
        )
        .custom(Section::new("see also").paragraph("bat(1), catdoc(1), pdftotext(1)"))
        .render();

    // Write to OUT_DIR (standard cargo output directory)
    let out_dir = std::env::var("OUT_DIR").unwrap();
    let out_path = Path::new(&out_dir).join("batdoc.1");
    std::fs::write(&out_path, &page).unwrap();

    // Also write to target/man/ so packaging scripts have a stable path
    // that doesn't depend on the hash-based OUT_DIR.
    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").unwrap();
    let man_dir = Path::new(&manifest_dir).join("target").join("man");
    std::fs::create_dir_all(&man_dir).unwrap();
    std::fs::write(man_dir.join("batdoc.1"), &page).unwrap();

    println!("cargo::rerun-if-changed=build.rs");
}
