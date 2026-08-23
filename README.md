# batdoc

`cat` had [catdoc](http://www.intevation.de/catdoc/). `bat` gets `batdoc`.

Dumps `.doc`, `.docx`, `.xls`, `.xlsx`, `.pptx`, and `.pdf` files to your
terminal as markdown — and OCRs image files (png/jpg/gif/webp/bmp). To a
tty it syntax-highlights and pages (using [bat](https://github.com/sharkdp/bat));
piped, it gives you plain text.

```
batdoc report.docx                     # highlighted markdown in terminal
batdoc financials.xlsx                  # each sheet becomes a markdown table
batdoc slides.pptx                     # per-slide headings with text
batdoc paper.pdf                       # multi-page PDF with page headers
batdoc photo.png                     # OCR — text from a photo or scan
batdoc --ocr scanned.pdf             # OCR pages that have no text layer
batdoc --plain legacy.doc > out.txt    # just the text
cat mystery.bin | batdoc               # stdin works, format detected by magic bytes
```

Format is detected by file signature, not extension. OLE2 files (`.doc`/`.xls`)
are distinguished by peeking at internal streams; ZIP files (`.docx`/`.xlsx`/`.pptx`)
by checking for `word/document.xml` vs `xl/workbook.xml` vs `ppt/presentation.xml`;
PDFs by the `%PDF-` header.

## For AI agents

Copy the block below to `~/.agents/skills/batdoc/SKILL.md` to give a local
agent a discoverable `batdoc` skill for processing documents and images:

````markdown
---
name: batdoc
description: Use when processing PDF, DOC, DOCX, XLS, XLSX, or PPTX files,
  or OCRing images (PNG/JPG/GIF/WEBP/BMP), to extract their text as
  markdown or plain text.
---

# batdoc

Dumps office documents and PDFs to the terminal as markdown, and OCRs
images. Format is detected by file signature, not extension — stdin works
too.

## When to use

- PDF, text layer or scanned → `batdoc paper.pdf` (textless pages auto-OCR)
- Office documents (doc/docx/xls/xlsx/pptx) → `batdoc report.docx`
- Images or scans of text → `batdoc photo.png` (image files are always OCR'd)

## Quick reference

    batdoc report.docx                  # markdown, highlighted, paged on a tty
    batdoc --plain legacy.doc > out.txt # plain text
    batdoc scanned.pdf                  # textless PDF pages auto-OCR
    batdoc --ocr report.docx            # OCR docx/pptx embedded images
    batdoc --images report.docx         # embed images as base64 data URIs
    cat mystery.bin | batdoc            # format detected from magic bytes

## Notes

- `--plain` for text-only output; markdown is the default on a tty.
- OCR needs no flag for PDFs or image files; `--ocr` covers docx/pptx
  embedded images and is a no-op for PDF/image input.
- OCR models (~12 MB) download on first use; pre-seed `$BATDOC_MODELS_DIR`
  for offline or package-managed installs.
- Use a release build — OCR is much slower in debug builds.
````

## Install

**Arch Linux (AUR):**

```
yay -S batdoc
```

**Homebrew:**

```
brew install daemonp/tap/batdoc
```

**Linux (x86_64, static musl):**

```
curl -sL https://github.com/daemonp/batdoc/releases/latest/download/batdoc-linux-x86_64.zst | zstd -d > batdoc && chmod +x batdoc
```

**macOS (Apple Silicon):**

```
curl -sL https://github.com/daemonp/batdoc/releases/latest/download/batdoc-darwin-aarch64.zst | zstd -d > batdoc && chmod +x batdoc
```

**From source:**

```
cargo build --release
cp target/release/batdoc ~/.local/bin/
```

## Formats

`.docx` and `.xlsx` are parsed structurally from their XML — headings,
bold/italic, lists, tables, and hyperlinks come through properly.
Spreadsheets render as markdown tables, one `##` section per sheet.
Hyperlinks in all formats are rendered as `[text](url)` in markdown.
Comments, footnotes, and endnotes are appended after the body when
present, with `[^1]` / `[^e1]` markers at each reference site.

`.doc` is trickier. The binary format buries style info in structures we
don't fully parse, so markdown structure is inferred heuristically from the
text: numbered headings, bold subheadings, tab-delimited tables. It works
well on typical business documents; your mileage varies on weirder layouts.

`.xls` gets a full BIFF8 parser — SST with CONTINUE record boundaries,
all the cell types (LABELSST, NUMBER, RK, MULRK, FORMULA, BOOLERR), hidden
sheet filtering, encryption detection. It shares the same rendering path
as `.xlsx`.

`.pptx` extracts text from all shapes on each slide. Font size is used to
infer heading levels. Hyperlinks on text runs are resolved and rendered as
markdown links. Multi-slide decks get `## Slide N` headings. Speaker
notes are appended after the deck under `## Notes` when present.

`.pdf` extracts text from text-based PDFs using `pdf-extract`. Multi-page
documents get `## Page N` headings in markdown mode. When a PDF has no text
layer at all (a scan), batdoc automatically falls back to OCR'ing its embedded
page images — no `--ocr` flag needed. A scanned PDF whose OCR also finds
nothing gets a clean error message. Malformed PDFs that would crash the
underlying library are caught and reported as errors rather than panics.

### PDF extraction notes

PDFs with broken or missing font mappings (garbled CID fonts) are recovered
via a vendored fork of `pdf-extract` under `crates/pdf-extract`, wired in with
`[patch.crates-io]`. Builds from this repository — releases, AUR, deb/rpm,
Homebrew — include the fix. `cargo install batdoc` resolves `pdf-extract` from
crates.io instead and gets upstream behavior: still safe (panics are caught), but
garbled documents stay garbled. Publishing the fork to close this gap is a
deferred follow-up.

## Options

```
batdoc [OPTIONS] [FILE...]
cat FILE | batdoc [OPTIONS]

  -p, --plain       plain text, no highlighting
  -m, --markdown    force markdown (default on tty)
  -i, --images      embed images as inline base64 data URIs
      --ocr         OCR embedded images (docx/pptx); textless PDFs auto-OCR
  -h, --help        help
```

`--images` extracts embedded images from `.docx`, `.pptx`, and `.xlsx`
files and includes them as `![](data:image/...;base64,...)` in the
markdown output. Most useful when piping to a file:

```
batdoc --images report.docx > report.md
```

The resulting markdown is self-contained — no external image files
needed. JPEG, PNG, GIF, WebP, and BMP images are supported; vector
formats (EMF/WMF) are silently skipped. Ignored in plain text mode
and for formats without OOXML image support (`.doc`, `.xls`, `.pdf`).

## Known limitations

- `--images` supports `.docx`/`.pptx`/`.xlsx` only. Legacy `.doc`/`.xls`
  images are in MSODRAW binary format and not extracted. No PDF images.
- `.doc` heading/table detection is heuristic. It's good, not perfect.
- Only BIFF8 (Excel 97+). Older BIFF5 `.xls` files won't parse.
- No legacy `.ppt` support — only modern `.pptx`.
- `.pptx` heading detection is font-size based (>=28pt = h1, >=24pt = h2,
  >=20pt = h3). Works well on typical slide decks.
- OCR (`--ocr`) recognizes Latin scripts only (the ocrs engine is English/
  European-language focused). CJK scans still won't produce output.
- OCR models (~12 MB) download on first use — see the OCR section below.
- Some CJK encodings in PDFs may not extract correctly.

## OCR

OCR runs the [ocrs](https://github.com/robertknight/ocrs) engine (pure
Rust, rten backend) over:

- **Image files** — `batdoc photo.png` always OCRs (no flag needed; text is
  the only output an image can produce). PNG, JPEG, GIF, WebP, and BMP.
- **DOCX/PPTX embedded images** — `batdoc --ocr report.docx` renders OCR'd
  text as a blockquote after each image (paragraphs in `--plain` mode).
- **PDF pages without a text layer** — automatic: a scanned PDF whose pages
  have no text is OCR'd from its embedded page images (typical one-image-per-
  page scans). Text-bearing pages and text-bearing documents are never OCR'd.

Models (~12 MB, two `.rten` files) are downloaded on first OCR use from the
ocrs upstream distribution and cached in `$BATDOC_MODELS_DIR`, else
`$XDG_CACHE_HOME/batdoc/models`, else `~/.cache/batdoc/models`. Set
`BATDOC_MODELS_DIR` to a pre-seeded directory for offline or
package-managed installs. OCR takes roughly 0.5–2 s per image on CPU.

Note: ocrs runs much slower in debug builds; always use a release build
(`cargo build --release`) when testing OCR.

## Dependencies

The CLI binary depends on `batdoc-core` (document extraction library),
`bat` (syntax highlighting), and `is-terminal` (tty detection).

The `batdoc-core` library depends on `cfb`, `encoding_rs`, `quick-xml`,
`zip`, `pdf-extract`, `lopdf`, `base64`, and `thiserror`, plus `ocrs`,
`image`, and `ureq` for OCR and `rten` for model inference. No C, no
system libs.

## Library

The extraction engine is available as a standalone Rust library for
programmatic use — no CLI or terminal dependencies:

```toml
[dependencies]
batdoc-core = { git = "https://github.com/daemonp/batdoc" }
```

```rust
let data = std::fs::read("report.docx")?;
let markdown = batdoc_core::to_markdown(&data, false)?;

// OCR an image file:
let img = std::fs::read("photo.png")?;
let text = batdoc_core::extract_plain_with(
    &img,
    batdoc_core::Format::Image,
    batdoc_core::ExtractOptions::default(),
)?;
```

See [batdoc-core/README.md](batdoc-core/README.md) for the full API.

## History

The original [catdoc](http://www.intevation.de/catdoc/) by Vitaliy Strochkov
has been converting `.doc` files to text on Unix since the 90s. The `.doc`
parser here borrows its 256-byte block Unicode/8-bit detection heuristic
from that project. `batdoc` extends the idea to all five Office formats plus PDF and
outputs markdown instead of plain text — same spirit, modern tooling.

## License

MIT
