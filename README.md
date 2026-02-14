# batdoc

`cat` had [catdoc](http://www.intevation.de/catdoc/). `bat` gets `batdoc`.

Dumps `.doc`, `.docx`, `.xls`, `.xlsx`, `.pptx`, and `.pdf` files to your
terminal as markdown. To a tty it syntax-highlights and pages (using
[bat](https://github.com/sharkdp/bat)); piped, it gives you plain text.

```
batdoc report.docx                     # highlighted markdown in terminal
batdoc financials.xlsx                  # each sheet becomes a markdown table
batdoc slides.pptx                     # per-slide headings with text
batdoc paper.pdf                       # multi-page PDF with page headers
batdoc --plain legacy.doc > out.txt    # just the text
cat mystery.bin | batdoc               # stdin works, format detected by magic bytes
```

Format is detected by file signature, not extension. OLE2 files (`.doc`/`.xls`)
are distinguished by peeking at internal streams; ZIP files (`.docx`/`.xlsx`/`.pptx`)
by checking for `word/document.xml` vs `xl/workbook.xml` vs `ppt/presentation.xml`;
PDFs by the `%PDF-` header.

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
markdown links. Multi-slide decks get `## Slide N` headings.

`.pdf` extracts text from text-based PDFs using `pdf-extract`. Multi-page
documents get `## Page N` headings in markdown mode. Scanned/image-only
PDFs that contain no extractable text get a clean error message. Malformed
PDFs that would crash the underlying library are caught and reported as
errors rather than panics.

## Options

```
batdoc [OPTIONS] [FILE...]
cat FILE | batdoc [OPTIONS]

  -p, --plain       plain text, no highlighting
  -m, --markdown    force markdown (default on tty)
  -i, --images      embed images as inline base64 data URIs
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
- PDFs must contain actual text — scanned/image-only PDFs won't produce
  output (no OCR). Some CJK encodings in PDFs may not extract correctly.

## Dependencies

Eight crates, no C, no system libs: `base64`, `bat`, `cfb`, `encoding_rs`,
`pdf-extract`, `quick-xml`, `zip`, `is-terminal`.

## History

The original [catdoc](http://www.intevation.de/catdoc/) by Vitaliy Strochkov
has been converting `.doc` files to text on Unix since the 90s. The `.doc`
parser here borrows its 256-byte block Unicode/8-bit detection heuristic
from that project. `batdoc` extends the idea to all five Office formats plus PDF and
outputs markdown instead of plain text — same spirit, modern tooling.

## License

MIT
