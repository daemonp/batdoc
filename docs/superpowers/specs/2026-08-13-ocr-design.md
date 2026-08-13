# OCR — Image Input, Embedded Images, Scanned-PDF Fallback

Date: 2026-08-13
Status: pending review
Scope: `batdoc-core` (new `ocr` module, image format detection, PDF fallback, DOCX/PPTX image OCR) + CLI (`--ocr` flag, help, man page, README)

## Problem

Three text-bearing inputs currently produce no output:

1. **Image files.** `batdoc photo.png` fails format detection — PNG/JPEG/GIF/WebP/BMP are not recognized top-level formats.
2. **Scanned/image-only PDFs.** `pdf-extract` reads only the text layer; image-only pages hit the clean error "PDF contains no extractable text (may be scanned/image-only)". The README lists this as a known limitation ("no OCR").
3. **Text inside embedded document images.** DOCX/PPTX images are extractable (`--images`) but their textual content is invisible to the dump.

## Goals

- `batdoc photo.png|jpg|gif|webp|bmp` extracts text via OCR, no flag required.
- `--ocr` flag enables OCR of DOCX/PPTX embedded images and PDF fallback for pages with an empty text layer.
- All new dependencies are pure Rust — preserves the "no C, no system libs" story that static-musl and Docker packaging rely on.
- WITHOUT `--ocr`, every existing input produces byte-identical output to today (share this acceptance criterion with the extraction-fidelity slice).
- Models are downloaded at first OCR use into a cache directory, with an environment-variable override. No model files ship in the binary or packages.
- Public API stays backwards compatible: existing functions unchanged; OCR is additive surface.

## Non-goals

- Non-Latin scripts. ocrs recognizes Latin only (English, European languages). CJK PDFs remain unsupported — existing README caveat stands.
- Full PDF page rasterization (MuPDF/PDFium). Native C/C++ deps; rejected on dependency grounds. PDF OCR sources embedded image XObjects only — covers the typical one-image-per-page scan.
- OCR for `.xls` / `.xlsx` / `.doc` — no embedded-image extraction exists there (`--images` does not cover them; unchanged).
- `--images` for PDF. Stays as-is: DOCX/XLSX/PPTX only.
- Confidence filtering, layout reconstruction beyond line order, output markers labeling text as OCR'd.
- Bundling models into the binary (deferred; `include_bytes!` behind a cargo feature if ever requested).
- Clipboard input, CLI JSON/layout output (ocrs-cli extras; not batdoc's concern).

## Approach

In-process **ocrs 0.12** (rten backend — pure-Rust ONNX inference, MIT OR Apache-2.0). Rejected alternatives: shelling out to `ocrs-cli`/`tesseract` (breaks the single-binary identity; tesseract adds system deps) and PDF page rasterization via MuPDF/PDFium (native C++).

### New dependencies

| crate | version | role |
|---|---|---|
| `ocrs` | 0.12 | OCR engine, default features (rten format, no wasm export) |
| `image` | 0.25.10 | decode; features restricted to `png, jpeg, gif, webp, bmp` — matches the `--images` supported set |
| `pdf` | 0.10.0 | PDF image XObject extraction for the fallback path; `pdf-extract` stays for text |

All pure Rust; packaging targets (AUR, Homebrew, static musl, deb/rpm/alpine, Docker) unaffected — no packaging changes at all.

### Trigger semantics

| Input | Behavior |
|---|---|
| Image file (PNG/JPEG/GIF/WebP/BMP) | Always OCR'd. Text is the only output an image can produce; no flag needed. |
| PDF + `--ocr` | Pages with empty text-layer output fall back to OCR of that page's embedded images. |
| PDF without `--ocr` | Byte-identical to today, including the "no extractable text" error. |
| DOCX/PPTX + `--ocr` | OCR'd text from embedded images inserted at the image's position. |
| DOCX/PPTX without `--ocr` | Unchanged. `--images` remains purely about embedding image bytes; the flags compose but do not depend on each other. Image-position tracking is enabled by `--ocr` alone. |
| XLS/XLSX/DOC | No OCR (no embedded-image path). `--ocr` is a no-op for these formats. |

### Models

Two files, downloaded at first OCR use:

- `https://ocrs-models.s3-accelerate.amazonaws.com/text-detection.rten` — 2.4 MB
- `https://ocrs-models.s3-accelerate.amazonaws.com/text-recognition.rten` — 9.3 MB

Cache directory resolution (first match wins):

1. `$BATDOC_MODELS_DIR`
2. `$XDG_CACHE_HOME/batdoc/models`
3. `~/.cache/batdoc/models`

Download is atomic: write to `<target>.tmp` in the same directory, then rename — concurrent batdoc runs cannot corrupt the cache. Partial/failed downloads leave only the `.tmp` (removed on error). Download failure produces an actionable error naming the cache path and the override variable. Existing model files are used as-is (no re-download, no checksum pinning — the S3 bucket is the upstream publisher's own distribution point).

The OCR engine is constructed once per process (`OnceLock`), never per image. Model loading also honors the same resolution, so `BATDOC_MODELS_DIR` serves both packagers (pre-seed `/usr/share/batdoc/models`) and offline users.

### Module layout (`batdoc-core`)

New `src/ocr.rs`:

```rust
/// Resolve + download models on first use; returns paths.
fn ensure_models() -> Result<ModelPaths>;

/// Load the OCR engine once per process.
fn engine() -> Result<&'static OcrEngine>; // OnceLock

/// Decode and OCR an image byte slice. None = no text detected
/// (or undecodable image), Some = extracted text, lines in reading order.
pub(crate) fn ocr_image_bytes(data: &[u8]) -> Result<Option<String>>;
```

- `Format` gains `Image`; `detect_format` sniffs PNG (`89 50 4E 47`), JPEG (`FF D8 FF`), GIF (`GIF8`), WebP (`RIFF` + `WEBP`), BMP (`BM`) magic bytes. `Display` prints `IMAGE`.
- `lib.rs` gains:

```rust
#[derive(Default)]
pub struct ExtractOptions { pub images: bool, pub ocr: bool }

pub fn extract_plain_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String>;
pub fn extract_markdown_with(data: &[u8], format: Format, opts: ExtractOptions) -> Result<String>;
```

Existing `extract_plain` / `extract_markdown` / `to_plain` / `to_markdown` become default-option wrappers — non-breaking. Internally the `images`/`ocr` booleans thread through the DOCX/PPTX walkers exactly like `images` does today.

- Errors: OCR engine, download, and decode failures map to `BatdocError::Document(String)` with context. No new variant; `thiserror` surface unchanged.

### PDF fallback detail

Current flow: `extract_pages` → per-page strings → render. New `--ocr` flow per page:

1. If the page's text-layer string is non-empty → unchanged.
2. Else extract that page's image XObjects via the `pdf` crate (`page.get_images()`), decode (DCTDecode and FlateDecode handled by `pdf` + `image`), OCR each.
3. Cap: OCR the largest 4 images per page (typical scans are one image per page; the cap prevents pathological pages with dozens of small icons). Vector-only content (EMF/WMF) silently skipped — same policy as `--images`.
4. OCR text for the page is appended as that page's content.

Whole-document OCR-empty under `--ocr` → the existing error, wording amended to "(no text layer; OCR found nothing)". Text-bearing pages are never OCR'd — faster and more accurate than re-reading a quality text layer.

### Rendering

**Image input:** OCR text is the document body. Markdown mode: paragraphs of plain text. If OCR detects nothing: `BatdocError::Document("no text found in image")`. Multi-file CLI dumps and stdin work like any other format. No image is re-embedded for image input — the input *is* the image.

**DOCX/PPTX + `--ocr`:** OCR'd text inserts at the image's position:

- Markdown: blockquote after the inline image reference — `> ` prefixed lines, blank line between image and quote.
- Plain (`--plain`): a paragraph after the surrounding text; no image reference exists in plain mode (`--images` is ignored there), so OCR text reads as a standalone paragraph. OCR text appears in plain mode even without `--images` — it is text, and plain mode is the text dump.
- PPTX: per-slide, after the slide's text shapes, before/after existing image refs per current ordering (images render after text today; OCR quote follows the image refs).

### CLI + docs surface

- `-o, --ocr` flag; `--help` usage text updated.
- Man page updated (repo has one; see `38fee59`).
- README: remove the "no OCR" limitation line; add an OCR section covering: `--ocr` semantics, direct image input, Latin-only restriction, first-run model download + `BATDOC_MODELS_DIR`, speed expectation (~0.5–2 s/image on CPU), and the dev-build note that ocrs requires release-mode builds (`cargo run` debug builds are extremely slow).
- No changes to `Dockerfile`, `build.rs`, or `pkg/*` packaging scripts.

## Verification

- Unit (no models): `detect_format` magic-byte matches for all five image signatures (and rejection of near-misses); OCR cache path resolution (env override → XDG → home fallback); `ocr_image_bytes` decode-failure path.
- E2E (models required, `#[ignore]`-gated; CI job downloads models once into cache): commit a small fixture PNG containing rendered text; assert `ocr_image_bytes` returns the expected substring. DOCX fixture with an embedded text image → OCR'd text appears at the right position. PDF fixture (image-only, one image per page) → `--ocr` yields page text; same PDF without `--ocr` still errors as today.
- CLI smoke (after models present): `batdoc photo.jpg` and `batdoc --ocr scanned.pdf`, plus `--plain` composition.
- Regression: byte-identical output for a DOCX/PPTX with images, `--ocr` off, versus today's binary.

## Risks / Limitations

- **Speed:** ~0.5–2 s per image; a 100-page scanned PDF under `--ocr` is a multi-minute operation. Mitigated by opt-in flag and by never OCR'ing text-bearing pages.
- **Latin-only:** CJK scans still error. Documented.
- **First-run surprise:** model download happens on first OCR use — under a pipe with no flags, this never fires (OCR is never implicit for containers); for direct image input it may download mid-pipe with a stderr notice.
- **Scanned-PDF coverage:** OCR depends on the PDF embedding page images as XObjects. Scans that are only fully rasterizable (vector-annotated pages) fall through to the existing error; acceptable and documented.
- **Build cost:** `ocrs` + `rten` are large compile-time additions; release binary grows by roughly the linked engine (models are not in the binary).
