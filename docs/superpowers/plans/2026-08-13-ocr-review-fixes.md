# OCR Review Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use subagent-driven-development (recommended) or executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Apply every accepted finding from the `f63bee0..HEAD` OCR code review (M1–M4, D1–D2, breaking-change decisions, and nits).

**Architecture:** Five tasks partitioned by file so three of them can run as parallel worktree-isolated workers without merge conflicts. Tasks 1–3 (pdf.rs, ocr.rs, docx/pptx/markup) touch disjoint files and run in parallel; the parent merges their branches. Task 4 (lib.rs/main.rs/build.rs/README/CI) depends on Task 2's `models_present()`. Task 5 (`ExtractOptions` threading) touches nearly everything, so it runs last, sequentially. A final parallel review + fix-worker pass closes the loop.

**Tech Stack:** Rust 2021, `lopdf`, `image` 0.25, `ocrs`/`rten`, `ureq` 2, pi-subagents (`worker`, `reviewer`).

**Approved decisions (from user):**

- Scope: everything from the review.
- Breaking changes approved: `#[non_exhaustive]` on `Format`; drop the `-o` short flag (long-only `--ocr`).
- M3: download **timeout only** (no checksum pinning).
- D2: remove `eprintln!` from the library; the CLI prints the download notice via a new `batdoc_core::models_present()` check.

---

## Orchestration Overview (for the parent orchestrator)

1. **Phase 0 — clean tree.** `worktree: true` requires a clean git state. The modified `docs/superpowers/specs/2026-08-13-extraction-fidelity-design.md` is unrelated to this work — commit it (or stash it) before dispatching Phase 1. Untracked files (`.vexp/`, `Workspace-refactor-plan.md`) can stay.
2. **Phase 1 — parallel workers** (`worktree: true`, `context: "fresh"`, `async: true`): Tasks 1, 2, 3. Files are disjoint (pdf.rs / ocr.rs / markup.rs+docx.rs+pptx.rs).
3. **Phase 2 — parent merges.** Merge the three worktree branches into `master` in order 1 → 2 → 3 (disjoint files → clean merges). Gate: `cargo fmt --check && cargo clippy --workspace --all-targets && cargo test --workspace` all green before Phase 3.
4. **Phase 3 — sequential workers** in the main tree: Task 4, then Task 5.
5. **Phase 4 — review loop.** Three parallel fresh-context `reviewer` agents (correctness/regressions; tests/validation; simplicity/DRY), `output: false`. Parent synthesizes, one `worker` applies fixes worth doing now, then final validation (see "Final validation" at the end).

Worker prompts must include: the exact task text from this plan, the constraint "only modify the files listed in your task", and the required handoff (changed files, commands run with exit codes, test output, surprises).

---

### Task 1: PDF — hoist document parse, bound OCR memory

**Findings addressed:** M1 (`ocr_page` re-parses the whole PDF per textless page), M2 (4-image cap bounds count, not bytes), plus the pdf-extract↔lopdf page-order invariant comment.

**Files:**

- Modify: `batdoc-core/src/pdf.rs`

- [ ] **Step 1: Write the failing test**

Add to the `mod tests` block in `batdoc-core/src/pdf.rs`:

```rust
    #[test]
    fn ocr_candidates_caps_and_ranks() {
        let dict = lopdf::Dictionary::new();
        let mk = |width: i64, height: i64| PdfImage {
            id: (1, 0),
            width,
            height,
            color_space: Some("DeviceRGB".to_string()),
            filters: Some(Vec::new()),
            bits_per_component: Some(8),
            content: &[],
            origin_dict: &dict,
        };
        let images = vec![
            mk(100, 100),        // 10k px
            mk(20_000, 20_000),  // 400 MP — over budget, excluded
            mk(1000, 1000),      // 1 MP — largest valid
            mk(500, 500),
            mk(200, 200),
            mk(50, 50),
            mk(0, 100),          // zero area — excluded
        ];
        let picked = ocr_candidates(&images);
        // 5 valid images → capped at 4, largest first; oversized/zero excluded.
        assert_eq!(picked.len(), 4);
        assert_eq!((picked[0].width, picked[0].height), (1000, 1000));
        assert_eq!((picked[1].width, picked[1].height), (500, 500));
        assert_eq!((picked[2].width, picked[2].height), (200, 200));
        assert_eq!((picked[3].width, picked[3].height), (100, 100));
    }
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cargo test -p batdoc-core pdf::tests::ocr_candidates 2>&1 | tail -5`
Expected: FAIL to compile — `cannot find function ocr_candidates in this scope`.

- [ ] **Step 3: Implement constants, `ocr_candidates`, bounded JPEG decode, and the hoisted parse**

Add `use std::io::Cursor;` to the imports at the top of `pdf.rs`. Add these items after the `clean_page` function:

```rust
/// Maximum number of embedded images OCR'd per page (largest first).
const MAX_OCR_IMAGES_PER_PAGE: usize = 4;
/// Maximum decoded pixels per embedded image (~100 MP ≈ 300 MB RGB).
/// Larger images are skipped so a single image cannot exhaust memory.
const MAX_OCR_IMAGE_PIXELS: u64 = 100_000_000;
/// Maximum width/height accepted when decoding embedded JPEGs. Strict
/// decoder-side guard against a JPEG whose real dimensions exceed the
/// dimensions declared in the PDF dictionary (10_000² = the pixel budget).
const MAX_OCR_IMAGE_DIM: u32 = 10_000;

/// Select up to [`MAX_OCR_IMAGES_PER_PAGE`] OCR candidates, largest by
/// declared pixel area first, skipping images whose decoded size would
/// exceed [`MAX_OCR_IMAGE_PIXELS`]. The cap must bound peak memory, which
/// the declared dimensions determine (compressed images expand when decoded).
fn ocr_candidates<'a>(images: &'a [PdfImage<'a>]) -> Vec<&'a PdfImage<'a>> {
    let mut candidates: Vec<(&PdfImage, u64)> = images
        .iter()
        .map(|img| {
            (
                img,
                u64::try_from(img.width.max(0)).unwrap_or(0)
                    * u64::try_from(img.height.max(0)).unwrap_or(0),
            )
        })
        .filter(|(_, area)| *area > 0 && *area <= MAX_OCR_IMAGE_PIXELS)
        .collect();
    candidates.sort_by_key(|(_, area)| std::cmp::Reverse(*area));
    candidates
        .into_iter()
        .take(MAX_OCR_IMAGES_PER_PAGE)
        .map(|(img, _)| img)
        .collect()
}

/// Decode JPEG bytes with strict dimension limits, so a stream whose real
/// dimensions exceed the PDF dictionary's claim cannot allocate beyond the
/// OCR pixel budget.
fn decode_jpeg_bounded(bytes: &[u8]) -> Option<image::RgbImage> {
    let mut reader = image::ImageReader::new(Cursor::new(bytes));
    reader.set_format(image::ImageFormat::Jpeg);
    let mut limits = image::Limits::default();
    limits.max_image_width = Some(MAX_OCR_IMAGE_DIM);
    limits.max_image_height = Some(MAX_OCR_IMAGE_DIM);
    reader.limits(limits);
    reader.decode().ok().map(image::DynamicImage::into_rgb8)
}
```

Replace the whole `extract_pages_with_ocr` function with:

```rust
/// Produce the per-page text, OCR'ing empty pages when `ocr` is set.
fn extract_pages_with_ocr(data: &[u8], ocr: bool) -> Result<Vec<String>> {
    let pages = extract_pages(data)?;
    if !ocr {
        return Ok(pages.iter().map(|p| clean_page(p)).collect());
    }
    // Parse the document with lopdf ONCE for image extraction; if that
    // fails, empty pages fall through as "no text" (same as non-OCR).
    let doc_pages = lopdf::Document::load_mem(data)
        .ok()
        .map(|doc| (doc.get_pages().into_values().collect::<Vec<_>>(), doc));
    let mut out = Vec::with_capacity(pages.len());
    for (i, page) in pages.iter().enumerate() {
        let cleaned = clean_page(page);
        if !cleaned.is_empty() {
            out.push(cleaned);
            continue;
        }
        let ocr_text = match &doc_pages {
            Some((page_ids, doc)) => ocr_page(doc, page_ids, i)?,
            None => None,
        };
        out.push(ocr_text.map_or_else(String::new, |t| clean_page(&t)));
    }
    Ok(out)
}
```

Replace the whole `ocr_page` function with:

```rust
/// OCR the embedded images of one page (0-based index into `page_ids`).
/// `None` when the page has no OCR-able images or no text was detected.
///
/// Page-order invariant: `pdf_extract` iterates the page tree in document
/// order and lopdf's `get_pages()` numbers pages from the same tree in
/// ascending order, so `page_ids[i]` corresponds to `pages[i]` produced by
/// [`extract_pages`].
fn ocr_page(
    doc: &lopdf::Document,
    page_ids: &[lopdf::ObjectId],
    page_index: usize,
) -> Result<Option<String>> {
    let Some(page_id) = page_ids.get(page_index) else {
        return Ok(None);
    };
    let Ok(images) = doc.get_page_images(*page_id) else {
        return Ok(None);
    };
    let mut texts = Vec::new();
    for img in ocr_candidates(&images) {
        let Some(rgb) = decode_pdf_image(img) else {
            continue;
        };
        if let Some(text) = crate::ocr::ocr_rgb_image(&rgb)? {
            texts.push(text);
        }
    }
    if texts.is_empty() {
        Ok(None)
    } else {
        Ok(Some(texts.join("\n")))
    }
}
```

In `decode_pdf_image`, replace the DCT branch:

```rust
    if filters.contains(&"DCTDecode") {
        return decode_jpeg_bounded(img.content);
    }
```

(The old `image::load_from_memory(img.content).ok().map(image::DynamicImage::into_rgb8)` body is deleted.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `cargo test -p batdoc-core pdf:: 2>&1 | tail -5`
Expected: PASS, including `ocr_candidates_caps_and_ranks` and all pre-existing pdf tests (`decode_pdf_image_*`, `extract_plain_ocr_flag_errors_differ`, etc.).

Run: `cargo clippy -p batdoc-core --all-targets -- -W clippy::pedantic 2>&1 | tail -3`
Expected: zero warnings.

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/pdf.rs
git commit -m "perf: parse PDF once for OCR; cap embedded-image decode size

Hoist lopdf::Document::load_mem out of the per-page loop (was one full
document parse per textless page) and bound OCR memory: at most 4 images
per page, each at most 100 MP, with strict decoder dimension limits on the
JPEG path as defense against dimensions that exceed the PDF dictionary's
claim. Documents the pdf-extract/lopdf page-order invariant."
```

---

### Task 2: OCR engine — timeout, no failure caching, tmp sweep, `models_present`

**Findings addressed:** M3 (ureq download has no timeout), M4 (`LazyLock` caches transient failures process-wide), tmp-file leak nit, D2 library half (remove `eprintln!`, expose `models_present()` — CLI wiring is Task 4).

**Files:**

- Modify: `batdoc-core/src/ocr.rs`

- [ ] **Step 1: Write the failing tests**

Add to the `mod tests` block in `batdoc-core/src/ocr.rs`:

```rust
    #[test]
    fn stale_tmp_file_detection() {
        use std::time::SystemTime;
        // Fresh download in progress: recent mtime → keep.
        assert!(!is_stale_tmp_file(
            "text-detection.rten.tmp.123",
            Some(SystemTime::now())
        ));
        // Old leftover from a crashed run → sweep.
        assert!(is_stale_tmp_file(
            "text-detection.rten.tmp.123",
            Some(SystemTime::now() - 2 * STALE_TMP_AGE)
        ));
        // Not a tmp file → never sweep, however old.
        assert!(!is_stale_tmp_file(
            "text-detection.rten",
            Some(SystemTime::now() - 2 * STALE_TMP_AGE)
        ));
        // Unknown mtime → leave it alone.
        assert!(!is_stale_tmp_file("x.tmp.1", None));
    }

    #[test]
    fn models_present_in_requires_both_files() {
        let tmp = std::env::temp_dir().join(format!("batdoc-ocr-t2-{}", std::process::id()));
        std::fs::create_dir_all(&tmp).unwrap();
        assert!(!models_present_in(&tmp));
        std::fs::write(tmp.join(DETECTION_MODEL_FILE), b"x").unwrap();
        assert!(!models_present_in(&tmp));
        std::fs::write(tmp.join(RECOGNITION_MODEL_FILE), b"x").unwrap();
        assert!(models_present_in(&tmp));
        std::fs::remove_dir_all(&tmp).ok();
    }
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cargo test -p batdoc-core ocr::tests 2>&1 | tail -5`
Expected: FAIL to compile — `cannot find function is_stale_tmp_file` / `cannot find function models_present_in`.

- [ ] **Step 3: Implement**

Change the imports: replace `use std::sync::LazyLock;` with:

```rust
use std::sync::OnceLock;
use std::time::Duration;
```

Add constants after the model-file constants:

```rust
/// Network timeout for model downloads.
const DOWNLOAD_TIMEOUT: Duration = Duration::from_secs(120);
/// Age after which leftover `*.tmp.<pid>` download files are swept.
const STALE_TMP_AGE: Duration = Duration::from_secs(3600);
```

In `ensure_file`, delete the `eprintln!(...)` block (the CLI now owns the download notice — Task 4) and add the timeout to the request:

```rust
    let response = ureq::get(url)
        .timeout(DOWNLOAD_TIMEOUT)
        .call()
        .map_err(|e| {
```

(keep the rest of the `map_err` body unchanged).

In `ensure_models`, add the sweep as the first statement:

```rust
fn ensure_models() -> Result<ModelPaths> {
    let dir = cache_dir();
    sweep_stale_tmp_files(&dir);
    ...
```

Add the sweep helpers after `ensure_file`:

```rust
/// Remove stale `*.tmp.<pid>` download leftovers in `dir` (from interrupted
/// runs). Only files older than [`STALE_TMP_AGE`] are touched, so an active
/// concurrent download is never disturbed. Best-effort: errors ignored.
fn sweep_stale_tmp_files(dir: &Path) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        let Some(name) = path.file_name().and_then(|n| n.to_str()) else {
            continue;
        };
        let modified = entry.metadata().ok().and_then(|m| m.modified().ok());
        if is_stale_tmp_file(name, modified) {
            std::fs::remove_file(&path).ok();
        }
    }
}

/// `true` for `*.tmp.<pid>` download leftovers older than [`STALE_TMP_AGE`].
fn is_stale_tmp_file(name: &str, modified: Option<std::time::SystemTime>) -> bool {
    name.contains(".tmp.")
        && modified
            .and_then(|m| m.elapsed().ok())
            .is_some_and(|age| age > STALE_TMP_AGE)
}
```

Delete the `type EngineResult = ...` alias, change `build_engine`'s return type to `std::result::Result<OcrEngine, String>`, and replace `engine()`:

```rust
/// Process-wide OCR engine, built once on first use.
///
/// Failures are NOT cached: a transient error (e.g. a network failure
/// during the first model download) is returned to the caller, and a later
/// call retries. In a concurrent first-build race the loser discards its
/// engine and shares the stored one.
fn engine() -> Result<&'static OcrEngine> {
    static ENGINE: OnceLock<OcrEngine> = OnceLock::new();
    if let Some(engine) = ENGINE.get() {
        return Ok(engine);
    }
    let built = build_engine().map_err(BatdocError::Document)?;
    Ok(ENGINE.get_or_init(|| built))
}
```

Add after `engine()`:

```rust
/// `true` when both OCR model files already exist in the cache directory
/// (i.e. OCR will not trigger a download). Used by the CLI to print a
/// first-use download notice.
#[must_use]
pub fn models_present() -> bool {
    models_present_in(&cache_dir())
}

fn models_present_in(dir: &Path) -> bool {
    dir.join(DETECTION_MODEL_FILE).exists() && dir.join(RECOGNITION_MODEL_FILE).exists()
}
```

Note: `pub fn models_present` in a private module is not yet reachable from outside the crate — Task 4 adds the re-export in `lib.rs`. That is expected; do not edit `lib.rs` in this task.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cargo test -p batdoc-core ocr:: 2>&1 | tail -5`
Expected: PASS — all pre-existing ocr tests plus the two new ones.

Run: `cargo clippy -p batdoc-core --all-targets -- -W clippy::pedantic 2>&1 | tail -3`
Expected: zero warnings.

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/ocr.rs
git commit -m "fix: harden OCR model download and engine caching

- 120s timeout on model downloads (ureq default is no timeout)
- Transient engine-build failures are no longer cached process-wide;
  only a successfully built engine is stored (OnceLock)
- Sweep stale *.tmp.<pid> download leftovers older than 1h
- Move the download notice out of the library (CLI wires it via the new
  models_present()); batdoc-core no longer writes to stderr"
```

---

### Task 3: DRY — shared OCR rendering helpers in markup.rs

**Findings addressed:** D1 (blockquote/plain OCR rendering copy-pasted between docx.rs and pptx.rs, already drifting in separator handling).

**Files:**

- Modify: `batdoc-core/src/markup.rs`
- Modify: `batdoc-core/src/docx.rs`
- Modify: `batdoc-core/src/pptx.rs`

- [ ] **Step 1: Write the failing tests**

Add to the `mod tests` block in `batdoc-core/src/markup.rs` (it starts at line ~173):

```rust
    #[test]
    fn ocr_blockquote_skips_blank_lines() {
        let mut out = String::new();
        push_ocr_blockquote(&mut out, "one\n\n  \ntwo  \n");
        assert_eq!(out, "> one\n> two\n\n");
    }

    #[test]
    fn ocr_plain_paragraphs_respect_first_flag() {
        let mut out = String::new();
        let mut first = true;
        push_ocr_plain(&mut out, "a\n\nb", &mut first);
        assert_eq!(out, "a\n\nb\n");
        push_ocr_plain(&mut out, "c", &mut first);
        assert_eq!(out, "a\n\nb\n\nc\n");
    }
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cargo test -p batdoc-core markup::tests 2>&1 | tail -5`
Expected: FAIL to compile — `cannot find function push_ocr_blockquote`.

- [ ] **Step 3: Implement the helpers**

Add to `batdoc-core/src/markup.rs`, after `image_to_base64_ref` (before the tests module):

```rust
/// Append OCR'd image text as a markdown blockquote: one `> ` line per
/// non-blank line of `text` (trailing whitespace trimmed), then a blank
/// line. Shared by the DOCX and PPTX renderers.
pub(crate) fn push_ocr_blockquote(out: &mut String, text: &str) {
    for line in text.lines() {
        if line.trim().is_empty() {
            continue;
        }
        out.push_str("> ");
        out.push_str(line.trim_end());
        out.push('\n');
    }
    out.push('\n');
}

/// Append OCR'd image text as plain-text paragraphs: each non-blank line
/// becomes a paragraph, blank-line separated. `first` tracks whether
/// anything has been written to `out` yet, so no leading blank line is
/// emitted for the very first paragraph. Shared by the DOCX and PPTX
/// renderers.
pub(crate) fn push_ocr_plain(out: &mut String, text: &str, first: &mut bool) {
    for para in text.split('\n') {
        let para = para.trim();
        if para.is_empty() {
            continue;
        }
        if !*first {
            out.push('\n');
        }
        out.push_str(para);
        out.push('\n');
        *first = false;
    }
}
```

- [ ] **Step 4: Run helper tests**

Run: `cargo test -p batdoc-core markup:: 2>&1 | tail -5`
Expected: PASS.

- [ ] **Step 5: Swap the docx call sites**

In `batdoc-core/src/docx.rs`, `render_block_plain`, replace the `Block::Image { ocr_text: Some(text), .. }` arm body with:

```rust
        Block::Image {
            ocr_text: Some(text),
            ..
        } => {
            crate::markup::push_ocr_plain(out, text, first);
        }
```

In `render_block_markdown`, replace the `Block::Image { markdown, ocr_text, .. }` arm with:

```rust
        Block::Image {
            markdown, ocr_text, ..
        } => {
            if let Some(md) = markdown {
                out.push_str(md);
                out.push_str("\n\n");
            }
            if let Some(text) = ocr_text {
                crate::markup::push_ocr_blockquote(out, text);
            }
        }
```

- [ ] **Step 6: Swap the pptx call sites**

In `batdoc-core/src/pptx.rs`, `render_plain`, replace the OCR block (`let mut first_ocr = true;` … loop) with:

```rust
        let mut first_ocr = true;
        for ocr in &slide.image_ocr {
            crate::markup::push_ocr_plain(&mut out, ocr, &mut first_ocr);
        }
```

In `render_markdown`, replace the `for ocr in &slide.image_ocr { ... }` block with:

```rust
        for ocr in &slide.image_ocr {
            crate::markup::push_ocr_blockquote(&mut out, ocr);
        }
```

- [ ] **Step 7: Run the full crate test suite**

Run: `cargo test -p batdoc-core 2>&1 | tail -5`
Expected: PASS — 333+ tests. The pre-existing exact-output tests (`render_block_markdown_image_with_ocr_quote`, `render_block_plain_image_with_ocr`, `render_markdown_slide_with_ocr_text`, `render_plain_slide_with_ocr_text`) prove the swap changed nothing observable.

Run: `cargo clippy -p batdoc-core --all-targets -- -W clippy::pedantic 2>&1 | tail -3`
Expected: zero warnings.

- [ ] **Step 8: Commit**

```bash
git add batdoc-core/src/markup.rs batdoc-core/src/docx.rs batdoc-core/src/pptx.rs
git commit -m "refactor: share OCR rendering between docx and pptx

Extract push_ocr_blockquote / push_ocr_plain into markup.rs; the two
call sites had already drifted in separator handling."
```

---

### Task 4: Public API + CLI — non_exhaustive Format, magic bytes, flag rename, notices

**Findings addressed:** A1 (`#[non_exhaustive]` on `Format` — approved breaking change), GIF magic nit, BMP weakness comment, `-o` rename (approved: long-only `--ocr`), TTY plain rendering for `Format::Image`, D2 CLI half (download notice via `models_present()`), CI model caching.

**Depends on:** Task 2 (`ocr::models_present` must exist and be merged).

**Files:**

- Modify: `batdoc-core/src/lib.rs`
- Modify: `src/main.rs`
- Modify: `build.rs`
- Modify: `README.md`
- Modify: `.github/workflows/ci.yml`

- [ ] **Step 1: Write the failing test (GIF strictness)**

Add to the `mod tests` block in `batdoc-core/src/lib.rs`:

```rust
    #[test]
    fn detect_format_requires_full_gif_magic() {
        assert_eq!(detect_format(b"GIF87a....").unwrap(), Format::Image);
        assert_eq!(detect_format(b"GIF89a....").unwrap(), Format::Image);
        assert!(detect_format(b"GIFzzz....").is_err());
    }
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cargo test -p batdoc-core detect_format_requires_full_gif_magic 2>&1 | tail -5`
Expected: FAIL — `GIFzzz....` currently detected as Image (assertion fails on line 3).

- [ ] **Step 3: lib.rs changes**

On the `Format` enum, add the attribute and update its doc comment:

```rust
/// Document format detected from magic bytes.
///
/// `#[non_exhaustive]`: new variants may be added in minor releases;
/// downstream matches must include a wildcard arm.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Format {
```

(Keep the existing derive line as-is — only add `#[non_exhaustive]` and the doc lines. If the derive list differs, preserve it.)

Replace the GIF check in `detect_format`:

```rust
    if data.len() >= 6 && (&data[..6] == b"GIF87a" || &data[..6] == b"GIF89a") {
        return Ok(Format::Image); // GIF
    }
```

Replace the BMP check with:

```rust
    // BMP's 2-byte "BM" signature is weak: any file starting with those
    // bytes is routed to OCR and fails with "no text found in image"
    // rather than "unrecognized format". Accepted trade-off — real-world
    // collisions are rare.
    if data.len() >= 2 && &data[..2] == b"BM" {
        return Ok(Format::Image); // BMP
    }
```

Add the re-export next to `pub use error::{BatdocError, Result};` (line ~25):

```rust
pub use ocr::models_present;
```

- [ ] **Step 4: Run lib tests**

Run: `cargo test -p batdoc-core 2>&1 | tail -5`
Expected: PASS.

- [ ] **Step 5: main.rs changes**

In `src/main.rs`:

1. USAGE string: replace the line `  -o, --ocr         OCR text from images (embedded doc images, textless PDF pages)` with:

```
      --ocr         OCR text from images (embedded doc images, textless PDF pages)
```

2. Argument parsing: replace `"-o" | "--ocr" => ocr = true,` with `"--ocr" => ocr = true,`.

3. In `run()`, after `let opts = ExtractOptions { images, ocr };`, add the download notice:

```rust
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
```

4. TTY rendering: OCR text of an image file is plain text, not markdown — don't pretty-print it as Markdown. In `Mode::Markdown`, change the condition to `if is_tty && format != Format::Image {`. In `Mode::Auto`, change `if is_tty {` to `if is_tty && format != Format::Image {`. (Both branches already exist; only the conditions change.)

- [ ] **Step 6: build.rs (man page)**

In `build.rs`, delete the `.short("-o")` line from the `--ocr` flag definition (keep `.long("--ocr")` and the `.help(...)`).

- [ ] **Step 7: README.md**

Line ~94: change the options-table row `  -o, --ocr         OCR text from images: ...` to `      --ocr         OCR text from images: ...`. Then `grep -n '\-o, --ocr\|batdoc -o' README.md docs/ -r` — fix any other occurrence of the short flag.

- [ ] **Step 8: CI model caching**

In `.github/workflows/ci.yml`, in the `ocr-e2e` job, insert before the "Download OCR models" step:

```yaml
- name: Cache OCR models
  id: model-cache
  uses: actions/cache@v4
  with:
    path: ~/.cache/batdoc/models
    key: ocr-models-v1
```

and gate the download step: add `if: steps.model-cache.outputs.cache-hit != 'true'` to "Download OCR models".

- [ ] **Step 9: Verify**

Run: `cargo fmt --check && cargo clippy --workspace --all-targets -- -W clippy::pedantic -D warnings && cargo test --workspace 2>&1 | grep "test result"`
Expected: fmt clean, zero clippy warnings, all tests pass.

Run: `cargo run --quiet -- --help | grep -A1 -- "--ocr"`
Expected: `--ocr` listed with no `-o`.

Run: `cargo run --quiet -- -o 2>&1 | head -2`
Expected: `batdoc: unknown option: -o` (exit 1).

- [ ] **Step 10: Commit (three commits)**

```bash
git add batdoc-core/src/lib.rs
git commit -m "feat!: mark Format non_exhaustive; strict GIF magic

BREAKING CHANGE: Format is now #[non_exhaustive]; downstream matches need
a wildcard arm. (Format::Image in 1.4.0 was already an additive break;
this makes future variants semver-minor.)"

git add src/main.rs build.rs README.md
git commit -m "feat!: drop -o short flag; render image OCR as plain text on TTY

BREAKING CHANGE: -o removed (reserved for a future --output); use --ocr.
Also: image-file OCR output is plain text, so the TTY path no longer
pipes it to bat as Markdown, and the first-use model-download notice is
printed by the CLI (batdoc-core no longer writes to stderr)."

git add .github/workflows/ci.yml
git commit -m "ci: cache OCR models between runs"
```

---

### Task 5: Thread ExtractOptions through internal format APIs

**Findings addressed:** boolean-parameter threading nit (`extract_markdown(data, images, ocr)`-style signatures accumulating positional bools).

**Depends on:** Tasks 1 and 3 merged (touches the same functions).

**Files:**

- Modify: `batdoc-core/src/lib.rs` (call sites in `extract_plain_with` / `extract_markdown_with`)
- Modify: `batdoc-core/src/docx.rs`
- Modify: `batdoc-core/src/pptx.rs`
- Modify: `batdoc-core/src/pdf.rs`

- [ ] **Step 1: docx.rs**

Add `use crate::ExtractOptions;` to the imports. Change signatures:

- `pub(crate) fn extract_plain(data: &[u8], ocr: bool)` → `pub(crate) fn extract_plain(data: &[u8], opts: ExtractOptions)`; body calls `parse_docx(data, false, opts.ocr)` → after the parse change below, `parse_docx(data, ExtractOptions { images: false, ..opts })`.
- `pub(crate) fn extract_markdown(data: &[u8], images: bool, ocr: bool)` → `pub(crate) fn extract_markdown(data: &[u8], opts: ExtractOptions)`; body calls `parse_docx(data, opts)`.
- `fn parse_docx(data: &[u8], images: bool, ocr: bool)` → `fn parse_docx(data: &[u8], opts: ExtractOptions)`; inside, replace `images` with `opts.images` and `ocr` with `opts.ocr`, and call `resolve_images(&mut blocks, &mut archive, opts)`.
- `fn resolve_images(blocks: &mut Vec<Block>, archive: &mut ZipArchive<Cursor<&[u8]>>, images: bool, ocr: bool)` → same params minus the bools, plus `opts: ExtractOptions`; inside, `if images` → `if opts.images`, `if ocr` → `if opts.ocr`.

Update test call sites mechanically: `parse_docx(&data, false, false)` → `parse_docx(&data, ExtractOptions::default())`; `extract_markdown(&data, false, false)` → `extract_markdown(&data, ExtractOptions::default())`; `extract_markdown(&data, true, false)` → `extract_markdown(&data, ExtractOptions { images: true, ..Default::default() })`; `extract_plain(&data, false)` → `extract_plain(&data, ExtractOptions::default())`.

- [ ] **Step 2: pptx.rs**

Add `use crate::ExtractOptions;`. Change:

- `pub(crate) fn extract_plain(data: &[u8], ocr: bool)` → `(data: &[u8], opts: ExtractOptions)`; body: `parse_pptx(data, ExtractOptions { images: false, ..opts })`.
- `pub(crate) fn extract_markdown(data: &[u8], images: bool, ocr: bool)` → `(data: &[u8], opts: ExtractOptions)`; body: `parse_pptx(data, opts)`.
- `fn parse_pptx(data: &[u8], extract_images: bool, ocr: bool)` → `fn parse_pptx(data: &[u8], opts: ExtractOptions)`; inside, `extract_images` → `opts.images`, `ocr` → `opts.ocr`.

Update test call sites the same way as docx (`extract_markdown(&data, false, false)` → `extract_markdown(&data, ExtractOptions::default())`, etc.).

- [ ] **Step 3: pdf.rs**

Add `use crate::ExtractOptions;`. Change:

- `pub(crate) fn extract_plain(data: &[u8], ocr: bool)` → `(data: &[u8], opts: ExtractOptions)`; body: `extract_pages_with_ocr(data, opts.ocr)` and `no_text_error(opts.ocr)`.
- `pub(crate) fn extract_markdown(data: &[u8], ocr: bool)` → `(data: &[u8], opts: ExtractOptions)`; same replacements.

(Keep `extract_pages_with_ocr(data, ocr: bool)` and `no_text_error(ocr: bool)` as-is — they are single-flag helpers, not accumulations.) Update test call sites: `extract_plain(garbage, false)` → `extract_plain(garbage, ExtractOptions::default())`, `extract_plain(data, true)` → `extract_plain(data, ExtractOptions { ocr: true, ..Default::default() })`, etc.

- [ ] **Step 4: lib.rs call sites**

In `extract_plain_with`: `docx::extract_plain(data, opts)`, `pptx::extract_plain(data, opts)`, `pdf::extract_plain(data, opts)`. In `extract_markdown_with`: `docx::extract_markdown(data, opts)`, `pptx::extract_markdown(data, opts)`, `pdf::extract_markdown(data, opts)`. (`xlsx::extract_markdown(data, opts.images)` stays as-is — xlsx has no OCR support.)

- [ ] **Step 5: Verify**

Run: `cargo fmt --check && cargo clippy --workspace --all-targets -- -W clippy::pedantic -D warnings && cargo test --workspace 2>&1 | grep "test result"`
Expected: all clean, all pass. This task is a pure refactor — no test bodies change except the call-site signatures.

- [ ] **Step 6: Commit**

```bash
git add batdoc-core/src/
git commit -m "refactor: pass ExtractOptions through internal format APIs

Replaces positional bool parameters (extract_markdown(data, images, ocr))
with the Copy options struct, so the next option doesn't widen every
signature. Public API unchanged."
```

---

## Final validation (parent, after the Phase 4 review loop)

- [ ] `git status --short` — only expected files changed
- [ ] `cargo fmt --check` — clean
- [ ] `cargo clippy --workspace --all-targets -- -W clippy::pedantic -D warnings` — zero warnings
- [ ] `cargo test --workspace` — all pass (333+ plus the new ones from Tasks 1–4)
- [ ] `cargo build --release` — succeeds
- [ ] OCR e2e (only if `~/.cache/batdoc/models` already exists locally — do NOT download ~12 MB of models without asking): `cargo test --release -p batdoc-core --test ocr_e2e -- --ignored`
- [ ] Smoke: `cargo run --release -- --help` shows `--ocr` without `-o`

## Self-review notes

- Spec coverage: M1/M2 → Task 1; M3/M4 + tmp sweep + D2-library → Task 2; D1 → Task 3; A1 + GIF/BMP + `-o` + TTY + D2-CLI + CI cache → Task 4; ExtractOptions threading → Task 5. All accepted findings covered.
- Type consistency: `models_present` (Task 2) ↔ `batdoc_core::models_present()` re-export + CLI call (Task 4); `ocr_candidates`/`ocr_page(doc, page_ids, i)` signatures match between definition and tests; `ExtractOptions` field names (`images`, `ocr`) match the existing struct in `lib.rs`.
- Merge order matters: Phase-1 branches touch disjoint files; Task 4 edits `lib.rs` and needs Task 2 merged; Task 5 edits everything and must run last.
