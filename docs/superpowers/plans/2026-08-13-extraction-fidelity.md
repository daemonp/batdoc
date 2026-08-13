# Extraction Fidelity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use subagent-driven-development (recommended) or executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extract PPTX speaker notes and DOCX comments/footnotes/endnotes into trailing sections, always on, with body markers for footnote/endnote refs.

**Architecture:** Extend private models in `pptx.rs` / `docx.rs`. Notes discovery uses full relationship parse (not hyperlink-only `load_rels`). Footnote/endnote markers are a distinct `Run` kind rendered per mode. Final assembly is `body → extras → image_defs`.

**Tech Stack:** Rust, `quick-xml`, `zip` (already in `batdoc-core`), unit tests with synthetic ZIP bytes via `zip::write::ZipWriter`.

**Spec:** `docs/superpowers/specs/2026-08-13-extraction-fidelity-design.md`

---

## File map

| File | Responsibility |
|------|----------------|
| `batdoc-core/src/xml_util.rs` | Parse all relationships; resolve relative ZIP targets |
| `batdoc-core/src/pptx.rs` | `Slide.notes`, notes discovery, body-placeholder filter, Notes trailer |
| `batdoc-core/src/docx.rs` | Marker runs, extra-part parse, Comments/Footnotes/Endnotes trailers, assembly |
| `README.md` | Formats blurb for notes/comments/footnotes/endnotes |

No public API, CLI, or new crates.

---

### Task 1: Internal relationship helpers in `xml_util`

**Files:**
- Modify: `batdoc-core/src/xml_util.rs`
- Test: same file `mod tests`

- [ ] **Step 1: Write failing tests for `parse_all_rels_xml` and `resolve_zip_target`**

Add to `xml_util.rs` tests:

```rust
#[test]
fn parse_all_rels_includes_notes_slide() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;
    let rels = parse_all_rels_xml(xml);
    assert_eq!(rels.get("rId2").map(String::as_str), Some("../notesSlides/notesSlide1.xml"));
    assert_eq!(rels.get("rId1").map(String::as_str), Some("../slideLayouts/slideLayout1.xml"));
    assert_eq!(rels.get("rId3").map(String::as_str), Some("https://example.com"));
}

#[test]
fn resolve_zip_target_relative_to_part_dir() {
    assert_eq!(
        resolve_zip_target("../notesSlides/notesSlide1.xml", "ppt/slides"),
        "ppt/notesSlides/notesSlide1.xml"
    );
    assert_eq!(
        resolve_zip_target("/ppt/notesSlides/notesSlide1.xml", "ppt/slides"),
        "ppt/notesSlides/notesSlide1.xml"
    );
    assert_eq!(
        resolve_zip_target("slides/slide1.xml", "ppt"),
        "ppt/slides/slide1.xml"
    );
}

#[test]
fn find_rel_target_by_type_suffix() {
    let xml = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"#;
    let target = find_rel_target_by_type_suffix(xml, "/notesSlide");
    assert_eq!(target.as_deref(), Some("../notesSlides/notesSlide1.xml"));
    assert!(find_rel_target_by_type_suffix(xml, "/image").is_none());
}
```

- [ ] **Step 2: Run tests — expect FAIL**

Run: `cargo test -p batdoc-core --lib parse_all_rels_includes_notes_slide resolve_zip_target_relative_to_part_dir find_rel_target_by_type_suffix -- --nocapture`

Expected: compile error — functions not found.

- [ ] **Step 3: Implement helpers**

Near existing rel parsers in `xml_util.rs`:

```rust
/// Parse every Relationship Id → Target (no type/mode filter).
pub(crate) fn parse_all_rels_xml(xml: &str) -> Rels {
    let mut rels = Rels::new();
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let id = get_attr(e, b"Id").unwrap_or_default();
                let target = get_attr(e, b"Target").unwrap_or_default();
                if !id.is_empty() && !target.is_empty() {
                    rels.insert(id, target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    rels
}

/// First relationship Target whose Type ends with `type_suffix`, if any.
pub(crate) fn find_rel_target_by_type_suffix(xml: &str, type_suffix: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let target = get_attr(e, b"Target").unwrap_or_default();
                let rel_type = get_attr(e, b"Type").unwrap_or_default();
                if !target.is_empty() && rel_type.ends_with(type_suffix) {
                    return Some(target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    None
}

/// Resolve a relationship Target to a ZIP entry path given the owning part's directory.
///
/// `base_dir` is e.g. `"ppt/slides"` for `ppt/slides/slide1.xml`.
pub(crate) fn resolve_zip_target(target: &str, base_dir: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        let raw = if base_dir.is_empty() {
            target.to_string()
        } else {
            format!("{base_dir}/{target}")
        };
        normalize_zip_path(&raw)
    }
}
```

Export `normalize_zip_path` remains private — `resolve_zip_target` uses it.

Optionally add:

```rust
pub(crate) fn load_part_xml(archive: &mut ZipArchive<Cursor<&[u8]>>, path: &str) -> Option<String> {
    let mut xml = String::new();
    let mut entry = archive.by_name(path).ok()?;
    entry.read_to_string(&mut xml).ok()?;
    Some(xml)
}
```

Only if it removes real duplication in later tasks; otherwise skip.

- [ ] **Step 4: Run tests — expect PASS**

Run: `cargo test -p batdoc-core --lib xml_util:: -- --nocapture`

Expected: all `xml_util` tests pass.

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/xml_util.rs
git commit -m "$(cat <<'EOF'
feat(xml_util): parse all relationships and resolve ZIP targets

Needed for PPTX notesSlide discovery; hyperlink-only load_rels drops internal parts.
EOF
)"
```

---

### Task 2: PPTX Notes trailer rendering (in-memory)

**Files:**
- Modify: `batdoc-core/src/pptx.rs` (`Slide`, `render_plain`, `render_markdown`, tests)

- [ ] **Step 1: Write failing render tests**

```rust
fn shape_with_text(text: &str) -> ShapeText {
    ShapeText {
        paragraphs: vec![Paragraph {
            runs: vec![TextRun {
                text: text.into(),
                bold: false,
                italic: false,
                link_url: None,
                font_size: None,
            }],
            heading_level: 0,
            bullet: BulletKind::None,
        }],
    }
}

#[test]
fn render_markdown_notes_after_deck() {
    let slides = vec![
        Slide {
            number: 1,
            shapes: vec![shape_with_text("Title")],
            images: vec![],
            notes: vec![],
        },
        Slide {
            number: 2,
            shapes: vec![shape_with_text("Body")],
            images: vec![],
            notes: vec![shape_with_text("Remember the demo")],
        },
    ];
    let md = render_markdown(&slides);
    assert!(md.contains("## Slide 2"));
    assert!(md.contains("Body"));
    let notes_at = md.find("## Notes").expect("notes section");
    let body_at = md.find("Body").unwrap();
    assert!(notes_at > body_at);
    assert!(md[notes_at..].contains("### Slide 2"));
    assert!(md[notes_at..].contains("Remember the demo"));
    assert!(!md[notes_at..].contains("### Slide 1"));
}

#[test]
fn render_plain_notes_after_deck() {
    let slides = vec![Slide {
        number: 3,
        shapes: vec![shape_with_text("Hi")],
        images: vec![],
        notes: vec![shape_with_text("aside")],
    }];
    let plain = render_plain(&slides);
    assert!(plain.contains("Hi"));
    assert!(plain.contains("--- Notes ---"));
    assert!(plain.contains("[Slide 3]"));
    assert!(plain.contains("aside"));
}

#[test]
fn render_notes_only_slide_no_body_heading() {
    let slides = vec![
        Slide {
            number: 1,
            shapes: vec![shape_with_text("Only body")],
            images: vec![],
            notes: vec![],
        },
        Slide {
            number: 2,
            shapes: vec![],
            images: vec![],
            notes: vec![shape_with_text("orphaned notes")],
        },
    ];
    let md = render_markdown(&slides);
    assert!(!md.contains("## Slide 2\n"));
    assert!(md.contains("## Notes"));
    assert!(md.contains("### Slide 2"));
    assert!(md.contains("orphaned notes"));
}

#[test]
fn render_no_notes_omits_section() {
    let slides = vec![Slide {
        number: 1,
        shapes: vec![shape_with_text("Hi")],
        images: vec![],
        notes: vec![],
    }];
    assert!(!render_markdown(&slides).contains("## Notes"));
    assert!(!render_plain(&slides).contains("--- Notes ---"));
}
```

- [ ] **Step 2: Run tests — expect FAIL**

Run: `cargo test -p batdoc-core --lib render_markdown_notes_after_deck -- --nocapture`

Expected: compile error — `notes` field missing on `Slide`.

- [ ] **Step 3: Minimal model + render**

1. Add `notes: Vec<ShapeText>` to `Slide`.
2. Update every `Slide { ... }` construction in the file (parse + tests) to include `notes: vec![]` (or real notes in new tests).
3. After the body loop in `render_plain` / `render_markdown`, append Notes trailer:

**Markdown:**

```rust
fn append_notes_markdown(slides: &[Slide], out: &mut String) {
    let mut any = false;
    for slide in slides {
        if !notes_nonempty(&slide.notes) {
            continue;
        }
        if !any {
            if !out.is_empty() && !out.ends_with("\n\n") {
                if out.ends_with('\n') {
                    out.push('\n');
                } else {
                    out.push_str("\n\n");
                }
            }
            out.push_str("## Notes\n\n");
            any = true;
        }
        let _ = write!(out, "### Slide {}\n\n", slide.number);
        for shape in &slide.notes {
            // reuse same paragraph rendering path as body shapes
            // (extract a small helper from the body loop or call render_para_markdown)
            ...
        }
    }
}
```

Reuse existing paragraph markdown/plain logic for note shapes (bold/italic/lists). Simplest approach: factor the per-shape paragraph loop into `render_shape_markdown(shape, out)` / plain equivalent used by both body and notes.

**Plain:**

```text
--- Notes ---
[Slide N]
note lines...
```

**Empty notes:** `notes_nonempty` = any paragraph run text with non-whitespace after trim.

Blank line before first trailer when body non-empty (spec).

4. Wire `append_notes_*` at end of `render_plain` / `render_markdown` **before** callers append `image_defs` (callers already append defs after render — keep that).

- [ ] **Step 4: Run tests — expect PASS**

Run: `cargo test -p batdoc-core --lib pptx:: -- --nocapture`

Expected: all pptx tests pass, including new notes render tests.

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/pptx.rs
git commit -m "$(cat <<'EOF'
feat(pptx): render speaker notes trailer from Slide.notes

In-memory Notes section after deck body; notes-only slides still list under Notes.
EOF
)"
```

---

### Task 3: PPTX notes discovery + body-placeholder filter

**Files:**
- Modify: `batdoc-core/src/pptx.rs` (`parse_pptx`, notes XML parse, ZIP helpers in tests)

- [ ] **Step 1: Write failing ZIP integration tests**

Add a test helper that builds a minimal pptx ZIP:

```rust
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn zip_entry(z: &mut ZipWriter<Cursor<Vec<u8>>>, name: &str, body: &str) {
    z.start_file(name, SimpleFileOptions::default()).unwrap();
    z.write_all(body.as_bytes()).unwrap();
}

fn minimal_pptx_with_notes() -> Vec<u8> {
    let buf = Cursor::new(Vec::new());
    let mut z = ZipWriter::new(buf);
    zip_entry(
        &mut z,
        "[Content_Types].xml",
        r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>"#,
    );
    zip_entry(
        &mut z,
        "ppt/presentation.xml",
        r#"<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>"#,
    );
    zip_entry(
        &mut z,
        "ppt/_rels/presentation.xml.rels",
        r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#,
    );
    zip_entry(
        &mut z,
        "ppt/slides/slide1.xml",
        r#"<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:txBody><a:p><a:r><a:t>Deck title</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#,
    );
    zip_entry(
        &mut z,
        "ppt/slides/_rels/slide1.xml.rels",
        r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"#,
    );
    zip_entry(
        &mut z,
        "ppt/notesSlides/notesSlide1.xml",
        r#"<?xml version="1.0"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>
      <p:txBody><a:p><a:r><a:t>SHOULD NOT APPEAR</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
      <p:txBody><a:p><a:r><a:t>Speak slowly</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:txBody><a:p><a:r><a:t>freeform ignored</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>"#,
    );
    z.finish().unwrap().into_inner()
}
```

Tests:

```rust
#[test]
fn extract_markdown_includes_speaker_notes() {
    let data = minimal_pptx_with_notes();
    let md = extract_markdown(&data, false).unwrap();
    assert!(md.contains("Deck title"));
    assert!(md.contains("## Notes"));
    assert!(md.contains("Speak slowly"));
    assert!(!md.contains("SHOULD NOT APPEAR"));
    assert!(!md.contains("freeform ignored"));
}

#[test]
fn extract_whitespace_notes_body_omits_section() {
    // same fixture but body placeholder only has "   "
    ...
    assert!(!md.contains("## Notes"));
}

#[test]
fn extract_missing_notes_target_ok() {
    // rel points at missing part → no Notes, no error
    ...
}
```

- [ ] **Step 2: Run tests — expect FAIL**

Run: `cargo test -p batdoc-core --lib extract_markdown_includes_speaker_notes -- --nocapture`

Expected: FAIL — notes not in output.

- [ ] **Step 3: Implement discovery + filter**

In `parse_pptx` after loading each slide:

```rust
let notes = load_slide_notes(&mut archive, &path);
slides.push(Slide { number: num, shapes, images, notes });
```

```rust
fn load_slide_notes(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    slide_path: &str,
) -> Vec<ShapeText> {
    let rels_path = xml_util::rels_path(slide_path);
    let mut rels_xml = String::new();
    match archive.by_name(&rels_path) {
        Ok(mut e) => {
            if e.read_to_string(&mut rels_xml).is_err() {
                return Vec::new();
            }
        }
        Err(_) => return Vec::new(),
    }
    let Some(target) = xml_util::find_rel_target_by_type_suffix(&rels_xml, "/notesSlide") else {
        return Vec::new();
    };
    let base_dir = slide_path.rsplit_once('/').map_or("", |(d, _)| d);
    let notes_path = xml_util::resolve_zip_target(&target, base_dir);
    let mut xml = String::new();
    match archive.by_name(&notes_path) {
        Ok(mut e) => {
            if e.read_to_string(&mut xml).is_err() {
                return Vec::new();
            }
        }
        Err(_) => return Vec::new(),
    }
    parse_notes_slide_xml(&xml)
}
```

`parse_notes_slide_xml`:

- Walk shapes like `parse_slide_xml`.
- For each `p:sp`, detect placeholder type while parsing: look for `p:ph` with `type` attribute inside the shape (before/with `txBody`).
- Keep only `type="body"`.
- Parse text with existing `parse_text_body` / empty `Rels` (no hyperlinks required).
- After parse, if all collected text is whitespace-only, return `vec![]`.

Implementation tip for placeholder type: extend shape parse with optional filter, e.g. `parse_shape_filtered(reader, rels, end_tag, only_ph: Option<&str>)` that records `ph_type` from Empty/Start `ph` via `get_attr(e, b"type")`.

Do **not** use `xml_util::load_rels` for notes.

- [ ] **Step 4: Run tests — expect PASS**

Run: `cargo test -p batdoc-core --lib pptx:: -- --nocapture`

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/pptx.rs
git commit -m "$(cat <<'EOF'
feat(pptx): discover speaker notes via notesSlide relationships

Only body placeholders contribute notes; missing targets and empty bodies omit Notes.
EOF
)"
```

---

### Task 4: DOCX marker `Run` model + body/render

**Files:**
- Modify: `batdoc-core/src/docx.rs` (`Run`, `parse_run`, plain/md run render, `InlineRun`, tests)

- [ ] **Step 1: Write failing marker render tests**

```rust
#[test]
fn render_markdown_footnote_marker_in_paragraph() {
    let blocks = vec![Block::Paragraph {
        style: ParaStyle::default(),
        runs: vec![
            Run::text("Hello"),
            Run::footnote_ref(1),
            Run::text(" world"),
        ],
    }];
    let md = render_markdown(&blocks);
    assert_eq!(md.trim_end(), "Hello[^1] world");
}

#[test]
fn render_plain_endnote_marker() {
    let blocks = vec![Block::Paragraph {
        style: ParaStyle::default(),
        runs: vec![Run::text("See"), Run::endnote_ref(2)],
    }];
    let plain = render_plain(&blocks);
    assert_eq!(plain.trim_end(), "See[e2]");
}
```

Helper constructors on `Run` (test + production):

```rust
impl Run {
    fn text(s: impl Into<String>) -> Self { /* ordinary */ }
    fn footnote_ref(display_index: usize) -> Self { ... }
    fn endnote_ref(display_index: usize) -> Self { ... }
}
```

- [ ] **Step 2: Run — expect FAIL** (no marker support)

- [ ] **Step 3: Implement model**

Replace plain `Run` fields with an enum or optional marker:

```rust
#[derive(Debug, Clone)]
enum RunKind {
    Text,
    FootnoteRef { display_index: usize },
    EndnoteRef { display_index: usize },
}

#[derive(Debug, Clone)]
struct Run {
    text: String, // unused for markers; empty
    bold: bool,
    italic: bool,
    link_url: Option<String>,
    kind: RunKind, // default Text
}
```

Or cleaner:

```rust
#[derive(Debug, Clone)]
enum Run {
    Text {
        text: String,
        bold: bool,
        italic: bool,
        link_url: Option<String>,
    },
    FootnoteRef { display_index: usize },
    EndnoteRef { display_index: usize },
}
```

**Preferred: enum `Run`** — forces exhaustive match at render sites.

Update:

- All `Run { text, bold, italic, link_url }` constructions → `Run::Text { ... }`
- `impl InlineRun for Run` — markers: `text()` returns formatted marker for markdown path **or** handle markers outside `markup::render_runs_markdown`

**Important:** `markup::render_runs_markdown` will bold-wrap marker text if it goes through `InlineRun` as normal text. Better:

```rust
fn render_runs_markdown(runs: &[Run]) -> String {
    let mut out = String::new();
    let mut i = 0;
    while i < runs.len() {
        match &runs[i] {
            Run::FootnoteRef { display_index } => {
                out.push_str(&format!("[^{display_index}]"));
                i += 1;
            }
            Run::EndnoteRef { display_index } => {
                out.push_str(&format!("[^e{display_index}]"));
                i += 1;
            }
            Run::Text { .. } => {
                // gather consecutive Text runs and pass slice to markup::render_runs_markdown
                // OR implement InlineRun only for Text and split
                ...
            }
        }
    }
    out
}
```

Simplest workable approach: keep `InlineRun` on a newtype or only call markup for contiguous `Text` runs. Alternatively implement `InlineRun` where markers return `text()` = `[^1]` and `bold/italic` false, `link_url` None — that works if markers are never inside hyperlinks (they aren't).

**Use InlineRun with marker text preformatted for markdown**, and separate plain path:

```rust
impl InlineRun for Run {
    fn text(&self) -> &str {
        match self {
            Run::Text { text, .. } => text,
            Run::FootnoteRef { .. } | Run::EndnoteRef { .. } => {
                // Can't return owned String from &str easily.
            }
        }
    }
}
```

Owned marker labels need storage. Practical pattern:

```rust
struct Run {
    text: String, // marker label filled at render-prep OR store display_index + kind
    bold: bool,
    italic: bool,
    link_url: Option<String>,
    marker: Option<NoteMarker>,
}

enum NoteMarker {
    Footnote(usize),
    Endnote(usize),
}
```

At markdown render: if `marker` is Some, emit label and skip formatting. At plain: emit `[n]` / `[en]`. Keep `text` empty for markers.

```rust
fn run_plain_text(r: &Run) -> String {
    match r.marker {
        Some(NoteMarker::Footnote(n)) => format!("[{n}]"),
        Some(NoteMarker::Endnote(n)) => format!("[e{n}]"),
        None => r.text.clone(),
    }
}
```

Plain paragraph join uses `run_plain_text`. Markdown uses custom loop or sets temporary — **custom loop is clearer**.

Update `cell_to_text` and any `r.text` joins.

- [ ] **Step 4: Tests PASS** for render-only markers + full `docx::` suite green.

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/docx.rs
git commit -m "$(cat <<'EOF'
feat(docx): model footnote/endnote marker runs in body rendering

Markers render as [^{n}] / [n] and [^e{n}] / [e{n}] without baking mode-specific text at parse time.
EOF
)"
```

---

### Task 5: DOCX parse body references + assign display_index

**Files:**
- Modify: `batdoc-core/src/docx.rs` (`parse_run`, `parse_docx` plumbing, tests with XML snippets)

- [ ] **Step 1: Failing unit test on XML paragraph parse**

Expose a test-only or package-private path. Easiest: full minimal docx ZIP (same ZipWriter pattern) once plumbing exists; for this task, test `parse_run` / paragraph with a small XML fragment by adding:

```rust
#[test]
fn parse_paragraph_footnote_reference_emits_marker() {
    // Build minimal document.xml body paragraph XML and call internal parse
    // after wiring parse_docx to load empty footnote maps is OK later.
}
```

Or integration via full ZIP in Task 6; this task focuses on `parse_run` detecting refs and a mutable `NoteIndex` passed down.

Define:

```rust
struct NoteIndex {
    /// id → display_index (assigned on first body ref)
    assigned: HashMap<String, usize>,
    next: usize,
    /// ids that have definitions (from footnotes/endnotes part)
    defined: HashSet<String>,
}

impl NoteIndex {
    fn marker_for(&mut self, id: &str) -> Option<usize> {
        if !self.defined.contains(id) {
            return None;
        }
        if let Some(&n) = self.assigned.get(id) {
            return Some(n);
        }
        let n = self.next;
        self.next += 1;
        self.assigned.insert(id.to_string(), n);
        Some(n)
    }
}
```

Two indexes: footnotes (next starts 1), endnotes (next starts 1).

- [ ] **Step 2: FAIL without parse support**

- [ ] **Step 3: Implement**

In `parse_run`, on Empty/Start `footnoteReference` / `endnoteReference`:

```rust
// Need NoteIndex params — change signature:
fn parse_run(
    reader: &mut Reader<&[u8]>,
    image_rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> (Vec<Run>, Option<Block>) // may return text run AND/OR marker run(s)
```

A single `w:r` can contain text + footnote ref. Emit multiple runs: text run if non-empty, plus marker runs when refs seen.

```rust
Ok(Event::Empty(ref e)) => {
    match e.local_name().as_ref() {
        b"footnoteReference" => {
            if let Some(id) = get_attr(e, b"id") { // w:id via get_attr(e, b"id") — OOXML local name "id"
                if let Some(n) = footnotes.marker_for(&id) {
                    markers.push(Run { marker: Some(NoteMarker::Footnote(n)), text: String::new(), ...});
                }
            }
        }
        b"endnoteReference" => { ... }
        ...
    }
}
```

`get_attr` matches local name — `id` works for `w:id`.

Thread `NoteIndex` through `parse_paragraph` → `parse_body` → `parse_table*` → `parse_hyperlink_runs`.

For this task, `defined` can be pre-filled in tests; production fill comes in Task 6.

Also update tab/br empty `Run` constructors for new fields.

- [ ] **Step 4: PASS unit tests for double-ref same id → same index; dangling id → no marker**

```rust
#[test]
fn repeated_footnote_ref_reuses_display_index() { ... }

#[test]
fn dangling_footnote_ref_emits_nothing() { ... }

#[test]
fn footnote_ref_in_table_cell() { ... }
```

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/docx.rs
git commit -m "$(cat <<'EOF'
feat(docx): assign footnote/endnote display indexes from body references

Repeated refs share an index; undefined ids emit no marker; works inside tables.
EOF
)"
```

---

### Task 6: DOCX parse footnotes, endnotes, comments parts

**Files:**
- Modify: `batdoc-core/src/docx.rs`

- [ ] **Step 1: Failing ZIP tests for extras parse + trailer content**

```rust
fn minimal_docx(parts: &[(&str, &str)]) -> Vec<u8> { /* Content_Types + listed parts */ }

#[test]
fn extract_footnote_markdown() {
    let data = minimal_docx(&[
        ("word/document.xml", r#"..."#), // body with footnoteReference w:id="1"
        ("word/_rels/document.xml.rels", r#"..."#), // may be empty Relationships
        ("word/footnotes.xml", r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="-1"><w:p><w:r><w:t>sep</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="0"><w:p><w:r><w:t>cont</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>Source A</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Unreferenced</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#),
    ]);
    let md = extract_markdown(&data, false).unwrap();
    assert!(md.contains("[^1]"));
    assert!(md.contains("## Footnotes"));
    assert!(md.contains("[^1]:"));
    assert!(md.contains("Source A"));
    assert!(!md.contains("Unreferenced"));
    assert!(!md.contains("sep"));
    assert!(!md.contains("## Endnotes"));
}
```

Similar tests: endnotes, comments by author, empty comments omit section, multi-paragraph footnote indented definition, comments order, Anonymous author.

- [ ] **Step 2: FAIL**

- [ ] **Step 3: Implement parsers**

```rust
struct Comment {
    id: String,
    author: String,
    blocks: Vec<Block>,
}

struct Note {
    id: String,
    display_index: usize, // 0 = unassigned / not referenced
    blocks: Vec<Block>,
}
```

`parse_docx` returns:

```rust
(Vec<Block>, Vec<String>, Vec<Comment>, Vec<Note>, Vec<Note>)
```

Flow:

1. Open archive; load hyperlink + image rels as today.
2. Optionally read `word/footnotes.xml` / `endnotes.xml` / `comments.xml` via `by_name` — missing → empty.
3. Parse definition maps:
   - footnotes/endnotes: each `w:footnote`/`w:endnote` with id not in `{"-1","0"}`; first id wins; body via shared block walker with **empty** image rels and **dummy** NoteIndexes that never assign (refs inside notes rare — ignore nested refs or use separate empty indexes).
   - Skip `w:footnoteRef`, `w:endnoteRef`, `w:annotationRef` in `parse_run` Empty/Start (consume, no text).
4. Parse comments: each `w:comment`; `w:author` attr; empty author → `"Anonymous"`; parse inner blocks like document body paragraphs/tables; drop if no non-whitespace text.
5. Parse `document.xml` body with real NoteIndexes whose `defined` = keys of definition maps.
6. After body walk, build footnote/endnote `Vec<Note>` for trailer: entries with assigned display_index, sorted by display_index.
7. `resolve_images` as today.

Shared block parsing: extract something like `parse_block_children(reader, end_tag, rels, image_rels, fn_idx, en_idx)` used by body, comments, footnotes. Avoid wholesale rewrite — comments/notes can use a simplified walker that handles `w:p` and `w:tbl` until parent end.

For footnote part structure, walker starts inside each footnote element.

- [ ] **Step 4: PASS parse-focused tests** (trailer may still be incomplete until Task 7 — if so, assert via `parse_docx` in tests with `#[cfg(test)]` visibility or complete trailer in same task if smaller).

**Preference:** implement parse + wire into extract in Task 7 if trailer is large; otherwise finish extract here with stub trailer and expand in 7.

Plan choice: **Task 6 delivers parse + populated structures from `parse_docx`; Task 7 renders trailers and assembly.**

Add `#[cfg(test)]` accessors or test through partial public behavior once Task 7 lands. For TDD red on Task 6, test internal `parse_footnotes_xml(xml) -> HashMap<String, Vec<Block>>` functions directly.

```rust
fn parse_footnotes_xml(xml: &str) -> HashMap<String, Vec<Block>> { ... }
fn parse_endnotes_xml(xml: &str) -> HashMap<String, Vec<Block>> { ... }
fn parse_comments_xml(xml: &str) -> Vec<Comment> { ... }
```

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/docx.rs
git commit -m "$(cat <<'EOF'
feat(docx): parse comments, footnotes, and endnotes parts

Skip separator ids and ref glyphs; comments keep author and multi-paragraph bodies.
EOF
)"
```

---

### Task 7: DOCX trailers + assembly order (`body → extras → image_defs`)

**Files:**
- Modify: `batdoc-core/src/docx.rs` (`extract_plain`, `extract_markdown`)

- [ ] **Step 1: Failing end-to-end tests** (full ZIP)

Cover:

- Footnotes + endnotes sections and markers
- Comments `### Author` grouping (consecutive merge; non-consecutive repeat)
- Multi-block footnote markdown indentation
- Plain multi-paragraph footnote continuation
- Missing extras → identical to body-only baseline string
- Images + footnotes: extras appear **before** image_defs (`[image1]:` after `## Footnotes`)
- Spacing: blank line before first trailer when body non-empty

- [ ] **Step 2: FAIL**

- [ ] **Step 3: Implement render helpers**

```rust
fn append_comments_markdown(comments: &[Comment], out: &mut String) { ... }
fn append_notes_definitions_markdown(title: &str, label: impl Fn(usize)->String, notes: &[Note], out: &mut String) { ... }
```

Comments markdown:

```markdown
## Comments

### Alice

comment blocks...

### Bob

...
```

Group consecutive same author.

Footnotes:

```rust
// single paragraph:
out.push_str(&format!("[^{n}]: {text}\n\n"));
// multi block:
out.push_str(&format!("[^{n}]:\n"));
let mut body = String::new();
for b in &note.blocks {
    render_block_markdown(b, &mut body);
}
for line in body.trim_end().lines() {
    out.push_str("    ");
    out.push_str(line);
    out.push('\n');
}
out.push('\n');
```

Endnotes: `[^e{n}]` labels via `format!("[^e{n}])`.

Plain counterparts with `--- Comments ---`, `[Author]`, `[n] body`.

`extract_markdown`:

```rust
let (blocks, image_defs, comments, footnotes, endnotes) = parse_docx(data, images)?;
let mut md = render_markdown(&blocks);
let had_body = !md.is_empty();
append_extras_markdown(&mut md, &comments, &footnotes, &endnotes, had_body);
for def in &image_defs {
    md.push_str(def);
    md.push('\n');
}
Ok(md)
```

`append_extras_*` inserts leading blank line once before first non-empty section when `had_body`.

- [ ] **Step 4: Full `cargo test -p batdoc-core --lib` PASS**

- [ ] **Step 5: Commit**

```bash
git add batdoc-core/src/docx.rs
git commit -m "$(cat <<'EOF'
feat(docx): append comments, footnotes, and endnotes after body

Assembly order is body, extras, then image definitions; empty sections omitted.
EOF
)"
```

---

### Task 8: README Formats blurb

**Files:**
- Modify: `README.md` (Formats section ~lines 51–70)

- [ ] **Step 1: Edit README**

After the `.docx` / `.pptx` paragraphs, ensure readers know:

- `.docx` dumps also include comments, footnotes, and endnotes when present (footnotes/endnotes as trailing definitions with body markers).
- `.pptx` dumps include speaker notes after the deck when present.

Keep tone terse like existing README. No new flags.

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "$(cat <<'EOF'
docs: mention speaker notes, comments, footnotes, and endnotes
EOF
)"
```

---

### Task 9: Final verification

- [ ] **Step 1: Run full library tests**

```bash
cargo test -p batdoc-core --lib
```

Expected: all pass (257 + new).

- [ ] **Step 2: Smoke with real extract if fixtures exist**

```bash
cargo test -p batdoc-core --lib extract_
```

- [ ] **Step 3: Spec checklist**

Confirm against design:

- [x] PPTX notes via rels, body ph only, whitespace empty
- [x] Notes-only slide
- [x] DOCX markers + shared display_index
- [x] Unreferenced notes omitted; all comments kept
- [x] Separator ids + ref glyphs skipped
- [x] Assembly body → extras → image_defs
- [x] No public API change
- [x] README

- [ ] **Step 4: Commit any leftover fixes** (or empty if clean)

---

## Self-review (plan vs spec)

| Spec item | Task |
|-----------|------|
| PPTX `Slide.notes` + Notes trailer | 2 |
| notesSlide rel discovery ≠ `load_rels` | 1, 3 |
| body placeholder only; freeform ignored | 3 |
| whitespace empty notes | 3 |
| notes-only slide | 2, 3 |
| DOCX marker model plain≠md | 4 |
| display_index first ref; reuse | 5 |
| dangling ref no marker | 5 |
| table cell refs | 5 |
| footnotes/endnotes/comments parse | 6 |
| skip -1/0 and ref glyphs | 6 |
| trailer membership | 6–7 |
| multi-block MD indent / plain cont. | 7 |
| comments author grouping | 7 |
| assembly vs image_defs | 7 |
| no extra hyperlink rels | 6 (empty rels in extras) |
| README | 8 |
| no API/CLI | all |

No TBD placeholders. Types consistent: `NoteIndex`, `Comment`, `Note`, `NoteMarker` / `Run` enum as chosen in Task 4 and used thereafter.
