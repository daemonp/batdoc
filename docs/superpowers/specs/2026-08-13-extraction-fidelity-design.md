# Extraction Fidelity — Speaker Notes, Comments, Footnotes, Endnotes

Date: 2026-08-13
Status: approved design
Scope: `batdoc-core` PPTX and DOCX extractors; CLI inherits automatically

## Problem

`batdoc` already dumps the main body of Office documents well. User-authored extras
in the same ZIP packages are silently dropped:

- PPTX speaker notes live in `ppt/notesSlides/` and are not parsed.
- DOCX comments live in `word/comments.xml` and are not parsed.
- DOCX footnotes live in `word/footnotes.xml` and are not parsed.
- DOCX endnotes live in `word/endnotes.xml` and are not parsed.

These are the highest-value leftovers in files we already open. Headers and
footers are out of scope: they are usually boilerplate (page numbers, running
titles) and would pollute a `cat` dump.

## Goals

- Extract PPTX speaker notes and DOCX comments, footnotes, and endnotes.
- Always on. No new CLI flags and no new public library options.
- Keep the document body clean: extras are trailing sections.
- Keep a small body marker at each footnote/endnote reference site so the
  trailing definitions remain findable.
- Missing extras must not change existing dumps.
- Empty extras omit their section heading entirely.

## Non-goals

- Headers, footers, text boxes, track changes, content controls.
- PPTX comments, slide-master inheritance, speaker-note images.
- Inline expansion of footnote/comment text at the reference site.
- New `ExtractOptions`, CLI flags, or public API surface.
- Shared `ExtraContent` trailer type (premature until more extras land).
- OCR, legacy `.ppt`, OpenDocument, library metadata.

## Approach

Extend the existing PPTX and DOCX models. Reuse the current XML walkers and
renderers. Do not introduce a new public type or options struct.

Public API stays:

```rust
pub fn extract_plain(data: &[u8], format: Format) -> Result<String>;
pub fn extract_markdown(data: &[u8], format: Format, images: bool) -> Result<String>;
```

## Data model

### PPTX

`Slide` gains an optional notes field that reuses the existing shape parser:

```rust
struct Slide {
    number: usize,
    shapes: Vec<ShapeText>,
    images: Vec<String>,
    notes: Vec<ShapeText>, // empty = no speaker notes
}
```

Notes are discovered from each slide's relationships (`notesSlide` type), not
by guessing `ppt/notesSlides/notesSlideN.xml`.

A notes slide contains more than speaker notes: typically a slide-image
placeholder, header/footer placeholders, and the notes body. Only text from
shapes whose placeholder type is `body` (`p:ph type="body"`) is kept. Other
shapes are ignored so slide titles do not get duplicated into the Notes
section. That body is then parsed with the existing text-body walker.

Placeholder chrome such as "Click to add notes" is treated as empty and
omitted.

### DOCX

`Block` is unchanged. The parser returns body blocks plus three extra streams:

```rust
struct Comment {
    id: String,
    author: String,
    blocks: Vec<Block>, // comments can be multi-paragraph
}

struct Note {
    id: String,          // document id from w:id
    display_index: usize, // 1-based order of first body reference
    blocks: Vec<Block>,
}

// parse_docx returns:
// (Vec<Block>, Vec<String> /* image defs */, Vec<Comment>, Vec<Note> /* footnotes */, Vec<Note> /* endnotes */)
```

Sources:

| Extra      | ZIP part              | Trigger in `word/document.xml`      |
|------------|-----------------------|-------------------------------------|
| Comments   | `word/comments.xml`   | none (not marked inline)            |
| Footnotes  | `word/footnotes.xml`  | `w:footnoteReference`               |
| Endnotes   | `word/endnotes.xml`   | `w:endnoteReference`                |

Skip separator and continuation notes (`w:id` of `-1` and `0`).

Body still only walks `word/document.xml`. When a run contains
`w:footnoteReference` or `w:endnoteReference`, emit a marker run. Comments
are not marked inline.

## Rendering

Extras are always appended after the body. Order is fixed:

1. Notes (PPTX only)
2. Comments (DOCX only)
3. Footnotes (DOCX only)
4. Endnotes (DOCX only)

If a category is empty, omit its heading entirely. Multi-file CLI dumps keep
each file's extras with that file.

### Markdown

```markdown
## Notes

### Slide 3
speaker note paragraphs here

## Comments

### Alice
comment text

## Footnotes

[^1]: footnote body

## Endnotes

[^e1]: endnote body
```

Body markers:

- Footnote reference → `[^1]`
- Endnote reference → `[^e1]`

The `e` prefix keeps footnote and endnote definition labels from colliding.
Display numbers are assigned in document order of first reference, not by the
OOXML `w:id` value.

Comment authors become `###` headings. Consecutive comments from the same
author stay under one heading. Missing/empty author becomes `Anonymous`.

Speaker-note paragraphs reuse the existing PPTX paragraph renderer (bold,
italic, lists). Footnote/endnote bodies reuse the existing DOCX block
renderer so a footnote that contains a list or a short table still renders.

### Plain text

```text
--- Notes ---
[Slide 3]
speaker note text

--- Comments ---
[Alice]
comment text

--- Footnotes ---
[1] footnote body

--- Endnotes ---
[e1] endnote body
```

Body markers: `[1]` for footnotes, `[e1]` for endnotes.

## Error handling

| Situation                                         | Behaviour                                      |
|---------------------------------------------------|------------------------------------------------|
| Optional part missing (`comments.xml`, notes rel) | Omit. Never fail.                              |
| Part exists but cannot be read from the ZIP       | Fail the extract (`BatdocError::Io` / `Zip`).  |
| Extra part XML is malformed                       | Same as body XML: stop walking that part.      |
| PPTX notes rel target missing from the ZIP        | Skip that slide's notes.                       |
| Footnote/endnote ref with no matching definition  | Omit the body marker. No dangling `[^5]`.      |
| Empty extras (whitespace, ids `-1`/`0`, chrome)   | Omit the section.                              |
| Duplicate footnote/comment ids                    | First wins.                                    |
| Encrypted / corrupt container                     | Unchanged: fail as today.                      |

No new `BatdocError` variants.

## Testing

Synthetic XML/ZIP fixtures, same style as the existing `docx.rs` / `pptx.rs`
unit tests. Cover both markdown and plain.

PPTX:

- Notes under one slide appear after the deck under `## Notes` / `--- Notes ---`.
- Mixed deck: only slides with real notes appear under `### Slide N`.
- Empty / placeholder notes omit the Notes section.
- Notes text does not leak into the slide body.
- Slide-image / title placeholders on the notes slide do not appear under Notes.

DOCX:

- Footnote reference emits `[^1]` / `[1]` and a matching trailing definition.
- Endnote reference emits `[^e1]` / `[e1]` and a matching trailing definition.
- Comments grouped by author; no inline marker. Multi-paragraph comments keep their paragraphs.
- Separator ids `-1` and `0` are skipped.
- Missing extras leave the dump unchanged.
- Dangling reference emits no marker.

Regression: existing dumps without extras stay byte-for-byte unchanged.

## Files to change

- `batdoc-core/src/pptx.rs` — notes discovery, `Slide.notes`, trailer render
- `batdoc-core/src/docx.rs` — extra-part parse, body markers, trailer render
- `README.md` — mention notes / comments / footnotes / endnotes under Formats

No CLI, public API, man page, or packaging changes.

## Implementation notes

PPTX notes discovery should follow slide relationships:

1. Load `ppt/slides/_rels/slideN.xml.rels`.
2. Find the relationship whose `Type` ends with `/notesSlide`.
3. Resolve `Target` relative to the slide path.
4. Parse that notes XML with the existing shape walker.

DOCX extra parts are siblings of `word/document.xml`. Parse them only if
`ZipArchive::by_name` succeeds. Build an id → note map, then assign
`display_index` while walking body references so markers and definitions
share the same number.

Keep image extraction out of extras for this slice. Speaker-note images and
images inside comments/footnotes are ignored.

## Success criteria

- `batdoc slides.pptx` shows speaker notes after the slides when they exist.
- `batdoc report.docx` shows comments, footnotes, and endnotes after the body
  when they exist, with `[^1]` / `[^e1]` markers at the reference sites.
- A document with none of these extras produces the same output as today.
- `cargo test -p batdoc-core` covers the cases above.
- Public `batdoc-core` API is unchanged.
