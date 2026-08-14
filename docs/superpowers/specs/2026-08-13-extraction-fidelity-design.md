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

- Headers, footers, track changes, content controls.
- PPTX comments, slide-master inheritance, speaker-note images.
- Freeform text boxes on PPTX notes pages outside the notes body placeholder.
- Modern threaded comments (`word/commentsExtended.xml`); only `word/comments.xml`.
- Hyperlink resolution inside comment/footnote/endnote parts (no per-extra
  `_rels` load this slice; link text still extracts as plain runs).
- Inline expansion of footnote/comment text at the reference site.
- New `ExtractOptions`, CLI flags, or public API surface.
- Shared `ExtraContent` trailer type (premature until more extras land).
- OCR, legacy `.ppt`, OpenDocument, library metadata.
- Image extraction from extras (speaker-note images; images inside
  comments/footnotes/endnotes).

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

`Slide` gains a notes field that reuses the existing shape parser:

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

**Rels loading:** `xml_util::load_rels` / `parse_rels_xml` only keep
external/hyperlink targets and will drop `notesSlide`. Notes discovery must
parse the slide `.rels` the same way `discover_slides` parses
`presentation.xml.rels` (all `Relationship` entries, filter by `Type` ending
in `/notesSlide`), or add a small internal-rel helper in `xml_util.rs`. Do not
call `load_rels` for this.

A notes slide contains more than speaker notes: typically a slide-image
placeholder, header/footer placeholders, and the notes body. Only text from
shapes whose placeholder type is `body` (`p:ph type="body"`) is kept. Other
shapes — including freeform text boxes on the notes page — are ignored so
slide titles do not get duplicated into the Notes section. That body is then
parsed with the existing text-body walker.

**Empty notes:** treat the notes body as absent when the body-placeholder
text is empty or whitespace-only after trim. Do not match localized chrome
strings such as "Click to add notes".

Notes are independent of whether the slide body printed. A slide with empty
`shapes` but non-empty `notes` still contributes a Notes entry (and still
does not print a body slide heading).

### DOCX

`Block` stays as the block enum. Footnote/endnote body markers are **not**
plain text baked into `Run::text` at parse time: plain and markdown markers
differ (`[1]` vs `[^1]`).

```rust
// Extend Run (or equivalent) with a marker distinct from ordinary text.
// Concrete shape is an implementation choice; behaviour is fixed:
// - Footnote { display_index }  → plain `[N]` / markdown `[^N]`
// - Endnote  { display_index }  → plain `[eN]` / markdown `[^eN]`
// Ordinary text runs unchanged: { text, bold, italic, link_url }

struct Comment {
    id: String,
    author: String,
    blocks: Vec<Block>, // comments can be multi-paragraph
}

struct Note {
    id: String,           // document id from w:id
    display_index: usize, // 1-based order of first body reference
    blocks: Vec<Block>,
}

// parse_docx returns:
// (Vec<Block>, Vec<String> /* image defs */, Vec<Comment>,
//  Vec<Note> /* footnotes */, Vec<Note> /* endnotes */)
```

Sources:

| Extra     | ZIP part             | Trigger in `word/document.xml` |
| --------- | -------------------- | ------------------------------ |
| Comments  | `word/comments.xml`  | none (not marked inline)       |
| Footnotes | `word/footnotes.xml` | `w:footnoteReference`          |
| Endnotes  | `word/endnotes.xml`  | `w:endnoteReference`           |

Skip separator and continuation notes (`w:id` of `-1` and `0`).

Body still only walks `word/document.xml`. When a run contains
`w:footnoteReference` or `w:endnoteReference` (typically an empty element
with `w:id`), emit a marker — including when the reference sits inside a
table cell paragraph. Comments are not marked inline; ignore
`w:commentRangeStart` / `w:commentRangeEnd` / `w:commentReference` for
output.

#### Trailer membership

- **Footnotes / endnotes:** only notes that are referenced at least once in
  the body, have a definition, and are not separator ids. Unreferenced
  definitions are omitted.
- **Comments:** every comment in `comments.xml` with non-empty text after
  parse, in `comments.xml` document order. Missing range markers in the body
  do not drop the comment.

#### Display index assignment

Assign `display_index` on first body reference in document order (tables and
nested blocks included). Footnotes and endnotes each have their own counter
starting at 1.

Later references to the same `w:id` **reuse** that index (both sites emit
`[^1]` / `[1]`). They do not allocate a new number.

Duplicate ids inside a definition part: first definition wins; later
duplicates ignored.

#### Chrome inside extra parts

When parsing comment/footnote/endnote bodies, skip auto-number / annotation
glyph elements so they do not leak into text:

- `w:footnoteRef`, `w:endnoteRef`
- `w:annotationRef`

## Rendering

### Output assembly order

Final string composition is fixed:

1. Body (including inline footnote/endnote markers)
2. Extras trailers (below), each omitted when empty
3. Image reference definitions (`image_defs`), unchanged from today

So: `body → Notes/Comments/Footnotes/Endnotes → image_defs`.

Documents with images but no extras stay byte-for-byte identical. Image
base64 blocks must not sit between body and footnotes.

Extras category order:

1. Notes (PPTX only)
2. Comments (DOCX only)
3. Footnotes (DOCX only)
4. Endnotes (DOCX only)

### Spacing

- If the body is non-empty and at least one extra section will print, insert
  one blank line before the first trailer heading.
- One blank line between consecutive trailer sections.
- Inside a section, follow existing markdown/plain block spacing from the
  reused renderers.

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
OOXML `w:id` value. Repeated refs share the label.

Comment authors become `###` headings. Consecutive comments from the same
author stay under one heading (non-consecutive same author → heading repeats).
Missing/empty author becomes `Anonymous`. Empty-bodied comments are dropped
before grouping; if none remain, omit `## Comments`.

Speaker-note paragraphs reuse the existing PPTX paragraph renderer (bold,
italic, lists).

#### Multi-block footnote / endnote / comment bodies

Comments render as normal blocks under the author heading (reuse
`render_block_markdown`).

Footnote/endnote definitions use CommonMark label syntax:

- **Single paragraph:** `[^1]: paragraph text` on one line (same for
  `[^e1]:`).
- **Multiple blocks** (extra paragraphs, lists, tables): put the definition
  marker on its own line, then render blocks on the following lines with each
  content line indented by 4 spaces so CommonMark treats them as part of the
  definition:

```markdown
[^1]:
    First paragraph.

    - list item

    | a | b |
    | --- | --- |
    | c | d |
```

Reuse the DOCX block renderer for the indented body. Indent every line of
its output by 4 spaces.

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

Body markers: `[1]` for footnotes, `[e1]` for endnotes (shared across repeat
refs).

Multi-block note/comment bodies: first line prefixed with `[1] ` / `[e1] ` /
author context as above; continuation blocks on following lines with no
extra bracket prefix (plain block renderer output, blank line between
blocks as today). Example:

```text
[1] First paragraph.
Second paragraph of the same footnote.
```

## Error handling

| Situation                                         | Behaviour                                     |
| ------------------------------------------------- | --------------------------------------------- |
| Optional part missing (`comments.xml`, notes rel) | Omit. Never fail.                             |
| Part exists but cannot be read from the ZIP       | Fail the extract (`BatdocError::Io` / `Zip`). |
| Extra part XML is malformed                       | Same as body XML: stop walking that part.     |
| PPTX notes rel target missing from the ZIP        | Skip that slide's notes.                      |
| Footnote/endnote ref with no matching definition  | Omit the body marker. No dangling `[^5]`.     |
| Empty extras (whitespace body, ids `-1`/`0`)      | Omit the section.                             |
| Duplicate footnote/comment definition ids         | First wins.                                   |
| Multiple body refs to one note id                 | Same `display_index`; marker at each site.    |
| Unreferenced footnote/endnote definitions         | Omit from trailer.                            |
| Encrypted / corrupt container                     | Unchanged: fail as today.                     |

No new `BatdocError` variants.

## Testing

Prefer small synthetic XML/ZIP fixtures where relationships and parts must
round-trip (notes rel discovery, optional parts). In-memory `Slide` / `Block`
render tests remain fine for pure trailer formatting. Cover both markdown and
plain.

PPTX:

- Notes under one slide appear after the deck under `## Notes` / `--- Notes ---`.
- Mixed deck: only slides with real notes appear under `### Slide N`.
- Empty / whitespace-only body placeholder omits the Notes section.
- Notes text does not leak into the slide body.
- Slide-image / title placeholders on the notes slide do not appear under Notes.
- Freeform (non-`body` placeholder) shapes on the notes page are ignored.
- Notes-only slide (empty body shapes, non-empty notes): no body slide
  section, but Notes trailer still lists that slide.
- Notes rel target missing: that slide contributes no notes; extract succeeds.

DOCX:

- Footnote reference emits `[^1]` / `[1]` and a matching trailing definition.
- Endnote reference emits `[^e1]` / `[e1]` and a matching trailing definition.
- Second body reference to the same footnote id reuses `[^1]` / `[1]`.
- Footnote reference inside a table cell still gets a marker and definition.
- Multi-paragraph footnote: markdown uses indented definition continuation;
  plain continues without a second `[N]` prefix.
- Comments grouped by author; no inline marker. Multi-paragraph comments keep
  their paragraphs. Order follows `comments.xml`.
- Separator ids `-1` and `0` are skipped; `footnoteRef` / `endnoteRef` /
  `annotationRef` glyphs do not appear in text.
- Unreferenced footnote definitions do not appear in the trailer.
- Missing extras leave the dump unchanged.
- Dangling reference emits no marker.
- With `--images`, assembly is body → extras → image_defs (extras not after
  base64). Images-only documents remain unchanged.

Regression: existing dumps without extras stay byte-for-byte unchanged.

## Files to change

- `batdoc-core/src/pptx.rs` — notes discovery, `Slide.notes`, trailer render
- `batdoc-core/src/docx.rs` — extra-part parse, marker runs, trailer render
- `batdoc-core/src/xml_util.rs` — optional internal-rel helper for notesSlide
  (or equivalent local parse in `pptx.rs`)
- `README.md` — mention notes / comments / footnotes / endnotes under Formats

No CLI, public API, man page, or packaging changes.

## Implementation notes

PPTX notes discovery should follow slide relationships:

1. Read `ppt/slides/_rels/slideN.xml.rels` (via `rels_path`).
2. Parse **all** relationships (not `load_rels`); find `Type` ending with
   `/notesSlide`.
3. Resolve `Target` relative to the slide directory (same join + `..`
   normalization style as image/slide path resolution).
4. Parse that notes XML; keep only shapes with `p:ph type="body"`.
5. If body-placeholder text is empty/whitespace, leave `notes` empty.

DOCX extra parts are siblings of `word/document.xml`. Parse them only if
`ZipArchive::by_name` succeeds.

1. Build id → note definition maps from footnotes/endnotes parts (skip
   `-1`/`0`; first id wins; strip ref glyphs).
2. Build comments list from `comments.xml` (document order; drop empty
   bodies; author default `Anonymous`).
3. Walk the body (including table cells). On each
   `w:footnoteReference` / `w:endnoteReference`:
   - If id has a definition and no `display_index` yet, allocate the next
     counter value and remember it on the note.
   - If id has a definition, emit a marker with that `display_index`.
   - If id has no definition, emit nothing.
4. Trailer lists footnotes/endnotes that received a `display_index`, sorted
   by `display_index` ascending.

Do not load `word/_rels/comments.xml.rels` (or footnote/endnote rels) this
slice.

Keep image extraction out of extras. After rendering body + extras, append
`image_defs` exactly as today.

## Success criteria

- `batdoc slides.pptx` shows speaker notes after the slides when they exist.
- `batdoc report.docx` shows comments, footnotes, and endnotes after the body
  when they exist, with `[^1]` / `[^e1]` markers at the reference sites.
- A document with none of these extras produces the same output as today.
- Documents with images but no extras still match today's image output.
- `cargo test -p batdoc-core` covers the cases above.
- Public `batdoc-core` API is unchanged.
