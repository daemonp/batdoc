//! OOXML `.docx` (Office Open XML) format parser.
//!
//! Unzips the `.docx` archive, parses `word/document.xml` with `quick-xml`
//! into structured [`Block`] types (paragraphs with heading/list styles and
//! runs with bold/italic/hyperlink, tables with rows and cells), then renders
//! to either plain text or markdown.

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::io::{Cursor, Read};
use zip::ZipArchive;

use crate::markup;
use crate::xml_util::{self, get_attr, Rels};
use crate::ExtractOptions;

/// Extracted document structure for rich output.
#[derive(Debug)]
enum Block {
    Paragraph {
        style: ParaStyle,
        runs: Vec<Run>,
    },
    Table {
        rows: Vec<Row>,
    }, // rows -> cells -> blocks
    /// An embedded image: the ZIP path of the image data, plus optional
    /// markdown reference (`--images`) and optional OCR text (`--ocr`).
    Image {
        /// ZIP path of the embedded image (e.g. `word/media/image1.png`).
        path: String,
        /// Inline markdown reference (`![][imageN]`) when `--images` is set.
        markdown: Option<String>,
        /// OCR'd text of the image when `--ocr` is set.
        ocr_text: Option<String>,
    },
}

#[derive(Debug, Clone, Default)]
struct ParaStyle {
    heading_level: u8, // 0 = normal, 1-9 = heading
    list_level: Option<u8>,
}

/// A footnote or endnote reference marker inside a paragraph.
///
/// The `usize` is the display index (1-based) to render in the output;
/// how that index is derived from the document's footnote/endnote parts
/// is handled by the parser (see the extraction-fidelity plan).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NoteMarker {
    Footnote(usize), // display_index
    Endnote(usize),  // display_index
}

#[derive(Debug, Clone)]
struct Run {
    text: String,
    bold: bool,
    italic: bool,
    /// If this run is inside a hyperlink, the resolved URL.
    link_url: Option<String>,
    /// None = ordinary text; Some = marker run with empty `text`.
    marker: Option<NoteMarker>,
}

impl Run {
    /// Create an ordinary text run.
    fn text(s: impl Into<String>) -> Self {
        Self {
            text: s.into(),
            bold: false,
            italic: false,
            link_url: None,
            marker: None,
        }
    }

    /// Create a footnote reference marker run (no visible text of its own).
    const fn footnote_ref(display_index: usize) -> Self {
        Self {
            text: String::new(),
            bold: false,
            italic: false,
            link_url: None,
            marker: Some(NoteMarker::Footnote(display_index)),
        }
    }

    /// Create an endnote reference marker run (no visible text of its own).
    const fn endnote_ref(display_index: usize) -> Self {
        Self {
            text: String::new(),
            bold: false,
            italic: false,
            link_url: None,
            marker: Some(NoteMarker::Endnote(display_index)),
        }
    }
}

/// A single table cell containing blocks.
type Cell = Vec<Block>;
/// A table row: a sequence of cells.
type Row = Vec<Cell>;

/// Tracks `display_index` assignment for footnote/endnote references.
///
/// One instance per note type (footnotes, endnotes), numbering from 1.
/// The first body reference to a defined `w:id` claims the next index;
/// later references to the same id reuse it.
struct NoteIndex {
    /// w:id → assigned `display_index` (first body reference wins).
    assigned: HashMap<String, usize>,
    /// Whether a definition exists for a w:id (seeded from the note parts by
    /// `parse_docx`; tests seed via `add_defined`).
    defined: HashSet<String>,
    /// Next `display_index` to assign.
    next: usize,
}

impl NoteIndex {
    fn new() -> Self {
        Self {
            assigned: HashMap::new(),
            defined: HashSet::new(),
            next: 1,
        }
    }

    /// Record that a definition exists for a note id, so body references to
    /// it emit a marker. Seeded by `parse_docx` from the note parts; tests
    /// seed it directly to exercise marker assignment.
    fn add_defined(&mut self, id: impl Into<String>) {
        self.defined.insert(id.into());
    }

    /// Returns `Some(display_index)` if defined; assigns next on first sight.
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

impl Default for NoteIndex {
    fn default() -> Self {
        Self::new()
    }
}

/// A document comment from `word/comments.xml`.
///
/// `author` and `blocks` are rendered into a `## Comments` /
/// `--- Comments ---` trailer, grouped into runs of consecutive comments
/// sharing an author. Empty-author comments map to `Anonymous` at parse
/// time. The `id` is parse-side bookkeeping (duplicate ids keep the first
/// occurrence) and is not rendered.
struct Comment {
    /// Dedup key; kept for tests and parse order, not shown in the trailers.
    #[allow(dead_code)]
    id: String,
    author: String,
    blocks: Vec<Block>, // comments can be multi-paragraph
}

/// A footnote or endnote definition from the note parts.
///
/// `display_index` is 0 until the body walk assigns the first reference to
/// this note's `w:id`; notes never referenced in the body keep 0 and are
/// omitted from the trailers, which render only referenced notes in
/// ascending `display_index` order.
struct Note {
    id: String,
    display_index: usize,
    blocks: Vec<Block>,
}

/// Whether any block holds non-whitespace text.
///
/// Used to drop empty extras (e.g. a comment whose body is only separator
/// paragraphs). Marker runs count as empty (they carry no text of their own).
fn blocks_have_text(blocks: &[Block]) -> bool {
    blocks.iter().any(|block| match block {
        Block::Paragraph { runs, .. } => runs.iter().any(|r| !r.text.trim().is_empty()),
        Block::Table { rows } => rows.iter().flatten().any(|cell| blocks_have_text(cell)),
        Block::Image { .. } => false,
    })
}

/// Copy the display indexes the body walk assigned (`index.assigned`) onto
/// the parsed definitions; notes never referenced in the body keep 0.
fn assign_display_indexes(notes: Vec<Note>, index: &NoteIndex) -> Vec<Note> {
    notes
        .into_iter()
        .map(|mut note| {
            note.display_index = index.assigned.get(&note.id).copied().unwrap_or(0);
            note
        })
        .collect()
}

/// Extract plain text from a .docx file.
pub(crate) fn extract_plain(data: &[u8], opts: ExtractOptions) -> crate::error::Result<String> {
    let (blocks, _, comments, footnotes, endnotes) = parse_docx(
        data,
        ExtractOptions {
            images: false,
            ..opts
        },
    )?;
    let mut out = render_plain(&blocks);
    append_extras_plain(&mut out, &comments, &footnotes, &endnotes);
    Ok(out)
}

/// Extract markdown-formatted text from a .docx file.
///
/// When `opts.images` is set, embedded images are extracted and included as
/// reference-style base64 images: `![][imageN]` inline with definitions
/// appended at the end of the document. When `opts.ocr` is set, embedded images
/// are OCR'd and their text is rendered as a blockquote after the image.
pub(crate) fn extract_markdown(data: &[u8], opts: ExtractOptions) -> crate::error::Result<String> {
    let (blocks, image_defs, comments, footnotes, endnotes) = parse_docx(data, opts)?;
    let mut md = render_markdown(&blocks);
    append_extras_markdown(&mut md, &comments, &footnotes, &endnotes);
    if !image_defs.is_empty() {
        for def in &image_defs {
            md.push_str(def);
            md.push('\n');
        }
    }
    Ok(md)
}

/// What [`parse_docx`] produces: body blocks, image reference definitions,
/// and the extras from the optional comment/footnote/endnote parts.
type DocxOutput = (Vec<Block>, Vec<String>, Vec<Comment>, Vec<Note>, Vec<Note>);

/// Read an optional ZIP part's full text.
///
/// `Ok(None)` is returned only when the part is genuinely absent from the
/// archive — a missing comment/footnote/endnote part is not an error and is
/// simply omitted. A part that is present but cannot be read from the ZIP
/// (corrupt deflate data, truncated entry, archive I/O failure) fails the
/// extract with a [`crate::error::BatdocError`]; it is NOT treated as
/// missing.
fn read_optional_part(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    name: &str,
) -> crate::error::Result<Option<String>> {
    let mut xml = String::new();
    match archive.by_name(name) {
        Ok(mut entry) => {
            entry.read_to_string(&mut xml)?; // propagate read/decompression errors
            Ok(Some(xml))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None), // genuinely missing → omit
        Err(e) => Err(e.into()),                              // other zip errors fail the extract
    }
}

/// Parse the docx XML into structured blocks, image reference definitions,
/// and the document extras (comments, footnotes, endnotes).
///
/// When `opts.images` is set, image relationships are loaded and `<w:drawing>`
/// elements are extracted as `Block::Image` entries with inline references.
/// The returned [`DocxOutput`] tuple is `(blocks, image_defs, comments,
/// footnotes, endnotes)`: `image_defs` holds the reference definitions to
/// append at the end of the document; `comments`/`footnotes`/`endnotes`
/// hold the extras parsed from the optional parts (empty when a part is
/// missing).
fn parse_docx(data: &[u8], opts: ExtractOptions) -> crate::error::Result<DocxOutput> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    // Load hyperlink relationships (rId → URL)
    let rels = xml_util::load_rels(&mut archive, "word/_rels/document.xml.rels");

    // Optionally load image relationships (needed for images and/or OCR)
    let image_rels = if opts.images || opts.ocr {
        xml_util::load_image_rels(&mut archive, "word/_rels/document.xml.rels")
    } else {
        xml_util::Rels::new()
    };

    // Optional extras parts: missing parts parse to empty vectors, keeping
    // docs without comments/footnotes/endnotes identical to before.
    let comments = read_optional_part(&mut archive, "word/comments.xml")?
        .map(|xml| parse_comments_xml(&xml))
        .unwrap_or_default();
    let footnotes = read_optional_part(&mut archive, "word/footnotes.xml")?
        .map(|xml| parse_footnotes_xml(&xml))
        .unwrap_or_default();
    let endnotes = read_optional_part(&mut archive, "word/endnotes.xml")?
        .map(|xml| parse_endnotes_xml(&xml))
        .unwrap_or_default();

    let mut xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut xml)?;

    let mut reader = Reader::from_str(&xml);
    let mut blocks = Vec::new();

    // Seed the note indexes with the parsed definitions so body references
    // emit markers exactly for defined ids (dangling refs stay silent).
    // Notes whose blocks render empty are NOT seeded: their definitions
    // would show nothing in the trailer, so a body reference would leave a
    // dangling `[^n]` marker and an empty `## Footnotes` section. Skipped
    // ids leave the display-index counter untouched (it starts at 1 and
    // only advances for markers actually emitted), so later ids still
    // number from 1.
    let mut footnotes_idx = NoteIndex::new();
    let mut endnotes_idx = NoteIndex::new();
    for note in &footnotes {
        if blocks_have_text(&note.blocks) {
            footnotes_idx.add_defined(&note.id);
        }
    }
    for note in &endnotes {
        if blocks_have_text(&note.blocks) {
            endnotes_idx.add_defined(&note.id);
        }
    }
    parse_body(
        &mut reader,
        &mut blocks,
        &rels,
        &image_rels,
        &mut footnotes_idx,
        &mut endnotes_idx,
    );

    // Record the display indexes the body walk assigned to each definition;
    // notes never referenced keep 0.
    let footnotes = assign_display_indexes(footnotes, &footnotes_idx);
    let endnotes = assign_display_indexes(endnotes, &endnotes_idx);

    // Read image data from the archive when either feature is active
    let image_defs = if opts.images || opts.ocr {
        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor)?;
        resolve_images(&mut blocks, &mut archive, opts)?
    } else {
        Vec::new()
    };

    Ok((blocks, image_defs, comments, footnotes, endnotes))
}

/// Walk the XML and collect blocks from the document body.
///
/// The block children of `<w:body>` are collected via the shared
/// [`parse_block_children`] walker; anything outside the body is ignored.
fn parse_body(
    reader: &mut Reader<&[u8]>,
    blocks: &mut Vec<Block>,
    rels: &Rels,
    image_rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) {
    loop {
        match &reader.read_event() {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"body" => {
                blocks.extend(parse_block_children(
                    reader, b"body", rels, image_rels, footnotes, endnotes,
                ));
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse the block-level children (`w:p`, `w:tbl`) of the element opened at
/// the reader's current position, stopping at the matching end tag (`stop`)
/// or at EOF.
///
/// Shared by the document-body walker ([`parse_body`]) and the extra-part
/// parsers (comments, footnotes, endnotes), so definitions reuse the body's
/// paragraph/table machinery. Extra parts pass empty `Rels` and empty note
/// indexes: hyperlinks are not resolved inside definitions and nested note
/// references produce no markers there.
fn parse_block_children(
    reader: &mut Reader<&[u8]>,
    stop: &[u8],
    rels: &Rels,
    image_rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> Vec<Block> {
    let mut blocks = Vec::new();
    loop {
        match &reader.read_event() {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"p" => {
                    let mut para_blocks =
                        parse_paragraph(reader, rels, image_rels, footnotes, endnotes);
                    blocks.append(&mut para_blocks);
                }
                b"tbl" => {
                    blocks.push(parse_table(reader, rels, footnotes, endnotes));
                }
                _ => {}
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == stop => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    blocks
}

/// Consume the ENTIRE subtree of the element whose `Start` was just read,
/// without interpreting its contents.
///
/// Tracks nesting depth like [`parse_run_props`] does (start 1 for the
/// opened element, +1 per nested start, −1 per end, exit at 0): the skip
/// ends only when the opened element itself closes, so nested elements
/// reusing the same local name are swallowed as part of the subtree instead
/// of ending the skip early.
fn skip_element(reader: &mut Reader<&[u8]>) {
    let mut depth = 1u32;
    loop {
        match &reader.read_event() {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Shared parse for the footnote/endnote parts, which share a shape
/// (`<w:footnotes>`/`<w:endnotes>` root with `w:footnote`/`w:endnote` note
/// elements).
///
/// Rules:
/// - Separator (`w:id` `-1`) and continuation (`w:id` `0`) notes are skipped.
/// - Duplicate ids keep the first occurrence.
/// - Bodies parse with the shared block walker; ref glyphs
///   (`w:footnoteRef`/`w:endnoteRef`) are skipped by `parse_run` so they
///   never leak text, and nested note references produce no markers (the
///   indexes passed here are empty by design).
fn parse_notes_xml(xml: &str, note_tag: &[u8]) -> Vec<Note> {
    let mut reader = Reader::from_str(xml);
    let mut notes = Vec::new();
    let mut seen = HashSet::new();
    // Empty rels and empty note indexes: no hyperlink resolution and no
    // nested note markers inside definitions.
    let empty_rels = Rels::new();
    let mut footnotes = NoteIndex::new();
    let mut endnotes = NoteIndex::new();

    loop {
        match &reader.read_event() {
            Ok(Event::Start(e)) if e.local_name().as_ref() == note_tag => {
                let Some(id) = get_attr(e, b"w:id").or_else(|| get_attr(e, b"id")) else {
                    // Missing id: consume the element without interpreting it.
                    skip_element(&mut reader);
                    continue;
                };
                if id == "-1" || id == "0" || !seen.insert(id.clone()) {
                    // Separator (-1), continuation (0), or duplicate id
                    // (first wins): skip.
                    skip_element(&mut reader);
                    continue;
                }
                let blocks = parse_block_children(
                    &mut reader,
                    note_tag,
                    &empty_rels,
                    &empty_rels,
                    &mut footnotes,
                    &mut endnotes,
                );
                notes.push(Note {
                    id,
                    display_index: 0,
                    blocks,
                });
                // parse_block_children consumed through End(note_tag).
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    notes
}

/// Parse `word/footnotes.xml` into note definitions in document order.
fn parse_footnotes_xml(xml: &str) -> Vec<Note> {
    parse_notes_xml(xml, b"footnote")
}

/// Parse `word/endnotes.xml` into note definitions (same shape and rules as
/// [`parse_footnotes_xml`]).
fn parse_endnotes_xml(xml: &str) -> Vec<Note> {
    parse_notes_xml(xml, b"endnote")
}

/// Parse `word/comments.xml` into comments in document order.
///
/// Comments keep their `w:author` (missing/empty → "Anonymous"); comments
/// whose body has no non-whitespace text are dropped; duplicate ids keep the
/// first occurrence. The `w:annotationRef` glyph inside comment bodies is
/// skipped by `parse_run`.
fn parse_comments_xml(xml: &str) -> Vec<Comment> {
    let mut reader = Reader::from_str(xml);
    let mut comments = Vec::new();
    let mut seen = HashSet::new();
    let empty_rels = Rels::new();
    let mut footnotes = NoteIndex::new();
    let mut endnotes = NoteIndex::new();

    loop {
        match &reader.read_event() {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"comment" => {
                let Some(id) = get_attr(e, b"w:id").or_else(|| get_attr(e, b"id")) else {
                    // Comment without an id: nothing to key it by, drop it.
                    skip_element(&mut reader);
                    continue;
                };
                if !seen.insert(id.clone()) {
                    // Duplicate id: first wins; consume the body unparsed so
                    // it can never leak text or markers.
                    skip_element(&mut reader);
                    continue;
                }
                let mut author = get_attr(e, b"w:author")
                    .or_else(|| get_attr(e, b"author"))
                    .unwrap_or_default();
                if author.trim().is_empty() {
                    author = "Anonymous".to_string();
                }
                let blocks = parse_block_children(
                    &mut reader,
                    b"comment",
                    &empty_rels,
                    &empty_rels,
                    &mut footnotes,
                    &mut endnotes,
                );
                if blocks_have_text(&blocks) {
                    comments.push(Comment { id, author, blocks });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    comments
}

/// Parse a `<w:p>` element into blocks.
///
/// Normally returns a single `Block::Paragraph`, but when image extraction
/// is active, any `<w:drawing>` elements (which live inside `<w:r>` runs)
/// produce additional `Block::Image` entries. The paragraph is always first,
/// followed by any images found.
fn parse_paragraph(
    reader: &mut Reader<&[u8]>,
    rels: &Rels,
    image_rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> Vec<Block> {
    let mut style = ParaStyle::default();
    let mut runs: Vec<Run> = Vec::new();
    let mut image_blocks: Vec<Block> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"pPr" => parse_para_props(reader, &mut style),
                    b"r" => {
                        let (run_runs, img_opt) =
                            parse_run(reader, image_rels, footnotes, endnotes);
                        runs.extend(run_runs);
                        if let Some(blk) = img_opt {
                            image_blocks.push(blk);
                        }
                    }
                    b"hyperlink" => {
                        // Resolve the hyperlink URL from r:id → rels map
                        let url = get_attr(e, b"r:id").and_then(|rid| rels.get(&rid).cloned());
                        parse_hyperlink_runs(
                            reader,
                            &mut runs,
                            url.as_deref(),
                            image_rels,
                            &mut image_blocks,
                            footnotes,
                            endnotes,
                        );
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"p" => {
                break;
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tab" {
                    runs.push(Run {
                        text: "\t".into(),
                        bold: false,
                        italic: false,
                        link_url: None,
                        marker: None,
                    });
                } else if name.as_ref() == b"br" {
                    runs.push(Run {
                        text: "\n".into(),
                        bold: false,
                        italic: false,
                        link_url: None,
                        marker: None,
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let mut result = vec![Block::Paragraph { style, runs }];
    result.append(&mut image_blocks);
    result
}

/// Parse `<w:pPr>` to extract heading level and list info.
fn parse_para_props(reader: &mut Reader<&[u8]>, style: &mut ParaStyle) {
    let mut depth = 1u32;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                depth += 1;
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"pStyle" => {
                        if let Some(val) = get_val_attr(e) {
                            if let Some(level) = parse_heading_level(&val) {
                                style.heading_level = level;
                            }
                        }
                    }
                    b"ilvl" => {
                        if let Some(val) = get_val_attr(e) {
                            if let Ok(n) = val.parse::<u8>() {
                                style.list_level = Some(n);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a heading style value like "Heading1" -> Some(1), "Title" -> Some(1).
fn parse_heading_level(val: &str) -> Option<u8> {
    // Standard: "Heading1" through "Heading9"
    if let Some(rest) = val.strip_prefix("Heading") {
        return rest.parse().ok();
    }
    // Also match lowercase variants like "heading 1", "heading1"
    let lower = val.to_lowercase();
    if let Some(rest) = lower.strip_prefix("heading") {
        let rest = rest.trim();
        return rest.parse().ok();
    }
    if lower == "title" {
        return Some(1);
    }
    if lower == "subtitle" {
        return Some(2);
    }
    None
}

/// Emit a marker `Run` for a `w:footnoteReference`/`w:endnoteReference` id.
///
/// The display index is assigned by the matching [`NoteIndex`] on first
/// sight; references to ids with no definition are dropped (no marker).
fn push_note_marker(
    runs: &mut Vec<Run>,
    name: &[u8],
    id: &str,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) {
    let marker = if name == b"footnoteReference" {
        footnotes.marker_for(id).map(NoteMarker::Footnote)
    } else {
        endnotes.marker_for(id).map(NoteMarker::Endnote)
    };
    if let Some(marker) = marker {
        let run = match marker {
            NoteMarker::Footnote(n) => Run::footnote_ref(n),
            NoteMarker::Endnote(n) => Run::endnote_ref(n),
        };
        runs.push(run);
    }
}

/// Parse a `<w:r>` element into text/marker `Run`s and/or an image `Block`.
///
/// A run may contain text, a drawing (image), a footnote/endnote reference
/// (`w:footnoteReference`/`w:endnoteReference`), or several of these. Note
/// references become marker runs carrying their display index. Text runs are
/// emitted in document order, split around marker runs so that text before a
/// reference stays before its marker. Ref glyphs (`w:footnoteRef`,
/// `w:endnoteRef`, `w:annotationRef` — the no-id markers in note/comment
/// bodies) are skipped entirely. When `image_rels` is non-empty and a
/// `<w:drawing>` is found inside the run, the image reference is extracted
/// and returned as a `Block::Image`.
fn parse_run(
    reader: &mut Reader<&[u8]>,
    image_rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> (Vec<Run>, Option<Block>) {
    let mut bold = false;
    let mut italic = false;
    let mut text = String::new();
    let mut image_block: Option<Block> = None;
    let mut runs: Vec<Run> = Vec::new();

    // Emit any accumulated text as a run before a marker run, so runs keep
    // their document order even when a run mixes text and references.
    let flush_text = |runs: &mut Vec<Run>, text: &mut String, bold: bool, italic: bool| {
        if !text.is_empty() {
            let mut run = Run::text(std::mem::take(text));
            run.bold = bold;
            run.italic = italic;
            runs.push(run);
        }
    };

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"rPr" => parse_run_props(reader, &mut bold, &mut italic),
                    b"t" => {
                        // Read text content
                        if let Ok(Event::Text(t)) = reader.read_event() {
                            if let Ok(s) = t.unescape() {
                                text.push_str(&s);
                            }
                        }
                        // Note: the </w:t> end tag will be consumed below
                    }
                    b"drawing" if !image_rels.is_empty() => {
                        if let Some(blk) = parse_drawing(reader, image_rels) {
                            image_block = Some(blk);
                        }
                    }
                    // Defensive Start form: note references are empty
                    // elements in practice, but consume through the end tag
                    // if a non-empty one appears.
                    b"footnoteReference" | b"endnoteReference" => {
                        if let Some(id) = get_attr(e, b"w:id").or_else(|| get_attr(e, b"id")) {
                            flush_text(&mut runs, &mut text, bold, italic);
                            push_note_marker(&mut runs, name.as_ref(), &id, footnotes, endnotes);
                        }
                        loop {
                            match reader.read_event() {
                                Ok(Event::End(ref e))
                                    if e.local_name().as_ref() == name.as_ref() =>
                                {
                                    break;
                                }
                                Ok(Event::Eof) | Err(_) => break,
                                _ => {}
                            }
                        }
                    }
                    // Ref glyphs (`w:footnoteRef`, `w:endnoteRef`,
                    // `w:annotationRef`): empty elements without an id that
                    // mark a reference site inside the note/comment bodies.
                    // Skipped (also in the Empty arm below) so they never
                    // leak text or produce marker runs. The Empty form needs
                    // no consumption; this defensive Start form swallows the
                    // whole subtree if a non-empty one ever appears.
                    b"footnoteRef" | b"endnoteRef" | b"annotationRef" => {
                        skip_element(reader);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tab" {
                    text.push('\t');
                } else if name.as_ref() == b"br" {
                    text.push('\n');
                } else if name.as_ref() == b"footnoteReference"
                    || name.as_ref() == b"endnoteReference"
                {
                    if let Some(id) = get_attr(e, b"w:id").or_else(|| get_attr(e, b"id")) {
                        flush_text(&mut runs, &mut text, bold, italic);
                        push_note_marker(&mut runs, name.as_ref(), &id, footnotes, endnotes);
                    }
                } else if name.as_ref() == b"b" || name.as_ref() == b"bCs" {
                    // Self-closing <w:b/> in rPr means bold on
                    bold = true;
                } else if name.as_ref() == b"i" || name.as_ref() == b"iCs" {
                    italic = true;
                } else if name.as_ref() == b"footnoteRef"
                    || name.as_ref() == b"endnoteRef"
                    || name.as_ref() == b"annotationRef"
                {
                    // Ref glyph (empty form): skipped; never text or marker.
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"r" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    flush_text(&mut runs, &mut text, bold, italic);

    (runs, image_block)
}

/// Parse <w:rPr> to extract bold/italic.
fn parse_run_props(reader: &mut Reader<&[u8]>, bold: &mut bool, italic: &mut bool) {
    let mut depth = 1u32;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                depth += 1;
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"b" | b"bCs" => {
                        // Check for val="false" or val="0"
                        let val = get_val_attr(e);
                        *bold = !matches!(val.as_deref(), Some("false" | "0"));
                    }
                    b"i" | b"iCs" => {
                        let val = get_val_attr(e);
                        *italic = !matches!(val.as_deref(), Some("false" | "0"));
                    }
                    _ => {}
                }
            }
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse runs inside a `<w:hyperlink>` element, tagging each run with the URL.
fn parse_hyperlink_runs(
    reader: &mut Reader<&[u8]>,
    runs: &mut Vec<Run>,
    url: Option<&str>,
    image_rels: &Rels,
    image_blocks: &mut Vec<Block>,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"r" => {
                let (run_runs, img_opt) = parse_run(reader, image_rels, footnotes, endnotes);
                for mut run in run_runs {
                    run.link_url = url.map(String::from);
                    runs.push(run);
                }
                if let Some(blk) = img_opt {
                    image_blocks.push(blk);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"hyperlink" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a `<w:drawing>` element to find an embedded image reference.
///
/// Walks into `<wp:inline>` or `<wp:anchor>` → `<a:graphic>` →
/// `<a:graphicData>` → `<pic:blipFill>` → `<a:blip r:embed="rIdN"/>`.
/// Returns a `Block::Image` with the image's ZIP path (to be resolved later)
/// stored in the `path` field.
fn parse_drawing(reader: &mut Reader<&[u8]>, image_rels: &Rels) -> Option<Block> {
    let mut embed_rid: Option<String> = None;
    let mut depth = 1u32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                if e.local_name().as_ref() == b"blip" {
                    if let Some(rid) = get_attr(e, b"r:embed") {
                        embed_rid = Some(rid);
                    }
                }
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"blip" => {
                if let Some(rid) = get_attr(e, b"r:embed") {
                    embed_rid = Some(rid);
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"drawing" {
                    break;
                }
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let rid = embed_rid?;
    let target = image_rels.get(&rid)?;

    // Store the resolved ZIP path; resolve_images() will read the data and
    // fill in the markdown reference and/or OCR text as configured.
    let zip_path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("word/{target}")
    };

    Some(Block::Image {
        path: zip_path,
        markdown: None,
        ocr_text: None,
    })
}

/// Resolve `Block::Image` entries by reading image data from the ZIP archive.
///
/// With `images`, produces reference-style base64 definitions. With `ocr`,
/// runs OCR on the image bytes. Blocks that end up with neither a markdown
/// reference nor OCR text (EMF/WMF, unreadable) are removed.
fn resolve_images(
    blocks: &mut Vec<Block>,
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    opts: ExtractOptions,
) -> crate::error::Result<Vec<String>> {
    let mut definitions = Vec::new();
    let mut counter = 0usize;

    for block in blocks.iter_mut() {
        if let Block::Image {
            path,
            markdown,
            ocr_text,
        } = block
        {
            if let Some(data) = xml_util::read_image_from_zip(archive, path, "") {
                if opts.images {
                    counter += 1;
                    let id = format!("image{counter}");
                    if let Some(img_ref) = crate::markup::image_to_base64_ref(&data, &id) {
                        *markdown = Some(img_ref.inline);
                        definitions.push(img_ref.definition);
                    }
                }
                if opts.ocr {
                    *ocr_text = crate::ocr::ocr_image_bytes(&data)?;
                }
            }
        }
    }
    // Remove image blocks with nothing to render (unsupported or unreadable)
    blocks.retain(|b| {
        !matches!(
            b,
            Block::Image {
                markdown: None,
                ocr_text: None,
                ..
            }
        )
    });

    Ok(definitions)
}

/// Parse a `<w:tbl>` element into a `Block::Table`.
fn parse_table(
    reader: &mut Reader<&[u8]>,
    rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> Block {
    let mut rows: Vec<Row> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tr" {
                    let row = parse_table_row(reader, rels, footnotes, endnotes);
                    rows.push(row);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tbl" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Block::Table { rows }
}

/// Parse a `<w:tr>` element into a row of cells.
fn parse_table_row(
    reader: &mut Reader<&[u8]>,
    rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> Row {
    let mut cells: Row = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tc" {
                    let cell = parse_table_cell(reader, rels, footnotes, endnotes);
                    cells.push(cell);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tr" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    cells
}

/// Parse a `<w:tc>` element into a list of blocks.
///
/// Images inside table cells are not extracted (impractical in markdown
/// tables), so an empty `image_rels` is used for paragraph parsing.
fn parse_table_cell(
    reader: &mut Reader<&[u8]>,
    rels: &Rels,
    footnotes: &mut NoteIndex,
    endnotes: &mut NoteIndex,
) -> Cell {
    let empty_image_rels = xml_util::Rels::new();
    let mut blocks = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"p" => {
                        let mut para_blocks =
                            parse_paragraph(reader, rels, &empty_image_rels, footnotes, endnotes);
                        blocks.append(&mut para_blocks);
                    }
                    b"tbl" => {
                        blocks.push(parse_table(reader, rels, footnotes, endnotes));
                        // nested table
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tc" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    blocks
}

/// Get the `w:val` (or `val`) attribute value from an XML element.
fn get_val_attr(e: &quick_xml::events::BytesStart) -> Option<String> {
    get_attr(e, b"w:val").or_else(|| get_attr(e, b"val"))
}

/// Plain-text rendering of a single run: the run's text, or the marker
/// label (`[n]` footnote / `[eN]` endnote) for marker runs.
fn run_plain_text(r: &Run) -> std::borrow::Cow<'_, str> {
    match r.marker {
        Some(NoteMarker::Footnote(n)) => std::borrow::Cow::Owned(format!("[{n}]")),
        Some(NoteMarker::Endnote(n)) => std::borrow::Cow::Owned(format!("[e{n}]")),
        None => std::borrow::Cow::Borrowed(&r.text),
    }
}

/// Extract text content from a cell's blocks, joining paragraphs with spaces.
fn cell_to_text(cell: &[Block], use_markdown: bool) -> String {
    cell.iter()
        .filter_map(|b| match b {
            Block::Paragraph { runs, .. } => {
                let t = if use_markdown {
                    render_runs_markdown(runs)
                } else {
                    runs.iter().map(run_plain_text).collect::<String>()
                };
                let t = t.trim().to_string();
                if t.is_empty() {
                    None
                } else {
                    Some(t)
                }
            }
            Block::Table { .. } | Block::Image { .. } => None,
        })
        .collect::<Vec<_>>()
        .join(" ")
}

// ── Plain text rendering ──────────────────────────────────────────

fn render_plain(blocks: &[Block]) -> String {
    let mut out = String::new();
    let mut first = true;

    for block in blocks {
        render_block_plain(block, &mut out, &mut first);
    }

    out
}

fn render_block_plain(block: &Block, out: &mut String, first: &mut bool) {
    match block {
        Block::Paragraph { runs, .. } => {
            let text: String = runs.iter().map(run_plain_text).collect();
            let text = text.trim_end();
            if !text.is_empty() {
                if !*first {
                    out.push('\n');
                }
                out.push_str(text);
                out.push('\n');
                *first = false;
            }
        }
        Block::Table { rows } => {
            for row in rows {
                let cells: Vec<String> = row.iter().map(|cell| cell_to_text(cell, false)).collect();

                let line = cells.join("\t");
                let line = line.trim_end();
                if !line.is_empty() {
                    if !*first {
                        out.push('\n');
                    }
                    out.push_str(line);
                    out.push('\n');
                    *first = false;
                }
            }
        }
        Block::Image {
            ocr_text: Some(text),
            ..
        } => {
            crate::markup::push_ocr_plain(out, text, first);
        }
        Block::Image { .. } => {
            // Images without OCR text are not rendered in plain text mode
        }
    }
}

// ── Markdown rendering ────────────────────────────────────────────

fn render_markdown(blocks: &[Block]) -> String {
    let mut out = String::new();

    for block in blocks {
        render_block_markdown(block, &mut out);
    }

    out
}

fn render_block_markdown(block: &Block, out: &mut String) {
    match block {
        Block::Paragraph { style, runs } => {
            let text = render_runs_markdown(runs);
            let text = text.trim_end();
            if text.is_empty() {
                return;
            }

            if style.heading_level > 0 && style.heading_level <= 6 {
                for _ in 0..style.heading_level {
                    out.push('#');
                }
                out.push(' ');
                out.push_str(text);
                out.push_str("\n\n");
            } else if let Some(level) = style.list_level {
                let indent = "  ".repeat(usize::from(level));
                out.push_str(&indent);
                out.push_str("- ");
                out.push_str(text);
                out.push('\n');
            } else {
                out.push_str(text);
                out.push_str("\n\n");
            }
        }
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
        Block::Table { rows } => {
            if rows.is_empty() {
                return;
            }

            let ncols = rows.iter().map(Vec::len).max().unwrap_or(0);
            if ncols == 0 {
                return;
            }

            let mut md_rows: Vec<Vec<String>> = Vec::new();
            for row in rows {
                let mut md_row = Vec::new();
                for cell in row {
                    let cell_text = cell_to_text(cell, true);
                    md_row.push(cell_text.replace('|', "\\|"));
                }
                while md_row.len() < ncols {
                    md_row.push(String::new());
                }
                md_rows.push(md_row);
            }

            if let Some(header) = md_rows.first() {
                out.push_str("| ");
                out.push_str(&header.join(" | "));
                out.push_str(" |\n");

                out.push('|');
                for _ in 0..ncols {
                    out.push_str(" --- |");
                }
                out.push('\n');

                for row in md_rows.iter().skip(1) {
                    out.push_str("| ");
                    out.push_str(&row.join(" | "));
                    out.push_str(" |\n");
                }
                out.push('\n');
            }
        }
    }
}

/// Render runs with markdown inline formatting (bold/italic/hyperlinks)
/// and footnote/endnote markers.
///
/// Marker runs are emitted directly as `[^n]` / `[^eN]` labels so they
/// are never wrapped in formatting markers or absorbed by hyperlink
/// grouping. Consecutive ordinary runs are delegated to the shared
/// markup renderer, which groups adjacent runs sharing the same
/// `link_url` so the markdown link wraps the entire visible text:
/// `[text](url)` instead of producing separate
/// `[part1](url)[part2](url)` fragments. Markers therefore act as
/// boundaries between hyperlink groups, but never alter the rendering
/// of ordinary runs.
fn render_runs_markdown(runs: &[Run]) -> String {
    let mut out = String::new();
    let mut i = 0;

    while i < runs.len() {
        match runs[i].marker {
            Some(NoteMarker::Footnote(n)) => {
                write!(out, "[^{n}]").expect("writing to a String cannot fail");
                i += 1;
            }
            Some(NoteMarker::Endnote(n)) => {
                write!(out, "[^e{n}]").expect("writing to a String cannot fail");
                i += 1;
            }
            None => {
                let start = i;
                while i < runs.len() && runs[i].marker.is_none() {
                    i += 1;
                }
                out.push_str(&markup::render_runs_markdown(&runs[start..i]));
            }
        }
    }

    out
}

/// Implement [`InlineRun`] for docx `Run` so the shared markup renderer
/// can inspect formatting without knowing the concrete type.
///
/// Marker runs expose empty text: their `[^n]` / `[^eN]` labels are
/// emitted by the docx-local `render_runs_markdown` (which only hands
/// contiguous ordinary-run slices to the markup renderer), so an empty
/// `text()` here keeps markers inert if one ever reaches the shared
/// renderer (e.g. inside a hyperlink slice).
impl markup::InlineRun for Run {
    fn text(&self) -> &str {
        match self.marker {
            Some(_) => "",
            None => &self.text,
        }
    }
    fn bold(&self) -> bool {
        self.bold
    }
    fn italic(&self) -> bool {
        self.italic
    }
    fn link_url(&self) -> Option<&str> {
        self.link_url.as_deref()
    }
}

// ── Extras trailers (comments / footnotes / endnotes) ───────────

/// Ensure the output ends with a blank line (two newlines) so a following
/// trailer section or group heading is separated from what precedes it.
/// No-op on empty output or when a blank line is already present.
fn ensure_trailer_blank_line(out: &mut String) {
    if !out.is_empty() && !out.ends_with("\n\n") {
        if out.ends_with('\n') {
            out.push('\n');
        } else {
            out.push_str("\n\n");
        }
    }
}

/// Append the extras trailers to a rendered markdown body in the fixed order
/// Comments, Footnotes, Endnotes. Empty sections are omitted entirely; a
/// single blank line separates the body from the first trailer.
fn append_extras_markdown(
    out: &mut String,
    comments: &[Comment],
    footnotes: &[Note],
    endnotes: &[Note],
) {
    append_comments_markdown(out, comments);
    append_footnotes_markdown(out, footnotes);
    append_endnotes_markdown(out, endnotes);
}

/// Append a `## Comments` section: consecutive same-author comments share one
/// `### Author` heading; a non-consecutive repeat of an author gets the
/// heading again, preserving comment order.
fn append_comments_markdown(out: &mut String, comments: &[Comment]) {
    if comments.is_empty() {
        return;
    }
    ensure_trailer_blank_line(out);
    out.push_str("## Comments\n\n");
    let mut i = 0;
    while i < comments.len() {
        let author = comments[i].author.as_str();
        out.push_str("### ");
        out.push_str(author);
        out.push_str("\n\n");
        let mut j = i;
        while j < comments.len() && comments[j].author == author {
            for block in &comments[j].blocks {
                render_block_markdown(block, out);
            }
            // Blank line between comments; also gives the next heading its
            // separation when the last block is a list item ending in '\n'.
            ensure_trailer_blank_line(out);
            j += 1;
        }
        i = j;
    }
}

/// The referenced notes of one kind, in ascending display-index order.
fn referenced_notes(notes: &[Note]) -> Vec<&Note> {
    let mut referenced: Vec<&Note> = notes.iter().filter(|n| n.display_index > 0).collect();
    referenced.sort_by_key(|n| n.display_index);
    referenced
}

/// Append a `## Footnotes` section with one definition per referenced note.
fn append_footnotes_markdown(out: &mut String, footnotes: &[Note]) {
    let notes = referenced_notes(footnotes);
    if notes.is_empty() {
        return;
    }
    ensure_trailer_blank_line(out);
    out.push_str("## Footnotes\n\n");
    for note in notes {
        render_note_definition_markdown(out, "", note);
    }
}

/// Append a `## Endnotes` section with one definition per referenced note.
fn append_endnotes_markdown(out: &mut String, endnotes: &[Note]) {
    let notes = referenced_notes(endnotes);
    if notes.is_empty() {
        return;
    }
    ensure_trailer_blank_line(out);
    out.push_str("## Endnotes\n\n");
    for note in notes {
        render_note_definition_markdown(out, "e", note);
    }
}

/// Render one note's markdown definition: `[^{prefix}{n}]: text` for a
/// note whose sole rendered block is a single paragraph (`prefix` is `""`
/// for footnotes and `"e"` for endnotes), or a definition-list form whose
/// continuation lines are indented 4 spaces (`CommonMark`) for every other
/// shape — multi-block bodies, or a single table/image (a `[^n]: | row |`
/// line would let the table's following lines escape the definition).
/// Blocks whose markdown renders empty are skipped; a note with nothing to
/// show emits nothing.
fn render_note_definition_markdown(out: &mut String, prefix: &str, note: &Note) {
    let mut body = String::new();
    let mut rendered = 0usize;
    let mut single_is_paragraph = false;
    for block in &note.blocks {
        let mut piece = String::new();
        render_block_markdown(block, &mut piece);
        if !piece.trim_end().is_empty() {
            rendered += 1;
            single_is_paragraph = matches!(block, Block::Paragraph { .. });
            body.push_str(&piece);
        }
    }
    if rendered == 0 {
        return;
    }
    out.push_str("[^");
    out.push_str(prefix);
    let _ = write!(out, "{}", note.display_index);
    out.push(']');
    if rendered == 1 && single_is_paragraph {
        // Single paragraph: inline label on the definition line.
        out.push_str(": ");
        out.push_str(body.trim_end());
        out.push_str("\n\n");
        return;
    }
    // Multi-block: indent every line of the body 4 spaces, leaving blank
    // lines between blocks empty.
    out.push_str(":\n");
    for line in body.trim_end().lines() {
        if !line.is_empty() {
            out.push_str("    ");
            out.push_str(line);
        }
        out.push('\n');
    }
    out.push('\n');
}

/// Append the extras trailers to a rendered plain-text body.
fn append_extras_plain(
    out: &mut String,
    comments: &[Comment],
    footnotes: &[Note],
    endnotes: &[Note],
) {
    append_comments_plain(out, comments);
    append_footnotes_plain(out, footnotes);
    append_endnotes_plain(out, endnotes);
}

/// Append a `--- Comments ---` section: per group of consecutive same-author
/// comments, a `[Author]` line followed by the comment paragraph lines, with
/// a blank line between groups.
fn append_comments_plain(out: &mut String, comments: &[Comment]) {
    if comments.is_empty() {
        return;
    }
    ensure_trailer_blank_line(out);
    out.push_str("--- Comments ---\n");
    let mut i = 0;
    while i < comments.len() {
        let author = comments[i].author.as_str();
        let _ = writeln!(out, "[{author}]");
        // Shared `first` flag joins adjacent paragraphs/comments the way the
        // plain body renderer does (single newline between lines).
        let mut first = true;
        let mut j = i;
        while j < comments.len() && comments[j].author == author {
            for block in &comments[j].blocks {
                render_block_plain(block, out, &mut first);
            }
            j += 1;
        }
        out.push('\n'); // blank line between groups
        i = j;
    }
}

/// Append a `--- Footnotes ---` section with one entry per referenced note.
fn append_footnotes_plain(out: &mut String, footnotes: &[Note]) {
    let notes = referenced_notes(footnotes);
    if notes.is_empty() {
        return;
    }
    ensure_trailer_blank_line(out);
    out.push_str("--- Footnotes ---\n");
    for note in notes {
        render_note_plain(out, "", note);
    }
}

/// Append a `--- Endnotes ---` section with one entry per referenced note.
fn append_endnotes_plain(out: &mut String, endnotes: &[Note]) {
    let notes = referenced_notes(endnotes);
    if notes.is_empty() {
        return;
    }
    ensure_trailer_blank_line(out);
    out.push_str("--- Endnotes ---\n");
    for note in notes {
        render_note_plain(out, "e", note);
    }
}

/// Render one note's plain-text entry: the first line carries the
/// `[{prefix}{n}]` bracket, continuation paragraphs follow on their own
/// lines without a repeated bracket. A note whose blocks all render empty
/// emits nothing. A blank line separates consecutive entries.
fn render_note_plain(out: &mut String, prefix: &str, note: &Note) {
    let mut body = String::new();
    let mut first = true;
    for block in &note.blocks {
        render_block_plain(block, &mut body, &mut first);
    }
    let body = body.trim_end();
    if body.is_empty() {
        return;
    }
    let mut lines = body.lines();
    if let Some(head) = lines.next() {
        out.push('[');
        out.push_str(prefix);
        let _ = write!(out, "{}", note.display_index);
        out.push_str("] ");
        out.push_str(head);
        out.push('\n');
    }
    for line in lines {
        out.push_str(line);
        out.push('\n');
    }
    out.push('\n');
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── parse_heading_level ──────────────────────────────────────

    #[test]
    fn heading_standard() {
        assert_eq!(parse_heading_level("Heading1"), Some(1));
        assert_eq!(parse_heading_level("Heading3"), Some(3));
        assert_eq!(parse_heading_level("Heading9"), Some(9));
    }

    #[test]
    fn heading_lowercase() {
        assert_eq!(parse_heading_level("heading2"), Some(2));
        assert_eq!(parse_heading_level("heading 4"), Some(4));
    }

    #[test]
    fn heading_title_subtitle() {
        assert_eq!(parse_heading_level("Title"), Some(1));
        assert_eq!(parse_heading_level("title"), Some(1));
        assert_eq!(parse_heading_level("Subtitle"), Some(2));
        assert_eq!(parse_heading_level("subtitle"), Some(2));
    }

    #[test]
    fn heading_not_heading() {
        assert_eq!(parse_heading_level("Normal"), None);
        assert_eq!(parse_heading_level("ListParagraph"), None);
        assert_eq!(parse_heading_level("BodyText"), None);
    }

    #[test]
    fn heading_invalid_number() {
        assert_eq!(parse_heading_level("HeadingX"), None);
        assert_eq!(parse_heading_level("Heading"), None);
    }

    // ── render_runs_markdown ─────────────────────────────────────

    /// Helper to create a plain Run without a hyperlink.
    fn run(text: &str, bold: bool, italic: bool) -> Run {
        Run {
            bold,
            italic,
            ..Run::text(text)
        }
    }

    #[test]
    fn runs_plain_text() {
        let runs = vec![run("Hello", false, false)];
        assert_eq!(render_runs_markdown(&runs), "Hello");
    }

    #[test]
    fn runs_bold() {
        let runs = vec![run("Bold", true, false)];
        assert_eq!(render_runs_markdown(&runs), "**Bold**");
    }

    #[test]
    fn runs_italic() {
        let runs = vec![run("Italic", false, true)];
        assert_eq!(render_runs_markdown(&runs), "*Italic*");
    }

    #[test]
    fn runs_bold_italic() {
        let runs = vec![run("Both", true, true)];
        assert_eq!(render_runs_markdown(&runs), "***Both***");
    }

    #[test]
    fn runs_mixed() {
        let runs = vec![
            run("Normal ", false, false),
            run("bold", true, false),
            run(" end", false, false),
        ];
        assert_eq!(render_runs_markdown(&runs), "Normal **bold** end");
    }

    #[test]
    fn runs_whitespace_only_not_formatted() {
        let runs = vec![run("   ", true, true)];
        // Whitespace-only runs should not be wrapped in formatting markers
        assert_eq!(render_runs_markdown(&runs), "   ");
    }

    #[test]
    fn runs_empty() {
        let runs: Vec<Run> = vec![];
        assert_eq!(render_runs_markdown(&runs), "");
    }

    // ── hyperlink rendering ──────────────────────────────────────

    #[test]
    fn runs_hyperlink_basic() {
        let runs = vec![Run {
            link_url: Some("https://example.com".into()),
            ..Run::text("click here")
        }];
        assert_eq!(
            render_runs_markdown(&runs),
            "[click here](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_bold() {
        let runs = vec![Run {
            bold: true,
            link_url: Some("https://example.com".into()),
            ..Run::text("bold link")
        }];
        assert_eq!(
            render_runs_markdown(&runs),
            "[**bold link**](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_multiple_runs_grouped() {
        // Two runs with the same URL should be grouped into one markdown link
        let runs = vec![
            Run {
                link_url: Some("https://example.com".into()),
                ..Run::text("part ")
            },
            Run {
                bold: true,
                link_url: Some("https://example.com".into()),
                ..Run::text("one")
            },
        ];
        assert_eq!(
            render_runs_markdown(&runs),
            "[part **one**](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_mixed_with_plain() {
        let runs = vec![
            run("See ", false, false),
            Run {
                link_url: Some("https://example.com".into()),
                ..Run::text("this link")
            },
            run(" for details", false, false),
        ];
        assert_eq!(
            render_runs_markdown(&runs),
            "See [this link](https://example.com) for details"
        );
    }

    // ── marker runs (footnote / endnote) ─────────────────────────

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
        assert_eq!(render_markdown(&blocks).trim_end(), "Hello[^1] world");
    }

    #[test]
    fn render_plain_endnote_marker() {
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![Run::text("See"), Run::endnote_ref(2)],
        }];
        assert_eq!(render_plain(&blocks).trim_end(), "See[e2]");
    }

    #[test]
    fn render_markdown_marker_between_plain_runs() {
        // A marker between plain (unlinked) runs is emitted inline.
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![Run::text("a"), Run::footnote_ref(3), Run::text("b")],
        }];
        assert_eq!(render_markdown(&blocks).trim_end(), "a[^3]b");
    }

    #[test]
    fn render_markdown_marker_splits_hyperlink_group() {
        // Markers act as boundaries between hyperlink groups: sibling runs
        // sharing a URL stay grouped, but a marker between them forces each
        // side into its own link instead of one `[xy](url)`.
        let make_link = |text: &str| Run {
            link_url: Some("https://e".into()),
            ..Run::text(text)
        };
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![make_link("x"), Run::footnote_ref(3), make_link("y")],
        }];
        assert_eq!(
            render_markdown(&blocks).trim_end(),
            "[x](https://e)[^3][y](https://e)"
        );
    }

    #[test]
    fn render_markdown_marker_only_paragraph() {
        // A paragraph consisting solely of a marker run still emits its label.
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![Run::footnote_ref(5)],
        }];
        assert_eq!(render_markdown(&blocks).trim_end(), "[^5]");
    }

    #[test]
    fn render_markdown_mixed_footnote_endnote_markers() {
        // Footnote and endnote markers interleaved with text keep their own
        // distinct labels: `[^n]` vs `[^eN]`.
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![
                Run::text("a"),
                Run::footnote_ref(2),
                Run::endnote_ref(3),
                Run::text("b"),
            ],
        }];
        assert_eq!(render_markdown(&blocks).trim_end(), "a[^2][^e3]b");
    }

    #[test]
    fn render_markdown_hyperlink_grouping_preserved_with_marker() {
        // Sibling guard: adjacent runs sharing a URL still group into one
        // link after the marker-aware render_runs_markdown refactor.
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![
                Run {
                    link_url: Some("https://e".into()),
                    ..Run::text("x")
                },
                Run {
                    link_url: Some("https://e".into()),
                    ..Run::text("y")
                },
            ],
        }];
        assert_eq!(render_markdown(&blocks).trim_end(), "[xy](https://e)");
    }

    #[test]
    fn render_plain_marker_inside_table_cell() {
        // Marker labels flow through the cell_to_text plain path.
        let blocks = vec![Block::Table {
            rows: vec![vec![vec![Block::Paragraph {
                style: ParaStyle::default(),
                runs: vec![Run::footnote_ref(1)],
            }]]],
        }];
        assert!(render_plain(&blocks).contains("[1]"));
    }

    #[test]
    fn render_plain_tab_br_runs_unchanged() {
        // Guard: tab/br runs (parse sites now construct them with
        // marker: None) still render verbatim.
        let tab = Run {
            text: "\t".into(),
            bold: false,
            italic: false,
            link_url: None,
            marker: None,
        };
        let br = Run {
            text: "\n".into(),
            bold: false,
            italic: false,
            link_url: None,
            marker: None,
        };
        assert_eq!(run_plain_text(&tab), "\t");
        assert_eq!(run_plain_text(&br), "\n");
        let blocks = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![Run::text("a"), tab, br, Run::text("b")],
        }];
        assert_eq!(render_plain(&blocks), "a\t\nb\n");
    }

    // ── NoteIndex (footnote/endnote display index assignment) ────

    #[test]
    fn note_index_first_ref_assigns_and_reuses() {
        let mut idx = NoteIndex::new();
        idx.add_defined("1");
        assert_eq!(idx.marker_for("1"), Some(1));
        assert_eq!(idx.marker_for("1"), Some(1)); // reused, not re-assigned
        assert_eq!(idx.marker_for("2"), None); // undefined → no index
        assert_eq!(idx.marker_for("1"), Some(1)); // still assigned
        idx.add_defined("2");
        assert_eq!(idx.marker_for("2"), Some(2)); // next counter after 1
    }

    // ── parse_run footnote/endnote reference markers ─────────────

    /// Wrap an XML fragment in `<w:r>…</w:r>` and feed it to `parse_run`.
    fn parse_run_xml(
        xml: &str,
        footnotes: &mut NoteIndex,
        endnotes: &mut NoteIndex,
    ) -> (Vec<Run>, Option<Block>) {
        let full = format!("<w:r>{xml}</w:r>");
        let mut reader = Reader::from_str(&full);
        assert!(matches!(reader.read_event(), Ok(Event::Start(_)))); // consume <w:r>
        parse_run(&mut reader, &Rels::new(), footnotes, endnotes)
    }

    #[test]
    fn parse_run_footnote_reference_emits_marker_run() {
        let mut footnotes = NoteIndex::new();
        footnotes.add_defined("1");
        let mut endnotes = NoteIndex::new();

        // Defined id → marker run (display index 1) plus the text run.
        let (runs, img) = parse_run_xml(
            "<w:footnoteReference w:id=\"1\"/><w:t>text</w:t>",
            &mut footnotes,
            &mut endnotes,
        );
        assert!(img.is_none());
        assert_eq!(runs.len(), 2);
        assert!(
            runs.iter()
                .any(|r| r.marker == Some(NoteMarker::Footnote(1))),
            "missing footnote marker run: {runs:?}"
        );
        assert!(
            runs.iter().any(|r| r.text == "text" && r.marker.is_none()),
            "missing text run: {runs:?}"
        );

        // Dangling id (no definition) → no marker run, only the text run.
        let (runs, _) = parse_run_xml(
            "<w:footnoteReference w:id=\"99\"/><w:t>text</w:t>",
            &mut footnotes,
            &mut endnotes,
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "text");
    }

    #[test]
    fn parse_run_endnote_reference_emits_marker_run() {
        let mut footnotes = NoteIndex::new();
        let mut endnotes = NoteIndex::new();
        endnotes.add_defined("7");

        // Endnote ids keep a separate counter from footnotes.
        let (runs, _) = parse_run_xml(
            "<w:endnoteReference w:id=\"7\"/><w:t>tail</w:t>",
            &mut footnotes,
            &mut endnotes,
        );
        assert_eq!(runs.len(), 2);
        assert!(runs
            .iter()
            .any(|r| r.marker == Some(NoteMarker::Endnote(1))));
        assert!(runs.iter().any(|r| r.text == "tail" && r.marker.is_none()));
    }

    #[test]
    fn parse_run_footnote_ref_after_text_keeps_order() {
        let mut footnotes = NoteIndex::new();
        footnotes.add_defined("1");
        let mut endnotes = NoteIndex::new();

        // A reference inside the same <w:r> after text must render after it:
        // <w:r><w:t>x</w:t><w:footnoteReference w:id="1"/></w:r> → [x, [^1]].
        let (runs, img) = parse_run_xml(
            "<w:t>x</w:t><w:footnoteReference w:id=\"1\"/>",
            &mut footnotes,
            &mut endnotes,
        );
        assert!(img.is_none());
        assert_eq!(runs.len(), 2);
        assert!(
            runs[0].marker.is_none() && runs[0].text == "x",
            "first run should be text \"x\": {runs:?}"
        );
        assert!(
            runs[1].marker == Some(NoteMarker::Footnote(1)),
            "second run should be the footnote marker: {runs:?}"
        );

        // Interleaved <w:t>a</w:t><ref/><w:t>b</w:t> → [a-text, marker, b-text].
        let (runs, _) = parse_run_xml(
            "<w:t>a</w:t><w:footnoteReference w:id=\"1\"/><w:t>b</w:t>",
            &mut footnotes,
            &mut endnotes,
        );
        assert_eq!(runs.len(), 3);
        assert!(
            runs[0].marker.is_none() && runs[0].text == "a",
            "first run should be text \"a\": {runs:?}"
        );
        assert_eq!(runs[1].marker, Some(NoteMarker::Footnote(1)));
        assert!(
            runs[2].marker.is_none() && runs[2].text == "b",
            "last run should be text \"b\": {runs:?}"
        );
    }

    #[test]
    fn parse_paragraph_footnote_ref_in_opening() {
        let mut footnotes = NoteIndex::new();
        footnotes.add_defined("1");
        let mut endnotes = NoteIndex::new();

        let xml = "<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r><w:r><w:t>X</w:t></w:r></w:p>";
        let full = format!("<w:body>{xml}</w:body>");
        let mut reader = Reader::from_str(&full);
        assert!(matches!(reader.read_event(), Ok(Event::Start(_)))); // consume <w:body>
        let blocks = parse_paragraph(
            &mut reader,
            &Rels::new(),
            &Rels::new(),
            &mut footnotes,
            &mut endnotes,
        );

        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            Block::Paragraph { runs, .. } => {
                assert!(runs
                    .iter()
                    .any(|r| r.marker == Some(NoteMarker::Footnote(1))));
                assert!(runs.iter().any(|r| r.text == "X"));
            }
            _ => panic!("expected a paragraph block"),
        }
        // End-to-end: parse → render yields the marker label right before X.
        assert_eq!(render_markdown(&blocks).trim_end(), "[^1]X");
    }

    #[test]
    fn parse_paragraph_footnote_ref_inside_hyperlink() {
        let mut footnotes = NoteIndex::new();
        footnotes.add_defined("1");
        let mut endnotes = NoteIndex::new();

        let xml = "<w:p><w:hyperlink r:id=\"rId9\"><w:r><w:t>see</w:t></w:r>\
                   <w:r><w:footnoteReference w:id=\"1\"/></w:r></w:hyperlink></w:p>";
        let full = format!("<w:body>{xml}</w:body>");
        let mut reader = Reader::from_str(&full);
        assert!(matches!(reader.read_event(), Ok(Event::Start(_)))); // consume <w:body>
        let mut rels = Rels::new();
        rels.insert("rId9".into(), "https://example.com".into());
        let blocks = parse_paragraph(
            &mut reader,
            &rels,
            &Rels::new(),
            &mut footnotes,
            &mut endnotes,
        );

        match &blocks[0] {
            Block::Paragraph { runs, .. } => {
                assert!(runs
                    .iter()
                    .any(|r| r.marker == Some(NoteMarker::Footnote(1))));
                assert!(runs.iter().any(
                    |r| r.text == "see" && r.link_url.as_deref() == Some("https://example.com")
                ));
            }
            _ => panic!("expected a paragraph block"),
        }
        // The marker splits the hyperlink group and renders its own label.
        assert_eq!(
            render_markdown(&blocks).trim_end(),
            "[see](https://example.com)[^1]"
        );
    }

    #[test]
    fn parse_table_cell_footnote_ref() {
        let mut footnotes = NoteIndex::new();
        footnotes.add_defined("1");
        let mut endnotes = NoteIndex::new();

        let xml = "<w:tbl><w:tr><w:tc><w:p><w:r><w:footnoteReference w:id=\"1\"/>\
                   <w:t>note</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        let full = format!("<w:body>{xml}</w:body>");
        let mut reader = Reader::from_str(&full);
        assert!(matches!(reader.read_event(), Ok(Event::Start(_)))); // consume <w:body>
        let table = parse_table(&mut reader, &Rels::new(), &mut footnotes, &mut endnotes);

        assert!(matches!(&table, Block::Table { rows } if rows.len() == 1));
        let md = render_markdown(std::slice::from_ref(&table));
        assert!(md.contains("[^1]"), "table markdown missing marker: {md}");
    }

    // ── cell_to_text ─────────────────────────────────────────────

    #[test]
    fn cell_to_text_plain() {
        let cell = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![run("Hello", false, false)],
        }];
        assert_eq!(cell_to_text(&cell, false), "Hello");
    }

    #[test]
    fn cell_to_text_markdown_bold() {
        let cell = vec![Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![run("Bold", true, false)],
        }];
        assert_eq!(cell_to_text(&cell, true), "**Bold**");
    }

    #[test]
    fn cell_to_text_multiple_paragraphs() {
        let cell = vec![
            Block::Paragraph {
                style: ParaStyle::default(),
                runs: vec![run("First", false, false)],
            },
            Block::Paragraph {
                style: ParaStyle::default(),
                runs: vec![run("Second", false, false)],
            },
        ];
        assert_eq!(cell_to_text(&cell, false), "First Second");
    }

    #[test]
    fn cell_to_text_empty_paragraphs_skipped() {
        let cell = vec![
            Block::Paragraph {
                style: ParaStyle::default(),
                runs: vec![],
            },
            Block::Paragraph {
                style: ParaStyle::default(),
                runs: vec![run("Content", false, false)],
            },
        ];
        assert_eq!(cell_to_text(&cell, false), "Content");
    }

    // ── render_block_markdown (unit-level) ────────────────────────

    #[test]
    fn render_heading_paragraph() {
        let block = Block::Paragraph {
            style: ParaStyle {
                heading_level: 2,
                list_level: None,
            },
            runs: vec![run("My Heading", false, false)],
        };
        let mut out = String::new();
        render_block_markdown(&block, &mut out);
        assert_eq!(out, "## My Heading\n\n");
    }

    #[test]
    fn render_list_item() {
        let block = Block::Paragraph {
            style: ParaStyle {
                heading_level: 0,
                list_level: Some(0),
            },
            runs: vec![run("Item one", false, false)],
        };
        let mut out = String::new();
        render_block_markdown(&block, &mut out);
        assert_eq!(out, "- Item one\n");
    }

    #[test]
    fn render_nested_list_item() {
        let block = Block::Paragraph {
            style: ParaStyle {
                heading_level: 0,
                list_level: Some(2),
            },
            runs: vec![run("Nested", false, false)],
        };
        let mut out = String::new();
        render_block_markdown(&block, &mut out);
        assert_eq!(out, "    - Nested\n");
    }

    #[test]
    fn render_table_markdown() {
        let table = Block::Table {
            rows: vec![
                vec![
                    vec![Block::Paragraph {
                        style: ParaStyle::default(),
                        runs: vec![run("Name", false, false)],
                    }],
                    vec![Block::Paragraph {
                        style: ParaStyle::default(),
                        runs: vec![run("Age", false, false)],
                    }],
                ],
                vec![
                    vec![Block::Paragraph {
                        style: ParaStyle::default(),
                        runs: vec![run("Alice", false, false)],
                    }],
                    vec![Block::Paragraph {
                        style: ParaStyle::default(),
                        runs: vec![run("30", false, false)],
                    }],
                ],
            ],
        };
        let mut out = String::new();
        render_block_markdown(&table, &mut out);
        assert!(out.contains("| Name | Age |"));
        assert!(out.contains("| --- | --- |"));
        assert!(out.contains("| Alice | 30 |"));
    }

    #[test]
    fn render_empty_paragraph_skipped() {
        let block = Block::Paragraph {
            style: ParaStyle::default(),
            runs: vec![],
        };
        let mut out = String::new();
        render_block_markdown(&block, &mut out);
        assert_eq!(out, "");
    }

    #[test]
    fn render_pipe_escaped_in_table() {
        let table = Block::Table {
            rows: vec![vec![vec![Block::Paragraph {
                style: ParaStyle::default(),
                runs: vec![run("A|B", false, false)],
            }]]],
        };
        let mut out = String::new();
        render_block_markdown(&table, &mut out);
        assert!(out.contains("A\\|B"));
    }

    #[test]
    fn render_block_markdown_image_with_ocr_quote() {
        let mut out = String::new();
        let block = Block::Image {
            path: "word/media/image1.png".into(),
            markdown: Some("![][image1]".into()),
            ocr_text: Some("BATDOC\nOCR 123".into()),
        };
        render_block_markdown(&block, &mut out);
        assert_eq!(out, "![][image1]\n\n> BATDOC\n> OCR 123\n\n");
    }

    #[test]
    fn render_block_markdown_image_ocr_only_no_tag() {
        let mut out = String::new();
        let block = Block::Image {
            path: "word/media/image1.png".into(),
            markdown: None,
            ocr_text: Some("just the text".into()),
        };
        render_block_markdown(&block, &mut out);
        assert_eq!(out, "> just the text\n\n");
    }

    #[test]
    fn render_block_plain_image_with_ocr() {
        let mut out = String::new();
        let mut first = true;
        let block = Block::Image {
            path: "word/media/image1.png".into(),
            markdown: None,
            ocr_text: Some("line one\nline two".into()),
        };
        render_block_plain(&block, &mut out, &mut first);
        assert_eq!(out, "line one\n\nline two\n");
    }

    // ── Extra parts: comments / footnotes / endnotes ────────────

    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn zip_entry(z: &mut ZipWriter<Cursor<Vec<u8>>>, name: &str, body: &str) {
        z.start_file(name, SimpleFileOptions::default()).unwrap();
        z.write_all(body.as_bytes()).unwrap();
    }

    /// Dummy package content-types part; no relationships are declared, so
    /// only ZIP presence matters to `parse_docx`.
    const CONTENT_TYPES: &str = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>"#;

    /// Build a minimal docx ZIP from the given parts plus a dummy
    /// `[Content_Types].xml`. `word/document.xml` must be present; the
    /// optional extra parts (footnotes/endnotes/comments) may be omitted.
    fn minimal_docx(parts: &[(&str, &str)]) -> Vec<u8> {
        let buf = Cursor::new(Vec::new());
        let mut z = ZipWriter::new(buf);
        zip_entry(&mut z, "[Content_Types].xml", CONTENT_TYPES);
        for (name, body) in parts {
            zip_entry(&mut z, name, body);
        }
        z.finish().unwrap().into_inner()
    }

    /// Image relationship mapping `rId5` → `word/media/image1.png`, plus the
    /// 8-byte PNG signature (base64 `iVBORw0KGgo=`), which `detect_image_mime`
    /// recognizes as `image/png`.
    const IMAGE_RELS_XML: &str = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;
    const PNG_SIGNATURE: [u8; 8] = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    /// Build a docx whose body embeds `media/image1.png` (via `rId5`), plus
    /// any extra parts. The drawing reference is written as
    /// `<w:drawing><a:blip r:embed="rId5"/></w:drawing>` inside a run.
    fn image_docx(body_children: &str, extras: &[(&str, &str)]) -> Vec<u8> {
        let buf = Cursor::new(Vec::new());
        let mut z = ZipWriter::new(buf);
        zip_entry(&mut z, "[Content_Types].xml", CONTENT_TYPES);
        zip_entry(&mut z, "word/document.xml", &document_xml(body_children));
        zip_entry(&mut z, "word/_rels/document.xml.rels", IMAGE_RELS_XML);
        z.start_file("word/media/image1.png", SimpleFileOptions::default())
            .unwrap();
        z.write_all(&PNG_SIGNATURE).unwrap();
        for (name, body) in extras {
            zip_entry(&mut z, name, body);
        }
        z.finish().unwrap().into_inner()
    }

    /// Wrap `body_children` in a minimal `word/document.xml`.
    fn document_xml(body_children: &str) -> String {
        format!(
            r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>{body_children}</w:body>
</w:document>"#
        )
    }

    const FOOTNOTES_XML: &str = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="-1"><w:p><w:r><w:t>separator</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="0"><w:p><w:r><w:t>continuation</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>Source A</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p/></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>duplicate</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;

    const ENDNOTES_XML: &str = r#"<?xml version="1.0"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="-1"><w:p><w:r><w:t>separator</w:t></w:r></w:p></w:endnote>
  <w:endnote w:id="1"><w:p><w:r><w:endnoteRef/><w:t>Endnote B</w:t></w:r></w:p></w:endnote>
</w:endnotes>"#;

    const COMMENTS_XML: &str = r#"<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Alice">
    <w:p><w:r><w:t>First para</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second para</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Alice"><w:p><w:r><w:t>Last word</w:t></w:r></w:p></w:comment>
  <w:comment w:id="2"><w:p><w:r><w:t>No author</w:t></w:r></w:p></w:comment>
  <w:comment w:id="3" w:author="Bob"><w:p><w:r><w:t>   </w:t></w:r></w:p></w:comment>
</w:comments>"#;

    #[test]
    fn parse_footnotes_xml_skips_separators_and_takes_blocks() {
        let notes = parse_footnotes_xml(FOOTNOTES_XML);
        let ids: Vec<&str> = notes.iter().map(|n| n.id.as_str()).collect();
        assert_eq!(ids, ["1", "2"]);

        // First occurrence of id 1 wins; its body is parsed; the glyph
        // contributes nothing; separators/continuations/duplicates leak none.
        assert_eq!(render_plain(&notes[0].blocks).trim_end(), "Source A");
        // id 2 is an empty paragraph: kept as a definition (the trailer in a
        // later task decides what to display), not filtered here.
        assert_eq!(render_plain(&notes[1].blocks).trim_end(), "");
    }

    #[test]
    fn parse_endnotes_xml_skips_separators_and_takes_blocks() {
        let notes = parse_endnotes_xml(ENDNOTES_XML);
        let ids: Vec<&str> = notes.iter().map(|n| n.id.as_str()).collect();
        assert_eq!(ids, ["1"]);
        assert_eq!(render_plain(&notes[0].blocks).trim_end(), "Endnote B");
    }

    #[test]
    fn parse_comments_xml_groups_and_defaults_author() {
        let comments = parse_comments_xml(COMMENTS_XML);
        let ids: Vec<&str> = comments.iter().map(|c| c.id.as_str()).collect();
        // Whitespace-only comment (id 3) is dropped; order preserved.
        assert_eq!(ids, ["0", "1", "2"]);
        assert_eq!(comments[0].author, "Alice");
        assert_eq!(comments[0].blocks.len(), 2); // multi-paragraph body
        assert_eq!(comments[1].author, "Alice");
        assert_eq!(comments[2].author, "Anonymous"); // missing author
        assert_eq!(render_plain(&comments[2].blocks).trim_end(), "No author");
    }

    #[test]
    fn parse_comments_xml_skips_annotation_glyph() {
        let xml = r#"<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="A"><w:p><w:r><w:annotationRef/><w:t>clean text</w:t></w:r></w:p></w:comment>
</w:comments>"#;
        let comments = parse_comments_xml(xml);
        assert_eq!(comments.len(), 1);
        assert_eq!(render_plain(&comments[0].blocks).trim_end(), "clean text");
    }

    #[test]
    fn parse_footnotes_xml_skips_ref_glyph() {
        let xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>Note text</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
        let notes = parse_footnotes_xml(xml);
        assert_eq!(notes.len(), 1);
        // Only "Note text": the glyph emits nothing (no stray text/marker).
        assert_eq!(render_plain(&notes[0].blocks).trim_end(), "Note text");
    }

    #[test]
    fn parse_docx_seeds_footnote_definitions_and_assigns_index() {
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t> tail</w:t></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", FOOTNOTES_XML),
        ]);
        let (blocks, _, _, footnotes, endnotes) =
            parse_docx(&data, crate::ExtractOptions::default()).unwrap();

        // Markers now appear in production output for defined ids.
        let md = render_markdown(&blocks);
        assert!(md.contains("[^1]"), "missing marker: {md}");
        // Trailer-included (Task 7): the extract-level output appends the
        // referenced note to a `## Footnotes` section with its text as a
        // definition; unreferenced notes stay out.
        let full = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert!(
            full.contains("## Footnotes"),
            "missing footnotes trailer: {full}"
        );
        assert!(
            full.contains("[^1]: Source A"),
            "missing footnote definition: {full}"
        );

        // Definitions kept in part order; the referenced note got the display
        // index assigned by the body walk, the unreferenced one stays 0.
        assert_eq!(footnotes.len(), 2);
        assert_eq!(footnotes[0].id, "1");
        assert_eq!(footnotes[0].display_index, 1);
        assert_eq!(footnotes[1].id, "2");
        assert_eq!(footnotes[1].display_index, 0);
        assert!(endnotes.is_empty());
    }

    #[test]
    fn parse_docx_seeds_endnote_definitions() {
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>See</w:t><w:endnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/endnotes.xml", ENDNOTES_XML),
        ]);
        let (blocks, _, _, footnotes, endnotes) =
            parse_docx(&data, crate::ExtractOptions::default()).unwrap();

        let md = render_markdown(&blocks);
        assert!(md.contains("[^e1]"), "missing endnote marker: {md}");
        assert!(footnotes.is_empty());
        assert_eq!(endnotes.len(), 1);
        assert_eq!(endnotes[0].id, "1");
        assert_eq!(endnotes[0].display_index, 1);
    }

    #[test]
    fn parse_docx_comments_populated() {
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(r"<w:p><w:r><w:t>Body only</w:t></w:r></w:p>"),
            ),
            ("word/comments.xml", COMMENTS_XML),
        ]);
        let (blocks, _, comments, footnotes, endnotes) =
            parse_docx(&data, crate::ExtractOptions::default()).unwrap();

        // Body unaffected; comments carry author + multi-paragraph blocks
        // (rendering into a trailer is Task 7).
        assert_eq!(render_markdown(&blocks).trim_end(), "Body only");
        assert_eq!(comments.len(), 3);
        assert_eq!(comments[0].author, "Alice");
        assert_eq!(comments[0].blocks.len(), 2);
        assert!(footnotes.is_empty());
        assert!(endnotes.is_empty());
    }

    #[test]
    fn extract_no_footnote_part_no_markers() {
        let data = minimal_docx(&[(
            "word/document.xml",
            &document_xml(r#"<w:p><w:r><w:t>Hi</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#),
        )]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert_eq!(md, "Hi\n\n");
        let plain = extract_plain(&data, crate::ExtractOptions::default()).unwrap();
        assert_eq!(plain, "Hi\n");
    }

    // ── extract_* trailers: comments / footnotes / endnotes ──────

    #[test]
    fn extract_markdown_footnotes_and_endnotes_trailers() {
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/><w:t> and </w:t><w:endnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", FOOTNOTES_XML),
            ("word/endnotes.xml", ENDNOTES_XML),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        // Body carries the inline markers for the referenced notes.
        assert!(md.contains("[^1]"), "footnote marker missing: {md}");
        assert!(md.contains("[^e1]"), "endnote marker missing: {md}");
        // Trailers follow the body in the fixed order footnotes → endnotes;
        // no comments part → no comments section.
        let body_at = md.find("Body").expect("body text");
        let footnotes_at = md.find("## Footnotes").expect("footnotes section");
        let endnotes_at = md.find("## Endnotes").expect("endnotes section");
        assert!(body_at < footnotes_at, "body must precede trailers: {md}");
        assert!(
            footnotes_at < endnotes_at,
            "footnotes must precede endnotes: {md}"
        );
        // Definitions carry the note text.
        assert!(md.contains("[^1]: Source A"), "footnote definition: {md}");
        assert!(md.contains("[^e1]: Endnote B"), "endnote definition: {md}");
        assert!(!md.contains("## Comments"), "no comments expected: {md}");
    }

    #[test]
    fn extract_markdown_comments_grouped_by_author() {
        let comments_xml = r#"<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Alice"><w:p><w:r><w:t>First thought</w:t></w:r></w:p><w:p><w:r><w:t>Second thought</w:t></w:r></w:p></w:comment>
  <w:comment w:id="1" w:author="Alice"><w:p><w:r><w:t>Third thought</w:t></w:r></w:p></w:comment>
  <w:comment w:id="2" w:author="Bob"><w:p><w:r><w:t>Bob's note</w:t></w:r></w:p></w:comment>
  <w:comment w:id="3" w:author="Alice"><w:p><w:r><w:t>Return thought</w:t></w:r></w:p></w:comment>
</w:comments>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml("<w:p><w:r><w:t>Body</w:t></w:r></w:p>"),
            ),
            ("word/comments.xml", comments_xml),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        // Consecutive Alice comments share one heading; the later Alice
        // comment is NOT consecutive, so the heading repeats. Order of the
        // headings matches the comment order.
        assert_eq!(
            md.matches("### Alice").count(),
            2,
            "expected two Alice groups: {md}"
        );
        assert_eq!(
            md.matches("### Bob").count(),
            1,
            "expected one Bob group: {md}"
        );
        let first_alice = md.find("### Alice").expect("first Alice heading");
        let bob = md.find("### Bob").expect("Bob heading");
        let second_alice = md.rfind("### Alice").expect("second Alice heading");
        assert!(
            first_alice < bob && bob < second_alice,
            "heading order: {md}"
        );
        for text in [
            "First thought",
            "Second thought",
            "Third thought",
            "Bob's note",
            "Return thought",
        ] {
            assert!(md.contains(text), "missing comment text {text:?}: {md}");
        }
    }

    #[test]
    fn extract_markdown_multi_block_footnote_indented() {
        let footnotes_xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>First paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Cited</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", footnotes_xml),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        // Definition-list form: `[^1]:` on the label line, then every line
        // of the body indented 4 spaces (CommonMark continuation).
        assert!(
            md.contains("[^1]:\n    First paragraph."),
            "indented definition: {md}"
        );
        assert!(
            md.contains("    Second paragraph."),
            "indented continuation: {md}"
        );
        // The inline single-line form is for single-paragraph notes only.
        assert!(
            !md.contains("[^1]: First paragraph."),
            "inline form used for multi-block note: {md}"
        );
    }

    #[test]
    fn extract_markdown_table_only_footnote_indented() {
        // A table-only footnote must use the indented (multi-block)
        // definition form: the inline single-line form would let the
        // table's second line escape the `[^1]:` definition, which is
        // invalid CommonMark.
        let footnotes_xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:footnote>
</w:footnotes>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Cited</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", footnotes_xml),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        // Label line, then every body line indented 4 spaces: the table
        // header starts with 4 spaces + `| `.
        assert!(
            md.contains("[^1]:\n    | Table cell |"),
            "indented table definition: {md}"
        );
        assert!(
            md.contains("\n    | --- |"),
            "indented table separator: {md}"
        );
        // The inline form would emit `[^1]: | …` on the label line — banned.
        assert!(
            !md.contains("[^1]: |"),
            "inline form used for table-only note: {md}"
        );
    }

    #[test]
    fn extract_markdown_empty_defined_note_no_marker_no_section() {
        // A defined note whose blocks render empty must not be seeded: the
        // body emits no marker and the trailer omits the whole section (no
        // dangling `[^1]`, no empty `## Footnotes`).
        let footnotes_xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p/><w:p/></w:footnote>
  <w:footnote w:id="2"><w:p/></w:footnote>
</w:footnotes>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Ref</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", footnotes_xml),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        assert!(md.contains("Ref"), "body text missing: {md}");
        assert!(!md.contains("[^1]"), "empty note got a marker: {md}");
        // Trailer: both notes are empty, so the section is omitted entirely
        // (the unreferenced empty note id 2 is trivially excluded too).
        assert!(
            !md.contains("## Footnotes"),
            "empty notes got a section: {md}"
        );
    }

    #[test]
    fn extract_plain_footnotes_multi_paragraph_continuation() {
        let footnotes_xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>First paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Cited</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", footnotes_xml),
        ]);
        let plain = extract_plain(&data, crate::ExtractOptions::default()).unwrap();

        assert!(
            plain.contains("\n\n--- Footnotes ---\n"),
            "blank line before trailer: {plain}"
        );
        // First line carries the [1] bracket; the second paragraph continues
        // on its own line WITHOUT a repeated bracket. The plain renderer
        // separates consecutive paragraph blocks with a blank line (its
        // regular between-paragraph convention), so both are expected here.
        assert!(
            plain.contains("[1] First paragraph.\n\nSecond paragraph."),
            "continuation without bracket: {plain}"
        );
        assert!(
            !plain.contains("[1] Second paragraph."),
            "repeated bracket on continuation: {plain}"
        );
    }

    #[test]
    fn extract_markdown_unreferenced_footnote_omitted() {
        let footnotes_xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>Referenced note</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Hidden note</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Only one</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", footnotes_xml),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        assert!(
            md.contains("[^1]: Referenced note"),
            "referenced note missing: {md}"
        );
        assert!(
            !md.contains("Hidden note"),
            "unreferenced note leaked: {md}"
        );
        assert!(!md.contains("[^2]"), "unreferenced note got a label: {md}");
    }

    #[test]
    fn extract_markdown_empty_extras_identical_to_body_only() {
        // No extras parts at all → output is exactly the body (regression:
        // trailer wiring must not alter body-only documents).
        let plain_doc = minimal_docx(&[(
            "word/document.xml",
            &document_xml("<w:p><w:r><w:t>Body only</w:t></w:r></w:p>"),
        )]);
        assert_eq!(
            extract_markdown(&plain_doc, crate::ExtractOptions::default()).unwrap(),
            "Body only\n\n"
        );
        assert_eq!(
            extract_plain(&plain_doc, crate::ExtractOptions::default()).unwrap(),
            "Body only\n"
        );

        // With images enabled and still no extras parts: body, inline image
        // reference, then the trailing definition — nothing else.
        let image_doc = image_docx(
            r#"<w:p><w:r><w:t>Pic</w:t></w:r><w:r><w:drawing><a:blip r:embed="rId5"/></w:drawing></w:r></w:p>"#,
            &[],
        );
        assert_eq!(
            extract_markdown(
                &image_doc,
                crate::ExtractOptions {
                    images: true,
                    ..crate::ExtractOptions::default()
                }
            )
            .unwrap(),
            "Pic\n\n![][image1]\n\n[image1]: <data:image/png;base64,iVBORw0KGgo=>\n"
        );
    }

    #[test]
    fn extract_markdown_images_extras_before_defs() {
        let footnotes_xml = r#"<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>Source A</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;
        let data = image_docx(
            r#"<w:p><w:r><w:t>Pic</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:drawing><a:blip r:embed="rId5"/></w:drawing></w:r></w:p>"#,
            &[("word/footnotes.xml", footnotes_xml)],
        );
        let md = extract_markdown(
            &data,
            crate::ExtractOptions {
                images: true,
                ..crate::ExtractOptions::default()
            },
        )
        .unwrap();

        // The inline image reference renders in the body…
        assert!(md.contains("![][image1]"), "inline image ref: {md}");
        // …but the definition block is appended AFTER the extras trailers.
        let footnotes_at = md.find("## Footnotes").expect("footnotes section");
        let defs_at = md
            .find("[image1]: <data:image/png;base64,")
            .expect("image definition");
        assert!(
            footnotes_at < defs_at,
            "extras must precede image defs: {md}"
        );
        assert!(md.contains("[^1]: Source A"), "footnote definition: {md}");
    }

    #[test]
    fn extract_plain_comments_trailer() {
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml("<w:p><w:r><w:t>Body only</w:t></w:r></w:p>"),
            ),
            ("word/comments.xml", COMMENTS_XML),
        ]);
        let plain = extract_plain(&data, crate::ExtractOptions::default()).unwrap();

        // Section heading after the body, one blank line apart; no other
        // trailer sections (empty sections are omitted).
        let body_at = plain.find("Body only").expect("body text");
        let comments_at = plain.find("--- Comments ---").expect("comments section");
        assert!(body_at < comments_at, "body must precede trailer: {plain}");
        assert!(
            plain.contains("\n\n--- Comments ---\n"),
            "blank line before trailer: {plain}"
        );
        assert!(
            !plain.contains("--- Footnotes ---"),
            "no footnotes: {plain}"
        );
        // Consecutive Alice comments share one [Alice] group; the missing
        // author maps to [Anonymous] at parse; Bob's whitespace-only comment
        // is dropped entirely.
        assert_eq!(
            plain.matches("[Alice]").count(),
            1,
            "Alice must be grouped: {plain}"
        );
        assert!(plain.contains("[Anonymous]"), "anonymous group: {plain}");
        for text in ["First para", "Second para", "Last word", "No author"] {
            assert!(
                plain.contains(text),
                "missing comment text {text:?}: {plain}"
            );
        }
        assert!(
            !plain.contains("[Bob]"),
            "empty comment must be dropped: {plain}"
        );
    }

    #[test]
    fn extract_markdown_empty_comments_section_omitted() {
        // A comments part whose every comment body is whitespace-only parses
        // to zero comments: both renderers must omit the section entirely
        // and stay byte-identical to the body-only document.
        let comments_xml = r#"<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Nobody"><w:p><w:r><w:t>   </w:t></w:r></w:p></w:comment>
</w:comments>"#;
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml("<w:p><w:r><w:t>Body</w:t></w:r></w:p>"),
            ),
            ("word/comments.xml", comments_xml),
        ]);
        assert_eq!(
            extract_markdown(&data, crate::ExtractOptions::default()).unwrap(),
            "Body\n\n",
            "empty comments section must be omitted: {}",
            extract_markdown(&data, crate::ExtractOptions::default()).unwrap()
        );
        assert_eq!(
            extract_plain(&data, crate::ExtractOptions::default()).unwrap(),
            "Body\n",
            "empty comments section must be omitted: {}",
            extract_plain(&data, crate::ExtractOptions::default()).unwrap()
        );
    }

    #[test]
    fn extract_plain_endnotes_trailer() {
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>See</w:t><w:endnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/endnotes.xml", ENDNOTES_XML),
        ]);
        let plain = extract_plain(&data, crate::ExtractOptions::default()).unwrap();

        let body_at = plain.find("See").expect("body text");
        let endnotes_at = plain.find("--- Endnotes ---").expect("endnotes section");
        assert!(body_at < endnotes_at, "body must precede trailer: {plain}");
        assert!(
            plain.contains("[e1] Endnote B"),
            "endnote definition: {plain}"
        );
    }

    #[test]
    fn extract_markdown_trailer_spacing_single_blank_line() {
        // Body ending in a paragraph already ends with a blank line; the
        // trailer must not add more than one blank line between them.
        let data = minimal_docx(&[
            (
                "word/document.xml",
                &document_xml(
                    r#"<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
                ),
            ),
            ("word/footnotes.xml", FOOTNOTES_XML),
        ]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();

        assert!(
            md.contains("\n\n## Footnotes"),
            "exactly one blank line before trailer: {md}"
        );
        assert!(!md.contains("\n\n\n"), "more than one blank line: {md}");
    }

    // ── read_optional_part: missing vs present-but-unreadable ─────────

    #[test]
    fn read_optional_part_returns_some_when_part_present() {
        let data = minimal_docx(&[("word/footnotes.xml", "<w:footnotes/>")]);
        let mut archive = ZipArchive::new(Cursor::new(data.as_slice())).unwrap();
        let got = read_optional_part(&mut archive, "word/footnotes.xml").unwrap();
        assert_eq!(got.as_deref(), Some("<w:footnotes/>"));
    }

    #[test]
    fn read_optional_part_returns_none_when_part_missing() {
        let data = minimal_docx(&[]);
        let mut archive = ZipArchive::new(Cursor::new(data.as_slice())).unwrap();
        let got = read_optional_part(&mut archive, "word/footnotes.xml").unwrap();
        assert!(
            got.is_none(),
            "a genuinely missing part must be omitted, never an error"
        );
    }

    #[test]
    fn read_optional_part_fails_when_present_part_is_corrupt() {
        // A single-entry zip whose deflated data is corrupted mid-stream
        // (local header and central directory stay intact): the part exists,
        // so extraction must fail instead of silently treating it as missing.
        let xml = r#"<w:footnotes><w:footnote w:id="1"><w:p><w:r><w:t>x</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
        let mut data = {
            let mut z = ZipWriter::new(Cursor::new(Vec::new()));
            z.start_file("word/footnotes.xml", SimpleFileOptions::default())
                .unwrap();
            z.write_all(xml.as_bytes()).unwrap();
            z.finish().unwrap().into_inner()
        };

        // The end-of-central-directory record occupies the trailing 22 bytes;
        // its central-directory offset lives at +16..+20.
        let eocd = data.len() - 22;
        assert_eq!(&data[eocd..eocd + 4], b"PK\x05\x06");
        let cd_offset = u32::from_le_bytes(data[eocd + 16..eocd + 20].try_into().unwrap()) as usize; // u32 → usize: lossless on 32+ bit
        assert_eq!(&data[cd_offset..cd_offset + 4], b"PK\x01\x02");

        // File data starts right after the 30-byte local header plus its
        // name and extra fields; compressed size comes from the central
        // directory entry (+20..+24) and the entry name length from +28..+30.
        let lh_name_len = usize::from(u16::from_le_bytes(data[26..28].try_into().unwrap()));
        let lh_extra_len = usize::from(u16::from_le_bytes(data[28..30].try_into().unwrap()));
        let name_len = usize::from(u16::from_le_bytes(
            data[cd_offset + 28..cd_offset + 30].try_into().unwrap(),
        ));
        assert_eq!(name_len, lh_name_len);
        let csize =
            u32::from_le_bytes(data[cd_offset + 20..cd_offset + 24].try_into().unwrap()) as usize; // u32 → usize: lossless on 32+ bit
        let mid = 30 + lh_name_len + lh_extra_len + csize / 2;
        data[mid] ^= 0xFF;

        let mut archive = ZipArchive::new(Cursor::new(data.as_slice())).unwrap();
        let got = read_optional_part(&mut archive, "word/footnotes.xml");
        assert!(
            got.is_err(),
            "a present-but-unreadable part must fail the extract"
        );
    }
}
