//! OOXML `.pptx` (`PowerPoint`) presentation parser.
//!
//! Unzips the `.pptx` archive, discovers slides from `ppt/presentation.xml`
//! and its relationships, then parses each slide's XML to extract text from
//! shapes. Hyperlinks on text runs are resolved from per-slide relationship
//! files. Output is either plain text or markdown with per-slide headings.

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::{Cursor, Read};
use zip::ZipArchive;

use crate::markup;
use crate::xml_util::{self, get_attr, Rels};

/// A parsed slide: its number and extracted text runs.
#[derive(Debug)]
struct Slide {
    number: usize,
    /// Each element is one shape's worth of text (paragraphs joined by newlines).
    shapes: Vec<ShapeText>,
    /// Inline image references for this slide (e.g., `![][image1]`).
    images: Vec<String>,
    /// OCR'd text of embedded images — `--ocr` only.
    image_ocr: Vec<String>,
    /// Speaker notes shapes. Empty means no speaker notes for this slide.
    notes: Vec<ShapeText>,
}

/// Text extracted from a single shape, preserving paragraph structure.
#[derive(Debug)]
struct ShapeText {
    paragraphs: Vec<Paragraph>,
}

/// Whether a paragraph is a bullet/numbered list item.
#[derive(Debug, Clone, PartialEq, Eq)]
enum BulletKind {
    /// Not a list item.
    None,
    /// Unordered (bullet) list item at the given nesting level (0-based).
    Bullet(u8),
    /// Ordered (numbered) list item at the given nesting level (0-based).
    Numbered(u8),
}

/// A paragraph inside a shape, with optional heading level inference.
#[derive(Debug)]
struct Paragraph {
    runs: Vec<TextRun>,
    /// 0 = normal, 1-6 = heading level (inferred from font size).
    heading_level: u8,
    /// Bullet/numbered list membership.
    bullet: BulletKind,
}

/// A single text run with optional formatting.
#[derive(Debug)]
struct TextRun {
    text: String,
    bold: bool,
    italic: bool,
    /// Resolved hyperlink URL, if any.
    link_url: Option<String>,
    /// Font size in half-points (OOXML stores as hundredths of a point,
    /// so 2400 = 24pt). Used for heading inference.
    font_size: Option<u32>,
}

/// Extract plain text from a .pptx file.
pub(crate) fn extract_plain(data: &[u8], ocr: bool) -> crate::error::Result<String> {
    let (slides, _) = parse_pptx(data, false, ocr)?;
    Ok(render_plain(&slides))
}

/// Extract markdown-formatted text from a .pptx file.
///
/// When `images` is true, embedded images are extracted and included as
/// reference-style base64 images with definitions appended at the end.
/// When `ocr` is true, embedded images are OCR'd and rendered as a
/// blockquote after each slide's images.
pub(crate) fn extract_markdown(
    data: &[u8],
    images: bool,
    ocr: bool,
) -> crate::error::Result<String> {
    let (slides, image_defs) = parse_pptx(data, images, ocr)?;
    let mut md = render_markdown(&slides);
    if !image_defs.is_empty() {
        for def in &image_defs {
            md.push_str(def);
            md.push('\n');
        }
    }
    Ok(md)
}

// ── Parsing ────────────────────────────────────────────────────────

/// Parse the pptx archive into slides and image reference definitions.
///
/// When `extract_images` is true, image relationships are loaded and
/// `<p:pic>` elements are extracted as reference-style images. When `ocr`
/// is true, the same images are additionally OCR'd.
fn parse_pptx(
    data: &[u8],
    extract_images: bool,
    ocr: bool,
) -> crate::error::Result<(Vec<Slide>, Vec<String>)> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    // Discover slides from presentation.xml + rels
    let slide_paths = discover_slides(&mut archive)?;

    let mut slides = Vec::new();
    let mut all_image_defs = Vec::new();
    let mut image_counter = 0usize;

    for (num, path) in slide_paths {
        let mut xml = String::new();
        match archive.by_name(&path) {
            Ok(mut entry) => {
                entry.read_to_string(&mut xml)?;
            }
            Err(_) => continue,
        }

        // Load per-slide hyperlink rels
        let slide_rels_path = xml_util::rels_path(&path);
        let rels = xml_util::load_rels(&mut archive, &slide_rels_path);

        // Optionally load image rels for this slide
        let image_rels = if extract_images || ocr {
            xml_util::load_image_rels(&mut archive, &slide_rels_path)
        } else {
            xml_util::Rels::new()
        };

        let shapes = parse_slide_xml(&xml, &rels);

        // Extract images from <p:pic> elements (for --images refs and/or OCR)
        let (images, image_ocr) = if (extract_images || ocr) && !image_rels.is_empty() {
            let pic_rids = parse_slide_pic_rids(&xml);
            let base_dir = path.rsplit_once('/').map_or("ppt", |(dir, _)| dir);
            let mut inline_refs = Vec::new();
            let mut ocr_texts = Vec::new();
            for rid in pic_rids {
                if let Some(target) = image_rels.get(&rid) {
                    if let Some(data) =
                        xml_util::read_image_from_zip(&mut archive, target, base_dir)
                    {
                        if extract_images {
                            image_counter += 1;
                            let id = format!("image{image_counter}");
                            if let Some(img_ref) = crate::markup::image_to_base64_ref(&data, &id) {
                                inline_refs.push(img_ref.inline);
                                all_image_defs.push(img_ref.definition);
                            }
                        }
                        if ocr {
                            if let Some(text) = crate::ocr::ocr_image_bytes(&data)? {
                                ocr_texts.push(text);
                            }
                        }
                    }
                }
            }
            (inline_refs, ocr_texts)
        } else {
            (Vec::new(), Vec::new())
        };

        // Discover speaker notes from the notesSlide relationship. Missing
        // rels, missing targets, and whitespace-only bodies yield no notes.
        let notes = load_slide_notes(&mut archive, &path);

        slides.push(Slide {
            number: num,
            shapes,
            images,
            image_ocr,
            notes,
        });
    }

    Ok((slides, all_image_defs))
}

/// Parse a slide's XML to extract rId values from `<p:pic>` → `<a:blip>` elements.
///
/// Returns the rIds in document order.
fn parse_slide_pic_rids(xml: &str) -> Vec<String> {
    let mut rids = Vec::new();
    let mut reader = Reader::from_str(xml);
    let mut in_pic = false;
    let mut depth = 0u32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"pic" {
                    in_pic = true;
                    depth = 1;
                } else if in_pic {
                    depth += 1;
                    if e.local_name().as_ref() == b"blip" {
                        if let Some(rid) = get_attr(e, b"r:embed") {
                            rids.push(rid);
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) if in_pic && e.local_name().as_ref() == b"blip" => {
                if let Some(rid) = get_attr(e, b"r:embed") {
                    rids.push(rid);
                }
            }
            Ok(Event::End(ref e)) if in_pic => {
                if e.local_name().as_ref() == b"pic" {
                    in_pic = false;
                } else {
                    depth -= 1;
                    if depth == 0 {
                        in_pic = false;
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    rids
}

/// Discover slide file paths from presentation.xml, in order.
///
/// Returns `(slide_number, zip_path)` pairs sorted by slide order.
fn discover_slides(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> crate::error::Result<Vec<(usize, String)>> {
    // Parse presentation.xml for slide rId ordering
    let mut pres_xml = String::new();
    archive
        .by_name("ppt/presentation.xml")?
        .read_to_string(&mut pres_xml)?;

    // Collect slide rIds in order
    let mut slide_rids: Vec<String> = Vec::new();
    let mut reader = Reader::from_str(&pres_xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"sldId" {
                    if let Some(rid) = get_attr(e, b"r:id") {
                        slide_rids.push(rid);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Parse presentation.xml.rels for rId → target path mapping
    let mut rels_xml = String::new();
    archive
        .by_name("ppt/_rels/presentation.xml.rels")?
        .read_to_string(&mut rels_xml)?;

    let mut rid_to_target: HashMap<String, String> = HashMap::new();
    let mut reader = Reader::from_str(&rels_xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let id = get_attr(e, b"Id").unwrap_or_default();
                let target = get_attr(e, b"Target").unwrap_or_default();
                if !id.is_empty() && !target.is_empty() {
                    rid_to_target.insert(id, target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Resolve: slide number → zip path
    let mut result = Vec::new();
    for (i, rid) in slide_rids.iter().enumerate() {
        if let Some(target) = rid_to_target.get(rid) {
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("ppt/{target}")
            };
            result.push((i + 1, path));
        }
    }

    Ok(result)
}

// ── Slide XML parsing ──────────────────────────────────────────────

/// Parse a single slide's XML, extracting text from all shapes.
fn parse_slide_xml(xml: &str, rels: &Rels) -> Vec<ShapeText> {
    let mut reader = Reader::from_str(xml);
    let mut shapes = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                // <p:sp> = shape, <p:graphicFrame> = table/chart, <p:grpSp> = group
                if name.as_ref() == b"sp" || name.as_ref() == b"graphicFrame" {
                    if let Some(shape) = parse_shape(&mut reader, rels, e.local_name().as_ref()) {
                        if !shape.paragraphs.is_empty() {
                            shapes.push(shape);
                        }
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    shapes
}

/// Parse a shape element (`<p:sp>` or `<p:graphicFrame>`), extracting its text body.
fn parse_shape(reader: &mut Reader<&[u8]>, rels: &Rels, end_tag: &[u8]) -> Option<ShapeText> {
    let mut paragraphs = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"txBody" => {
                parse_text_body(reader, rels, &mut paragraphs);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(ShapeText { paragraphs })
    }
}

/// Discover speaker notes for a slide from its `notesSlide` relationship.
///
/// Reads the slide's relationship file, finds the first `notesSlide`
/// target, resolves it to a ZIP path relative to the slide's directory,
/// and parses the notes part. Missing or unreadable rels/notes parts and
/// whitespace-only notes bodies all yield no notes.
fn load_slide_notes(archive: &mut ZipArchive<Cursor<&[u8]>>, slide_path: &str) -> Vec<ShapeText> {
    let mut rels_xml = String::new();
    match archive.by_name(&xml_util::rels_path(slide_path)) {
        Ok(mut entry) => {
            if entry.read_to_string(&mut rels_xml).is_err() {
                return Vec::new();
            }
        }
        Err(_) => return Vec::new(),
    }

    let Some(target) = xml_util::find_rel_target_by_type_suffix(&rels_xml, "/notesSlide") else {
        return Vec::new();
    };
    let base_dir = slide_path.rsplit_once('/').map_or("", |(dir, _)| dir);
    let notes_path = xml_util::resolve_zip_target(&target, base_dir);

    let mut xml = String::new();
    match archive.by_name(&notes_path) {
        Ok(mut entry) => {
            if entry.read_to_string(&mut xml).is_err() {
                return Vec::new();
            }
        }
        Err(_) => return Vec::new(),
    }

    let shapes = parse_notes_slide_xml(&xml);
    if notes_nonempty(&shapes) {
        shapes
    } else {
        Vec::new()
    }
}

/// Parse a notes slide's XML, extracting the text of `body` placeholders.
///
/// Only shapes anchored to a `body` placeholder contribute speaker notes;
/// slide-image placeholders (`sldImg`), titles, and freeform shapes are
/// ignored. Run hyperlinks are not resolved — notes rels are not loaded,
/// so the run walker runs against an empty `Rels`.
fn parse_notes_slide_xml(xml: &str) -> Vec<ShapeText> {
    let mut reader = Reader::from_str(xml);
    let rels = Rels::new();
    let mut shapes = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sp" => {
                if let Some(shape) = parse_notes_shape(&mut reader, &rels, b"sp") {
                    shapes.push(shape);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    shapes
}

/// Parse a single notes `<p:sp>` shape, keeping only `body` placeholders.
///
/// Records the placeholder type from `<p:ph type="…"/>` (which appears in
/// `<p:nvSpPr>/<p:nvPr>` before the text body), then parses `txBody` with
/// the same walker used for slides. Shapes whose placeholder type is not
/// `body` — including freeform shapes with no `<p:ph>` at all — are dropped.
fn parse_notes_shape(reader: &mut Reader<&[u8]>, rels: &Rels, end_tag: &[u8]) -> Option<ShapeText> {
    let mut ph_type: Option<String> = None;
    let mut paragraphs = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"ph" && ph_type.is_none() =>
            {
                ph_type = get_attr(e, b"type");
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"txBody" => {
                parse_text_body(reader, rels, &mut paragraphs);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Only shapes anchored to a `body` placeholder carry speaker notes.
    // Empty text bodies are dropped here; whitespace-only notes are
    // filtered out by load_slide_notes.
    if ph_type.as_deref() != Some("body") || paragraphs.is_empty() {
        None
    } else {
        Some(ShapeText { paragraphs })
    }
}

/// Parse a `<p:txBody>` (or `<a:txBody>`) element.
fn parse_text_body(reader: &mut Reader<&[u8]>, rels: &Rels, paragraphs: &mut Vec<Paragraph>) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"p" => {
                let para = parse_para(reader, rels);
                if !para.runs.is_empty() {
                    paragraphs.push(para);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"txBody" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a `<a:p>` paragraph element within a text body.
fn parse_para(reader: &mut Reader<&[u8]>, rels: &Rels) -> Paragraph {
    let mut runs = Vec::new();
    let mut max_font_size: Option<u32> = None;
    let mut bullet = BulletKind::None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"pPr" => {
                        parse_para_props(reader, e, &mut bullet);
                    }
                    b"r" => {
                        let run = parse_text_run(reader, rels, None);
                        if let Some(fs) = run.font_size {
                            max_font_size = Some(max_font_size.map_or(fs, |prev| prev.max(fs)));
                        }
                        if !run.text.is_empty() {
                            runs.push(run);
                        }
                    }
                    b"fld" => {
                        // Field element (slide number, date, etc.) — extract text
                        let run = parse_text_run(reader, rels, Some(b"fld"));
                        if !run.text.is_empty() {
                            runs.push(run);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"br" => {
                        runs.push(TextRun {
                            text: "\n".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        });
                    }
                    b"pPr" => {
                        // Self-closing <a:pPr lvl="1"/> with no bullet children
                        // means default bullet for body placeholders
                        let lvl = get_attr(e, b"lvl")
                            .and_then(|v| v.parse::<u8>().ok())
                            .unwrap_or(0);
                        // Self-closing pPr has no child elements, so we can't tell
                        // bullet vs. no-bullet — leave as None (will be plain text).
                        // In practice, self-closing pPr without buNone in a body
                        // placeholder still gets a bullet from the master, but we
                        // can't know that without parsing the slide layout/master.
                        // We'll rely on the <a:pPr> Start form for bullet info.
                        let _ = lvl;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"p" => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Infer heading level from font size (in hundredths of a point):
    // >= 2800 (28pt) → h1, >= 2400 (24pt) → h2, >= 2000 (20pt) → h3
    let heading_level = max_font_size.map_or(0, |fs| {
        if fs >= 2800 {
            1
        } else if fs >= 2400 {
            2
        } else if fs >= 2000 {
            3
        } else {
            0
        }
    });

    // Bullets and headings are mutually exclusive — headings win.
    if heading_level > 0 {
        bullet = BulletKind::None;
    }

    Paragraph {
        runs,
        heading_level,
        bullet,
    }
}

/// Parse `<a:pPr>` element for bullet/numbering and nesting level.
///
/// Looks at the `lvl` attribute (0-based nesting level) and child elements:
/// - `<a:buChar char="●"/>` → unordered bullet
/// - `<a:buAutoNum type="arabicPeriod"/>` (any type) → numbered
/// - `<a:buNone/>` → explicitly no bullet
///
/// When no explicit bullet child is found but `lvl` is present, we default
/// to an unordered bullet — `PowerPoint` body placeholders inherit bullets
/// from the slide layout/master, and the `lvl` attribute alone indicates
/// list membership in practice.
fn parse_para_props(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    bullet: &mut BulletKind,
) {
    let lvl = get_attr(start, b"lvl")
        .and_then(|v| v.parse::<u8>().ok())
        .unwrap_or(0);

    let mut found_bu_char = false;
    let mut found_bu_auto = false;
    let mut found_bu_none = false;
    let mut depth = 1u32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                depth += 1;
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"buChar" | b"buBlip" | b"buFont" => found_bu_char = true,
                    b"buAutoNum" => found_bu_auto = true,
                    b"buNone" => found_bu_none = true,
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

    if found_bu_none {
        *bullet = BulletKind::None;
    } else if found_bu_auto {
        *bullet = BulletKind::Numbered(lvl);
    } else if found_bu_char {
        *bullet = BulletKind::Bullet(lvl);
    }
    // If none of the explicit bullet markers were found, leave bullet as None.
    // We intentionally don't infer bullets from `lvl` alone — that would
    // require parsing the slide master/layout to know whether the placeholder
    // has default bullets.
}

/// Parse a `<a:r>` text run (or `<a:fld>` field) element.
///
/// Extracts text, bold/italic, font size, and hyperlink URL.
fn parse_text_run(reader: &mut Reader<&[u8]>, rels: &Rels, end_tag: Option<&[u8]>) -> TextRun {
    let end_name = end_tag.unwrap_or(b"r");
    let mut text = String::new();
    let mut bold = false;
    let mut italic = false;
    let mut link_url: Option<String> = None;
    let mut font_size: Option<u32> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"rPr" => {
                        // Read attributes from the <a:rPr> start tag
                        read_rpr_attrs(e, &mut bold, &mut italic, &mut font_size);
                        // Parse children for hyperlinks
                        parse_run_props_children(reader, &mut link_url, rels);
                    }
                    b"t" => {
                        if let Ok(Event::Text(t)) = reader.read_event() {
                            if let Ok(s) = t.unescape() {
                                text.push_str(&s);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"rPr" {
                    // Self-closing <a:rPr b="1" i="1" sz="2400"/>
                    read_rpr_attrs(e, &mut bold, &mut italic, &mut font_size);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_name => {
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    TextRun {
        text,
        bold,
        italic,
        link_url,
        font_size,
    }
}

/// Read bold, italic, and font size from `<a:rPr>` element attributes.
fn read_rpr_attrs(
    e: &quick_xml::events::BytesStart,
    bold: &mut bool,
    italic: &mut bool,
    font_size: &mut Option<u32>,
) {
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"b" => *bold = attr.value.as_ref() == b"1",
            b"i" => *italic = attr.value.as_ref() == b"1",
            b"sz" => {
                if let Ok(s) = std::str::from_utf8(&attr.value) {
                    *font_size = s.parse().ok();
                }
            }
            _ => {}
        }
    }
}

/// Parse children of `<a:rPr>` to find hyperlink references.
fn parse_run_props_children(
    reader: &mut Reader<&[u8]>,
    link_url: &mut Option<String>,
    rels: &Rels,
) {
    let mut depth = 1u32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"hlinkClick" {
                    if let Some(rid) = get_attr(e, b"r:id") {
                        if let Some(url) = rels.get(&rid) {
                            *link_url = Some(url.clone());
                        }
                    }
                }
                depth += 1;
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"hlinkClick" => {
                if let Some(rid) = get_attr(e, b"r:id") {
                    if let Some(url) = rels.get(&rid) {
                        *link_url = Some(url.clone());
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"rPr" {
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
}

// ── Rendering ──────────────────────────────────────────────────────

/// Render slides as plain text.
fn render_plain(slides: &[Slide]) -> String {
    let mut out = String::new();
    let multiple = slides.len() > 1;

    for (i, slide) in slides.iter().enumerate() {
        if slide.shapes.is_empty() && slide.image_ocr.is_empty() {
            continue;
        }

        if multiple {
            if i > 0 {
                out.push('\n');
            }
            let _ = writeln!(out, "--- Slide {} ---", slide.number);
        }

        for shape in &slide.shapes {
            render_shape_plain(shape, &mut out);
        }

        let mut first_ocr = true;
        for ocr in &slide.image_ocr {
            crate::markup::push_ocr_plain(&mut out, ocr, &mut first_ocr);
        }
    }

    append_notes_plain(slides, &mut out);

    out
}

/// Render slides as markdown.
fn render_markdown(slides: &[Slide]) -> String {
    let mut out = String::new();
    let multiple = slides.len() > 1;

    for slide in slides {
        if slide.shapes.is_empty() && slide.images.is_empty() && slide.image_ocr.is_empty() {
            continue;
        }

        if multiple {
            let _ = write!(out, "## Slide {}\n\n", slide.number);
        }

        let heading_offset = if multiple { 2 } else { 0 };
        let mut first_shape = true;
        for shape in &slide.shapes {
            if !first_shape {
                out.push('\n');
            }
            first_shape = false;
            render_shape_markdown(shape, &mut out, heading_offset);
        }

        // Render embedded images after text content
        for img_md in &slide.images {
            out.push_str(img_md);
            out.push_str("\n\n");
        }

        for ocr in &slide.image_ocr {
            crate::markup::push_ocr_blockquote(&mut out, ocr);
        }
    }

    append_notes_markdown(slides, &mut out);

    out
}

/// Render a shape's paragraphs as markdown into `out`.
///
/// `heading_offset` is added to a paragraph's inferred heading level (capped
/// at 6) so headings nest under the `## Slide N` body headings / `### Slide N`
/// notes sub-headings. Shared by the body and notes loops.
fn render_shape_markdown(shape: &ShapeText, out: &mut String, heading_offset: u8) {
    let mut prev_was_list = false;
    for para in &shape.paragraphs {
        let text = render_para_markdown(para);
        let text = text.trim();
        if text.is_empty() {
            continue;
        }

        let is_list = !matches!(&para.bullet, BulletKind::None);

        if para.heading_level > 0 && para.heading_level <= 6 {
            // Blank line after a list block before a heading
            if prev_was_list {
                out.push('\n');
            }
            let level = (para.heading_level + heading_offset).min(6);
            for _ in 0..level {
                out.push('#');
            }
            out.push(' ');
            out.push_str(text);
            out.push_str("\n\n");
        } else if is_list {
            let lvl = match &para.bullet {
                BulletKind::Bullet(l) | BulletKind::Numbered(l) => *l,
                BulletKind::None => 0,
            };
            let indent = "  ".repeat(usize::from(lvl));
            let marker = if matches!(&para.bullet, BulletKind::Numbered(_)) {
                "1."
            } else {
                "-"
            };
            out.push_str(&indent);
            out.push_str(marker);
            out.push(' ');
            out.push_str(text);
            out.push('\n');
        } else {
            // Blank line after a list block before regular text
            if prev_was_list {
                out.push('\n');
            }
            out.push_str(text);
            out.push_str("\n\n");
        }

        prev_was_list = is_list;
    }
    // If the shape ended with a list, add trailing blank line
    if prev_was_list {
        out.push('\n');
    }
}

/// Render a shape's paragraphs as plain text into `out`.
fn render_shape_plain(shape: &ShapeText, out: &mut String) {
    for para in &shape.paragraphs {
        let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        let text = text.trim();
        if !text.is_empty() {
            match &para.bullet {
                BulletKind::None => {
                    out.push_str(text);
                    out.push('\n');
                }
                BulletKind::Bullet(lvl) | BulletKind::Numbered(lvl) => {
                    let indent = "  ".repeat(usize::from(*lvl));
                    let marker = if matches!(&para.bullet, BulletKind::Numbered(_)) {
                        "1."
                    } else {
                        "-"
                    };
                    out.push_str(&indent);
                    out.push_str(marker);
                    out.push(' ');
                    out.push_str(text);
                    out.push('\n');
                }
            }
        }
    }
}

/// Whether any note shape contains non-whitespace run text.
fn notes_nonempty(notes: &[ShapeText]) -> bool {
    notes.iter().any(|shape| {
        shape
            .paragraphs
            .iter()
            .any(|para| para.runs.iter().any(|r| !r.text.trim().is_empty()))
    })
}

/// Ensure a blank line separates the body from a trailer section: adds the
/// missing newline(s) unless the output already ends with two newlines.
fn ensure_trailer_blank_line(out: &mut String) {
    if !out.is_empty() && !out.ends_with("\n\n") {
        if out.ends_with('\n') {
            out.push('\n');
        } else {
            out.push_str("\n\n");
        }
    }
}

/// Append a `## Notes` trailer listing slides with speaker notes. Only slides
/// with non-empty notes appear, under `### Slide N` sub-headings.
fn append_notes_markdown(slides: &[Slide], out: &mut String) {
    let mut any = false;
    for slide in slides {
        if !notes_nonempty(&slide.notes) {
            continue;
        }
        if !any {
            ensure_trailer_blank_line(out);
            out.push_str("## Notes\n\n");
            any = true;
        }
        let _ = write!(out, "### Slide {}\n\n", slide.number);
        for shape in &slide.notes {
            render_shape_markdown(shape, out, 3);
        }
    }
}

/// Append a `--- Notes ---` trailer listing slides with speaker notes.
fn append_notes_plain(slides: &[Slide], out: &mut String) {
    let mut any = false;
    for slide in slides {
        if !notes_nonempty(&slide.notes) {
            continue;
        }
        if !any {
            ensure_trailer_blank_line(out);
            out.push_str("--- Notes ---\n");
            any = true;
        }
        let _ = writeln!(out, "[Slide {}]", slide.number);
        for shape in &slide.notes {
            render_shape_plain(shape, out);
        }
    }
}

/// Render a paragraph's runs as markdown, handling bold/italic/hyperlinks.
fn render_para_markdown(para: &Paragraph) -> String {
    markup::render_runs_markdown(&para.runs)
}

/// Implement [`InlineRun`] for pptx `TextRun` so the shared markup renderer
/// can inspect formatting without knowing the concrete type.
impl markup::InlineRun for TextRun {
    fn text(&self) -> &str {
        &self.text
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

#[cfg(test)]
mod tests {
    use super::*;

    // ── render_para_markdown ─────────────────────────────────────

    #[test]
    fn render_plain_text_run() {
        let para = Paragraph {
            runs: vec![TextRun {
                text: "Hello".into(),
                bold: false,
                italic: false,
                link_url: None,
                font_size: None,
            }],
            heading_level: 0,
            bullet: BulletKind::None,
        };
        assert_eq!(render_para_markdown(&para), "Hello");
    }

    #[test]
    fn render_bold_run() {
        let para = Paragraph {
            runs: vec![TextRun {
                text: "Important".into(),
                bold: true,
                italic: false,
                link_url: None,
                font_size: None,
            }],
            heading_level: 0,
            bullet: BulletKind::None,
        };
        assert_eq!(render_para_markdown(&para), "**Important**");
    }

    #[test]
    fn render_hyperlink_run() {
        let para = Paragraph {
            runs: vec![TextRun {
                text: "click me".into(),
                bold: false,
                italic: false,
                link_url: Some("https://example.com".into()),
                font_size: None,
            }],
            heading_level: 0,
            bullet: BulletKind::None,
        };
        assert_eq!(
            render_para_markdown(&para),
            "[click me](https://example.com)"
        );
    }

    #[test]
    fn render_bold_hyperlink_run() {
        let para = Paragraph {
            runs: vec![TextRun {
                text: "link".into(),
                bold: true,
                italic: false,
                link_url: Some("https://example.com".into()),
                font_size: None,
            }],
            heading_level: 0,
            bullet: BulletKind::None,
        };
        assert_eq!(
            render_para_markdown(&para),
            "[**link**](https://example.com)"
        );
    }

    // ── heading inference ────────────────────────────────────────

    #[test]
    fn heading_from_large_font() {
        let slides = vec![Slide {
            number: 1,
            shapes: vec![ShapeText {
                paragraphs: vec![
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Title".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: Some(2800),
                        }],
                        heading_level: 1,
                        bullet: BulletKind::None,
                    },
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Body text".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: Some(1800),
                        }],
                        heading_level: 0,
                        bullet: BulletKind::None,
                    },
                ],
            }],
            images: Vec::new(),
            image_ocr: Vec::new(),
            notes: vec![],
        }];

        let md = render_markdown(&slides);
        assert!(md.contains("# Title"));
        assert!(md.contains("Body text"));
    }

    // ── render_plain ─────────────────────────────────────────────

    #[test]
    fn plain_single_slide() {
        let slides = vec![Slide {
            number: 1,
            shapes: vec![ShapeText {
                paragraphs: vec![Paragraph {
                    runs: vec![TextRun {
                        text: "Hello World".into(),
                        bold: false,
                        italic: false,
                        link_url: None,
                        font_size: None,
                    }],
                    heading_level: 0,
                    bullet: BulletKind::None,
                }],
            }],
            images: Vec::new(),
            image_ocr: Vec::new(),
            notes: vec![],
        }];

        let text = render_plain(&slides);
        assert_eq!(text, "Hello World\n");
    }

    #[test]
    fn plain_multi_slide() {
        let slides = vec![
            Slide {
                number: 1,
                shapes: vec![ShapeText {
                    paragraphs: vec![Paragraph {
                        runs: vec![TextRun {
                            text: "Slide one".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::None,
                    }],
                }],
                images: Vec::new(),
                image_ocr: Vec::new(),
                notes: vec![],
            },
            Slide {
                number: 2,
                shapes: vec![ShapeText {
                    paragraphs: vec![Paragraph {
                        runs: vec![TextRun {
                            text: "Slide two".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::None,
                    }],
                }],
                images: Vec::new(),
                image_ocr: Vec::new(),
                notes: vec![],
            },
        ];

        let text = render_plain(&slides);
        assert!(text.contains("--- Slide 1 ---"));
        assert!(text.contains("Slide one"));
        assert!(text.contains("--- Slide 2 ---"));
        assert!(text.contains("Slide two"));
    }

    // ── render_markdown multi-slide ──────────────────────────────

    #[test]
    fn markdown_multi_slide_headings_offset() {
        let slides = vec![
            Slide {
                number: 1,
                shapes: vec![ShapeText {
                    paragraphs: vec![Paragraph {
                        runs: vec![TextRun {
                            text: "Title".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: Some(2800),
                        }],
                        heading_level: 1,
                        bullet: BulletKind::None,
                    }],
                }],
                images: Vec::new(),
                image_ocr: Vec::new(),
                notes: vec![],
            },
            Slide {
                number: 2,
                shapes: vec![ShapeText {
                    paragraphs: vec![Paragraph {
                        runs: vec![TextRun {
                            text: "Another".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::None,
                    }],
                }],
                images: Vec::new(),
                image_ocr: Vec::new(),
                notes: vec![],
            },
        ];

        let md = render_markdown(&slides);
        // Multi-slide: slide headings are ##, shape headings offset to ###
        assert!(md.contains("## Slide 1"));
        assert!(md.contains("### Title"));
        assert!(md.contains("## Slide 2"));
    }

    // ── parse_slide_xml ──────────────────────────────────────────

    #[test]
    fn parse_slide_basic_shape() {
        let xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr lang="en-US" b="1"/>
                                    <a:t>Hello World</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"#;

        let rels = HashMap::new();
        let shapes = parse_slide_xml(xml, &rels);
        assert_eq!(shapes.len(), 1);
        assert_eq!(shapes[0].paragraphs.len(), 1);
        assert_eq!(shapes[0].paragraphs[0].runs[0].text, "Hello World");
    }

    #[test]
    fn parse_slide_with_hyperlink() {
        let xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr lang="en-US">
                                        <a:hlinkClick r:id="rId2"/>
                                    </a:rPr>
                                    <a:t>Click here</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"#;

        let rels: HashMap<String, String> = [("rId2".into(), "https://example.com".into())].into();
        let shapes = parse_slide_xml(xml, &rels);
        assert_eq!(shapes.len(), 1);
        assert_eq!(
            shapes[0].paragraphs[0].runs[0].link_url.as_deref(),
            Some("https://example.com")
        );
    }

    #[test]
    fn parse_slide_empty() {
        let xml = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree/></p:cSld>
        </p:sld>"#;
        let shapes = parse_slide_xml(xml, &HashMap::new());
        assert!(shapes.is_empty());
    }

    // ── bullet parsing ────────────────────────────────────────────

    #[test]
    fn parse_slide_bullet_char() {
        let xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree><p:sp><p:txBody>
                <a:p>
                    <a:pPr lvl="0"><a:buChar char="●"/></a:pPr>
                    <a:r><a:t>Top level</a:t></a:r>
                </a:p>
                <a:p>
                    <a:pPr lvl="1"><a:buChar char="○"/></a:pPr>
                    <a:r><a:t>Sub item</a:t></a:r>
                </a:p>
            </p:txBody></p:sp></p:spTree></p:cSld>
        </p:sld>"#;

        let shapes = parse_slide_xml(xml, &HashMap::new());
        assert_eq!(shapes[0].paragraphs[0].bullet, BulletKind::Bullet(0));
        assert_eq!(shapes[0].paragraphs[1].bullet, BulletKind::Bullet(1));
    }

    #[test]
    fn parse_slide_bullet_auto_num() {
        let xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree><p:sp><p:txBody>
                <a:p>
                    <a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr>
                    <a:r><a:t>First</a:t></a:r>
                </a:p>
            </p:txBody></p:sp></p:spTree></p:cSld>
        </p:sld>"#;

        let shapes = parse_slide_xml(xml, &HashMap::new());
        assert_eq!(shapes[0].paragraphs[0].bullet, BulletKind::Numbered(0));
    }

    #[test]
    fn parse_slide_bu_none_not_bullet() {
        let xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree><p:sp><p:txBody>
                <a:p>
                    <a:pPr lvl="0"><a:buNone/></a:pPr>
                    <a:r><a:t>Not a bullet</a:t></a:r>
                </a:p>
            </p:txBody></p:sp></p:spTree></p:cSld>
        </p:sld>"#;

        let shapes = parse_slide_xml(xml, &HashMap::new());
        assert_eq!(shapes[0].paragraphs[0].bullet, BulletKind::None);
    }

    // ── bullet rendering (markdown) ────────────────────────────────

    #[test]
    fn render_bullets_markdown() {
        let slides = vec![Slide {
            number: 1,
            shapes: vec![ShapeText {
                paragraphs: vec![
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Attitude".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Bullet(0),
                    },
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Excited about tech".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Bullet(1),
                    },
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Making impact".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Bullet(1),
                    },
                ],
            }],
            images: Vec::new(),
            image_ocr: Vec::new(),
            notes: vec![],
        }];

        let md = render_markdown(&slides);
        assert!(md.contains("- Attitude\n"));
        assert!(md.contains("  - Excited about tech\n"));
        assert!(md.contains("  - Making impact\n"));
    }

    #[test]
    fn render_numbered_list_markdown() {
        let slides = vec![Slide {
            number: 1,
            shapes: vec![ShapeText {
                paragraphs: vec![
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Step one".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Numbered(0),
                    },
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Step two".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Numbered(0),
                    },
                ],
            }],
            images: Vec::new(),
            image_ocr: Vec::new(),
            notes: vec![],
        }];

        let md = render_markdown(&slides);
        assert!(md.contains("1. Step one\n"));
        assert!(md.contains("1. Step two\n"));
    }

    #[test]
    fn render_bullets_plain() {
        let slides = vec![Slide {
            number: 1,
            shapes: vec![ShapeText {
                paragraphs: vec![
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Top".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Bullet(0),
                    },
                    Paragraph {
                        runs: vec![TextRun {
                            text: "Sub".into(),
                            bold: false,
                            italic: false,
                            link_url: None,
                            font_size: None,
                        }],
                        heading_level: 0,
                        bullet: BulletKind::Bullet(1),
                    },
                ],
            }],
            images: Vec::new(),
            image_ocr: Vec::new(),
            notes: vec![],
        }];

        let text = render_plain(&slides);
        assert!(text.contains("- Top\n"));
        assert!(text.contains("  - Sub\n"));
    }

    // ── notes trailer rendering ─────────────────────────────────

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
                image_ocr: Vec::new(),
                notes: vec![],
            },
            Slide {
                number: 2,
                shapes: vec![shape_with_text("Body")],
                images: vec![],
                image_ocr: Vec::new(),
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
            image_ocr: Vec::new(),
            notes: vec![shape_with_text("aside")],
        }];
        let plain = render_plain(&slides);
        assert!(plain.contains("Hi"));
        let notes_at = plain.find("--- Notes ---").expect("notes section");
        let body_at = plain.find("Hi").unwrap();
        assert!(notes_at > body_at);
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
                image_ocr: Vec::new(),
                notes: vec![],
            },
            Slide {
                number: 2,
                shapes: vec![],
                images: vec![],
                image_ocr: Vec::new(),
                notes: vec![shape_with_text("orphaned notes")],
            },
        ];
        let md = render_markdown(&slides);
        let notes_at = md.find("## Notes").expect("notes section");
        // No body heading for the notes-only slide (### Slide 2 sub-heading
        // would match a naive `contains("## Slide 2")`, hence the region cut).
        assert!(!md[..notes_at].contains("## Slide 2"));
        assert!(md[notes_at..].contains("### Slide 2"));
        assert!(md[notes_at..].contains("orphaned notes"));
    }

    #[test]
    fn render_markdown_blank_line_before_notes_after_list() {
        // Body's last shape ends in a bullet list, which leaves the body
        // ending with a single '\n' before ensure_trailer_blank_line runs.
        // The trailer must be separated by exactly one blank line.
        let slides = vec![Slide {
            number: 1,
            shapes: vec![ShapeText {
                paragraphs: vec![Paragraph {
                    runs: vec![TextRun {
                        text: "list item".into(),
                        bold: false,
                        italic: false,
                        link_url: None,
                        font_size: None,
                    }],
                    heading_level: 0,
                    bullet: BulletKind::Bullet(0),
                }],
            }],
            images: vec![],
            image_ocr: Vec::new(),
            notes: vec![shape_with_text("speaker notes")],
        }];
        let md = render_markdown(&slides);
        assert!(md.contains("list item\n\n## Notes"), "md: {md:?}");
    }

    #[test]
    fn render_whitespace_only_notes_omit_section() {
        // Notes whose run text is all whitespace count as empty: no Notes
        // trailer in either renderer, but the body still renders.
        let slides = vec![Slide {
            number: 1,
            shapes: vec![shape_with_text("Body")],
            images: vec![],
            image_ocr: Vec::new(),
            notes: vec![shape_with_text("   ")],
        }];
        let md = render_markdown(&slides);
        let plain = render_plain(&slides);
        assert!(!md.contains("## Notes"), "md: {md:?}");
        assert!(!plain.contains("--- Notes ---"), "plain: {plain:?}");
        assert!(md.contains("Body"));
        assert!(plain.contains("Body"));
    }

    #[test]
    fn render_no_notes_omits_section() {
        let slides = vec![Slide {
            number: 1,
            shapes: vec![shape_with_text("Hi")],
            images: vec![],
            image_ocr: Vec::new(),
            notes: vec![],
        }];
        assert!(!render_markdown(&slides).contains("## Notes"));
        assert!(!render_plain(&slides).contains("--- Notes ---"));
    }

    #[test]
    fn render_markdown_slide_with_ocr_text() {
        let slides = vec![Slide {
            number: 1,
            shapes: Vec::new(),
            images: Vec::new(),
            notes: Vec::new(),
            image_ocr: vec!["BATDOC\nOCR 123".into()],
        }];
        let md = render_markdown(&slides);
        assert!(md.contains("> BATDOC"), "got: {md}");
        assert!(md.contains("> OCR 123"), "got: {md}");
    }

    #[test]
    fn render_plain_slide_with_ocr_text() {
        let slides = vec![Slide {
            number: 1,
            shapes: Vec::new(),
            images: Vec::new(),
            notes: Vec::new(),
            image_ocr: vec!["line one\nline two".into()],
        }];
        let text = render_plain(&slides);
        assert!(text.contains("line one"), "got: {text}");
        assert!(text.contains("line two"), "got: {text}");
    }

    // ── ZIP integration: speaker notes discovery ──────────────────

    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn zip_entry(z: &mut ZipWriter<Cursor<Vec<u8>>>, name: &str, body: &str) {
        z.start_file(name, SimpleFileOptions::default()).unwrap();
        z.write_all(body.as_bytes()).unwrap();
    }

    /// Minimal one-slide pptx whose notes slide's `body` placeholder says
    /// `"Speak slowly"` alongside a `sldImg` placeholder and a freeform shape.
    fn minimal_pptx_with_notes() -> Vec<u8> {
        minimal_pptx("Speak slowly", "../notesSlides/notesSlide1.xml")
    }

    /// Same deck, but the notes `body` placeholder text is `notes_body`.
    fn minimal_pptx_with_notes_body(notes_body: &str) -> Vec<u8> {
        minimal_pptx(notes_body, "../notesSlides/notesSlide1.xml")
    }

    /// Build a one-slide pptx whose slide rels point the notesSlide at
    /// `notes_target`, with `notes_body` as the body placeholder text.
    fn minimal_pptx(notes_body: &str, notes_target: &str) -> Vec<u8> {
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
            &format!(
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="{notes_target}"/>
</Relationships>"#
            ),
        );
        zip_entry(
            &mut z,
            "ppt/notesSlides/notesSlide1.xml",
            &format!(
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
      <p:txBody><a:p><a:r><a:t>{notes_body}</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:txBody><a:p><a:r><a:t>freeform ignored</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>"#
            ),
        );
        z.finish().unwrap().into_inner()
    }

    #[test]
    fn extract_markdown_includes_speaker_notes() {
        let data = minimal_pptx_with_notes();
        let md = extract_markdown(&data, false, false).unwrap();
        assert!(md.contains("Deck title"), "md: {md:?}");
        assert!(md.contains("## Notes"), "md: {md:?}");
        assert!(md.contains("Speak slowly"), "md: {md:?}");
        assert!(!md.contains("SHOULD NOT APPEAR"), "md: {md:?}");
        assert!(!md.contains("freeform ignored"), "md: {md:?}");
    }

    #[test]
    fn extract_whitespace_notes_body_omits_section() {
        let data = minimal_pptx_with_notes_body("   ");
        let md = extract_markdown(&data, false, false).unwrap();
        let plain = extract_plain(&data, false).unwrap();
        assert!(!md.contains("## Notes"), "md: {md:?}");
        assert!(!plain.contains("--- Notes ---"), "plain: {plain:?}");
    }

    #[test]
    fn extract_missing_notes_target_ok() {
        let data = minimal_pptx("Speak slowly", "../notesSlides/missing.xml");
        let md = extract_markdown(&data, false, false).unwrap();
        assert!(md.contains("Deck title"), "md: {md:?}");
        assert!(!md.contains("## Notes"), "md: {md:?}");
    }

    #[test]
    fn extract_plain_includes_speaker_notes() {
        let data = minimal_pptx_with_notes();
        let plain = extract_plain(&data, false).unwrap();
        assert!(plain.contains("--- Notes ---"), "plain: {plain:?}");
        assert!(plain.contains("[Slide 1]"), "plain: {plain:?}");
        assert!(plain.contains("Speak slowly"), "plain: {plain:?}");
    }
}
