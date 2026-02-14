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
pub(crate) fn extract_plain(data: &[u8]) -> crate::error::Result<String> {
    let slides = parse_pptx(data)?;
    Ok(render_plain(&slides))
}

/// Extract markdown-formatted text from a .pptx file.
pub(crate) fn extract_markdown(data: &[u8]) -> crate::error::Result<String> {
    let slides = parse_pptx(data)?;
    Ok(render_markdown(&slides))
}

// ── Parsing ────────────────────────────────────────────────────────

/// Parse the pptx archive into a list of slides.
fn parse_pptx(data: &[u8]) -> crate::error::Result<Vec<Slide>> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    // Discover slides from presentation.xml + rels
    let slide_paths = discover_slides(&mut archive)?;

    let mut slides = Vec::new();
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

        let shapes = parse_slide_xml(&xml, &rels);
        slides.push(Slide {
            number: num,
            shapes,
        });
    }

    Ok(slides)
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
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"txBody" {
                    parse_text_body(reader, rels, &mut paragraphs);
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    break;
                }
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

/// Parse a `<p:txBody>` (or `<a:txBody>`) element.
fn parse_text_body(reader: &mut Reader<&[u8]>, rels: &Rels, paragraphs: &mut Vec<Paragraph>) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"p" {
                    let para = parse_para(reader, rels);
                    if !para.runs.is_empty() {
                        paragraphs.push(para);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"txBody" {
                    break;
                }
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
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"p" {
                    break;
                }
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
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_name {
                    break;
                }
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
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"hlinkClick" {
                    if let Some(rid) = get_attr(e, b"r:id") {
                        if let Some(url) = rels.get(&rid) {
                            *link_url = Some(url.clone());
                        }
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
        if slide.shapes.is_empty() {
            continue;
        }

        if multiple {
            if i > 0 {
                out.push('\n');
            }
            let _ = writeln!(out, "--- Slide {} ---", slide.number);
        }

        for shape in &slide.shapes {
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
    }

    out
}

/// Render slides as markdown.
fn render_markdown(slides: &[Slide]) -> String {
    let mut out = String::new();
    let multiple = slides.len() > 1;

    for slide in slides {
        if slide.shapes.is_empty() {
            continue;
        }

        if multiple {
            let _ = write!(out, "## Slide {}\n\n", slide.number);
        }

        let mut first_shape = true;
        for shape in &slide.shapes {
            if !first_shape {
                out.push('\n');
            }
            first_shape = false;

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
                    let level = if multiple {
                        (para.heading_level + 2).min(6)
                    } else {
                        para.heading_level
                    };
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
    }

    out
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
        }];

        let text = render_plain(&slides);
        assert!(text.contains("- Top\n"));
        assert!(text.contains("  - Sub\n"));
    }
}
