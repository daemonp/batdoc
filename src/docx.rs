//! OOXML `.docx` (Office Open XML) format parser.
//!
//! Unzips the `.docx` archive, parses `word/document.xml` with `quick-xml`
//! into structured [`Block`] types (paragraphs with heading/list styles and
//! runs with bold/italic/hyperlink, tables with rows and cells), then renders
//! to either plain text or markdown.

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::io::{Cursor, Read};
use zip::ZipArchive;

use crate::markup;
use crate::xml_util::{self, get_attr, Rels};

/// Extracted document structure for rich output.
#[derive(Debug)]
enum Block {
    Paragraph { style: ParaStyle, runs: Vec<Run> },
    Table { rows: Vec<Row> }, // rows -> cells -> blocks
}

#[derive(Debug, Clone, Default)]
struct ParaStyle {
    heading_level: u8, // 0 = normal, 1-9 = heading
    list_level: Option<u8>,
}

#[derive(Debug, Clone)]
struct Run {
    text: String,
    bold: bool,
    italic: bool,
    /// If this run is inside a hyperlink, the resolved URL.
    link_url: Option<String>,
}

/// A single table cell containing blocks.
type Cell = Vec<Block>;
/// A table row: a sequence of cells.
type Row = Vec<Cell>;

/// Extract plain text from a .docx file.
pub(crate) fn extract_plain(data: &[u8]) -> crate::error::Result<String> {
    let blocks = parse_docx(data)?;
    Ok(render_plain(&blocks))
}

/// Extract markdown-formatted text from a .docx file.
pub(crate) fn extract_markdown(data: &[u8]) -> crate::error::Result<String> {
    let blocks = parse_docx(data)?;
    Ok(render_markdown(&blocks))
}

/// Parse the docx XML into structured blocks.
fn parse_docx(data: &[u8]) -> crate::error::Result<Vec<Block>> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    // Load hyperlink relationships (rId → URL)
    let rels = xml_util::load_rels(&mut archive, "word/_rels/document.xml.rels");

    let mut xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut xml)?;

    let mut reader = Reader::from_str(&xml);
    let mut blocks = Vec::new();
    let mut in_body = false;

    parse_body(&mut reader, &mut blocks, &mut in_body, &rels);

    Ok(blocks)
}

/// Walk the XML and collect blocks from the document body.
fn parse_body(
    reader: &mut Reader<&[u8]>,
    blocks: &mut Vec<Block>,
    in_body: &mut bool,
    rels: &Rels,
) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"body" => *in_body = true,
                    b"p" if *in_body => {
                        let block = parse_paragraph(reader, rels);
                        blocks.push(block);
                    }
                    b"tbl" if *in_body => {
                        let table = parse_table(reader, rels);
                        blocks.push(table);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"body" {
                    *in_body = false;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a `<w:p>` element into a `Block::Paragraph`.
fn parse_paragraph(reader: &mut Reader<&[u8]>, rels: &Rels) -> Block {
    let mut style = ParaStyle::default();
    let mut runs: Vec<Run> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"pPr" => parse_para_props(reader, &mut style),
                    b"r" => {
                        if let Some(run) = parse_run(reader) {
                            runs.push(run);
                        }
                    }
                    b"hyperlink" => {
                        // Resolve the hyperlink URL from r:id → rels map
                        let url = get_attr(e, b"r:id").and_then(|rid| rels.get(&rid).cloned());
                        parse_hyperlink_runs(reader, &mut runs, url.as_deref());
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"p" {
                    break;
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tab" {
                    runs.push(Run {
                        text: "\t".into(),
                        bold: false,
                        italic: false,
                        link_url: None,
                    });
                } else if name.as_ref() == b"br" {
                    runs.push(Run {
                        text: "\n".into(),
                        bold: false,
                        italic: false,
                        link_url: None,
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Block::Paragraph { style, runs }
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

/// Parse a `<w:r>` element into a `Run`.
fn parse_run(reader: &mut Reader<&[u8]>) -> Option<Run> {
    let mut bold = false;
    let mut italic = false;
    let mut text = String::new();

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
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tab" {
                    text.push('\t');
                } else if name.as_ref() == b"br" {
                    text.push('\n');
                } else if name.as_ref() == b"b" || name.as_ref() == b"bCs" {
                    // Self-closing <w:b/> in rPr means bold on
                    bold = true;
                } else if name.as_ref() == b"i" || name.as_ref() == b"iCs" {
                    italic = true;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"r" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if text.is_empty() {
        None
    } else {
        Some(Run {
            text,
            bold,
            italic,
            link_url: None,
        })
    }
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
fn parse_hyperlink_runs(reader: &mut Reader<&[u8]>, runs: &mut Vec<Run>, url: Option<&str>) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"r" {
                    if let Some(mut run) = parse_run(reader) {
                        run.link_url = url.map(String::from);
                        runs.push(run);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"hyperlink" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a `<w:tbl>` element into a `Block::Table`.
fn parse_table(reader: &mut Reader<&[u8]>, rels: &Rels) -> Block {
    let mut rows: Vec<Row> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tr" {
                    let row = parse_table_row(reader, rels);
                    rows.push(row);
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"tbl" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Block::Table { rows }
}

/// Parse a `<w:tr>` element into a row of cells.
fn parse_table_row(reader: &mut Reader<&[u8]>, rels: &Rels) -> Row {
    let mut cells: Row = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"tc" {
                    let cell = parse_table_cell(reader, rels);
                    cells.push(cell);
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"tr" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    cells
}

/// Parse a `<w:tc>` element into a list of blocks.
fn parse_table_cell(reader: &mut Reader<&[u8]>, rels: &Rels) -> Cell {
    let mut blocks = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"p" => blocks.push(parse_paragraph(reader, rels)),
                    b"tbl" => blocks.push(parse_table(reader, rels)), // nested table
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"tc" {
                    break;
                }
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

/// Extract text content from a cell's blocks, joining paragraphs with spaces.
fn cell_to_text(cell: &[Block], use_markdown: bool) -> String {
    cell.iter()
        .filter_map(|b| match b {
            Block::Paragraph { runs, .. } => {
                let t = if use_markdown {
                    render_runs_markdown(runs)
                } else {
                    runs.iter().map(|r| r.text.as_str()).collect::<String>()
                };
                let t = t.trim().to_string();
                if t.is_empty() {
                    None
                } else {
                    Some(t)
                }
            }
            Block::Table { .. } => None,
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
            let text: String = runs.iter().map(|r| r.text.as_str()).collect();
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

/// Render runs with markdown inline formatting (bold/italic/hyperlinks).
///
/// Adjacent runs sharing the same `link_url` are grouped so the markdown
/// link wraps the entire visible text: `[text](url)` instead of producing
/// separate `[part1](url)[part2](url)` fragments.
fn render_runs_markdown(runs: &[Run]) -> String {
    markup::render_runs_markdown(runs)
}

/// Implement [`InlineRun`] for docx `Run` so the shared markup renderer
/// can inspect formatting without knowing the concrete type.
impl markup::InlineRun for Run {
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
            text: text.into(),
            bold,
            italic,
            link_url: None,
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
            text: "click here".into(),
            bold: false,
            italic: false,
            link_url: Some("https://example.com".into()),
        }];
        assert_eq!(
            render_runs_markdown(&runs),
            "[click here](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_bold() {
        let runs = vec![Run {
            text: "bold link".into(),
            bold: true,
            italic: false,
            link_url: Some("https://example.com".into()),
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
                text: "part ".into(),
                bold: false,
                italic: false,
                link_url: Some("https://example.com".into()),
            },
            Run {
                text: "one".into(),
                bold: true,
                italic: false,
                link_url: Some("https://example.com".into()),
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
                text: "this link".into(),
                bold: false,
                italic: false,
                link_url: Some("https://example.com".into()),
            },
            run(" for details", false, false),
        ];
        assert_eq!(
            render_runs_markdown(&runs),
            "See [this link](https://example.com) for details"
        );
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
}
