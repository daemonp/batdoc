//! OOXML `.xlsx` (Excel) spreadsheet parser.
//!
//! Unzips the `.xlsx` archive, parses the shared string table and each
//! worksheet's XML, then renders every sheet as either tab-separated
//! plain text or a markdown table with a heading per sheet. Hyperlinks
//! are resolved from sheet relationship files and rendered as markdown links.

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::io::{Cursor, Read};
use zip::ZipArchive;

use crate::dateconv;
use crate::sheet::Sheet;
use crate::xml_util::{self, get_attr, Rels};

/// Extract plain text (TSV) from an .xlsx file.
pub(crate) fn extract_plain(data: &[u8]) -> crate::error::Result<String> {
    let sheets = parse_xlsx(data)?;
    Ok(crate::sheet::render_plain(&sheets))
}

/// Extract markdown-formatted text from an .xlsx file.
///
/// When `images` is true, embedded images from drawings are extracted
/// and appended as reference-style base64 images with definitions at the end.
pub(crate) fn extract_markdown(data: &[u8], images: bool) -> crate::error::Result<String> {
    let sheets = parse_xlsx(data)?;
    let mut md = crate::sheet::render_markdown(&sheets);

    if images {
        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor)?;
        let sheet_info = discover_sheets(&mut archive)?;
        append_sheet_images(&mut md, &sheet_info, &mut archive);
    }

    Ok(md)
}

// ── Parsing ────────────────────────────────────────────────────────

/// Parse the xlsx archive into a list of sheets.
fn parse_xlsx(data: &[u8]) -> crate::error::Result<Vec<Sheet>> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    // 1. Load shared strings table (optional — some files use inline strings)
    let shared_strings = parse_shared_strings(&mut archive);

    // 2. Load styles (for date format detection)
    let styles = parse_styles(&mut archive);

    // 3. Discover sheets: name + file path
    let sheet_info = discover_sheets(&mut archive)?;

    // 4. Parse each sheet
    let mut sheets = Vec::new();
    for (name, path) in &sheet_info {
        let mut xml = String::new();
        match archive.by_name(path) {
            Ok(mut entry) => {
                entry.read_to_string(&mut xml)?;
            }
            Err(_) => continue,
        }

        // Load hyperlink relationships for this sheet
        let sheet_rels_path = xml_util::rels_path(path);
        let rels = xml_util::load_rels(&mut archive, &sheet_rels_path);

        let mut rows = parse_sheet_xml(&xml, &shared_strings, &styles);

        // Apply hyperlinks: parse <hyperlinks> from sheet XML and
        // resolve URLs from the rels map
        apply_hyperlinks(&xml, &rels, &mut rows);

        sheets.push(Sheet {
            name: name.clone(),
            rows,
        });
    }

    Ok(sheets)
}

// ── Style / date format detection ──────────────────────────────────

/// Resolved style information: for each cell style index (`s` attribute),
/// whether the number format is a date format.
#[derive(Debug, Default)]
struct Styles {
    /// For each xf index, true if the numFmtId is a date format.
    is_date: Vec<bool>,
}

impl Styles {
    /// Check if a cell style index corresponds to a date format.
    fn is_date_style(&self, style_idx: usize) -> bool {
        self.is_date.get(style_idx).copied().unwrap_or(false)
    }
}

/// Parse `xl/styles.xml` to determine which cell styles are date formats.
///
/// Reads `<numFmt>` elements for custom format strings and `<xf>` elements
/// in `<cellXfs>` for the numFmtId associated with each style index.
fn parse_styles(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Styles {
    let mut xml = String::new();
    match archive.by_name("xl/styles.xml") {
        Ok(mut entry) => {
            if entry.read_to_string(&mut xml).is_err() {
                return Styles::default();
            }
        }
        Err(_) => return Styles::default(),
    }

    parse_styles_xml(&xml)
}

/// Parse styles XML into resolved style info (separated for testability).
fn parse_styles_xml(xml: &str) -> Styles {
    let mut reader = Reader::from_str(xml);

    // Custom number formats: numFmtId → format string
    let mut custom_formats: Vec<(u16, String)> = Vec::new();
    // Cell xf entries: each entry's numFmtId
    let mut cell_xf_fmt_ids: Vec<u16> = Vec::new();
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"cellXfs" => in_cell_xfs = true,
                    b"xf" if in_cell_xfs => {
                        let fmt_id: u16 = get_attr(e, b"numFmtId")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                        cell_xf_fmt_ids.push(fmt_id);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"numFmt" => {
                        if let (Some(id_str), Some(code)) =
                            (get_attr(e, b"numFmtId"), get_attr(e, b"formatCode"))
                        {
                            if let Ok(id) = id_str.parse::<u16>() {
                                custom_formats.push((id, code));
                            }
                        }
                    }
                    b"xf" if in_cell_xfs => {
                        let fmt_id: u16 = get_attr(e, b"numFmtId")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                        cell_xf_fmt_ids.push(fmt_id);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => {
                in_cell_xfs = false;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Styles {
        is_date: dateconv::resolve_date_styles(&cell_xf_fmt_ids, &custom_formats),
    }
}

// ── Hyperlink resolution ────────────────────────────────────────────

/// Parse `<hyperlinks>` from a sheet XML and apply URLs to cell values.
///
/// Each `<hyperlink ref="A1" r:id="rId1"/>` maps a cell reference to
/// a relationship ID. We resolve the rId to a URL and wrap the existing
/// cell value as `[value](url)`.
fn apply_hyperlinks(xml: &str, rels: &Rels, rows: &mut [Vec<String>]) {
    if rels.is_empty() {
        return;
    }

    let mut reader = Reader::from_str(xml);
    let mut in_hyperlinks = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hyperlinks" => {
                in_hyperlinks = true;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"hyperlinks" => {
                break;
            }
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if in_hyperlinks && e.local_name().as_ref() == b"hyperlink" =>
            {
                let cell_ref = get_attr(e, b"ref").unwrap_or_default();
                let rid = get_attr(e, b"r:id").unwrap_or_default();

                if let Some(url) = rels.get(&rid) {
                    let col = col_ref_to_index(&cell_ref);
                    let row_num = cell_ref_to_row(&cell_ref);
                    if let Some(row) = rows.get_mut(row_num) {
                        if let Some(cell) = row.get_mut(col) {
                            if !cell.is_empty() {
                                *cell = format!("[{cell}]({url})");
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

// ── Image extraction ─────────────────────────────────────────────

/// Append embedded images from drawing overlays to the markdown output.
///
/// For each sheet that has a drawing relationship, parses the drawing XML
/// to find `<a:blip>` references, reads the images from the ZIP, and
/// appends them as reference-style markdown images. Inline refs go in the
/// text flow; definitions are collected and appended at the document end.
fn append_sheet_images(
    md: &mut String,
    sheet_info: &[(String, String)],
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) {
    let mut image_counter = 0usize;
    let mut inline_refs = Vec::new();
    let mut definitions = Vec::new();

    for (_name, path) in sheet_info {
        // Load the sheet's relationships to find drawing references
        let sheet_rels_path = xml_util::rels_path(path);
        let mut rels_xml = String::new();
        if let Ok(mut entry) = archive.by_name(&sheet_rels_path) {
            let _ = entry.read_to_string(&mut rels_xml);
        } else {
            continue;
        }

        // Find drawing relationships (Type ends with /drawing)
        let drawing_targets = parse_drawing_rels(&rels_xml);
        if drawing_targets.is_empty() {
            continue;
        }

        let base_dir = path.rsplit_once('/').map_or("xl", |(dir, _)| dir);

        for drawing_target in &drawing_targets {
            // Resolve drawing path relative to the sheet
            let drawing_path = if drawing_target.starts_with('/') {
                drawing_target.trim_start_matches('/').to_string()
            } else {
                let raw = format!("{base_dir}/{drawing_target}");
                // Normalize ../
                normalize_dotdot(&raw)
            };

            // Read drawing XML
            let mut drawing_xml = String::new();
            if let Ok(mut entry) = archive.by_name(&drawing_path) {
                let _ = entry.read_to_string(&mut drawing_xml);
            } else {
                continue;
            }

            // Load image rels for the drawing
            let drawing_rels_path = xml_util::rels_path(&drawing_path);
            let image_rels = xml_util::load_image_rels(archive, &drawing_rels_path);
            if image_rels.is_empty() {
                continue;
            }

            // Extract blip rIds from drawing XML
            let rids = parse_drawing_blip_rids(&drawing_xml);
            let drawing_base = drawing_path.rsplit_once('/').map_or("xl", |(dir, _)| dir);

            for rid in &rids {
                if let Some(target) = image_rels.get(rid) {
                    if let Some(data) = xml_util::read_image_from_zip(archive, target, drawing_base)
                    {
                        image_counter += 1;
                        let id = format!("image{image_counter}");
                        if let Some(img_ref) = crate::markup::image_to_base64_ref(&data, &id) {
                            inline_refs.push(img_ref.inline);
                            definitions.push(img_ref.definition);
                        }
                    }
                }
            }
        }
    }

    // Append inline references in the text flow
    for inline in &inline_refs {
        md.push_str(inline);
        md.push_str("\n\n");
    }

    // Append definitions at the end
    for def in &definitions {
        md.push_str(def);
        md.push('\n');
    }
}

/// Parse relationships XML to find drawing targets.
fn parse_drawing_rels(xml: &str) -> Vec<String> {
    let mut targets = Vec::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let rel_type = get_attr(e, b"Type").unwrap_or_default();
                let target = get_attr(e, b"Target").unwrap_or_default();
                if rel_type.ends_with("/drawing") && !target.is_empty() {
                    targets.push(target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    targets
}

/// Extract blip rIds from drawing XML (`<a:blip r:embed="rIdN"/>`).
fn parse_drawing_blip_rids(xml: &str) -> Vec<String> {
    let mut rids = Vec::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"blip" {
                    if let Some(rid) = get_attr(e, b"r:embed") {
                        rids.push(rid);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    rids
}

/// Normalize a path by resolving `..` segments.
fn normalize_dotdot(path: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for segment in path.split('/') {
        if segment == ".." {
            parts.pop();
        } else if !segment.is_empty() && segment != "." {
            parts.push(segment);
        }
    }
    parts.join("/")
}

/// Extract the 0-based row number from a cell reference like "B3" → 2.
fn cell_ref_to_row(cell_ref: &str) -> usize {
    let digits: String = cell_ref
        .chars()
        .skip_while(char::is_ascii_alphabetic)
        .collect();
    digits.parse::<usize>().unwrap_or(1).saturating_sub(1)
}

/// Parse `xl/sharedStrings.xml` into a lookup table.
///
/// Each `<si>` element contributes one string at its positional index.
/// Strings may be plain `<t>` text or rich text with multiple `<r><t>` runs.
fn parse_shared_strings(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let mut xml = String::new();
    match archive.by_name("xl/sharedStrings.xml") {
        Ok(mut entry) => {
            if entry.read_to_string(&mut xml).is_err() {
                return Vec::new();
            }
        }
        Err(_) => return Vec::new(),
    }

    parse_shared_strings_xml(&xml)
}

/// Parse shared string table XML into a list of strings.
///
/// Separated from `parse_shared_strings` for testability (avoids needing
/// a ZIP archive in tests).
fn parse_shared_strings_xml(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut strings = Vec::new();
    let mut in_si = false;
    let mut current = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"si" {
                    in_si = true;
                    current.clear();
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"si" {
                    strings.push(std::mem::take(&mut current));
                    in_si = false;
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_si {
                    if let Ok(s) = t.unescape() {
                        current.push_str(&s);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    strings
}

/// Discover sheet names and their file paths from workbook.xml and relationships.
///
/// Returns `(sheet_name, zip_path)` pairs in workbook order.
fn discover_sheets(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> crate::error::Result<Vec<(String, String)>> {
    // Parse workbook.xml for sheet name → rId mapping
    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;

    let mut sheet_entries: Vec<(String, String)> = Vec::new(); // (name, rId)
    let mut reader = Reader::from_str(&workbook_xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"sheet" =>
            {
                let name = get_attr(e, b"name").unwrap_or_default();
                let rid = get_attr(e, b"r:id").unwrap_or_default();
                let state = get_attr(e, b"state").unwrap_or_default();
                // Skip hidden sheets
                if state != "hidden" && !name.is_empty() && !rid.is_empty() {
                    sheet_entries.push((name, rid));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Parse workbook.xml.rels for rId → Target path mapping
    let mut rels_xml = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;

    let mut rid_to_target: Vec<(String, String)> = Vec::new();
    let mut reader = Reader::from_str(&rels_xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let id = get_attr(e, b"Id").unwrap_or_default();
                let target = get_attr(e, b"Target").unwrap_or_default();
                if !id.is_empty() && !target.is_empty() {
                    rid_to_target.push((id, target));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Resolve: sheet name → zip path
    let mut result = Vec::new();
    for (name, rid) in &sheet_entries {
        if let Some((_, target)) = rid_to_target.iter().find(|(id, _)| id == rid) {
            // Target is relative to xl/, may have leading /
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("xl/{target}")
            };
            result.push((name.clone(), path));
        }
    }

    Ok(result)
}

/// Parse a single worksheet XML into a 2D grid of string values.
///
/// Handles three cell types:
/// - `t="s"`: shared string reference (value is an index into `shared_strings`)
/// - `t="inlineStr"`: inline string with `<is><t>` content
/// - Otherwise: raw value from `<v>` (numbers, dates, formulas with cached values)
///
/// Numeric cells whose style maps to a date format are converted to ISO dates.
fn parse_sheet_xml(xml: &str, shared_strings: &[String], styles: &Styles) -> Vec<Vec<String>> {
    let mut reader = Reader::from_str(xml);
    let mut sparse_rows: Vec<Vec<(usize, String)>> = Vec::new();
    let mut max_col = 0usize;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => {
                let row = parse_row(&mut reader, shared_strings, styles);
                for &(col, _) in &row {
                    if col + 1 > max_col {
                        max_col = col + 1;
                    }
                }
                sparse_rows.push(row);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // Convert sparse (col_index, value) pairs into a dense rectangular grid
    let mut rows: Vec<Vec<String>> = Vec::with_capacity(sparse_rows.len());
    for sparse_row in sparse_rows {
        let mut dense = vec![String::new(); max_col];
        for (col, val) in sparse_row {
            if col < max_col {
                dense[col] = val;
            }
        }
        rows.push(dense);
    }

    rows
}

/// Parse a `<row>` element, returning `(column_index, value)` pairs.
fn parse_row(
    reader: &mut Reader<&[u8]>,
    shared_strings: &[String],
    styles: &Styles,
) -> Vec<(usize, String)> {
    let mut cells: Vec<(usize, String)> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c" => {
                let col_idx = get_attr(e, b"r")
                    .as_deref()
                    .map_or(cells.len(), col_ref_to_index);
                let cell_type = get_attr(e, b"t").unwrap_or_default();
                let style_idx: usize = get_attr(e, b"s").and_then(|s| s.parse().ok()).unwrap_or(0);
                let value = parse_cell(reader, &cell_type, shared_strings, style_idx, styles);
                cells.push((col_idx, value));
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"c" => {
                // Self-closing <c/> — empty cell, skip
                let col_idx = get_attr(e, b"r")
                    .as_deref()
                    .map_or(cells.len(), col_ref_to_index);
                cells.push((col_idx, String::new()));
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"row" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    cells
}

/// Parse a single `<c>` cell element and return its text value.
///
/// For numeric cells (no `t` attribute or `t="n"`), checks the style
/// to see if the number format is a date — if so, converts the serial
/// number to an ISO date string.
fn parse_cell(
    reader: &mut Reader<&[u8]>,
    cell_type: &str,
    shared_strings: &[String],
    style_idx: usize,
    styles: &Styles,
) -> String {
    let mut value = String::new();
    let mut inline_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"v" => {
                        // Read the <v> text content
                        if let Ok(Event::Text(t)) = reader.read_event() {
                            if let Ok(s) = t.unescape() {
                                value = s.into_owned();
                            }
                        }
                    }
                    b"is" => {
                        // Inline string: collect all <t> text within <is>
                        inline_text = parse_inline_string(reader);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    match cell_type {
        "s" => {
            // Shared string reference
            value
                .parse::<usize>()
                .ok()
                .and_then(|idx| shared_strings.get(idx).cloned())
                .unwrap_or_default()
        }
        "inlineStr" => inline_text,
        // Numeric or untyped cells: check for date format
        "" | "n" => maybe_convert_date(&value, style_idx, styles),
        _ => value, // booleans ("b"), errors ("e"), formula strings ("str")
    }
}

/// If the cell's style is a date format and the value parses as a number,
/// convert it to an ISO date string. Otherwise return the raw value.
fn maybe_convert_date(value: &str, style_idx: usize, styles: &Styles) -> String {
    if styles.is_date_style(style_idx) {
        if let Ok(serial) = value.parse::<f64>() {
            return dateconv::serial_to_iso(serial);
        }
    }
    value.to_string()
}

/// Parse an `<is>` inline string element, collecting all `<t>` text.
fn parse_inline_string(reader: &mut Reader<&[u8]>) -> String {
    let mut text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Text(ref t)) => {
                if let Ok(s) = t.unescape() {
                    text.push_str(&s);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"is" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    text
}

/// Convert a cell reference like "B3" or "AA1" to a 0-based column index.
///
/// Extracts the letter prefix and converts it: A=0, B=1, ..., Z=25, AA=26, etc.
fn col_ref_to_index(cell_ref: &str) -> usize {
    let mut col = 0usize;
    for ch in cell_ref.bytes() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + usize::from(ch.to_ascii_uppercase() - b'A') + 1;
        } else {
            break;
        }
    }
    col.saturating_sub(1) // convert from 1-based to 0-based
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── col_ref_to_index ─────────────────────────────────────────

    #[test]
    fn col_ref_a1() {
        assert_eq!(col_ref_to_index("A1"), 0);
    }

    #[test]
    fn col_ref_b5() {
        assert_eq!(col_ref_to_index("B5"), 1);
    }

    #[test]
    fn col_ref_z1() {
        assert_eq!(col_ref_to_index("Z1"), 25);
    }

    #[test]
    fn col_ref_aa1() {
        assert_eq!(col_ref_to_index("AA1"), 26);
    }

    #[test]
    fn col_ref_az1() {
        assert_eq!(col_ref_to_index("AZ1"), 51);
    }

    #[test]
    fn col_ref_ba1() {
        assert_eq!(col_ref_to_index("BA1"), 52);
    }

    #[test]
    fn col_ref_lowercase() {
        assert_eq!(col_ref_to_index("c3"), 2);
    }

    // ── parse_shared_strings ─────────────────────────────────────

    #[test]
    fn shared_strings_simple() {
        let xml = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>Hello</t></si>
            <si><t>World</t></si>
        </sst>"#;
        assert_eq!(parse_shared_strings_xml(xml), vec!["Hello", "World"]);
    }

    #[test]
    fn shared_strings_rich_text() {
        // Rich text: <si><r><t>Part1</t></r><r><t>Part2</t></r></si>
        let xml = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><r><rPr><b/></rPr><t>Bold</t></r><r><t> Normal</t></r></si>
        </sst>"#;
        assert_eq!(parse_shared_strings_xml(xml), vec!["Bold Normal"]);
    }

    // ── parse_sheet_xml ──────────────────────────────────────────

    #[test]
    fn parse_sheet_shared_strings() {
        let shared = vec!["Name".to_string(), "Age".to_string(), "Alice".to_string()];
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1" t="s"><v>1</v></c>
                </row>
                <row r="2">
                    <c r="A2" t="s"><v>2</v></c>
                    <c r="B2"><v>30</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let no_styles = Styles::default();
        let rows = parse_sheet_xml(xml, &shared, &no_styles);
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0], vec!["Name", "Age"]);
        assert_eq!(rows[1], vec!["Alice", "30"]);
    }

    #[test]
    fn parse_sheet_inline_strings() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="inlineStr"><is><t>Status</t></is></c>
                    <c r="B1" t="inlineStr"><is><t>Task</t></is></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let no_styles = Styles::default();
        let rows = parse_sheet_xml(xml, &[], &no_styles);
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0], vec!["Status", "Task"]);
    }

    #[test]
    fn parse_sheet_sparse_columns() {
        // Row has A1 and C1 but no B1 — should produce 3 columns with gap
        let shared = vec!["First".to_string(), "Third".to_string()];
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="C1" t="s"><v>1</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let no_styles = Styles::default();
        let rows = parse_sheet_xml(xml, &shared, &no_styles);
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].len(), 3);
        assert_eq!(rows[0][0], "First");
        assert_eq!(rows[0][1], ""); // gap
        assert_eq!(rows[0][2], "Third");
    }

    #[test]
    fn parse_sheet_empty() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
        </worksheet>"#;

        let no_styles = Styles::default();
        let rows = parse_sheet_xml(xml, &[], &no_styles);
        assert!(rows.is_empty());
    }

    // ── styles / date detection ───────────────────────────────────

    #[test]
    fn parse_styles_builtin_date() {
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cellXfs count="2">
                <xf numFmtId="0"/>
                <xf numFmtId="14"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_styles_xml(xml);
        assert!(!styles.is_date_style(0));
        assert!(styles.is_date_style(1));
    }

    #[test]
    fn parse_styles_custom_date() {
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="1">
                <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
            </numFmts>
            <cellXfs count="2">
                <xf numFmtId="0"/>
                <xf numFmtId="164"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_styles_xml(xml);
        assert!(!styles.is_date_style(0));
        assert!(styles.is_date_style(1));
    }

    #[test]
    fn parse_styles_custom_number_not_date() {
        let xml = r##"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="1">
                <numFmt numFmtId="164" formatCode="#,##0.00"/>
            </numFmts>
            <cellXfs count="1">
                <xf numFmtId="164"/>
            </cellXfs>
        </styleSheet>"##;
        let styles = parse_styles_xml(xml);
        assert!(!styles.is_date_style(0));
    }

    #[test]
    fn parse_sheet_date_cell_converted() {
        // Style index 1 maps to numFmtId 14 (builtin date)
        let styles = Styles {
            is_date: vec![false, true],
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" s="0"><v>42</v></c>
                    <c r="B1" s="1"><v>45292</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let rows = parse_sheet_xml(xml, &[], &styles);
        assert_eq!(rows[0][0], "42");
        assert_eq!(rows[0][1], "2024-01-01");
    }

    // ── hyperlink resolution ───────────────────────────────────────

    #[test]
    fn cell_ref_to_row_basic() {
        assert_eq!(cell_ref_to_row("A1"), 0);
        assert_eq!(cell_ref_to_row("B3"), 2);
        assert_eq!(cell_ref_to_row("AA100"), 99);
    }

    #[test]
    fn apply_hyperlinks_basic() {
        let rels: Rels = [("rId1".into(), "https://example.com".into())].into();
        let sheet_xml = r#"<worksheet>
            <sheetData>
                <row r="1"><c r="A1" t="s"><v>0</v></c></row>
            </sheetData>
            <hyperlinks>
                <hyperlink ref="A1" r:id="rId1"/>
            </hyperlinks>
        </worksheet>"#;
        let mut rows = vec![vec!["Click here".to_string()]];
        apply_hyperlinks(sheet_xml, &rels, &mut rows);
        assert_eq!(rows[0][0], "[Click here](https://example.com)");
    }

    #[test]
    fn apply_hyperlinks_empty_rels_noop() {
        let rels = Rels::new();
        let mut rows = vec![vec!["Hello".to_string()]];
        apply_hyperlinks("<worksheet><hyperlinks/></worksheet>", &rels, &mut rows);
        assert_eq!(rows[0][0], "Hello");
    }
}
