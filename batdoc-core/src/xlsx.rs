//! OOXML `.xlsx` (Excel) spreadsheet parser.
//!
//! Unzips the `.xlsx` archive, parses the shared string table and each
//! worksheet's XML, then renders every sheet as either tab-separated
//! plain text or a markdown table with a heading per sheet. Hyperlinks
//! are resolved from sheet relationship files and rendered as markdown links.

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::HashMap;
use std::io::{BufRead, Cursor, Read};
use zip::ZipArchive;

use crate::arena::StringArena;
use crate::dateconv;
use crate::error::BatdocError;
use crate::sheet::{
    write_markdown_data_row, write_markdown_header, write_markdown_separator, write_plain_row,
    TableShape, MAX_COLS,
};
use crate::xml_util::{self, get_attr, Rels};
use crate::ExtractSink;

/// Extract plain text (TSV) from an .xlsx file.
pub(crate) fn extract_plain(data: &[u8]) -> crate::error::Result<String> {
    let mut out = String::new();
    extract_plain_to(data, &mut out)?;
    Ok(out)
}

/// Stream plain text (TSV) from an .xlsx file into `sink`.
pub(crate) fn extract_plain_to(
    data: &[u8],
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    let shared_strings = parse_shared_strings_arena(&mut archive);
    let styles = parse_styles(&mut archive);
    let sheet_info = discover_sheets(&mut archive)?;

    let mut emitted_any = false;
    for (name, path) in &sheet_info {
        let sheet_rels_path = xml_util::rels_path(path);
        let rels = xml_util::load_rels(&mut archive, &sheet_rels_path);
        let hyperlinks = if rels.is_empty() {
            HashMap::new()
        } else {
            load_sheet_hyperlinks(&mut archive, path, &rels)
        };
        emit_sheet_plain(
            &mut archive,
            path,
            name,
            &shared_strings,
            &styles,
            &hyperlinks,
            &mut emitted_any,
            sink,
        )?;
    }

    Ok(())
}

/// Extract markdown-formatted text from an .xlsx file.
///
/// When `images` is true, embedded images from drawings are extracted
/// and appended as reference-style base64 images with definitions at the end.
pub(crate) fn extract_markdown(data: &[u8], images: bool) -> crate::error::Result<String> {
    let mut out = String::new();
    extract_markdown_to(data, images, &mut out)?;
    Ok(out)
}

/// Stream markdown from an .xlsx file into `sink`.
pub(crate) fn extract_markdown_to(
    data: &[u8],
    images: bool,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    let shared_strings = parse_shared_strings_arena(&mut archive);
    let styles = parse_styles(&mut archive);
    let sheet_info = discover_sheets(&mut archive)?;

    let mut shapes = Vec::with_capacity(sheet_info.len());
    for (_, path) in &sheet_info {
        shapes.push(scan_sheet_shape(
            &mut archive,
            path,
            &shared_strings,
            &styles,
        )?);
    }
    let multiple = shapes.iter().filter(|s| s.is_some()).count() > 1;

    for ((name, path), shape) in sheet_info.iter().zip(shapes.iter()) {
        let Some(shape) = shape else {
            continue;
        };
        let sheet_rels_path = xml_util::rels_path(path);
        let rels = xml_util::load_rels(&mut archive, &sheet_rels_path);
        let hyperlinks = if rels.is_empty() {
            HashMap::new()
        } else {
            load_sheet_hyperlinks(&mut archive, path, &rels)
        };
        if multiple {
            sink.write_str("## ")?;
            sink.write_str(name)?;
            sink.write_str("\n\n")?;
        }
        emit_sheet_markdown(
            &mut archive,
            path,
            shape,
            &shared_strings,
            &styles,
            &hyperlinks,
            sink,
        )?;
    }

    if images {
        let mut extra = String::new();
        append_sheet_images(&mut extra, &sheet_info, &mut archive);
        if !extra.is_empty() {
            sink.write_str(&extra)?;
        }
    }

    Ok(())
}

// ── Parsing ────────────────────────────────────────────────────────

fn parse_shared_strings_arena(archive: &mut ZipArchive<Cursor<&[u8]>>) -> StringArena {
    let mut arena = StringArena::new();
    let Ok(mut reader) = xml_util::open_xml(archive, "xl/sharedStrings.xml") else {
        return arena;
    };
    let mut buf = Vec::new();
    fill_shared_strings(&mut reader, &mut buf, &mut arena);
    arena
}

fn fill_shared_strings<R: BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    arena: &mut StringArena,
) {
    let mut in_si = false;
    let mut current = String::new();

    loop {
        match reader.read_event_into(buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => {
                in_si = true;
                current.clear();
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"si" => {
                arena.push(&current);
                in_si = false;
            }
            Ok(Event::Text(ref t)) if in_si => {
                if let Ok(s) = t.unescape() {
                    current.push_str(&s);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn load_sheet_hyperlinks(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
    rels: &Rels,
) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let Ok(mut reader) = xml_util::open_xml(archive, path) else {
        return map;
    };
    let mut buf = Vec::new();
    let mut in_hyperlinks = false;

    loop {
        match reader.read_event_into(&mut buf) {
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
                    map.insert(cell_ref, url.clone());
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    map
}

#[allow(clippy::too_many_arguments)]
fn emit_sheet_plain(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
    name: &str,
    shared_strings: &StringArena,
    styles: &Styles,
    hyperlinks: &HashMap<String, String>,
    emitted_any: &mut bool,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let Ok(mut reader) = xml_util::open_xml(archive, path) else {
        return Ok(());
    };
    let mut buf = Vec::new();
    let mut wrote_header = false;

    loop {
        let kind = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => 1u8,
            Ok(Event::Eof) | Err(_) => 2,
            _ => 0,
        };
        buf.clear();
        match kind {
            1 => {
                let cells = parse_row_arena(&mut reader, &mut buf, shared_strings, styles)?;
                let dense = densify_plain_row(&cells, hyperlinks);
                let mut line = String::new();
                write_plain_row(&mut line, dense.iter().map(String::as_str))?;
                if !line.is_empty() {
                    if !wrote_header {
                        if *emitted_any {
                            sink.write_str("\n--- ")?;
                            sink.write_str(name)?;
                            sink.write_str(" ---\n")?;
                        }
                        wrote_header = true;
                        *emitted_any = true;
                    }
                    sink.write_str(&line)?;
                }
            }
            2 => break,
            _ => {}
        }
    }

    Ok(())
}

fn densify_plain_row(
    cells: &[(usize, String, String)],
    hyperlinks: &HashMap<String, String>,
) -> Vec<String> {
    let Some(max_col) = cells.iter().map(|(col, _, _)| *col).max() else {
        return Vec::new();
    };
    let mut dense = vec![String::new(); max_col + 1];
    for (col, value, cell_ref) in cells {
        if *col >= dense.len() {
            continue;
        }
        dense[*col] = apply_cell_hyperlink(value, cell_ref, hyperlinks);
    }
    dense
}

fn apply_cell_hyperlink(
    value: &str,
    cell_ref: &str,
    hyperlinks: &HashMap<String, String>,
) -> String {
    if value.is_empty() {
        return String::new();
    }
    match hyperlinks.get(cell_ref) {
        Some(url) => format!("[{value}]({url})"),
        None => value.to_string(),
    }
}

fn check_col(col: usize) -> crate::error::Result<()> {
    if col >= MAX_COLS {
        Err(BatdocError::Document("sheet exceeds 512 columns".into()))
    } else {
        Ok(())
    }
}

fn scan_sheet_shape(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
    shared_strings: &StringArena,
    styles: &Styles,
) -> crate::error::Result<Option<TableShape>> {
    let Ok(mut reader) = xml_util::open_xml(archive, path) else {
        return Ok(None);
    };
    let mut buf = Vec::new();
    let mut used = vec![false; MAX_COLS];
    let mut last_nonempty_row = None;
    let mut row_idx = 0usize;

    loop {
        let kind = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => 1u8,
            Ok(Event::Eof) | Err(_) => 2,
            _ => 0,
        };
        buf.clear();
        match kind {
            1 => {
                let cells = parse_row_arena(&mut reader, &mut buf, shared_strings, styles)?;
                let mut nonempty = false;
                for (col, value, _) in &cells {
                    if !value.trim().is_empty() {
                        used[*col] = true;
                        nonempty = true;
                    }
                }
                if nonempty {
                    last_nonempty_row = Some(row_idx);
                }
                row_idx += 1;
            }
            2 => break,
            _ => {}
        }
    }

    let Some(last_row) = last_nonempty_row else {
        return Ok(None);
    };
    let Some(first_col) = used.iter().position(|&u| u) else {
        return Ok(None);
    };
    let last_col = used.iter().rposition(|&u| u).unwrap_or(first_col);
    Ok(Some(TableShape {
        first_col,
        last_col,
        last_row,
    }))
}

fn emit_sheet_markdown(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
    shape: &TableShape,
    shared_strings: &StringArena,
    styles: &Styles,
    hyperlinks: &HashMap<String, String>,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let Ok(mut reader) = xml_util::open_xml(archive, path) else {
        return Ok(());
    };
    let mut buf = Vec::new();
    let mut row_idx = 0usize;
    let mut is_header = true;
    let ncols = shape.last_col - shape.first_col + 1;

    loop {
        let kind = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => 1u8,
            Ok(Event::Eof) | Err(_) => 2,
            _ => 0,
        };
        buf.clear();
        match kind {
            1 => {
                if row_idx > shape.last_row {
                    break;
                }
                let cells = parse_row_arena(&mut reader, &mut buf, shared_strings, styles)?;
                let slice = densify_markdown_row(&cells, shape, hyperlinks);
                if is_header {
                    write_markdown_header(sink, &slice)?;
                    write_markdown_separator(sink, ncols)?;
                    is_header = false;
                } else {
                    write_markdown_data_row(sink, &slice)?;
                }
                row_idx += 1;
            }
            2 => break,
            _ => {}
        }
    }

    sink.write_str("\n")
}

fn densify_markdown_row(
    cells: &[(usize, String, String)],
    shape: &TableShape,
    hyperlinks: &HashMap<String, String>,
) -> Vec<String> {
    let ncols = shape.last_col - shape.first_col + 1;
    let mut dense = vec![String::new(); ncols];
    for (col, value, cell_ref) in cells {
        if *col < shape.first_col || *col > shape.last_col {
            continue;
        }
        dense[*col - shape.first_col] = apply_cell_hyperlink(value, cell_ref, hyperlinks);
    }
    dense
}

fn parse_row_arena<R: BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    shared_strings: &StringArena,
    styles: &Styles,
) -> crate::error::Result<Vec<(usize, String, String)>> {
    let mut cells: Vec<(usize, String, String)> = Vec::new();
    buf.clear();

    loop {
        let mut start_cell = None;
        let mut empty_cell = None;
        match reader.read_event_into(buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c" => {
                let cell_ref = get_attr(e, b"r").unwrap_or_default();
                let col_idx = if cell_ref.is_empty() {
                    cells.len()
                } else {
                    col_ref_to_index(&cell_ref)
                };
                check_col(col_idx)?;
                let cell_type = get_attr(e, b"t").unwrap_or_default();
                let style_idx: usize = get_attr(e, b"s").and_then(|s| s.parse().ok()).unwrap_or(0);
                start_cell = Some((col_idx, cell_ref, cell_type, style_idx));
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"c" => {
                let cell_ref = get_attr(e, b"r").unwrap_or_default();
                let col_idx = if cell_ref.is_empty() {
                    cells.len()
                } else {
                    col_ref_to_index(&cell_ref)
                };
                check_col(col_idx)?;
                empty_cell = Some((col_idx, cell_ref));
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"row" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
        if let Some((col_idx, cell_ref, cell_type, style_idx)) = start_cell {
            let value =
                parse_cell_arena(reader, buf, &cell_type, shared_strings, style_idx, styles);
            cells.push((col_idx, value, cell_ref));
        } else if let Some((col_idx, cell_ref)) = empty_cell {
            cells.push((col_idx, String::new(), cell_ref));
        }
    }

    Ok(cells)
}

fn parse_cell_arena<R: BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
    cell_type: &str,
    shared_strings: &StringArena,
    style_idx: usize,
    styles: &Styles,
) -> String {
    let mut value = String::new();
    let mut inline_text = String::new();
    buf.clear();

    loop {
        let mut read_v = false;
        let mut read_is = false;
        match reader.read_event_into(buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"v" => read_v = true,
                b"is" => read_is = true,
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
        if read_v {
            if let Ok(Event::Text(t)) = reader.read_event_into(buf) {
                if let Ok(s) = t.unescape() {
                    value = s.into_owned();
                }
            }
            buf.clear();
        } else if read_is {
            inline_text = parse_inline_string_into(reader, buf);
        }
    }

    match cell_type {
        "s" => value
            .parse::<usize>()
            .ok()
            .and_then(|idx| shared_strings.get(idx).map(str::to_string))
            .unwrap_or_default(),
        "inlineStr" => inline_text,
        "" | "n" => maybe_convert_date(&value, style_idx, styles),
        _ => value,
    }
}

fn parse_inline_string_into<R: BufRead>(reader: &mut Reader<R>, buf: &mut Vec<u8>) -> String {
    let mut text = String::new();
    buf.clear();

    loop {
        match reader.read_event_into(buf) {
            Ok(Event::Text(ref t)) => {
                if let Ok(s) = t.unescape() {
                    text.push_str(&s);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"is" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    text
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
    let Ok(mut reader) = xml_util::open_xml(archive, "xl/styles.xml") else {
        return Styles::default();
    };
    let mut buf = Vec::new();
    fill_styles(&mut reader, &mut buf)
}

/// Parse styles XML into resolved style info (separated for testability).
#[cfg(test)]
fn parse_styles_xml(xml: &str) -> Styles {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    fill_styles(&mut reader, &mut buf)
}

fn fill_styles<R: BufRead>(reader: &mut Reader<R>, buf: &mut Vec<u8>) -> Styles {
    let mut custom_formats: Vec<(u16, String)> = Vec::new();
    let mut cell_xf_fmt_ids: Vec<u16> = Vec::new();
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event_into(buf) {
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
        buf.clear();
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
#[cfg(test)]
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
            Ok(Event::Start(ref e) | Event::Empty(ref e)) if e.local_name().as_ref() == b"blip" => {
                if let Some(rid) = get_attr(e, b"r:embed") {
                    rids.push(rid);
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
#[cfg(test)]
fn cell_ref_to_row(cell_ref: &str) -> usize {
    let digits: String = cell_ref
        .chars()
        .skip_while(char::is_ascii_alphabetic)
        .collect();
    digits.parse::<usize>().unwrap_or(1).saturating_sub(1)
}

/// Parse shared string table XML into a list of strings.
///
/// Separated from `parse_shared_strings` for testability (avoids needing
/// a ZIP archive in tests).
#[cfg(test)]
fn parse_shared_strings_xml(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut arena = StringArena::new();
    fill_shared_strings(&mut reader, &mut buf, &mut arena);
    (0..arena.len())
        .map(|i| arena.get(i).unwrap_or("").to_string())
        .collect()
}

/// Discover sheet names and their file paths from workbook.xml and relationships.
///
/// Returns `(sheet_name, zip_path)` pairs in workbook order.
fn discover_sheets(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> crate::error::Result<Vec<(String, String)>> {
    let sheet_entries = {
        let mut reader = xml_util::open_xml(archive, "xl/workbook.xml")?;
        let mut buf = Vec::new();
        let mut sheet_entries: Vec<(String, String)> = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e) | Event::Start(ref e))
                    if e.local_name().as_ref() == b"sheet" =>
                {
                    let name = get_attr(e, b"name").unwrap_or_default();
                    let rid = get_attr(e, b"r:id").unwrap_or_default();
                    let state = get_attr(e, b"state").unwrap_or_default();
                    if state != "hidden" && !name.is_empty() && !rid.is_empty() {
                        sheet_entries.push((name, rid));
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        sheet_entries
    };

    let rid_to_target = {
        let mut reader = xml_util::open_xml(archive, "xl/_rels/workbook.xml.rels")?;
        let mut buf = Vec::new();
        let mut rid_to_target: Vec<(String, String)> = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
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
            buf.clear();
        }
        rid_to_target
    };

    let mut result = Vec::new();
    for (name, rid) in &sheet_entries {
        if let Some((_, target)) = rid_to_target.iter().find(|(id, _)| id == rid) {
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
#[cfg(test)]
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
#[cfg(test)]
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
#[cfg(test)]
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
#[cfg(test)]
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

    // ── streaming plain (zip fixtures) ───────────────────────────

    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn zip_xlsx(parts: &[(&str, &str)]) -> Vec<u8> {
        let buf = Cursor::new(Vec::new());
        let mut z = ZipWriter::new(buf);
        for (name, body) in parts {
            z.start_file(*name, SimpleFileOptions::default()).unwrap();
            z.write_all(body.as_bytes()).unwrap();
        }
        z.finish().unwrap().into_inner()
    }

    fn minimal_xlsx_two_rows() -> Vec<u8> {
        zip_xlsx(&[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Name</t></si>
  <si><t>Age</t></si>
  <si><t>Alice</t></si>
  <si><t>30</t></si>
</sst>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2" t="s"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>"#,
            ),
        ])
    }

    #[test]
    fn xlsx_plain_to_matches_buffered_and_has_no_trailing_tab_row() {
        let data = minimal_xlsx_two_rows();
        let buffered = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        let mut streamed = String::new();
        crate::extract_plain_to(
            &data,
            crate::Format::Xlsx,
            crate::ExtractOptions::default(),
            &mut streamed,
        )
        .unwrap();
        assert_eq!(buffered, streamed);
        assert_eq!(buffered, "Name\tAge\nAlice\t30\n");
    }

    #[test]
    fn xlsx_plain_to_respects_max_output_bytes() {
        let data = minimal_xlsx_two_rows();
        let mut out = String::new();
        let err = crate::extract_plain_to(
            &data,
            crate::Format::Xlsx,
            crate::ExtractOptions {
                max_output_bytes: Some(16),
                ..Default::default()
            },
            &mut out,
        )
        .unwrap_err()
        .to_string();
        assert_eq!(err, "output exceeded 16 bytes");
        assert!(out.len() <= 16);
    }

    fn workbook_two_sheets() -> &'static str {
        r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="People" sheetId="1" r:id="rId1"/>
    <sheet name="Places" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#
    }

    fn workbook_two_sheet_rels() -> &'static str {
        r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#
    }

    fn sheet_xml(body: &str) -> String {
        format!(
            r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>{body}</sheetData>
</worksheet>"#
        )
    }

    #[test]
    fn xlsx_plain_skips_empty_sheet_and_omits_single_sheet_header() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            ("xl/workbook.xml", workbook_two_sheets()),
            ("xl/_rels/workbook.xml.rels", workbook_two_sheet_rels()),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(r#"<row r="1"><c r="A1"><v></v></c></row>"#),
            ),
            (
                "xl/worksheets/sheet2.xml",
                &sheet_xml(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>NYC</t></is></c></row>"#),
            ),
        ]);
        let text = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        assert_eq!(text, "NYC\n");
    }

    #[test]
    fn xlsx_plain_multi_sheet_header_only_before_non_first_emitted() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            ("xl/workbook.xml", workbook_two_sheets()),
            ("xl/_rels/workbook.xml.rels", workbook_two_sheet_rels()),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Alice</t></is></c></row>"#),
            ),
            (
                "xl/worksheets/sheet2.xml",
                &sheet_xml(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>NYC</t></is></c></row>"#),
            ),
        ]);
        let text = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        assert_eq!(text, "Alice\n\n--- Places ---\nNYC\n");
    }

    #[test]
    fn xlsx_plain_sparse_row_keeps_middle_gap() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<sst><si><t>First</t></si><si><t>Third</t></si></sst>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(
                    r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>"#,
                ),
            ),
        ]);
        let text = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        assert_eq!(text, "First\t\tThird\n");
    }

    #[test]
    fn xlsx_plain_missing_sst_index_is_empty_cell() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            ("xl/sharedStrings.xml", r#"<sst><si><t>Only</t></si></sst>"#),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(
                    r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>99</v></c></row>"#,
                ),
            ),
        ]);
        let text = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        assert_eq!(text, "Only\n");
    }

    #[test]
    fn xlsx_plain_date_styled_numeric_is_iso() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0"/>
    <xf numFmtId="14"/>
  </cellXfs>
</styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(
                    r#"<row r="1"><c r="A1" s="0"><v>42</v></c><c r="B1" s="1"><v>45292</v></c></row>"#,
                ),
            ),
        ]);
        let text = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        assert_eq!(text, "42\t2024-01-01\n");
    }

    #[test]
    fn xlsx_plain_hyperlink_wraps_cell() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Click here</t></is></c></row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1"/>
  </hyperlinks>
</worksheet>"#,
            ),
        ]);
        let text = crate::extract_plain(&data, crate::Format::Xlsx).unwrap();
        assert_eq!(text, "[Click here](https://example.com)\n");
    }

    fn xlsx_with_cell_at(cell_ref: &str) -> Vec<u8> {
        let body =
            format!(r#"<row r="1"><c r="{cell_ref}" t="inlineStr"><is><t>X</t></is></c></row>"#);
        zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", &sheet_xml(&body)),
        ])
    }

    #[test]
    fn xlsx_markdown_rejects_col_513() {
        let data = xlsx_with_cell_at("SS1");
        let err = crate::extract_markdown(&data, crate::Format::Xlsx, false)
            .unwrap_err()
            .to_string();
        assert_eq!(err, "sheet exceeds 512 columns");
    }

    #[test]
    fn xlsx_markdown_accepts_col_ja() {
        let data = xlsx_with_cell_at("JA1");
        let md = crate::extract_markdown(&data, crate::Format::Xlsx, false).unwrap();
        assert_eq!(md, "| X |\n| --- |\n\n");
    }

    #[test]
    fn xlsx_markdown_two_rows_locked() {
        let data = minimal_xlsx_two_rows();
        let md = crate::extract_markdown(&data, crate::Format::Xlsx, false).unwrap();
        assert_eq!(md, "| Name | Age |\n| --- | --- |\n| Alice | 30 |\n\n");
    }

    #[test]
    fn xlsx_markdown_to_respects_max_output_bytes() {
        let data = minimal_xlsx_two_rows();
        let mut out = String::new();
        let err = crate::extract_markdown_to(
            &data,
            crate::Format::Xlsx,
            crate::ExtractOptions {
                max_output_bytes: Some(16),
                ..Default::default()
            },
            &mut out,
        )
        .unwrap_err()
        .to_string();
        assert_eq!(err, "output exceeded 16 bytes");
        assert!(out.len() <= 16);
    }

    #[test]
    fn xlsx_markdown_skips_empty_sheet_and_omits_single_sheet_heading() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            ("xl/workbook.xml", workbook_two_sheets()),
            ("xl/_rels/workbook.xml.rels", workbook_two_sheet_rels()),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(r#"<row r="1"><c r="A1"><v></v></c></row>"#),
            ),
            (
                "xl/worksheets/sheet2.xml",
                &sheet_xml(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>NYC</t></is></c></row>"#),
            ),
        ]);
        let md = crate::extract_markdown(&data, crate::Format::Xlsx, false).unwrap();
        assert_eq!(md, "| NYC |\n| --- |\n\n");
    }

    #[test]
    fn xlsx_markdown_multi_sheet_headings() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            ("xl/workbook.xml", workbook_two_sheets()),
            ("xl/_rels/workbook.xml.rels", workbook_two_sheet_rels()),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Alice</t></is></c></row>"#),
            ),
            (
                "xl/worksheets/sheet2.xml",
                &sheet_xml(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>NYC</t></is></c></row>"#),
            ),
        ]);
        let md = crate::extract_markdown(&data, crate::Format::Xlsx, false).unwrap();
        assert_eq!(
            md,
            "## People\n\n| Alice |\n| --- |\n\n## Places\n\n| NYC |\n| --- |\n\n"
        );
    }

    #[test]
    fn xlsx_markdown_strips_leading_empty_col_and_trailing_empty_row() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                &sheet_xml(
                    r#"<row r="1"><c r="B1" t="inlineStr"><is><t>Name</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>Alice</t></is></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>  </t></is></c></row>"#,
                ),
            ),
        ]);
        let md = crate::extract_markdown(&data, crate::Format::Xlsx, false).unwrap();
        assert_eq!(md, "| Name |\n| --- |\n| Alice |\n\n");
    }

    #[test]
    fn xlsx_markdown_hyperlink_wraps_cell() {
        let data = zip_xlsx(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types/>"#),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Click here</t></is></c></row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1"/>
  </hyperlinks>
</worksheet>"#,
            ),
        ]);
        let md = crate::extract_markdown(&data, crate::Format::Xlsx, false).unwrap();
        assert_eq!(md, "| [Click here](https://example.com) |\n| --- |\n\n");
    }
}
