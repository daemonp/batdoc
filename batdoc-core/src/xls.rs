//! Legacy BIFF8 `.xls` (Excel 97+) binary format parser.
//!
//! Reads the `Workbook` (or `Book`) stream from the OLE2 compound file,
//! parses the BIFF8 record stream to extract the Shared String Table (SST),
//! sheet metadata (`BoundSheet8`), and cell records (LABELSST, NUMBER, RK,
//! MULRK, FORMULA, LABEL, BOOLERR). Emits rows through the shared sheet writers.

use cfb::CompoundFile;
use std::io::{Cursor, Read};

use crate::arena::StringArena;
use crate::codepage;
use crate::dateconv;
use crate::error::BatdocError;
use crate::sheet::{
    write_markdown_data_row, write_markdown_header, write_markdown_separator, write_plain_row,
    TableShape, MAX_COLS,
};
use crate::sheets::{finalize_sheet_row, SheetSink};
use crate::ExtractSink;

// ── BIFF8 record types ────────────────────────────────────────────

const REC_BOF: u16 = 0x0809;
const REC_EOF: u16 = 0x000A;
const REC_BOUNDSHEET: u16 = 0x0085;
const REC_SST: u16 = 0x00FC;
const REC_CONTINUE: u16 = 0x003C;
const REC_LABELSST: u16 = 0x00FD;
const REC_LABEL: u16 = 0x0204;
const REC_RSTRING: u16 = 0x00D6;
const REC_NUMBER: u16 = 0x0203;
const REC_RK: u16 = 0x027E;
const REC_MULRK: u16 = 0x00BD;
const REC_FORMULA: u16 = 0x0006;
const REC_STRING: u16 = 0x0207;
const REC_BOOLERR: u16 = 0x0205;
const REC_FILEPASS: u16 = 0x002F;
const REC_FORMAT: u16 = 0x041E;
const REC_XF: u16 = 0x00E0;
const REC_CODEPAGE: u16 = 0x0042;

/// Extract plain text (TSV) from a BIFF8 .xls file.
pub(crate) fn extract_plain(data: &[u8]) -> crate::error::Result<String> {
    let mut out = String::new();
    extract_plain_to(data, &mut out)?;
    Ok(out)
}

/// Stream plain text (TSV) from a BIFF8 .xls file into `sink`.
pub(crate) fn extract_plain_to(
    data: &[u8],
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let wb = open_workbook(data)?;
    let visible: Vec<&SheetEntry> = wb
        .sheets
        .iter()
        .filter(|e| e.sheet_type == 0 && e.visibility == 0)
        .collect();
    let mut emitted_any = false;
    for entry in visible {
        emit_sheet_plain(
            &wb.buf,
            entry,
            &wb.sst,
            &wb.xf_styles,
            wb.cp,
            &mut emitted_any,
            sink,
        )?;
    }
    Ok(())
}

/// Extract markdown-formatted text from a BIFF8 .xls file.
pub(crate) fn extract_markdown(data: &[u8]) -> crate::error::Result<String> {
    let mut out = String::new();
    extract_markdown_to(data, &mut out)?;
    Ok(out)
}

/// Stream structured sheets into `sink` (visible worksheets only).
#[allow(dead_code)] // consumed by write_sheets entry in task 4
pub(crate) fn extract_sheets_to(
    data: &[u8],
    sink: &mut impl SheetSink,
) -> crate::error::Result<()> {
    let wb = open_workbook(data)?;
    let visible: Vec<&SheetEntry> = wb
        .sheets
        .iter()
        .filter(|e| e.sheet_type == 0 && e.visibility == 0)
        .collect();
    for entry in visible {
        sink.begin_sheet(&entry.name)?;
        if let Some(mut walk) = SheetWalk::start(&wb.buf, entry.bof_offset) {
            let mut current = CurrentRow::new();
            while let Some((row, col, value)) = walk.next_cell(&wb.sst, &wb.xf_styles, wb.cp)? {
                if let Some(prev) = current.push(row, col, value)? {
                    if let Some(dense) = finalize_sheet_row(densify_plain_row(&prev)) {
                        sink.row(dense)?;
                    }
                }
            }
            let last = current.take();
            if !last.is_empty() {
                if let Some(dense) = finalize_sheet_row(densify_plain_row(&last)) {
                    sink.row(dense)?;
                }
            }
        }
        sink.end_sheet()?;
    }
    Ok(())
}

/// Stream markdown from a BIFF8 .xls file into `sink`.
pub(crate) fn extract_markdown_to(
    data: &[u8],
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let wb = open_workbook(data)?;
    let visible: Vec<&SheetEntry> = wb
        .sheets
        .iter()
        .filter(|e| e.sheet_type == 0 && e.visibility == 0)
        .collect();
    let mut shapes = Vec::with_capacity(visible.len());
    for entry in &visible {
        shapes.push(scan_sheet_shape(
            &wb.buf,
            entry.bof_offset,
            &wb.sst,
            &wb.xf_styles,
            wb.cp,
        )?);
    }
    let multiple = shapes.iter().filter(|s| s.is_some()).count() > 1;
    for (entry, shape) in visible.iter().zip(shapes.iter()) {
        let Some(shape) = shape else {
            continue;
        };
        if multiple {
            sink.write_str("## ")?;
            sink.write_str(&entry.name)?;
            sink.write_str("\n\n")?;
        }
        emit_sheet_markdown(
            &wb.buf,
            entry.bof_offset,
            shape,
            &wb.sst,
            &wb.xf_styles,
            wb.cp,
            sink,
        )?;
    }
    Ok(())
}

// ── Record-level types ─────────────────────────────────────────────

/// A raw BIFF8 record: type, offset in stream, and data bytes.
#[derive(Debug)]
struct Record<'a> {
    rec_type: u16,
    data: &'a [u8],
}

/// Sheet metadata from `BoundSheet8` record.
#[derive(Debug)]
struct SheetEntry {
    name: String,
    bof_offset: u32,
    visibility: u8,
    sheet_type: u8,
}

// ── Main parser ────────────────────────────────────────────────────

struct Workbook {
    buf: Vec<u8>,
    sst: StringArena,
    sheets: Vec<SheetEntry>,
    xf_styles: XfStyles,
    cp: u16,
}

fn open_workbook(data: &[u8]) -> crate::error::Result<Workbook> {
    let cursor = Cursor::new(data);
    let mut cfb = CompoundFile::open(cursor)?;

    let stream_name = if cfb.exists("/Workbook") {
        "/Workbook"
    } else if cfb.exists("/Book") {
        "/Book"
    } else {
        return Err(BatdocError::Document(
            "not an Excel file (no Workbook or Book stream)".into(),
        ));
    };

    let mut stream = cfb.open_stream(stream_name)?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf)?;

    let records = parse_records(&buf);
    let (sst, sheets, xf_styles, cp) = parse_globals(&records)?;
    Ok(Workbook {
        buf,
        sst,
        sheets,
        xf_styles,
        cp,
    })
}

/// Maximum number of BIFF8 records to parse (defense-in-depth against
/// degenerate files with millions of tiny records).
const MAX_RECORDS: usize = 2_000_000;

/// Parse the raw byte stream into a flat list of BIFF8 records.
///
/// Records borrow their data directly from the input slice, avoiding
/// per-record allocations.
fn parse_records(data: &[u8]) -> Vec<Record<'_>> {
    let mut records = Vec::new();
    let mut offset = 0;

    while offset + 4 <= data.len() {
        let rec_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let rec_len = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));

        if rec_type == 0 && rec_len == 0 {
            break;
        }

        let end = (offset + 4 + rec_len).min(data.len());

        records.push(Record {
            rec_type,
            data: &data[offset + 4..end],
        });

        if records.len() >= MAX_RECORDS {
            break;
        }

        offset = end;
    }

    records
}

/// Resolved XF style information for date detection.
///
/// Maps each XF record index to whether its number format is a date format,
/// analogous to the `Styles` struct in the xlsx parser.
#[derive(Debug, Default)]
struct XfStyles {
    /// For each XF index, true if the numFmtId is a date format.
    is_date: Vec<bool>,
}

impl XfStyles {
    /// Check if an XF index corresponds to a date format.
    fn is_date_xf(&self, xf_idx: u16) -> bool {
        self.is_date
            .get(usize::from(xf_idx))
            .copied()
            .unwrap_or(false)
    }
}

/// Parse workbook globals: extract SST, `BoundSheet8` entries, XF styles,
/// and codepage.
///
/// Detects encrypted files early via the FILEPASS record, returning
/// an error before doing any further parsing.
fn parse_globals(
    records: &[Record<'_>],
) -> crate::error::Result<(StringArena, Vec<SheetEntry>, XfStyles, u16)> {
    let mut sst = StringArena::new();
    let mut sheet_entries = Vec::new();
    // Custom FORMAT records: numFmtId → format string
    let mut custom_formats: Vec<(u16, String)> = Vec::new();
    // XF records: each entry's numFmtId
    let mut xf_fmt_ids: Vec<u16> = Vec::new();
    // Codepage from CODEPAGE record (default: 1252 = Western European)
    let mut cp: u16 = 1252;

    let mut i = 0;
    while i < records.len() {
        let rec = &records[i];

        match rec.rec_type {
            REC_FILEPASS => {
                return Err(BatdocError::Document("document is encrypted".into()));
            }
            REC_CODEPAGE
                if rec.data.len() >= 2 => {
                    cp = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                }
            REC_FORMAT => {
                if let Some((id, code)) = parse_format_record(rec.data, cp) {
                    custom_formats.push((id, code));
                }
            }
            REC_XF
                // XF record: bytes 2-3 are numFmtId
                if rec.data.len() >= 4 => {
                    let fmt_id = u16::from_le_bytes([rec.data[2], rec.data[3]]);
                    xf_fmt_ids.push(fmt_id);
                }
            REC_SST => {
                // Collect SST + following CONTINUE records
                let mut combined = rec.data.to_vec();
                let mut continue_boundaries = vec![combined.len()];
                let mut j = i + 1;
                while j < records.len() && records[j].rec_type == REC_CONTINUE {
                    continue_boundaries.push(combined.len() + records[j].data.len());
                    combined.extend_from_slice(records[j].data);
                    j += 1;
                }
                sst = StringArena::new();
                parse_sst(&combined, &continue_boundaries, cp, &mut sst);
                i = j;
                continue;
            }
            REC_BOUNDSHEET => {
                if let Some(entry) = parse_boundsheet(rec.data, cp) {
                    sheet_entries.push(entry);
                }
            }
            REC_EOF => break, // End of workbook globals
            _ => {}
        }

        i += 1;
    }

    Ok((
        sst,
        sheet_entries,
        XfStyles {
            is_date: dateconv::resolve_date_styles(&xf_fmt_ids, &custom_formats),
        },
        cp,
    ))
}

/// Parse a FORMAT record (0x041E) into (`numFmtId`, `format_string`).
///
/// Record format: 2 bytes numFmtId + BIFF8 unicode string (the format code).
fn parse_format_record(data: &[u8], cp: u16) -> Option<(u16, String)> {
    if data.len() < 5 {
        return None;
    }
    let id = u16::from_le_bytes([data[0], data[1]]);
    let (s, _) = read_biff8_string(data, 2, &[], cp);
    Some((id, s))
}

// ── SST parsing ────────────────────────────────────────────────────

/// Parse the Shared String Table from combined SST + CONTINUE data.
///
/// The SST record format:
///   - 4 bytes: total string references in workbook
///   - 4 bytes: number of unique strings
///   - Variable: string data (may span CONTINUE record boundaries)
///
/// BIFF8 strings that span a CONTINUE boundary have a special encoding:
/// at the CONTINUE boundary, a new "grbit" byte indicates whether the
/// remaining characters are compressed (0) or uncompressed (1).
fn parse_sst(data: &[u8], continue_boundaries: &[usize], cp: u16, arena: &mut StringArena) {
    if data.len() < 8 {
        return;
    }

    let unique_count = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
    let mut pos = 8;

    for _ in 0..unique_count {
        if pos + 3 > data.len() {
            break;
        }

        let (s, new_pos) = read_biff8_string(data, pos, continue_boundaries, cp);
        arena.push(&s);
        pos = new_pos;
    }
}

/// Read a BIFF8 unicode string from a buffer, handling CONTINUE boundaries.
///
/// String format:
///   - 2 bytes: character count (not byte count)
///   - 1 byte: flags (bit 0 = unicode, bit 2 = extended, bit 3 = rich text)
///   - If rich: 2 bytes run count
///   - If extended: 4 bytes extension size
///   - Character data (either 1 byte/char compressed or 2 bytes/char UTF-16LE)
///   - Rich text runs (4 bytes each)
///   - Extended data
fn read_biff8_string(
    data: &[u8],
    start: usize,
    continue_boundaries: &[usize],
    cp: u16,
) -> (String, usize) {
    if start + 3 > data.len() {
        return (String::new(), data.len());
    }

    let char_count = usize::from(u16::from_le_bytes([data[start], data[start + 1]]));
    let flags = data[start + 2];
    let mut pos = start + 3;

    let is_unicode = flags & 0x01 != 0;
    let has_ext = flags & 0x04 != 0;
    let has_rich = flags & 0x08 != 0;

    let rich_runs = if has_rich {
        if pos + 2 > data.len() {
            return (String::new(), data.len());
        }
        let n = usize::from(u16::from_le_bytes([data[pos], data[pos + 1]]));
        pos += 2;
        n
    } else {
        0
    };

    let ext_len = if has_ext {
        if pos + 4 > data.len() {
            return (String::new(), data.len());
        }
        let n =
            u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]) as usize;
        pos += 4;
        n
    } else {
        0
    };

    // Read characters, handling CONTINUE boundaries.
    // We keep a cursor `bi` into the sorted `continue_boundaries` slice so that
    // boundary lookups are O(1) amortised instead of O(n) per character.
    let mut result = String::with_capacity(char_count);
    let mut chars_remaining = char_count;
    let mut current_unicode = is_unicode;
    let mut bi = 0; // boundary index cursor

    while chars_remaining > 0 && pos < data.len() {
        // Advance cursor past boundaries we've already passed
        while bi < continue_boundaries.len() && continue_boundaries[bi] < pos {
            bi += 1;
        }

        // Check if we're exactly at a CONTINUE boundary
        if bi < continue_boundaries.len() && continue_boundaries[bi] == pos {
            if pos >= data.len() {
                break;
            }
            current_unicode = data[pos] & 0x01 != 0;
            pos += 1;
            bi += 1;
        }

        // Next boundary (or end of data)
        let next_boundary = if bi < continue_boundaries.len() {
            continue_boundaries[bi]
        } else {
            data.len()
        };

        let bytes_available = next_boundary.saturating_sub(pos);

        if current_unicode {
            let chars_available = bytes_available / 2;
            let chars_to_read = chars_remaining.min(chars_available);

            for _ in 0..chars_to_read {
                if pos + 2 > data.len() {
                    break;
                }
                let code = u16::from_le_bytes([data[pos], data[pos + 1]]);
                if let Some(ch) = char::from_u32(u32::from(code)) {
                    result.push(ch);
                }
                pos += 2;
            }
            chars_remaining -= chars_to_read;
        } else {
            // Compressed: 1 byte per character using workbook codepage
            let chars_to_read = chars_remaining.min(bytes_available);

            for _ in 0..chars_to_read {
                if pos >= data.len() {
                    break;
                }
                result.push(codepage::decode_byte(data[pos], cp));
                pos += 1;
            }
            chars_remaining -= chars_to_read;
        }
    }

    // Skip rich text formatting runs
    pos += rich_runs * 4;
    // Skip extended data
    pos += ext_len;

    (result, pos)
}

// ── BoundSheet8 parsing ────────────────────────────────────────────

fn parse_boundsheet(data: &[u8], cp: u16) -> Option<SheetEntry> {
    if data.len() < 8 {
        return None;
    }

    let bof_offset = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let visibility = data[4];
    let sheet_type = data[5];
    let str_len = usize::from(data[6]);
    let options = data[7];

    let name = if options & 0x01 != 0 {
        // Unicode string
        let byte_len = str_len * 2;
        if data.len() < 8 + byte_len {
            return None;
        }
        decode_utf16le(&data[8..8 + byte_len])
    } else {
        // Compressed: use workbook codepage
        if data.len() < 8 + str_len {
            return None;
        }
        data[8..8 + str_len]
            .iter()
            .map(|&b| codepage::decode_byte(b, cp))
            .collect()
    };

    Some(SheetEntry {
        name,
        bof_offset,
        visibility,
        sheet_type,
    })
}

// ── Sheet substream parsing ────────────────────────────────────────

struct CurrentRow {
    row: Option<u16>,
    cells: Vec<(u16, String)>,
}

impl CurrentRow {
    const fn new() -> Self {
        Self {
            row: None,
            cells: Vec::new(),
        }
    }

    fn push(
        &mut self,
        row: u16,
        col: u16,
        value: String,
    ) -> crate::error::Result<Option<Vec<(u16, String)>>> {
        check_col(col)?;
        if self.row.is_some_and(|r| r != row) {
            let flushed = std::mem::take(&mut self.cells);
            self.row = Some(row);
            self.cells.push((col, value));
            return Ok(Some(flushed));
        }
        self.row = Some(row);
        self.cells.push((col, value));
        Ok(None)
    }

    fn take(&mut self) -> Vec<(u16, String)> {
        self.row = None;
        std::mem::take(&mut self.cells)
    }
}

fn check_col(col: u16) -> crate::error::Result<()> {
    if usize::from(col) >= MAX_COLS {
        Err(BatdocError::Document("sheet exceeds 512 columns".into()))
    } else {
        Ok(())
    }
}

fn skip_sheet_bof(data: &[u8], bof_offset: u32) -> Option<usize> {
    let offset = bof_offset as usize;
    if offset + 4 > data.len() {
        return None;
    }
    let rec_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
    if rec_type != REC_BOF {
        return None;
    }
    let rec_len = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));
    Some(offset + 4 + rec_len)
}

struct SheetWalk<'a> {
    data: &'a [u8],
    offset: usize,
    pending_string_cell: Option<(u16, u16)>,
    pending_cells: Vec<(u16, u16, String)>,
}

impl<'a> SheetWalk<'a> {
    fn start(data: &'a [u8], bof_offset: u32) -> Option<Self> {
        Some(Self {
            data,
            offset: skip_sheet_bof(data, bof_offset)?,
            pending_string_cell: None,
            pending_cells: Vec::new(),
        })
    }

    #[allow(clippy::unnecessary_wraps)]
    fn next_cell(
        &mut self,
        sst: &StringArena,
        xf_styles: &XfStyles,
        cp: u16,
    ) -> crate::error::Result<Option<(u16, u16, String)>> {
        if !self.pending_cells.is_empty() {
            return Ok(Some(self.pending_cells.remove(0)));
        }
        while self.offset + 4 <= self.data.len() {
            let rec_type = u16::from_le_bytes([self.data[self.offset], self.data[self.offset + 1]]);
            let rec_len = usize::from(u16::from_le_bytes([
                self.data[self.offset + 2],
                self.data[self.offset + 3],
            ]));
            let rec_end = (self.offset + 4 + rec_len).min(self.data.len());
            let rec_data = &self.data[self.offset + 4..rec_end];
            self.offset = rec_end;

            let emitted = match rec_type {
                REC_EOF => return Ok(None),
                REC_LABELSST => handle_labelsst(rec_data, sst),
                REC_LABEL | REC_RSTRING => handle_label(rec_data, cp),
                REC_NUMBER => handle_number(rec_data, xf_styles),
                REC_RK => handle_rk(rec_data, xf_styles),
                REC_MULRK => {
                    let mut cells = handle_mulrk(rec_data, xf_styles);
                    if cells.is_empty() {
                        None
                    } else {
                        let first = cells.remove(0);
                        self.pending_cells = cells;
                        Some(first)
                    }
                }
                REC_FORMULA => handle_formula(rec_data, &mut self.pending_string_cell, xf_styles),
                REC_STRING => handle_string(rec_data, &mut self.pending_string_cell, cp),
                REC_BOOLERR => handle_boolerr(rec_data),
                _ => {
                    if rec_type != REC_CONTINUE {
                        self.pending_string_cell = None;
                    }
                    None
                }
            };
            if emitted.is_some() {
                return Ok(emitted);
            }
        }
        Ok(None)
    }
}

fn densify_plain_row(cells: &[(u16, String)]) -> Vec<String> {
    let Some(max_col) = cells.iter().map(|(col, _)| *col).max() else {
        return Vec::new();
    };
    let mut dense = vec![String::new(); usize::from(max_col) + 1];
    for (col, value) in cells {
        let idx = usize::from(*col);
        if idx < dense.len() {
            dense[idx].clone_from(value);
        }
    }
    dense
}

fn densify_markdown_row(cells: &[(u16, String)], shape: &TableShape) -> Vec<String> {
    let ncols = shape.last_col - shape.first_col + 1;
    let mut dense = vec![String::new(); ncols];
    for (col, value) in cells {
        let col = usize::from(*col);
        if col < shape.first_col || col > shape.last_col {
            continue;
        }
        dense[col - shape.first_col].clone_from(value);
    }
    dense
}

fn row_has_content(cells: &[(u16, String)]) -> bool {
    cells.iter().any(|(_, v)| !v.trim().is_empty())
}

fn emit_sheet_plain(
    data: &[u8],
    entry: &SheetEntry,
    sst: &StringArena,
    xf_styles: &XfStyles,
    cp: u16,
    emitted_any: &mut bool,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let Some(mut walk) = SheetWalk::start(data, entry.bof_offset) else {
        return Ok(());
    };
    let mut current = CurrentRow::new();
    let mut wrote_header = false;

    while let Some((row, col, value)) = walk.next_cell(sst, xf_styles, cp)? {
        if let Some(prev) = current.push(row, col, value)? {
            flush_plain_row(&prev, entry, &mut wrote_header, emitted_any, sink)?;
        }
    }
    let last = current.take();
    if !last.is_empty() {
        flush_plain_row(&last, entry, &mut wrote_header, emitted_any, sink)?;
    }
    Ok(())
}

fn flush_plain_row(
    cells: &[(u16, String)],
    entry: &SheetEntry,
    wrote_header: &mut bool,
    emitted_any: &mut bool,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let dense = densify_plain_row(cells);
    let mut line = String::new();
    write_plain_row(&mut line, dense.iter().map(String::as_str))?;
    if line.is_empty() {
        return Ok(());
    }
    if !*wrote_header {
        if *emitted_any {
            sink.write_str("\n--- ")?;
            sink.write_str(&entry.name)?;
            sink.write_str(" ---\n")?;
        }
        *wrote_header = true;
        *emitted_any = true;
    }
    sink.write_str(&line)
}

fn scan_sheet_shape(
    data: &[u8],
    bof_offset: u32,
    sst: &StringArena,
    xf_styles: &XfStyles,
    cp: u16,
) -> crate::error::Result<Option<TableShape>> {
    let Some(mut walk) = SheetWalk::start(data, bof_offset) else {
        return Ok(None);
    };
    let mut current = CurrentRow::new();
    let mut used = vec![false; MAX_COLS];
    let mut last_nonempty_row = None;
    let mut row_idx = 0usize;

    let mut mark = |cells: &[(u16, String)]| {
        if !row_has_content(cells) {
            row_idx += 1;
            return;
        }
        for (col, value) in cells {
            if !value.trim().is_empty() {
                used[usize::from(*col)] = true;
            }
        }
        last_nonempty_row = Some(row_idx);
        row_idx += 1;
    };

    while let Some((row, col, value)) = walk.next_cell(sst, xf_styles, cp)? {
        if let Some(prev) = current.push(row, col, value)? {
            mark(&prev);
        }
    }
    let last = current.take();
    if !last.is_empty() {
        mark(&last);
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
    data: &[u8],
    bof_offset: u32,
    shape: &TableShape,
    sst: &StringArena,
    xf_styles: &XfStyles,
    cp: u16,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    let Some(mut walk) = SheetWalk::start(data, bof_offset) else {
        return Ok(());
    };
    let mut current = CurrentRow::new();
    let mut row_idx = 0usize;
    let mut is_header = true;
    let ncols = shape.last_col - shape.first_col + 1;

    while let Some((row, col, value)) = walk.next_cell(sst, xf_styles, cp)? {
        if let Some(prev) = current.push(row, col, value)? {
            emit_markdown_row(&prev, shape, ncols, &mut row_idx, &mut is_header, sink)?;
        }
    }
    let last = current.take();
    if !last.is_empty() {
        emit_markdown_row(&last, shape, ncols, &mut row_idx, &mut is_header, sink)?;
    }
    sink.write_str("\n")
}

fn emit_markdown_row(
    cells: &[(u16, String)],
    shape: &TableShape,
    ncols: usize,
    row_idx: &mut usize,
    is_header: &mut bool,
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
    if *row_idx > shape.last_row {
        return Ok(());
    }
    let slice = densify_markdown_row(cells, shape);
    if *is_header {
        write_markdown_header(sink, &slice)?;
        write_markdown_separator(sink, ncols)?;
        *is_header = false;
    } else {
        write_markdown_data_row(sink, &slice)?;
    }
    *row_idx += 1;
    Ok(())
}

// ── Cell record handlers ───────────────────────────────────────────

fn handle_labelsst(rec_data: &[u8], sst: &StringArena) -> Option<(u16, u16, String)> {
    if rec_data.len() < 10 {
        return None;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let sst_idx = u32::from_le_bytes([rec_data[6], rec_data[7], rec_data[8], rec_data[9]]) as usize;
    let value = sst.get(sst_idx).unwrap_or("").to_string();
    Some((row, col, value))
}

fn handle_label(rec_data: &[u8], cp: u16) -> Option<(u16, u16, String)> {
    if rec_data.len() < 8 {
        return None;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let (s, _) = read_biff8_string(rec_data, 6, &[], cp);
    Some((row, col, s))
}

fn handle_number(rec_data: &[u8], xf_styles: &XfStyles) -> Option<(u16, u16, String)> {
    if rec_data.len() < 14 {
        return None;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let ixfe = u16::from_le_bytes([rec_data[4], rec_data[5]]);
    let val = f64::from_le_bytes([
        rec_data[6],
        rec_data[7],
        rec_data[8],
        rec_data[9],
        rec_data[10],
        rec_data[11],
        rec_data[12],
        rec_data[13],
    ]);
    Some((row, col, format_maybe_date(val, ixfe, xf_styles)))
}

fn handle_rk(rec_data: &[u8], xf_styles: &XfStyles) -> Option<(u16, u16, String)> {
    if rec_data.len() < 10 {
        return None;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let ixfe = u16::from_le_bytes([rec_data[4], rec_data[5]]);
    let rk = u32::from_le_bytes([rec_data[6], rec_data[7], rec_data[8], rec_data[9]]);
    Some((row, col, format_maybe_date(decode_rk(rk), ixfe, xf_styles)))
}

fn handle_mulrk(rec_data: &[u8], xf_styles: &XfStyles) -> Vec<(u16, u16, String)> {
    let mut cells = Vec::new();
    if rec_data.len() < 6 {
        return cells;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let first_col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let last_col = u16::from_le_bytes([rec_data[rec_data.len() - 2], rec_data[rec_data.len() - 1]]);
    let mut pos = 4;
    for c in first_col..=last_col {
        if pos + 6 > rec_data.len() - 2 {
            break;
        }
        let ixfe = u16::from_le_bytes([rec_data[pos], rec_data[pos + 1]]);
        let rk = u32::from_le_bytes([
            rec_data[pos + 2],
            rec_data[pos + 3],
            rec_data[pos + 4],
            rec_data[pos + 5],
        ]);
        cells.push((row, c, format_maybe_date(decode_rk(rk), ixfe, xf_styles)));
        pos += 6;
    }
    cells
}

fn handle_formula(
    rec_data: &[u8],
    pending_string_cell: &mut Option<(u16, u16)>,
    xf_styles: &XfStyles,
) -> Option<(u16, u16, String)> {
    if rec_data.len() < 20 {
        return None;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let ixfe = u16::from_le_bytes([rec_data[4], rec_data[5]]);
    let result_bytes = &rec_data[6..14];

    if result_bytes[6] == 0xFF && result_bytes[7] == 0xFF {
        match result_bytes[0] {
            0 => {
                *pending_string_cell = Some((row, col));
                None
            }
            1 => {
                let val = if result_bytes[2] != 0 {
                    "TRUE"
                } else {
                    "FALSE"
                };
                Some((row, col, val.to_string()))
            }
            3 => Some((row, col, String::new())),
            _ => None,
        }
    } else {
        let val = f64::from_le_bytes([
            result_bytes[0],
            result_bytes[1],
            result_bytes[2],
            result_bytes[3],
            result_bytes[4],
            result_bytes[5],
            result_bytes[6],
            result_bytes[7],
        ]);
        Some((row, col, format_maybe_date(val, ixfe, xf_styles)))
    }
}

fn handle_string(
    rec_data: &[u8],
    pending_string_cell: &mut Option<(u16, u16)>,
    cp: u16,
) -> Option<(u16, u16, String)> {
    let (row, col) = pending_string_cell.take()?;
    if rec_data.len() < 3 {
        return None;
    }
    let (s, _) = read_biff8_string(rec_data, 0, &[], cp);
    Some((row, col, s))
}

fn handle_boolerr(rec_data: &[u8]) -> Option<(u16, u16, String)> {
    if rec_data.len() < 8 {
        return None;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let is_error = rec_data[7];
    if is_error != 0 {
        return None;
    }
    let val = if rec_data[6] != 0 { "TRUE" } else { "FALSE" };
    Some((row, col, val.to_string()))
}

// ── Date-aware number formatting ────────────────────────────────────

/// Format a numeric value, converting to ISO date if the XF style is a date format.
fn format_maybe_date(val: f64, ixfe: u16, xf_styles: &XfStyles) -> String {
    if xf_styles.is_date_xf(ixfe) {
        dateconv::serial_to_iso(val)
    } else {
        format_number(val)
    }
}

// ── RK value decoding ──────────────────────────────────────────────

/// Decode an RK (compressed number) value.
///
/// RK encoding uses 4 bytes to store either an integer or a truncated IEEE 754
/// double. Bit 0 indicates /100 scaling, bit 1 indicates integer vs float.
fn decode_rk(rk: u32) -> f64 {
    let val = if rk & 0x02 != 0 {
        // Integer: bits 2..31 are a signed 30-bit integer.
        // The wrapping cast is intentional — RK uses the sign bit.
        #[allow(clippy::cast_possible_wrap)]
        let ival = rk.cast_signed() >> 2;
        f64::from(ival)
    } else {
        // IEEE 754 double with low 32 bits zeroed, bottom 2 bits of high word masked
        let hi = u64::from(rk & 0xFFFF_FFFC);
        let bits = hi << 32;
        f64::from_bits(bits)
    };

    if rk & 0x01 != 0 {
        val / 100.0
    } else {
        val
    }
}

// ── Number formatting ──────────────────────────────────────────────

/// Format a floating-point number for display.
/// Integers display without decimal point; others use default formatting.
fn format_number(val: f64) -> String {
    if val.is_nan() || val.is_infinite() {
        return val.to_string();
    }
    // If the value is an integer (and within f64's exact integer range), display without decimal.
    // 2^53 is the largest integer where f64 can represent all integers exactly.
    #[allow(clippy::cast_precision_loss, clippy::cast_possible_truncation)]
    if val.fract() == 0.0 && val.abs() < (1i64 << 53) as f64 {
        format!("{}", val as i64)
    } else {
        // Use a reasonable precision, stripping trailing zeros
        let s = format!("{val:.10}");
        let s = s.trim_end_matches('0');
        let s = s.trim_end_matches('.');
        s.to_string()
    }
}

// ── UTF-16LE decoding ──────────────────────────────────────────────

/// Decode a UTF-16LE byte slice into a String.
///
/// Uses `char::decode_utf16` to correctly handle surrogate pairs for
/// supplementary plane characters (emoji, CJK Extension B, etc.).
/// Invalid surrogates are replaced with U+FFFD.
pub(crate) fn decode_utf16le(data: &[u8]) -> String {
    let iter = data
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes(*pair));
    char::decode_utf16(iter)
        .map(|r| r.unwrap_or('\u{FFFD}'))
        .collect()
}

#[cfg(test)]
#[allow(clippy::float_cmp)]
mod tests {
    use super::*;

    // ── decode_rk ─────────────────────────────────────────────────

    #[test]
    fn rk_integer() {
        // Integer 1: (1 << 2) | 0x02 = 6
        let rk = (1u32 << 2) | 0x02;
        assert_eq!(decode_rk(rk), 1.0);
    }

    #[test]
    fn rk_integer_large() {
        // Integer 42: (42 << 2) | 0x02 = 170
        let rk = (42u32 << 2) | 0x02;
        assert_eq!(decode_rk(rk), 42.0);
    }

    #[test]
    fn rk_integer_zero() {
        let rk = 0x02u32; // (0 << 2) | 0x02
        assert_eq!(decode_rk(rk), 0.0);
    }

    #[test]
    fn rk_integer_negative() {
        // -1 as signed 30-bit: all bits set, shifted left 2, OR with 0x02
        let rk = ((-1i32).cast_unsigned() & 0xFFFF_FFFC) | 0x02;
        assert_eq!(decode_rk(rk), -1.0);
    }

    #[test]
    fn rk_integer_div100() {
        // Integer 150 / 100 = 1.5
        let rk = (150u32 << 2) | 0x02 | 0x01;
        assert_eq!(decode_rk(rk), 1.5);
    }

    #[test]
    fn rk_float() {
        // Encode 1.0 as IEEE double: 0x3FF0_0000_0000_0000
        // High 32 bits: 0x3FF00000, low 32 bits zeroed
        // RK stores the high 32 bits with bottom 2 bits masked
        let rk = 0x3FF0_0000u32; // bit 0 and 1 are 0
        assert_eq!(decode_rk(rk), 1.0);
    }

    #[test]
    fn rk_float_div100() {
        // 100.0 as double: 0x4059_0000_0000_0000
        // High 32 bits: 0x40590000
        // /100 => 1.0
        let rk = 0x4059_0000u32 | 0x01; // div100
        assert_eq!(decode_rk(rk), 1.0);
    }

    // ── format_number ─────────────────────────────────────────────

    #[test]
    fn format_integer() {
        assert_eq!(format_number(42.0), "42");
    }

    #[test]
    fn format_zero() {
        assert_eq!(format_number(0.0), "0");
    }

    #[test]
    fn format_negative_integer() {
        assert_eq!(format_number(-7.0), "-7");
    }

    #[test]
    fn format_float() {
        assert_eq!(format_number(3.125), "3.125");
    }

    #[test]
    fn format_float_trailing_zeros() {
        assert_eq!(format_number(1.5), "1.5");
    }

    // ── decode_utf16le ────────────────────────────────────────────

    #[test]
    fn utf16le_ascii() {
        let data = [0x48, 0x00, 0x69, 0x00]; // "Hi"
        assert_eq!(decode_utf16le(&data), "Hi");
    }

    #[test]
    fn utf16le_empty() {
        assert_eq!(decode_utf16le(&[]), "");
    }

    #[test]
    fn utf16le_surrogate_pair() {
        // U+1F600 (😀) = D83D DE00 in UTF-16LE
        let data = [0x3D, 0xD8, 0x00, 0xDE];
        assert_eq!(decode_utf16le(&data), "\u{1F600}");
    }

    #[test]
    fn utf16le_unpaired_surrogate() {
        // Lone high surrogate → U+FFFD
        let data = [0x3D, 0xD8, 0x48, 0x00]; // D83D then 'H'
        assert_eq!(decode_utf16le(&data), "\u{FFFD}H");
    }

    // ── read_biff8_string ─────────────────────────────────────────

    #[test]
    fn biff8_string_compressed() {
        // char_count=3, flags=0 (compressed), "ABC"
        let data = [0x03, 0x00, 0x00, b'A', b'B', b'C'];
        let (s, pos) = read_biff8_string(&data, 0, &[], 1252);
        assert_eq!(s, "ABC");
        assert_eq!(pos, 6);
    }

    #[test]
    fn biff8_string_unicode() {
        // char_count=2, flags=1 (unicode), "Hi" in UTF-16LE
        let data = [0x02, 0x00, 0x01, 0x48, 0x00, 0x69, 0x00];
        let (s, pos) = read_biff8_string(&data, 0, &[], 1252);
        assert_eq!(s, "Hi");
        assert_eq!(pos, 7);
    }

    #[test]
    fn biff8_string_with_offset() {
        // Some prefix data, then string at offset 3
        let data = [0xFF, 0xFF, 0xFF, 0x02, 0x00, 0x00, b'O', b'K'];
        let (s, pos) = read_biff8_string(&data, 3, &[], 1252);
        assert_eq!(s, "OK");
        assert_eq!(pos, 8);
    }

    #[test]
    fn biff8_string_with_rich_text() {
        // char_count=2, flags=0x08 (has rich), rich_runs=1, "AB", + 4 bytes rich data
        let data = [
            0x02, 0x00, 0x08, // header: 2 chars, rich flag
            0x01, 0x00, // 1 rich run
            b'A', b'B', // characters
            0x00, 0x00, 0x00, 0x00, // rich run data (4 bytes)
        ];
        let (s, pos) = read_biff8_string(&data, 0, &[], 1252);
        assert_eq!(s, "AB");
        assert_eq!(pos, 11);
    }

    #[test]
    fn biff8_string_empty() {
        let data = [0x00, 0x00, 0x00]; // 0 chars
        let (s, pos) = read_biff8_string(&data, 0, &[], 1252);
        assert_eq!(s, "");
        assert_eq!(pos, 3);
    }

    // ── parse_boundsheet ──────────────────────────────────────────

    #[test]
    fn boundsheet_compressed_name() {
        let mut data = vec![
            0x00, 0x10, 0x00, 0x00, // bof_offset = 0x1000
            0x00, // visible
            0x00, // worksheet
            0x05, // name length = 5
            0x00, // compressed
        ];
        data.extend_from_slice(b"Sheet");
        let entry = parse_boundsheet(&data, 1252).unwrap();
        assert_eq!(entry.name, "Sheet");
        assert_eq!(entry.bof_offset, 0x1000);
        assert_eq!(entry.visibility, 0);
        assert_eq!(entry.sheet_type, 0);
    }

    #[test]
    fn boundsheet_unicode_name() {
        let mut data = vec![
            0x00, 0x20, 0x00, 0x00, // bof_offset = 0x2000
            0x01, // hidden
            0x00, // worksheet
            0x02, // name length = 2 characters
            0x01, // unicode
        ];
        // "Hi" in UTF-16LE
        data.extend_from_slice(&[0x48, 0x00, 0x69, 0x00]);
        let entry = parse_boundsheet(&data, 1252).unwrap();
        assert_eq!(entry.name, "Hi");
        assert_eq!(entry.visibility, 1);
    }

    #[test]
    fn boundsheet_too_short() {
        let data = vec![0x00, 0x00, 0x00];
        assert!(parse_boundsheet(&data, 1252).is_none());
    }

    // ── parse_sst ─────────────────────────────────────────────────

    #[test]
    fn sst_basic() {
        // SST with 2 unique strings: "Hi" and "Go"
        let data = vec![
            0x02, 0x00, 0x00, 0x00, // total refs = 2
            0x02, 0x00, 0x00, 0x00, // unique = 2
            // String 1: "Hi" compressed
            0x02, 0x00, 0x00, b'H', b'i', // String 2: "Go" compressed
            0x02, 0x00, 0x00, b'G', b'o',
        ];
        let mut arena = StringArena::new();
        parse_sst(&data, &[data.len()], 1252, &mut arena);
        assert_eq!(arena.get(0), Some("Hi"));
        assert_eq!(arena.get(1), Some("Go"));
        assert_eq!(arena.len(), 2);
    }

    #[test]
    fn sst_unicode_string() {
        let data = vec![
            0x01, 0x00, 0x00, 0x00, // total refs = 1
            0x01, 0x00, 0x00, 0x00, // unique = 1
            // String: "A" in unicode
            0x01, 0x00, 0x01, 0x41, 0x00,
        ];
        let mut arena = StringArena::new();
        parse_sst(&data, &[data.len()], 1252, &mut arena);
        assert_eq!(arena.get(0), Some("A"));
        assert_eq!(arena.len(), 1);
    }

    #[test]
    fn sst_empty() {
        let data = vec![0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let mut arena = StringArena::new();
        parse_sst(&data, &[data.len()], 1252, &mut arena);
        assert_eq!(arena.len(), 0);
    }

    // ── SST with CONTINUE boundary ────────────────────────────────

    #[test]
    fn sst_continue_boundary_compressed() {
        // A string that spans a CONTINUE boundary
        // First part: "He" (2 of 5 chars), then CONTINUE boundary, then "llo"
        let mut data = vec![
            0x01, 0x00, 0x00, 0x00, // total refs = 1
            0x01, 0x00, 0x00, 0x00, // unique = 1
            // String header: 5 chars, compressed
            0x05, 0x00, 0x00, b'H', b'e', // 2 chars before boundary
        ];
        let boundary = data.len(); // CONTINUE starts here
                                   // At CONTINUE boundary: grbit byte (0 = still compressed)
        data.push(0x00);
        data.extend_from_slice(b"llo");

        let mut arena = StringArena::new();
        parse_sst(&data, &[boundary], 1252, &mut arena);
        assert_eq!(arena.get(0), Some("Hello"));
    }

    #[test]
    fn sst_continue_boundary_encoding_switch() {
        // String starts compressed, switches to unicode at CONTINUE boundary
        let mut data = vec![
            0x01, 0x00, 0x00, 0x00, // total refs = 1
            0x01, 0x00, 0x00, 0x00, // unique = 1
            // String header: 3 chars, compressed
            0x03, 0x00, 0x00, b'A', // 1 char before boundary
        ];
        let boundary = data.len();
        // At CONTINUE: grbit=1 (switch to unicode)
        data.push(0x01);
        // "BC" in UTF-16LE
        data.extend_from_slice(&[0x42, 0x00, 0x43, 0x00]);

        let mut arena = StringArena::new();
        parse_sst(&data, &[boundary], 1252, &mut arena);
        assert_eq!(arena.get(0), Some("ABC"));
    }

    // ── streaming extract (OLE fixtures) ──────────────────────────

    use std::io::Write;

    fn rec(rec_type: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + data.len());
        out.extend_from_slice(&rec_type.to_le_bytes());
        out.extend_from_slice(&(u16::try_from(data.len()).unwrap()).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    fn bof(dt: u16) -> Vec<u8> {
        let mut data = vec![0; 16];
        data[0..2].copy_from_slice(&0x0600u16.to_le_bytes());
        data[2..4].copy_from_slice(&dt.to_le_bytes());
        rec(REC_BOF, &data)
    }

    fn eof() -> Vec<u8> {
        rec(REC_EOF, &[])
    }

    fn boundsheet(bof_offset: u32, name: &str) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&bof_offset.to_le_bytes());
        data.push(0);
        data.push(0);
        data.push(u8::try_from(name.len()).unwrap());
        data.push(0);
        data.extend_from_slice(name.as_bytes());
        rec(REC_BOUNDSHEET, &data)
    }

    fn label(row: u16, col: u16, text: &str) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&row.to_le_bytes());
        data.extend_from_slice(&col.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&(u16::try_from(text.len()).unwrap()).to_le_bytes());
        data.push(0);
        data.extend_from_slice(text.as_bytes());
        rec(REC_LABEL, &data)
    }

    fn ole_workbook(workbook: &[u8]) -> Vec<u8> {
        let mut cfb = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        {
            let mut stream = cfb.create_stream("/Workbook").unwrap();
            stream.write_all(workbook).unwrap();
        }
        cfb.flush().unwrap();
        cfb.into_inner().into_inner()
    }

    fn xls_with_label(row: u16, col: u16, text: &str) -> Vec<u8> {
        xls_with_sheets(&[("Sheet1", vec![label(row, col, text)])])
    }

    fn xls_with_sheets(sheets: &[(&str, Vec<Vec<u8>>)]) -> Vec<u8> {
        let sheet_payloads: Vec<Vec<u8>> = sheets
            .iter()
            .map(|(_, cells)| {
                let mut sheet = bof(0x0010);
                for cell in cells {
                    sheet.extend_from_slice(cell);
                }
                sheet.extend_from_slice(&eof());
                sheet
            })
            .collect();

        let boundsheet_lens: Vec<usize> = sheets.iter().map(|(name, _)| 12 + name.len()).collect();
        let globals_len: usize = 20 + boundsheet_lens.iter().sum::<usize>() + 4;

        let mut workbook = bof(0x0005);
        let mut offset = globals_len;
        for ((name, _), payload) in sheets.iter().zip(sheet_payloads.iter()) {
            workbook.extend_from_slice(&boundsheet(u32::try_from(offset).unwrap(), name));
            offset += payload.len();
        }
        workbook.extend_from_slice(&eof());
        for payload in &sheet_payloads {
            workbook.extend_from_slice(payload);
        }
        ole_workbook(&workbook)
    }

    fn two_row_xls() -> Vec<u8> {
        xls_with_sheets(&[(
            "Sheet1",
            vec![
                label(0, 0, "Name"),
                label(0, 1, "Age"),
                label(1, 0, "Alice"),
                label(1, 1, "30"),
            ],
        )])
    }

    #[test]
    fn xls_plain_rejects_col_512() {
        let data = xls_with_label(0, 512, "X");
        let err = crate::extract_plain(&data, crate::Format::Xls)
            .unwrap_err()
            .to_string();
        assert_eq!(err, "sheet exceeds 512 columns");
    }

    #[test]
    fn xls_plain_accepts_col_511() {
        let data = xls_with_label(0, 511, "X");
        let text = crate::extract_plain(&data, crate::Format::Xls).unwrap();
        assert!(text.contains("X"));
    }

    #[test]
    fn xls_plain_to_equals_extract_plain_on_current_fixtures() {
        let data = two_row_xls();
        let buffered = extract_plain(&data).unwrap();
        let mut streamed = String::new();
        extract_plain_to(&data, &mut streamed).unwrap();
        assert_eq!(buffered, streamed);
        assert_eq!(buffered, "Name\tAge\nAlice\t30\n");
    }

    #[test]
    fn xls_markdown_two_rows_locked() {
        let data = two_row_xls();
        let md = extract_markdown(&data).unwrap();
        assert_eq!(md, "| Name | Age |\n| --- | --- |\n| Alice | 30 |\n\n");
    }

    #[test]
    fn xls_plain_multi_sheet_headers() {
        let data = xls_with_sheets(&[
            ("People", vec![label(0, 0, "Alice")]),
            ("Places", vec![label(0, 0, "NYC")]),
        ]);
        let text = extract_plain(&data).unwrap();
        assert_eq!(text, "Alice\n\n--- Places ---\nNYC\n");
    }

    #[test]
    fn xls_markdown_multi_sheet_headings() {
        let data = xls_with_sheets(&[
            ("People", vec![label(0, 0, "Alice")]),
            ("Places", vec![label(0, 0, "NYC")]),
        ]);
        let md = extract_markdown(&data).unwrap();
        assert_eq!(
            md,
            "## People\n\n| Alice |\n| --- |\n\n## Places\n\n| NYC |\n| --- |\n\n"
        );
    }

    #[test]
    fn xls_markdown_skips_empty_sheet_and_omits_single_heading() {
        let data = xls_with_sheets(&[
            ("Empty", vec![label(0, 0, "  ")]),
            ("Data", vec![label(0, 0, "Hello")]),
        ]);
        let md = extract_markdown(&data).unwrap();
        assert_eq!(md, "| Hello |\n| --- |\n\n");
    }

    #[test]
    fn xls_sheets_two_rows() {
        let data = two_row_xls();
        let mut sheets = Vec::<crate::Sheet>::new();
        extract_sheets_to(&data, &mut sheets).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        let expected: Vec<Vec<String>> = vec![
            vec!["Name".into(), "Age".into()],
            vec!["Alice".into(), "30".into()],
        ];
        assert_eq!(sheets[0].rows, expected);
    }

    #[allow(clippy::assert_is_empty)] // intentionally brief-verbatim assertion
    #[test]
    fn xls_sheets_multi_and_empty_sheet_still_begins() {
        let data = xls_with_sheets(&[
            ("Empty", vec![label(0, 0, "  ")]),
            ("Data", vec![label(0, 0, "Hello")]),
        ]);
        let mut sheets = Vec::<crate::Sheet>::new();
        extract_sheets_to(&data, &mut sheets).unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Empty");
        assert!(sheets[0].rows.is_empty()); // whitespace-only row dropped
        let expected: Vec<Vec<String>> = vec![vec!["Hello".into()]];
        assert_eq!(sheets[1].rows, expected);
    }

    #[test]
    fn xls_sheets_rejects_col_512() {
        let data = xls_with_label(0, 512, "X");
        let mut sheets = Vec::<crate::Sheet>::new();
        let err = extract_sheets_to(&data, &mut sheets)
            .unwrap_err()
            .to_string();
        assert_eq!(err, "sheet exceeds 512 columns");
    }

    #[test]
    fn xls_sheets_sparse_gap() {
        let data = xls_with_sheets(&[("S", vec![label(0, 0, "First"), label(0, 2, "Third")])]);
        let mut sheets = Vec::<crate::Sheet>::new();
        extract_sheets_to(&data, &mut sheets).unwrap();
        assert_eq!(
            sheets[0].rows,
            vec![vec!["First".into(), String::new(), "Third".into()]]
        );
    }
}
