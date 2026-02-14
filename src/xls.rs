//! Legacy BIFF8 `.xls` (Excel 97+) binary format parser.
//!
//! Reads the `Workbook` (or `Book`) stream from the OLE2 compound file,
//! parses the BIFF8 record stream to extract the Shared String Table (SST),
//! sheet metadata (`BoundSheet8`), and cell records (LABELSST, NUMBER, RK,
//! MULRK, FORMULA, LABEL, BOOLERR). Produces the same `Sheet` type used
//! by the `.xlsx` parser for rendering.

use cfb::CompoundFile;
use std::io::{Cursor, Read};

use crate::codepage;
use crate::dateconv;
use crate::error::BatdocError;
use crate::sheet::Sheet;

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
    let sheets = parse_xls(data)?;
    Ok(crate::sheet::render_plain(&sheets))
}

/// Extract markdown-formatted text from a BIFF8 .xls file.
pub(crate) fn extract_markdown(data: &[u8]) -> crate::error::Result<String> {
    let sheets = parse_xls(data)?;
    Ok(crate::sheet::render_markdown(&sheets))
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

/// A cell being placed into the grid.
#[derive(Debug)]
struct Cell {
    row: u16,
    col: u16,
    value: String,
}

// ── Main parser ────────────────────────────────────────────────────

fn parse_xls(data: &[u8]) -> crate::error::Result<Vec<Sheet>> {
    let cursor = Cursor::new(data);
    let mut cfb = CompoundFile::open(cursor)?;

    // Try "Workbook" first (BIFF8), then "Book" (BIFF5 compat)
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

    // Parse all records
    let records = parse_records(&buf);

    // Phase 1: Parse workbook globals (SST + sheet entries + XF styles + codepage)
    // This also detects encryption (FILEPASS record) early.
    let (sst, sheet_entries, xf_styles, cp) = parse_globals(&records)?;

    // Phase 2: Parse each worksheet substream
    let mut sheets = Vec::new();
    for entry in &sheet_entries {
        // Skip non-worksheet types (charts, macros, VB modules)
        if entry.sheet_type != 0 {
            continue;
        }
        // Skip hidden sheets
        if entry.visibility != 0 {
            continue;
        }
        let rows = parse_sheet_substream(&buf, entry.bof_offset, &sst, &xf_styles, cp);
        sheets.push(Sheet {
            name: entry.name.clone(),
            rows,
        });
    }

    Ok(sheets)
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
) -> crate::error::Result<(Vec<String>, Vec<SheetEntry>, XfStyles, u16)> {
    let mut sst = Vec::new();
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
            REC_CODEPAGE => {
                if rec.data.len() >= 2 {
                    cp = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                }
            }
            REC_FORMAT => {
                if let Some((id, code)) = parse_format_record(rec.data, cp) {
                    custom_formats.push((id, code));
                }
            }
            REC_XF => {
                // XF record: bytes 2-3 are numFmtId
                if rec.data.len() >= 4 {
                    let fmt_id = u16::from_le_bytes([rec.data[2], rec.data[3]]);
                    xf_fmt_ids.push(fmt_id);
                }
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
                sst = parse_sst(&combined, &continue_boundaries, cp);
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
fn parse_sst(data: &[u8], continue_boundaries: &[usize], cp: u16) -> Vec<String> {
    if data.len() < 8 {
        return Vec::new();
    }

    let unique_count = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
    let mut strings = Vec::with_capacity(unique_count.min(65536));
    let mut pos = 8;

    for _ in 0..unique_count {
        if pos + 3 > data.len() {
            break;
        }

        let (s, new_pos) = read_biff8_string(data, pos, continue_boundaries, cp);
        strings.push(s);
        pos = new_pos;
    }

    strings
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

/// Accumulator for collecting cells while parsing a sheet substream.
struct GridBuilder {
    cells: Vec<Cell>,
    max_row: usize,
    max_col: usize,
}

impl GridBuilder {
    const fn new() -> Self {
        Self {
            cells: Vec::new(),
            max_row: 0,
            max_col: 0,
        }
    }

    fn push(&mut self, row: u16, col: u16, value: String) {
        let r = usize::from(row);
        let c = usize::from(col);
        if r + 1 > self.max_row {
            self.max_row = r + 1;
        }
        if c + 1 > self.max_col {
            self.max_col = c + 1;
        }
        self.cells.push(Cell { row, col, value });
    }

    fn into_grid(self) -> Vec<Vec<String>> {
        cells_to_grid(self.cells, self.max_row, self.max_col)
    }
}

/// Parse a sheet substream starting at `bof_offset` in the raw data,
/// extracting cell values into a 2D grid.
fn parse_sheet_substream(
    data: &[u8],
    bof_offset: u32,
    sst: &[String],
    xf_styles: &XfStyles,
    cp: u16,
) -> Vec<Vec<String>> {
    let mut grid = GridBuilder::new();
    let mut offset = bof_offset as usize; // u32 → usize: lossless on 32+ bit
    let mut pending_string_cell: Option<(u16, u16)> = None;

    // Verify BOF
    if offset + 4 > data.len() {
        return Vec::new();
    }
    let rec_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
    if rec_type != REC_BOF {
        return Vec::new();
    }

    // Skip BOF record
    let rec_len = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));
    offset += 4 + rec_len;

    while offset + 4 <= data.len() {
        let rec_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let rec_len = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));
        let rec_end = (offset + 4 + rec_len).min(data.len());
        let rec_data = &data[offset + 4..rec_end];

        match rec_type {
            REC_EOF => break,
            REC_LABELSST => handle_labelsst(rec_data, sst, &mut grid),
            REC_LABEL | REC_RSTRING => handle_label(rec_data, &mut grid, cp),
            REC_NUMBER => handle_number(rec_data, &mut grid, xf_styles),
            REC_RK => handle_rk(rec_data, &mut grid, xf_styles),
            REC_MULRK => handle_mulrk(rec_data, &mut grid, xf_styles),
            REC_FORMULA => {
                handle_formula(rec_data, &mut grid, &mut pending_string_cell, xf_styles);
            }
            REC_STRING => handle_string(rec_data, &mut grid, &mut pending_string_cell, cp),
            REC_BOOLERR => handle_boolerr(rec_data, &mut grid),
            _ => {
                // Clear pending string cell on any non-STRING record
                // (STRING must immediately follow FORMULA)
                if rec_type != REC_CONTINUE {
                    pending_string_cell = None;
                }
            }
        }

        offset = rec_end;
    }

    grid.into_grid()
}

// ── Cell record handlers ───────────────────────────────────────────

fn handle_labelsst(rec_data: &[u8], sst: &[String], grid: &mut GridBuilder) {
    if rec_data.len() >= 10 {
        let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
        let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
        let sst_idx =
            u32::from_le_bytes([rec_data[6], rec_data[7], rec_data[8], rec_data[9]]) as usize;
        let value = sst.get(sst_idx).cloned().unwrap_or_default();
        grid.push(row, col, value);
    }
}

fn handle_label(rec_data: &[u8], grid: &mut GridBuilder, cp: u16) {
    if rec_data.len() >= 8 {
        let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
        let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
        let (s, _) = read_biff8_string(rec_data, 6, &[], cp);
        grid.push(row, col, s);
    }
}

fn handle_number(rec_data: &[u8], grid: &mut GridBuilder, xf_styles: &XfStyles) {
    if rec_data.len() >= 14 {
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
        grid.push(row, col, format_maybe_date(val, ixfe, xf_styles));
    }
}

fn handle_rk(rec_data: &[u8], grid: &mut GridBuilder, xf_styles: &XfStyles) {
    if rec_data.len() >= 10 {
        let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
        let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
        let ixfe = u16::from_le_bytes([rec_data[4], rec_data[5]]);
        let rk = u32::from_le_bytes([rec_data[6], rec_data[7], rec_data[8], rec_data[9]]);
        grid.push(row, col, format_maybe_date(decode_rk(rk), ixfe, xf_styles));
    }
}

fn handle_mulrk(rec_data: &[u8], grid: &mut GridBuilder, xf_styles: &XfStyles) {
    if rec_data.len() >= 6 {
        let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
        let first_col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
        let last_col =
            u16::from_le_bytes([rec_data[rec_data.len() - 2], rec_data[rec_data.len() - 1]]);
        let mut pos = 4;
        for c in first_col..=last_col {
            if pos + 6 > rec_data.len() - 2 {
                break;
            }
            // Each MULRK entry: 2 bytes ixfe + 4 bytes RK value
            let ixfe = u16::from_le_bytes([rec_data[pos], rec_data[pos + 1]]);
            let rk = u32::from_le_bytes([
                rec_data[pos + 2],
                rec_data[pos + 3],
                rec_data[pos + 4],
                rec_data[pos + 5],
            ]);
            grid.push(row, c, format_maybe_date(decode_rk(rk), ixfe, xf_styles));
            pos += 6;
        }
    }
}

fn handle_formula(
    rec_data: &[u8],
    grid: &mut GridBuilder,
    pending_string_cell: &mut Option<(u16, u16)>,
    xf_styles: &XfStyles,
) {
    if rec_data.len() < 20 {
        return;
    }
    let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
    let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
    let ixfe = u16::from_le_bytes([rec_data[4], rec_data[5]]);
    let result_bytes = &rec_data[6..14];

    // Special type (string, bool, error, empty) indicated by 0xFFFF marker
    if result_bytes[6] == 0xFF && result_bytes[7] == 0xFF {
        match result_bytes[0] {
            0 => {
                // String result — follows in a STRING record
                *pending_string_cell = Some((row, col));
            }
            1 => {
                let val = if result_bytes[2] != 0 {
                    "TRUE"
                } else {
                    "FALSE"
                };
                grid.push(row, col, val.to_string());
            }
            // 2 = error, skip. 3 = empty string.
            3 => grid.push(row, col, String::new()),
            _ => {}
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
        grid.push(row, col, format_maybe_date(val, ixfe, xf_styles));
    }
}

fn handle_string(
    rec_data: &[u8],
    grid: &mut GridBuilder,
    pending_string_cell: &mut Option<(u16, u16)>,
    cp: u16,
) {
    if let Some((row, col)) = pending_string_cell.take() {
        if rec_data.len() >= 3 {
            let (s, _) = read_biff8_string(rec_data, 0, &[], cp);
            grid.push(row, col, s);
        }
    }
}

fn handle_boolerr(rec_data: &[u8], grid: &mut GridBuilder) {
    if rec_data.len() >= 8 {
        let row = u16::from_le_bytes([rec_data[0], rec_data[1]]);
        let col = u16::from_le_bytes([rec_data[2], rec_data[3]]);
        let is_error = rec_data[7];
        if is_error == 0 {
            let val = if rec_data[6] != 0 { "TRUE" } else { "FALSE" };
            grid.push(row, col, val.to_string());
        }
    }
}

/// Maximum grid cells to allocate (defense-in-depth against crafted files
/// with extreme row/col indices that would cause OOM).
const MAX_GRID_CELLS: usize = 1_000_000;

/// Convert sparse cell list into a dense 2D grid.
///
/// Returns an empty grid if the dimensions would exceed `MAX_GRID_CELLS`.
fn cells_to_grid(cells: Vec<Cell>, max_row: usize, max_col: usize) -> Vec<Vec<String>> {
    if max_row
        .checked_mul(max_col)
        .is_none_or(|n| n > MAX_GRID_CELLS)
    {
        // Dimensions too large — return empty rather than OOM
        return Vec::new();
    }

    let mut grid: Vec<Vec<String>> = vec![vec![String::new(); max_col]; max_row];

    for cell in cells {
        let r = usize::from(cell.row);
        let c = usize::from(cell.col);
        if r < max_row && c < max_col {
            grid[r][c] = cell.value;
        }
    }

    grid
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
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
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

    // ── cells_to_grid ─────────────────────────────────────────────

    #[test]
    fn grid_basic() {
        let cells = vec![
            Cell {
                row: 0,
                col: 0,
                value: "A".into(),
            },
            Cell {
                row: 0,
                col: 1,
                value: "B".into(),
            },
            Cell {
                row: 1,
                col: 0,
                value: "C".into(),
            },
        ];
        let grid = cells_to_grid(cells, 2, 2);
        assert_eq!(grid.len(), 2);
        assert_eq!(grid[0], vec!["A", "B"]);
        assert_eq!(grid[1], vec!["C", ""]);
    }

    #[test]
    fn grid_sparse() {
        let cells = vec![Cell {
            row: 0,
            col: 2,
            value: "X".into(),
        }];
        let grid = cells_to_grid(cells, 1, 3);
        assert_eq!(grid[0], vec!["", "", "X"]);
    }

    #[test]
    fn grid_empty() {
        let grid = cells_to_grid(Vec::new(), 0, 0);
        assert!(grid.is_empty());
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
        let strings = parse_sst(&data, &[data.len()], 1252);
        assert_eq!(strings, vec!["Hi", "Go"]);
    }

    #[test]
    fn sst_unicode_string() {
        let data = vec![
            0x01, 0x00, 0x00, 0x00, // total refs = 1
            0x01, 0x00, 0x00, 0x00, // unique = 1
            // String: "A" in unicode
            0x01, 0x00, 0x01, 0x41, 0x00,
        ];
        let strings = parse_sst(&data, &[data.len()], 1252);
        assert_eq!(strings, vec!["A"]);
    }

    #[test]
    fn sst_empty() {
        let data = vec![0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let strings = parse_sst(&data, &[data.len()], 1252);
        assert!(strings.is_empty());
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

        let strings = parse_sst(&data, &[boundary], 1252);
        assert_eq!(strings, vec!["Hello"]);
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

        let strings = parse_sst(&data, &[boundary], 1252);
        assert_eq!(strings, vec!["ABC"]);
    }
}
