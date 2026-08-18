//! Shared spreadsheet row writers used by `.xlsx` and `.xls` parsers.

use crate::error::Result;
use crate::ExtractSink;

pub(crate) const MAX_COLS: usize = 512;

/// Occupancy from pass 1.
pub(crate) struct TableShape {
    pub first_col: usize,
    pub last_col: usize,
    pub last_row: usize, // inclusive, 0-based among observed rows in order
}

pub(crate) fn write_plain_row(
    sink: &mut impl ExtractSink,
    cells: impl IntoIterator<Item = impl AsRef<str>>,
) -> Result<()> {
    let mut line = String::new();
    let mut first = true;
    for cell in cells {
        if !first {
            line.push('\t');
        }
        first = false;
        line.push_str(cell.as_ref());
    }
    let line = line.trim_end();
    if line.is_empty() {
        return Ok(());
    }
    sink.write_str(line)?;
    sink.write_str("\n")
}

pub(crate) fn write_markdown_header<C: AsRef<str>>(
    sink: &mut impl ExtractSink,
    cells: &[C],
) -> Result<()> {
    write_markdown_cells(sink, cells)
}

pub(crate) fn write_markdown_separator(sink: &mut impl ExtractSink, ncols: usize) -> Result<()> {
    sink.write_str("|")?;
    for _ in 0..ncols {
        sink.write_str(" --- |")?;
    }
    sink.write_str("\n")
}

pub(crate) fn write_markdown_data_row<C: AsRef<str>>(
    sink: &mut impl ExtractSink,
    cells: &[C],
) -> Result<()> {
    write_markdown_cells(sink, cells)
}

fn write_markdown_cells<C: AsRef<str>>(sink: &mut impl ExtractSink, cells: &[C]) -> Result<()> {
    sink.write_str("| ")?;
    for (i, cell) in cells.iter().enumerate() {
        if i > 0 {
            sink.write_str(" | ")?;
        }
        write_escaped_pipes(sink, cell.as_ref())?;
    }
    sink.write_str(" |\n")
}

fn write_escaped_pipes(sink: &mut impl ExtractSink, s: &str) -> Result<()> {
    let mut rest = s;
    while let Some(idx) = rest.find('|') {
        sink.write_str(&rest[..idx])?;
        sink.write_str("\\|")?;
        rest = &rest[idx + 1..];
    }
    sink.write_str(rest)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn write_plain_row_skips_empty() {
        let mut s = String::new();
        write_plain_row(&mut s, ["", "  "]).unwrap();
        assert_eq!(s, "");
    }

    #[test]
    fn write_plain_row_tabs_and_trim() {
        let mut s = String::new();
        write_plain_row(&mut s, ["a", "b", ""]).unwrap();
        assert_eq!(s, "a\tb\n");
    }

    #[test]
    fn write_markdown_escapes_pipe() {
        let mut s = String::new();
        write_markdown_header(&mut s, &["A|B"]).unwrap();
        write_markdown_separator(&mut s, 1).unwrap();
        write_markdown_data_row(&mut s, &["C"]).unwrap();
        assert_eq!(s, "| A\\|B |\n| --- |\n| C |\n");
    }

    #[test]
    fn max_cols_constant_is_512() {
        assert_eq!(MAX_COLS, 512);
    }
}
