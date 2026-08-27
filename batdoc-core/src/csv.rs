//! RFC 4180 CSV helpers and a `SheetSink` → byte-stream adapter.

use crate::error::Result;
use crate::sheets::SheetSink;
use crate::ExtractSink;

/// Quote per RFC 4180 iff field contains comma, `"`, CR, or LF.
pub fn escape_field(field: &str, out: &mut String) {
    let needs = field
        .bytes()
        .any(|b| b == b',' || b == b'"' || b == b'\n' || b == b'\r');
    if !needs {
        out.push_str(field);
        return;
    }
    out.push('"');
    for ch in field.chars() {
        if ch == '"' {
            out.push('"');
            out.push('"');
        } else {
            out.push(ch);
        }
    }
    out.push('"');
}

/// Escape each field, join with `,`, terminate with CRLF.
#[must_use]
pub fn to_csv_row<C: AsRef<str>>(cells: &[C]) -> String {
    let mut s = String::new();
    for (i, c) in cells.iter().enumerate() {
        if i > 0 {
            s.push(',');
        }
        escape_field(c.as_ref(), &mut s);
    }
    s.push('\r');
    s.push('\n');
    s
}

/// Escape hatch: one CSV stream for a whole workbook. Prefer per-sheet sinks.
pub struct CsvSink<W: ExtractSink> {
    inner: W,
    pending_name: Option<String>,
    content_emitted: bool,
}

impl<W: ExtractSink> CsvSink<W> {
    pub const fn new(inner: W) -> Self {
        Self {
            inner,
            pending_name: None,
            content_emitted: false,
        }
    }
}

impl<W: ExtractSink> SheetSink for CsvSink<W> {
    fn begin_sheet(&mut self, name: &str) -> Result<()> {
        self.pending_name = Some(name.to_string());
        Ok(())
    }

    fn row(&mut self, cells: Vec<String>) -> Result<()> {
        if let Some(name) = self.pending_name.take() {
            if self.content_emitted {
                self.inner.write_str("\n--- ")?;
                self.inner.write_str(&name)?;
                self.inner.write_str(" ---\n")?;
            }
        }
        let line = to_csv_row(&cells);
        self.inner.write_str(&line)?;
        self.content_emitted = true;
        Ok(())
    }

    fn end_sheet(&mut self) -> Result<()> {
        self.pending_name = None;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::SheetSink;

    #[test]
    fn escape_field_cases() {
        let mut s = String::new();
        escape_field("plain", &mut s);
        assert_eq!(s, "plain");
        s.clear();
        escape_field("a,b", &mut s);
        assert_eq!(s, "\"a,b\"");
        s.clear();
        escape_field("say \"hi\"", &mut s);
        assert_eq!(s, "\"say \"\"hi\"\"\"");
        s.clear();
        escape_field("a\nb", &mut s);
        assert_eq!(s, "\"a\nb\"");
        s.clear();
        escape_field("a\rb", &mut s);
        assert_eq!(s, "\"a\rb\"");
        s.clear();
        escape_field("", &mut s);
        assert_eq!(s, "");
    }

    #[test]
    fn to_csv_row_crlf() {
        assert_eq!(to_csv_row(&["a", "b"]), "a,b\r\n");
        assert_eq!(to_csv_row(&["a,b", "c"]), "\"a,b\",c\r\n");
    }

    #[test]
    fn csv_sink_single_sheet() {
        let mut out = String::new();
        let mut sink = CsvSink::new(&mut out);
        sink.begin_sheet("S").unwrap();
        sink.row(vec!["x".into(), "y".into()]).unwrap();
        sink.end_sheet().unwrap();
        assert_eq!(out, "x,y\r\n");
    }

    #[test]
    fn csv_sink_two_content_sheets_separator() {
        let mut out = String::new();
        let mut sink = CsvSink::new(&mut out);
        sink.begin_sheet("People").unwrap();
        sink.row(vec!["Alice".into()]).unwrap();
        sink.end_sheet().unwrap();
        sink.begin_sheet("Places").unwrap();
        sink.row(vec!["NYC".into()]).unwrap();
        sink.end_sheet().unwrap();
        assert_eq!(out, "Alice\r\n\n--- Places ---\nNYC\r\n");
    }

    #[test]
    fn csv_sink_empty_first_sheet_no_separator() {
        let mut out = String::new();
        let mut sink = CsvSink::new(&mut out);
        sink.begin_sheet("Empty").unwrap();
        sink.end_sheet().unwrap();
        sink.begin_sheet("Data").unwrap();
        sink.row(vec!["Hello".into()]).unwrap();
        sink.end_sheet().unwrap();
        assert_eq!(out, "Hello\r\n");
    }
}
