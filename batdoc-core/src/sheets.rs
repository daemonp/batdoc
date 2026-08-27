//! Structured tabular extraction sinks (XLS / XLSX).
//!
//! Streaming `SheetSink` is the primary API. `Sheet` / `Vec<Sheet>` is a
//! collecting convenience that is O(total cells) — prefer the sink on
//! large workbooks (128 MiB Worker isolates).

use crate::error::{BatdocError, Result};

/// One worksheet, collected. Convenience only — holds the whole sheet.
#[derive(Debug, Clone, Default)]
pub struct Sheet {
    pub name: String,
    pub rows: Vec<Vec<String>>,
}

/// Row-granular sink. Lifecycle: `begin_sheet` → `row`* → `end_sheet`.
/// On `Err`, the extractor stops (no further calls, including `end_sheet`
/// after a failed `row`). No `Send`/`Sync` bound in v1.
pub trait SheetSink {
    /// # Errors
    /// Implementation-defined sink failures.
    fn begin_sheet(&mut self, name: &str) -> Result<()>;

    /// One densified, trailing-trimmed row with at least one trim-nonempty cell.
    ///
    /// # Errors
    /// Implementation-defined sink failures.
    fn row(&mut self, cells: Vec<String>) -> Result<()>;

    /// # Errors
    /// Implementation-defined sink failures.
    fn end_sheet(&mut self) -> Result<()> {
        Ok(())
    }
}

impl SheetSink for Vec<Sheet> {
    fn begin_sheet(&mut self, name: &str) -> Result<()> {
        self.push(Sheet {
            name: name.to_string(),
            rows: Vec::new(),
        });
        Ok(())
    }

    fn row(&mut self, cells: Vec<String>) -> Result<()> {
        let Some(last) = self.last_mut() else {
            return Err(BatdocError::Document(
                "sheet row without begin_sheet".into(),
            ));
        };
        last.rows.push(cells);
        Ok(())
    }
}

impl<S: SheetSink + ?Sized> SheetSink for &mut S {
    fn begin_sheet(&mut self, name: &str) -> Result<()> {
        (**self).begin_sheet(name)
    }
    fn row(&mut self, cells: Vec<String>) -> Result<()> {
        (**self).row(cells)
    }
    fn end_sheet(&mut self) -> Result<()> {
        (**self).end_sheet()
    }
}

/// Output-payload budget. Counts name UTF-8 len on begin; on each row
/// `Σ(cell.len() + 1)`; `end_sheet` adds 0. Not wire/JS/heap size.
pub struct BudgetSheetSink<S: SheetSink> {
    inner: S,
    written: u64,
    max: u64,
}

impl<S: SheetSink> BudgetSheetSink<S> {
    pub const fn new(inner: S, max: u64) -> Self {
        Self {
            inner,
            written: 0,
            max,
        }
    }

    fn charge(&mut self, add: u64) -> Result<()> {
        if self.written.saturating_add(add) > self.max {
            return Err(BatdocError::Document(format!(
                "output exceeded {} bytes",
                self.max
            )));
        }
        self.written += add;
        Ok(())
    }
}

impl<S: SheetSink> SheetSink for BudgetSheetSink<S> {
    fn begin_sheet(&mut self, name: &str) -> Result<()> {
        self.charge(name.len() as u64)?;
        self.inner.begin_sheet(name)
    }

    fn row(&mut self, cells: Vec<String>) -> Result<()> {
        let add = cells.iter().map(|c| c.len() as u64 + 1).sum();
        self.charge(add)?;
        self.inner.row(cells)
    }

    fn end_sheet(&mut self) -> Result<()> {
        self.inner.end_sheet()
    }
}

/// Apply plain-mode trailing trim semantics to a densified row.
#[allow(dead_code)] // consumed by xls/xlsx extract_sheets_to in tasks 2/3
pub(crate) fn finalize_sheet_row(mut dense: Vec<String>) -> Option<Vec<String>> {
    while dense.last().is_some_and(|c| c.trim().is_empty()) {
        dense.pop();
    }
    if dense.iter().any(|c| !c.trim().is_empty()) {
        Some(dense)
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finalize_keeps_interior_gap_trims_trailing_empty_and_ws() {
        let row = finalize_sheet_row(vec![
            "a".into(),
            String::new(),
            "b".into(),
            String::new(),
            "  ".into(),
        ])
        .unwrap();
        assert_eq!(row, vec!["a", "", "b"]);
    }

    #[test]
    fn finalize_drops_whitespace_only_row() {
        assert!(finalize_sheet_row(vec!["  ".into(), "\t".into()]).is_none());
    }

    #[test]
    fn finalize_keeps_leading_ws_when_later_content() {
        let row = finalize_sheet_row(vec!["  ".into(), "x".into()]).unwrap();
        assert_eq!(row, vec!["  ", "x"]);
    }

    #[allow(clippy::assert_is_empty)] // intentionally brief-verbatim assertion
    #[test]
    fn vec_sink_collects_multi_sheet() {
        let mut sheets = Vec::<Sheet>::new();
        sheets.begin_sheet("A").unwrap();
        sheets.row(vec!["1".into()]).unwrap();
        sheets.end_sheet().unwrap();
        sheets.begin_sheet("B").unwrap();
        sheets.end_sheet().unwrap(); // empty sheet still present
        sheets.begin_sheet("C").unwrap();
        sheets.row(vec!["x".into(), "y".into()]).unwrap();
        sheets.end_sheet().unwrap();
        assert_eq!(sheets.len(), 3);
        assert_eq!(sheets[0].name, "A");
        assert_eq!(sheets[0].rows, vec![vec!["1".to_string()]]);
        assert_eq!(sheets[1].name, "B");
        assert!(sheets[1].rows.is_empty());
        assert_eq!(sheets[2].rows[0], vec!["x", "y"]);
    }

    #[test]
    fn vec_sink_row_without_begin_errors() {
        let mut sheets = Vec::<Sheet>::new();
        let err = sheets.row(vec!["x".into()]).unwrap_err().to_string();
        assert_eq!(err, "sheet row without begin_sheet");
    }

    #[test]
    fn budget_begin_counts_name_and_blocks_before_inner() {
        let mut inner = Vec::<Sheet>::new();
        let mut sink = BudgetSheetSink::new(&mut inner, 3);
        // "abcd".len() == 4 > 3
        let err = sink.begin_sheet("abcd").unwrap_err().to_string();
        assert_eq!(err, "output exceeded 3 bytes");
        assert!(inner.is_empty());
    }

    #[allow(clippy::drop_non_drop)] // drop() ends the &mut borrow so inner can be read
    #[test]
    fn budget_row_formula_and_no_delivery_on_reject() {
        let mut inner = Vec::<Sheet>::new();
        let mut sink = BudgetSheetSink::new(&mut inner, 10);
        sink.begin_sheet("S").unwrap(); // +1 → written=1
                                        // cells "hi"(2+1) + "there"(5+1) = 9 → total 10 ok
        sink.row(vec!["hi".into(), "there".into()]).unwrap();
        // next row "x"(1+1)=2 → would be 12 > 10
        let err = sink.row(vec!["x".into()]).unwrap_err().to_string();
        assert_eq!(err, "output exceeded 10 bytes");
        drop(sink);
        assert_eq!(inner[0].rows.len(), 1); // exactly one row delivered; rejected row not pushed
    }

    #[test]
    fn budget_end_sheet_adds_zero() {
        let mut inner = Vec::<Sheet>::new();
        let mut sink = BudgetSheetSink::new(&mut inner, 1);
        sink.begin_sheet("S").unwrap(); // uses full budget
        sink.end_sheet().unwrap(); // must succeed
    }

    #[test]
    fn mut_ref_delegates() {
        let mut sheets = Vec::<Sheet>::new();
        {
            let r = &mut sheets;
            r.begin_sheet("Z").unwrap();
            r.row(vec!["q".into()]).unwrap();
            r.end_sheet().unwrap();
        }
        assert_eq!(sheets[0].rows[0][0], "q");
    }
}
