//! Shared spreadsheet rendering used by both `.xlsx` and `.xls` parsers.
//!
//! Provides the `Sheet` struct (a named 2D grid of cell values) and renderers
//! that produce either tab-separated plain text or markdown tables.

/// A parsed worksheet: a name and a 2D grid of cell values.
#[derive(Debug)]
pub(crate) struct Sheet {
    pub(crate) name: String,
    pub(crate) rows: Vec<Vec<String>>,
}

// ── Plain text rendering ──────────────────────────────────────────

pub(crate) fn render_plain(sheets: &[Sheet]) -> String {
    let mut out = String::new();
    let multiple = sheets.len() > 1;

    for (i, sheet) in sheets.iter().enumerate() {
        if skip_empty_sheet(sheet) {
            continue;
        }

        if multiple {
            if i > 0 {
                out.push('\n');
            }
            out.push_str("--- ");
            out.push_str(&sheet.name);
            out.push_str(" ---\n");
        }

        for row in &sheet.rows {
            let line: String = row.join("\t");
            let line = line.trim_end();
            if !line.is_empty() {
                out.push_str(line);
                out.push('\n');
            }
        }
    }

    out
}

// ── Markdown rendering ────────────────────────────────────────────

pub(crate) fn render_markdown(sheets: &[Sheet]) -> String {
    let mut out = String::new();
    let multiple = sheets.len() > 1;

    for sheet in sheets {
        if skip_empty_sheet(sheet) {
            continue;
        }

        // Strip trailing empty rows
        let rows = strip_trailing_empty_rows(&sheet.rows);
        if rows.is_empty() {
            continue;
        }

        // Strip leading empty columns and trailing empty columns
        let (rows, ncols) = strip_empty_cols(&rows);
        if ncols == 0 {
            continue;
        }

        if multiple {
            out.push_str("## ");
            out.push_str(&sheet.name);
            out.push_str("\n\n");
        }

        // First row is the header
        if let Some(header) = rows.first() {
            out.push_str("| ");
            out.push_str(
                &header
                    .iter()
                    .map(|c| escape_pipe(c))
                    .collect::<Vec<_>>()
                    .join(" | "),
            );
            out.push_str(" |\n");

            // Separator
            out.push('|');
            for _ in 0..ncols {
                out.push_str(" --- |");
            }
            out.push('\n');

            // Data rows
            for row in rows.iter().skip(1) {
                out.push_str("| ");
                out.push_str(
                    &row.iter()
                        .map(|c| escape_pipe(c))
                        .collect::<Vec<_>>()
                        .join(" | "),
                );
                out.push_str(" |\n");
            }
            out.push('\n');
        }
    }

    out
}

/// Returns true if the sheet has no non-empty cells.
pub(crate) fn skip_empty_sheet(sheet: &Sheet) -> bool {
    sheet
        .rows
        .iter()
        .all(|row| row.iter().all(|cell| cell.trim().is_empty()))
}

/// Strip trailing rows that are entirely empty.
fn strip_trailing_empty_rows(rows: &[Vec<String>]) -> Vec<Vec<String>> {
    let last_nonempty = rows
        .iter()
        .rposition(|row| row.iter().any(|cell| !cell.trim().is_empty()));
    last_nonempty.map_or_else(Vec::new, |idx| rows[..=idx].to_vec())
}

/// Strip leading and trailing columns that are entirely empty.
/// Returns the trimmed rows and the new column count.
fn strip_empty_cols(rows: &[Vec<String>]) -> (Vec<Vec<String>>, usize) {
    if rows.is_empty() {
        return (Vec::new(), 0);
    }

    let ncols = rows.iter().map(Vec::len).max().unwrap_or(0);
    if ncols == 0 {
        return (Vec::new(), 0);
    }

    // Find first non-empty column
    let first_col = (0..ncols)
        .find(|&c| {
            rows.iter()
                .any(|r| r.get(c).is_some_and(|v| !v.trim().is_empty()))
        })
        .unwrap_or(0);

    // Find last non-empty column
    let last_col = (0..ncols)
        .rfind(|&c| {
            rows.iter()
                .any(|r| r.get(c).is_some_and(|v| !v.trim().is_empty()))
        })
        .unwrap_or(0);

    if first_col > last_col {
        return (Vec::new(), 0);
    }

    let trimmed: Vec<Vec<String>> = rows
        .iter()
        .map(|row| {
            (first_col..=last_col)
                .map(|c| row.get(c).cloned().unwrap_or_default())
                .collect()
        })
        .collect();

    let new_ncols = last_col - first_col + 1;
    (trimmed, new_ncols)
}

/// Escape pipe characters for markdown table cells.
fn escape_pipe(s: &str) -> String {
    s.replace('|', "\\|")
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── render_markdown ──────────────────────────────────────────

    #[test]
    fn render_single_sheet_markdown() {
        let sheets = vec![Sheet {
            name: "Sheet1".into(),
            rows: vec![
                vec!["Name".into(), "Age".into()],
                vec!["Alice".into(), "30".into()],
                vec!["Bob".into(), "25".into()],
            ],
        }];

        let md = render_markdown(&sheets);
        assert!(md.contains("| Name | Age |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| Alice | 30 |"));
        assert!(md.contains("| Bob | 25 |"));
        // Single sheet — no heading
        assert!(!md.contains("## Sheet1"));
    }

    #[test]
    fn render_multi_sheet_markdown() {
        let sheets = vec![
            Sheet {
                name: "People".into(),
                rows: vec![vec!["Name".into()], vec!["Alice".into()]],
            },
            Sheet {
                name: "Places".into(),
                rows: vec![vec!["City".into()], vec!["NYC".into()]],
            },
        ];

        let md = render_markdown(&sheets);
        assert!(md.contains("## People"));
        assert!(md.contains("## Places"));
        assert!(md.contains("| Name |"));
        assert!(md.contains("| City |"));
    }

    #[test]
    fn render_empty_sheet_skipped() {
        let sheets = vec![
            Sheet {
                name: "Empty".into(),
                rows: vec![vec![String::new(), String::new()]],
            },
            Sheet {
                name: "Data".into(),
                rows: vec![vec!["Hello".into()]],
            },
        ];

        let md = render_markdown(&sheets);
        assert!(!md.contains("Empty"));
        assert!(md.contains("| Hello |"));
    }

    #[test]
    fn render_pipe_escaped() {
        let sheets = vec![Sheet {
            name: "Sheet1".into(),
            rows: vec![vec!["A|B".into()], vec!["C".into()]],
        }];

        let md = render_markdown(&sheets);
        assert!(md.contains("A\\|B"));
    }

    // ── render_plain ─────────────────────────────────────────────

    #[test]
    fn render_plain_tsv() {
        let sheets = vec![Sheet {
            name: "Sheet1".into(),
            rows: vec![
                vec!["Name".into(), "Age".into()],
                vec!["Alice".into(), "30".into()],
            ],
        }];

        let text = render_plain(&sheets);
        assert!(text.contains("Name\tAge"));
        assert!(text.contains("Alice\t30"));
    }

    #[test]
    fn render_plain_multi_sheet() {
        let sheets = vec![
            Sheet {
                name: "People".into(),
                rows: vec![vec!["Alice".into()]],
            },
            Sheet {
                name: "Places".into(),
                rows: vec![vec!["NYC".into()]],
            },
        ];

        let text = render_plain(&sheets);
        assert!(text.contains("--- People ---"));
        assert!(text.contains("--- Places ---"));
    }

    // ── skip_empty_sheet ─────────────────────────────────────────

    #[test]
    fn empty_sheet_detected() {
        let sheet = Sheet {
            name: "Empty".into(),
            rows: vec![
                vec![String::new(), "  ".into()],
                vec![String::new(), String::new()],
            ],
        };
        assert!(skip_empty_sheet(&sheet));
    }

    #[test]
    fn nonempty_sheet_not_skipped() {
        let sheet = Sheet {
            name: "Data".into(),
            rows: vec![vec![String::new(), "Hello".into()]],
        };
        assert!(!skip_empty_sheet(&sheet));
    }

    // ── strip_trailing_empty_rows ────────────────────────────────

    #[test]
    fn strips_trailing_empty() {
        let rows = vec![vec!["A".into()], vec![String::new()], vec![String::new()]];
        let result = strip_trailing_empty_rows(&rows);
        assert_eq!(result.len(), 1);
        assert_eq!(result[0], vec!["A"]);
    }

    #[test]
    fn keeps_all_nonempty() {
        let rows = vec![vec!["A".into()], vec!["B".into()]];
        let result = strip_trailing_empty_rows(&rows);
        assert_eq!(result.len(), 2);
    }

    // ── strip_empty_cols ─────────────────────────────────────────

    #[test]
    fn strips_leading_trailing_empty_cols() {
        let rows = vec![
            vec![String::new(), "A".into(), "B".into(), String::new()],
            vec![String::new(), "C".into(), "D".into(), String::new()],
        ];
        let (result, ncols) = strip_empty_cols(&rows);
        assert_eq!(ncols, 2);
        assert_eq!(result[0], vec!["A", "B"]);
        assert_eq!(result[1], vec!["C", "D"]);
    }

    // ── escape_pipe ──────────────────────────────────────────────

    #[test]
    fn escape_pipe_in_text() {
        assert_eq!(escape_pipe("A|B|C"), "A\\|B\\|C");
    }

    #[test]
    fn escape_pipe_no_pipes() {
        assert_eq!(escape_pipe("hello"), "hello");
    }
}
