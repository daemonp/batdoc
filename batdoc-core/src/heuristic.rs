//! Heuristic plain-text → markdown inference engine.
//!
//! Since `.doc` binary format doesn't expose style information through
//! the text stream (headings, bold, tables are in separate binary
//! structures we don't parse), this module infers document structure
//! from textual patterns:
//!
//! - Numbered lines like `"1. Foo"` or `"2.3 Bar"` → markdown headings
//! - `"Appendix N:"` / `"Scenario N:"` → `## headings`
//! - Short standalone lines (< 80 chars, no sentence punctuation) → `**bold**`
//! - Tab-separated lines with consistent columns → markdown tables

/// Convert plain text into markdown using heuristics.
pub(crate) fn plain_to_markdown(text: &str) -> String {
    let lines: Vec<&str> = text.lines().collect();
    let mut out = String::new();
    let mut i = 0;

    while i < lines.len() {
        let line = lines[i].trim();

        if line.is_empty() {
            i += 1;
            continue;
        }

        // Detect tab-separated tabular data
        if line.contains('\t') {
            let table_start = i;
            while i < lines.len() && lines[i].trim().contains('\t') {
                i += 1;
            }
            let run = &lines[table_start..i];
            // Multi-line runs always get table treatment.
            // Single lines: only if they have enough cells to be a real table
            // (>6 cells = likely a mega-line .doc table), otherwise plain text.
            let tab_count = run[0].matches('\t').count();
            if run.len() >= 2 || tab_count >= 5 {
                render_table(run, &mut out);
            } else {
                // Few tabs on a single line: TOC entry, key-value, etc.
                let text = line.replace('\t', " ");
                let text = text.trim();
                if !text.is_empty() {
                    out.push_str(text);
                    out.push_str("\n\n");
                }
            }
            continue;
        }

        // Detect numbered headings: "1. Foo", "2.3 Bar", "Appendix 1: Foo"
        if let Some(heading) = detect_numbered_heading(line) {
            out.push_str(&heading);
            out.push_str("\n\n");
            i += 1;
            continue;
        }

        // Detect short standalone lines as subheadings (bold)
        // Must be: short, not ending in sentence punctuation, not a
        // single word, and surrounded by blank/different content
        if is_likely_subheading(line, i, &lines) {
            out.push_str("**");
            out.push_str(line);
            out.push_str("**\n\n");
            i += 1;
            continue;
        }

        // Regular paragraph
        out.push_str(line);
        out.push_str("\n\n");
        i += 1;
    }

    out
}

/// Try to detect a numbered heading like "1. Introduction" or "Appendix 1: Server Analysis".
/// Returns the markdown heading string if detected.
pub(crate) fn detect_numbered_heading(line: &str) -> Option<String> {
    let trimmed = line.trim();

    // Too long for a heading
    if trimmed.len() > 120 {
        return None;
    }

    // "1. Foo", "2. Foo", "10. Foo"
    // "1.1 Foo", "2.3.1 Foo"
    if let Some(rest) = try_strip_section_number(trimmed) {
        if !rest.is_empty() && !rest.ends_with('.') {
            let depth = trimmed[..trimmed.len() - rest.len()].matches('.').count();
            let level = depth.min(3); // cap at ###
            let hashes = "#".repeat(level);
            return Some(format!("{hashes} {trimmed}"));
        }
    }

    // "Appendix N: Title" or "Scenario N: Title"
    let lower = trimmed.to_lowercase();
    if (lower.starts_with("appendix") || lower.starts_with("scenario"))
        && trimmed.contains(':')
        && trimmed.len() < 100
    {
        return Some(format!("## {trimmed}"));
    }

    None
}

/// Try to strip a leading section number like "1. ", "2.3 ", "10.1.2 ".
/// Returns the remaining text after the number, or None.
pub(crate) fn try_strip_section_number(s: &str) -> Option<&str> {
    let bytes = s.as_bytes();
    let mut i = 0;

    // Must start with a digit
    if i >= bytes.len() || !bytes[i].is_ascii_digit() {
        return None;
    }

    // Consume digits and dots: "1.", "2.3.", "10.1.2."
    while i < bytes.len() && (bytes[i].is_ascii_digit() || bytes[i] == b'.') {
        i += 1;
    }

    // Must have consumed at least one dot (to distinguish from plain numbers)
    if !s[..i].contains('.') {
        return None;
    }

    // Skip whitespace
    while i < bytes.len() && bytes[i] == b' ' {
        i += 1;
    }

    Some(&s[i..])
}

/// Check if a line looks like a subheading: short, title-like, not a sentence.
pub(crate) fn is_likely_subheading(line: &str, idx: usize, lines: &[&str]) -> bool {
    let trimmed = line.trim();

    // Must be reasonably short
    if trimmed.len() > 80 || trimmed.len() < 3 {
        return false;
    }

    // Must not end with sentence punctuation
    if trimmed.ends_with('.') || trimmed.ends_with(',') || trimmed.ends_with(';') {
        return false;
    }

    // Must not be a single word
    if !trimmed.contains(' ') && !trimmed.contains('\t') {
        return false;
    }

    // Should have a blank line (or be at boundary) before and after
    let blank_before = idx == 0 || lines[idx - 1].trim().is_empty();
    let blank_after = idx + 1 >= lines.len() || lines[idx + 1].trim().is_empty();

    if !blank_before || !blank_after {
        return false;
    }

    // First letter should be uppercase
    if let Some(c) = trimmed.chars().next() {
        if !c.is_uppercase() {
            return false;
        }
    }

    true
}

/// Render a run of tab-separated lines as markdown.
///
/// Handles two cases from .doc text:
///   1. Normal multi-line tables (each line = one row)
///   2. Mega-lines where Word concatenated an entire table into one line
///      with tab separators and empty-string separators between rows.
///      We detect this by looking for a repeating column count pattern.
fn render_table(lines: &[&str], out: &mut String) {
    if lines.is_empty() {
        return;
    }

    // Collect all cells from all lines, splitting by tab
    let mut all_cells: Vec<Vec<&str>> = Vec::new();
    for line in lines {
        let cells: Vec<&str> = line.trim().split('\t').map(str::trim).collect();
        all_cells.push(cells);
    }

    // For each input line, try to reshape it into proper table rows
    for cells in &all_cells {
        let ncols = detect_column_count(cells);

        if ncols <= 1 {
            // Not really a table — just text with a stray tab (e.g. TOC: "1. Intro\t3")
            let text = cells.join(" ").trim().to_string();
            if !text.is_empty() {
                out.push_str(&text);
                out.push_str("\n\n");
            }
            continue;
        }

        // Split cells into rows of ncols each
        let rows: Vec<&[&str]> = cells.chunks(ncols).collect();
        if rows.is_empty() {
            continue;
        }

        // Skip rows that are entirely empty (row separators in the stream)
        let rows: Vec<&[&str]> = rows
            .into_iter()
            .filter(|row| row.iter().any(|c| !c.is_empty()))
            .collect();

        if rows.is_empty() {
            continue;
        }

        // Emit markdown table
        for (ri, row) in rows.iter().enumerate() {
            out.push('|');
            for ci in 0..ncols {
                let cell = row.get(ci).copied().unwrap_or("");
                out.push(' ');
                out.push_str(&cell.replace('|', "\\|"));
                out.push_str(" |");
            }
            out.push('\n');

            if ri == 0 {
                out.push('|');
                for _ in 0..ncols {
                    out.push_str(" --- |");
                }
                out.push('\n');
            }
        }
        out.push('\n');
    }
}

/// Detect the column count for a flat array of cells from a .doc table.
///
/// Word .doc tables often come as one long tab-separated line where each
/// "row" of N cells is followed by an empty cell separator. We look for
/// the repeating pattern.
///
/// Strategy: if the total cell count is small (≤ 6), use it directly.
/// Otherwise, look for empty cells that act as row separators at regular
/// intervals. If we find a consistent pattern, that's our column count.
/// Fall back to a reasonable guess.
pub(crate) fn detect_column_count(cells: &[&str]) -> usize {
    let n = cells.len();

    // 2 cells = likely a TOC entry ("Section title\t3") or key-value pair
    // Treat as non-table so it renders as plain text
    if n <= 2 {
        return 1;
    }

    // Small number of cells — treat as a single row
    if n <= 6 {
        return n;
    }

    // Look for empty-cell separators at regular intervals.
    // In .doc tables, rows are often separated by empty cells.
    // Find positions of empty cells.
    let empties: Vec<usize> = cells
        .iter()
        .enumerate()
        .filter(|(_, c)| c.is_empty())
        .map(|(i, _)| i)
        .collect();

    if empties.len() >= 2 {
        // Check if empty cells appear at regular intervals
        // The interval would be the column count (including the empty separator)
        let first_gap = empties[0] + 1; // distance from start to first empty (inclusive)

        // Verify this gap is consistent
        let consistent = empties.windows(2).all(|w| w[1] - w[0] == first_gap);

        if consistent && (2..=10).contains(&first_gap) {
            // The actual data columns are first_gap - 1 (the empty one is the separator)
            // But we include it in the chunking so it gets filtered out
            return first_gap;
        }
    }

    // Fallback: try common small column counts that divide evenly
    for ncols in [3, 4, 5, 2] {
        if n.is_multiple_of(ncols) && n / ncols >= 2 {
            return ncols;
        }
    }

    // Last resort: just dump as-is with all cells
    n.min(6)
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── try_strip_section_number ─────────────────────────────────

    #[test]
    fn strip_simple_number() {
        assert_eq!(
            try_strip_section_number("1. Introduction"),
            Some("Introduction")
        );
    }

    #[test]
    fn strip_multi_level() {
        assert_eq!(try_strip_section_number("2.3 Details"), Some("Details"));
    }

    #[test]
    fn strip_deep_number() {
        assert_eq!(
            try_strip_section_number("10.1.2 Deep Section"),
            Some("Deep Section")
        );
    }

    #[test]
    fn strip_no_dot() {
        assert_eq!(try_strip_section_number("42 The Answer"), None);
    }

    #[test]
    fn strip_no_digit() {
        assert_eq!(try_strip_section_number("Hello World"), None);
    }

    #[test]
    fn strip_empty() {
        assert_eq!(try_strip_section_number(""), None);
    }

    #[test]
    fn strip_bare_number() {
        assert_eq!(try_strip_section_number("42"), None);
    }

    #[test]
    fn strip_just_number() {
        assert_eq!(try_strip_section_number("1."), Some(""));
    }

    #[test]
    fn strip_trailing_dot_number() {
        assert_eq!(try_strip_section_number("1.2."), Some(""));
    }

    // ── detect_numbered_heading ──────────────────────────────────

    #[test]
    fn heading_simple() {
        assert_eq!(
            detect_numbered_heading("1. Introduction"),
            Some("# 1. Introduction".to_string())
        );
    }

    #[test]
    fn heading_subsection() {
        // "2.3" has 1 dot → depth=1, min(3) = 1 → "#"
        assert_eq!(
            detect_numbered_heading("2.3 Architecture"),
            Some("# 2.3 Architecture".to_string())
        );
    }

    #[test]
    fn heading_deep() {
        // "1.2.3" has 2 dots → depth=2, min(3) = 2 → "##"
        assert_eq!(
            detect_numbered_heading("1.2.3 Deep"),
            Some("## 1.2.3 Deep".to_string())
        );
    }

    #[test]
    fn heading_too_long() {
        let long = format!("1. {}", "x".repeat(120));
        assert_eq!(detect_numbered_heading(&long), None);
    }

    #[test]
    fn heading_appendix() {
        assert_eq!(
            detect_numbered_heading("Appendix 1: Server Analysis"),
            Some("## Appendix 1: Server Analysis".to_string())
        );
    }

    #[test]
    fn heading_scenario() {
        assert_eq!(
            detect_numbered_heading("Scenario 2: Failover"),
            Some("## Scenario 2: Failover".to_string())
        );
    }

    #[test]
    fn heading_not_a_heading() {
        assert_eq!(detect_numbered_heading("Hello World"), None);
    }

    #[test]
    fn heading_just_number_dot() {
        // "1." with nothing after → rest is empty → not a heading
        assert_eq!(detect_numbered_heading("1."), None);
    }

    // ── is_likely_subheading ─────────────────────────────────────

    #[test]
    fn subheading_basic() {
        let lines = vec!["", "Executive Summary", ""];
        assert!(is_likely_subheading("Executive Summary", 1, &lines));
    }

    #[test]
    fn subheading_at_start() {
        let lines = vec!["Executive Summary", ""];
        assert!(is_likely_subheading("Executive Summary", 0, &lines));
    }

    #[test]
    fn subheading_at_end() {
        let lines = vec!["", "Executive Summary"];
        assert!(is_likely_subheading("Executive Summary", 1, &lines));
    }

    #[test]
    fn not_subheading_ends_with_period() {
        let lines = vec!["", "This is a sentence.", ""];
        assert!(!is_likely_subheading("This is a sentence.", 1, &lines));
    }

    #[test]
    fn not_subheading_single_word() {
        let lines = vec!["", "Summary", ""];
        assert!(!is_likely_subheading("Summary", 1, &lines));
    }

    #[test]
    fn not_subheading_lowercase() {
        let lines = vec!["", "executive summary", ""];
        assert!(!is_likely_subheading("executive summary", 1, &lines));
    }

    #[test]
    fn not_subheading_no_blank_before() {
        let lines = vec!["Some text", "Executive Summary", ""];
        assert!(!is_likely_subheading("Executive Summary", 1, &lines));
    }

    #[test]
    fn not_subheading_no_blank_after() {
        let lines = vec!["", "Executive Summary", "Some text"];
        assert!(!is_likely_subheading("Executive Summary", 1, &lines));
    }

    #[test]
    fn not_subheading_too_long() {
        let long = format!("A {}", "word ".repeat(20));
        let lines = vec!["", &long, ""];
        assert!(!is_likely_subheading(&long, 1, &lines));
    }

    #[test]
    fn not_subheading_too_short() {
        let lines = vec!["", "AB", ""];
        assert!(!is_likely_subheading("AB", 1, &lines));
    }

    // ── detect_column_count ──────────────────────────────────────

    #[test]
    fn colcount_two_cells() {
        let cells = vec!["A", "B"];
        assert_eq!(detect_column_count(&cells), 1);
    }

    #[test]
    fn colcount_three_cells() {
        let cells = vec!["A", "B", "C"];
        assert_eq!(detect_column_count(&cells), 3);
    }

    #[test]
    fn colcount_six_cells() {
        let cells = vec!["A", "B", "C", "D", "E", "F"];
        assert_eq!(detect_column_count(&cells), 6);
    }

    #[test]
    fn colcount_regular_empties() {
        // Pattern: 3 data cells then 1 empty, repeated
        let cells = vec!["A", "B", "C", "", "D", "E", "F", ""];
        assert_eq!(detect_column_count(&cells), 4);
    }

    #[test]
    fn colcount_divisible_fallback() {
        // 9 cells, no empties → try 3 first (9 % 3 == 0, 9/3 >= 2)
        let cells = vec!["A", "B", "C", "D", "E", "F", "G", "H", "I"];
        assert_eq!(detect_column_count(&cells), 3);
    }

    #[test]
    fn colcount_single_cell() {
        let cells = vec!["A"];
        assert_eq!(detect_column_count(&cells), 1);
    }

    #[test]
    fn colcount_empty() {
        // n=0 ≤ 2, returns 1
        let cells: Vec<&str> = vec![];
        assert_eq!(detect_column_count(&cells), 1);
    }

    // ── plain_to_markdown (integration) ──────────────────────────

    #[test]
    fn markdown_heading_and_paragraph() {
        let input = "1. Introduction\n\nThis is the body.\n";
        let result = plain_to_markdown(input);
        assert!(result.contains("# 1. Introduction"));
        assert!(result.contains("This is the body."));
    }

    #[test]
    fn markdown_subheading() {
        let input = "\nExecutive Summary\n\nDetails here.\n";
        let result = plain_to_markdown(input);
        assert!(result.contains("**Executive Summary**"));
    }

    #[test]
    fn markdown_tab_table() {
        let input = "Name\tAge\tCity\nAlice\t30\tNY\n";
        let result = plain_to_markdown(input);
        assert!(result.contains('|'));
        assert!(result.contains("---"));
    }

    #[test]
    fn markdown_toc_not_table() {
        let input = "Introduction\t3\n";
        let result = plain_to_markdown(input);
        // Should render as plain text, not a table
        assert!(!result.contains('|'));
        assert!(result.contains("Introduction 3"));
    }
}
