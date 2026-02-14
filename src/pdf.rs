//! PDF text extraction.
//!
//! Uses [`pdf_extract`] to pull text from PDF files. Since `pdf_extract` can
//! panic on malformed input (rather than returning errors), all calls are
//! wrapped in [`std::panic::catch_unwind`] to convert panics into
//! [`BatdocError::Document`] errors.

use crate::error::{BatdocError, Result};
use std::fmt::Write as _;
use std::panic::{self, AssertUnwindSafe};

/// Extract pages of text from a PDF byte slice, returning one `String` per
/// page.
///
/// Panics from the underlying library are caught and converted to errors.
fn extract_pages(data: &[u8]) -> Result<Vec<String>> {
    let data = data.to_vec(); // owned copy for the unwind boundary
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        pdf_extract::extract_text_from_mem_by_pages(&data)
    }));
    match result {
        Ok(Ok(pages)) => Ok(pages),
        Ok(Err(e)) => Err(BatdocError::Document(format!("PDF extraction failed: {e}"))),
        Err(_) => Err(BatdocError::Document(
            "PDF extraction panicked (malformed document)".into(),
        )),
    }
}

/// Clean up a page of extracted text: trim trailing whitespace from each line,
/// collapse runs of 3+ blank lines down to 2, and trim leading/trailing
/// blank lines from the whole page.
fn clean_page(raw: &str) -> String {
    let lines: Vec<&str> = raw.lines().map(str::trim_end).collect();

    let mut out = String::with_capacity(raw.len());
    let mut blank_run = 0_u32;
    for line in &lines {
        if line.is_empty() {
            blank_run += 1;
            if blank_run <= 2 {
                out.push('\n');
            }
        } else {
            blank_run = 0;
            out.push_str(line);
            out.push('\n');
        }
    }

    // Trim leading/trailing blank lines
    let trimmed = out.trim_matches('\n');
    if trimmed.is_empty() {
        String::new()
    } else {
        let mut s = trimmed.to_string();
        s.push('\n');
        s
    }
}

/// Extract plain text from a PDF.
pub(crate) fn extract_plain(data: &[u8]) -> Result<String> {
    let pages = extract_pages(data)?;
    let cleaned: Vec<String> = pages.iter().map(|p| clean_page(p)).collect();

    // Filter out completely empty pages
    let nonempty: Vec<&str> = cleaned
        .iter()
        .map(String::as_str)
        .filter(|s| !s.is_empty())
        .collect();

    if nonempty.is_empty() {
        return Err(BatdocError::Document(
            "PDF contains no extractable text (may be scanned/image-only)".into(),
        ));
    }

    Ok(nonempty.join("\n"))
}

/// Extract markdown from a PDF.
///
/// Each page gets a `## Page N` heading. Single-page documents omit the
/// heading since it would be redundant.
pub(crate) fn extract_markdown(data: &[u8]) -> Result<String> {
    let pages = extract_pages(data)?;
    let cleaned: Vec<String> = pages.iter().map(|p| clean_page(p)).collect();

    let nonempty: Vec<(usize, &str)> = cleaned
        .iter()
        .enumerate()
        .filter_map(|(i, s)| {
            if s.is_empty() {
                None
            } else {
                Some((i + 1, s.as_str()))
            }
        })
        .collect();

    if nonempty.is_empty() {
        return Err(BatdocError::Document(
            "PDF contains no extractable text (may be scanned/image-only)".into(),
        ));
    }

    let mut out = String::new();

    if nonempty.len() == 1 {
        // Single page — no heading needed
        out.push_str(nonempty[0].1);
    } else {
        for (i, (page_num, text)) in nonempty.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            let _ = write!(out, "## Page {page_num}\n\n");
            out.push_str(text);
        }
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn clean_page_trims_trailing_whitespace() {
        let input = "hello   \nworld  \n";
        let result = clean_page(input);
        assert_eq!(result, "hello\nworld\n");
    }

    #[test]
    fn clean_page_collapses_blank_lines() {
        let input = "a\n\n\n\n\nb\n";
        let result = clean_page(input);
        assert_eq!(result, "a\n\n\nb\n");
    }

    #[test]
    fn clean_page_trims_leading_trailing_blanks() {
        let input = "\n\n\nhello\n\n\n";
        let result = clean_page(input);
        assert_eq!(result, "hello\n");
    }

    #[test]
    fn clean_page_empty_input() {
        assert_eq!(clean_page(""), String::new());
        assert_eq!(clean_page("\n\n\n"), String::new());
    }

    #[test]
    fn malformed_data_returns_error() {
        let garbage = b"not a pdf at all";
        let result = extract_plain(garbage);
        assert!(result.is_err());
    }

    #[test]
    fn empty_pdf_header_returns_error() {
        // A minimal PDF header with no real content
        let data = b"%PDF-1.4\n%%EOF\n";
        let result = extract_plain(data);
        assert!(result.is_err());
    }
}
