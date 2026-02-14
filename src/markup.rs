//! Shared markdown inline formatting and hyperlink grouping.
//!
//! Both `docx.rs` and `pptx.rs` need to render text runs with bold/italic
//! formatting and group consecutive runs sharing the same hyperlink URL.
//! This module provides a single implementation via the [`InlineRun`] trait.

/// Trait for a text run that can be rendered as markdown inline formatting.
///
/// Implemented by `docx::Run` and `pptx::TextRun` to allow shared rendering
/// logic without coupling the two modules.
pub(crate) trait InlineRun {
    /// The run's text content.
    fn text(&self) -> &str;
    /// Whether the run is bold.
    fn bold(&self) -> bool;
    /// Whether the run is italic.
    fn italic(&self) -> bool;
    /// The resolved hyperlink URL, if any.
    fn link_url(&self) -> Option<&str>;
}

/// Render a slice of runs as markdown with inline formatting and grouped
/// hyperlinks.
///
/// Adjacent runs sharing the same `link_url` are grouped so the markdown
/// link wraps the entire visible text: `[text](url)` instead of producing
/// separate `[part1](url)[part2](url)` fragments.
pub(crate) fn render_runs_markdown<R: InlineRun>(runs: &[R]) -> String {
    let mut out = String::new();
    let mut i = 0;

    while i < runs.len() {
        let run = &runs[i];

        // Group consecutive runs that share the same hyperlink URL
        if let Some(url) = run.link_url() {
            let mut link_text = String::new();
            while i < runs.len() && runs[i].link_url() == Some(url) {
                let r = &runs[i];
                if r.text().trim().is_empty() {
                    link_text.push_str(r.text());
                } else {
                    format_run_inline(r, &mut link_text);
                }
                i += 1;
            }
            let link_text = link_text.trim();
            if !link_text.is_empty() {
                out.push('[');
                out.push_str(link_text);
                out.push_str("](");
                out.push_str(url);
                out.push(')');
            }
            continue;
        }

        if run.text().trim().is_empty() {
            out.push_str(run.text());
            i += 1;
            continue;
        }

        format_run_inline(run, &mut out);
        i += 1;
    }

    out
}

/// Apply bold/italic formatting to a single run and append to `out`.
///
/// Whitespace-only runs are never wrapped in formatting markers.
pub(crate) fn format_run_inline<R: InlineRun>(run: &R, out: &mut String) {
    if run.text().trim().is_empty() {
        out.push_str(run.text());
        return;
    }

    match (run.bold(), run.italic()) {
        (true, true) => {
            out.push_str("***");
            out.push_str(run.text());
            out.push_str("***");
        }
        (true, false) => {
            out.push_str("**");
            out.push_str(run.text());
            out.push_str("**");
        }
        (false, true) => {
            out.push('*');
            out.push_str(run.text());
            out.push('*');
        }
        (false, false) => {
            out.push_str(run.text());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Test implementation of `InlineRun`.
    struct TestRun {
        text: String,
        bold: bool,
        italic: bool,
        link_url: Option<String>,
    }

    impl InlineRun for TestRun {
        fn text(&self) -> &str {
            &self.text
        }
        fn bold(&self) -> bool {
            self.bold
        }
        fn italic(&self) -> bool {
            self.italic
        }
        fn link_url(&self) -> Option<&str> {
            self.link_url.as_deref()
        }
    }

    fn run(text: &str, bold: bool, italic: bool) -> TestRun {
        TestRun {
            text: text.into(),
            bold,
            italic,
            link_url: None,
        }
    }

    fn link_run(text: &str, bold: bool, italic: bool, url: &str) -> TestRun {
        TestRun {
            text: text.into(),
            bold,
            italic,
            link_url: Some(url.into()),
        }
    }

    // ── format_run_inline ────────────────────────────────────────

    #[test]
    fn format_plain() {
        let mut out = String::new();
        format_run_inline(&run("Hello", false, false), &mut out);
        assert_eq!(out, "Hello");
    }

    #[test]
    fn format_bold() {
        let mut out = String::new();
        format_run_inline(&run("Bold", true, false), &mut out);
        assert_eq!(out, "**Bold**");
    }

    #[test]
    fn format_italic() {
        let mut out = String::new();
        format_run_inline(&run("Italic", false, true), &mut out);
        assert_eq!(out, "*Italic*");
    }

    #[test]
    fn format_bold_italic() {
        let mut out = String::new();
        format_run_inline(&run("Both", true, true), &mut out);
        assert_eq!(out, "***Both***");
    }

    #[test]
    fn format_whitespace_not_formatted() {
        let mut out = String::new();
        format_run_inline(&run("   ", true, true), &mut out);
        assert_eq!(out, "   ");
    }

    // ── render_runs_markdown ─────────────────────────────────────

    #[test]
    fn runs_plain_text() {
        let runs = vec![run("Hello", false, false)];
        assert_eq!(render_runs_markdown(&runs), "Hello");
    }

    #[test]
    fn runs_mixed() {
        let runs = vec![
            run("Normal ", false, false),
            run("bold", true, false),
            run(" end", false, false),
        ];
        assert_eq!(render_runs_markdown(&runs), "Normal **bold** end");
    }

    #[test]
    fn runs_hyperlink_basic() {
        let runs = vec![link_run("click here", false, false, "https://example.com")];
        assert_eq!(
            render_runs_markdown(&runs),
            "[click here](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_bold() {
        let runs = vec![link_run("bold link", true, false, "https://example.com")];
        assert_eq!(
            render_runs_markdown(&runs),
            "[**bold link**](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_grouped() {
        let runs = vec![
            link_run("part ", false, false, "https://example.com"),
            link_run("one", true, false, "https://example.com"),
        ];
        assert_eq!(
            render_runs_markdown(&runs),
            "[part **one**](https://example.com)"
        );
    }

    #[test]
    fn runs_hyperlink_mixed_with_plain() {
        let runs = vec![
            run("See ", false, false),
            link_run("this link", false, false, "https://example.com"),
            run(" for details", false, false),
        ];
        assert_eq!(
            render_runs_markdown(&runs),
            "See [this link](https://example.com) for details"
        );
    }

    #[test]
    fn runs_empty() {
        let runs: Vec<TestRun> = vec![];
        assert_eq!(render_runs_markdown(&runs), "");
    }
}
