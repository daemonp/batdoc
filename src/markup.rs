//! Shared markdown inline formatting, hyperlink grouping, and image helpers.
//!
//! Both `docx.rs` and `pptx.rs` need to render text runs with bold/italic
//! formatting and group consecutive runs sharing the same hyperlink URL.
//! This module provides a single implementation via the [`InlineRun`] trait.
//!
//! The [`image_to_base64_md`] function encodes raw image bytes into a
//! self-contained markdown `![](data:...)` image tag for `--images` support.

use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};

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

// ── Image helpers ──────────────────────────────────────────────────

/// Detect the MIME type of an image from its magic bytes.
///
/// Returns `None` for unsupported formats (EMF, WMF, TIFF, etc.)
/// since they can't be rendered in markdown viewers / browsers.
pub(crate) fn detect_image_mime(data: &[u8]) -> Option<&'static str> {
    if data.len() < 4 {
        return None;
    }
    // JPEG: FF D8 FF
    if data[..3] == [0xFF, 0xD8, 0xFF] {
        return Some("image/jpeg");
    }
    // PNG: 89 50 4E 47
    if data[..4] == [0x89, 0x50, 0x4E, 0x47] {
        return Some("image/png");
    }
    // GIF: GIF87a or GIF89a
    if data.len() >= 6 && &data[..3] == b"GIF" {
        return Some("image/gif");
    }
    // WebP: RIFF....WEBP
    if data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WEBP" {
        return Some("image/webp");
    }
    // BMP: BM
    if data[..2] == [0x42, 0x4D] {
        return Some("image/bmp");
    }
    // SVG: starts with '<' (heuristic — check for <?xml or <svg)
    if data[0] == b'<' {
        let prefix = std::str::from_utf8(&data[..data.len().min(256)]).unwrap_or("");
        if prefix.contains("<svg") {
            return Some("image/svg+xml");
        }
    }
    // EMF, WMF, TIFF, etc. — not supported in browsers/markdown
    None
}

/// A reference-style markdown image: an inline tag and a definition.
///
/// The inline tag (`![][image1]`) goes in the text flow; the definition
/// (`[image1]: <data:image/png;base64,...>`) goes at the end of the document.
/// This avoids extremely long lines that break some markdown renderers.
pub(crate) struct ImageRef {
    /// The inline reference to place in the text flow, e.g. `![][image1]`.
    pub(crate) inline: String,
    /// The definition to append at the document end, e.g. `[image1]: <data:...>`.
    pub(crate) definition: String,
}

/// Encode image data as a reference-style markdown image.
///
/// Returns `None` if the image format is unsupported (e.g., EMF/WMF).
/// The `id` is used for the reference label (e.g., `"image1"`).
pub(crate) fn image_to_base64_ref(data: &[u8], id: &str) -> Option<ImageRef> {
    let mime = detect_image_mime(data)?;
    let encoded = BASE64.encode(data);
    Some(ImageRef {
        inline: format!("![][{id}]"),
        definition: format!("[{id}]: <data:{mime};base64,{encoded}>"),
    })
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

    // ── image helpers ─────────────────────────────────────────────

    #[test]
    fn detect_jpeg() {
        assert_eq!(
            detect_image_mime(&[0xFF, 0xD8, 0xFF, 0xE0]),
            Some("image/jpeg")
        );
    }

    #[test]
    fn detect_png() {
        assert_eq!(
            detect_image_mime(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]),
            Some("image/png")
        );
    }

    #[test]
    fn detect_gif() {
        assert_eq!(detect_image_mime(b"GIF89a"), Some("image/gif"));
    }

    #[test]
    fn detect_webp() {
        assert_eq!(
            detect_image_mime(b"RIFF\x00\x00\x00\x00WEBP"),
            Some("image/webp")
        );
    }

    #[test]
    fn detect_bmp() {
        assert_eq!(
            detect_image_mime(&[0x42, 0x4D, 0x00, 0x00]),
            Some("image/bmp")
        );
    }

    #[test]
    fn detect_unsupported() {
        // EMF magic bytes — should be None
        assert_eq!(detect_image_mime(&[0x01, 0x00, 0x00, 0x00]), None);
    }

    #[test]
    fn detect_too_short() {
        assert_eq!(detect_image_mime(&[0xFF, 0xD8]), None);
    }

    #[test]
    fn image_to_base64_ref_jpeg() {
        let data = &[0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10];
        let img = image_to_base64_ref(data, "image1").unwrap();
        assert_eq!(img.inline, "![][image1]");
        assert!(img
            .definition
            .starts_with("[image1]: <data:image/jpeg;base64,"));
        assert!(img.definition.ends_with('>'));
    }

    #[test]
    fn image_to_base64_ref_unsupported() {
        let data = &[0x01, 0x00, 0x00, 0x00]; // not a recognized format
        assert!(image_to_base64_ref(data, "image1").is_none());
    }
}
