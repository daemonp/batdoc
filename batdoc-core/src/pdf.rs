//! PDF text extraction.
//!
//! Uses [`pdf_extract`] to pull text from PDF files. Since `pdf_extract` can
//! panic on malformed input (rather than returning errors), all calls are
//! wrapped in [`std::panic::catch_unwind`] to convert panics into
//! [`BatdocError::Document`] errors.

use crate::error::{BatdocError, Result};
use crate::ExtractOptions;
use crate::ExtractSink;
use std::panic::{self, AssertUnwindSafe};

/// Extract pages of text from a PDF byte slice, returning one `String` per
/// page.
///
/// Panics from the underlying library are caught and converted to errors.
fn extract_pages(data: &[u8]) -> Result<Vec<String>> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        pdf_extract::extract_text_from_mem_by_pages(data)
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

/// Produce the per-page text, OCR'ing empty pages when `ocr` is set.
fn extract_pages_with_ocr(data: &[u8], ocr: bool) -> Result<Vec<String>> {
    let pages = extract_pages(data)?;
    if !ocr {
        return Ok(pages.iter().map(|p| clean_page(p)).collect());
    }
    // Parse the document with lopdf ONCE for image extraction; if that
    // fails, empty pages fall through as "no text" (same as non-OCR).
    let doc_pages = lopdf::Document::load_mem(data)
        .ok()
        .map(|doc| (doc.get_pages().into_values().collect::<Vec<_>>(), doc));
    let mut out = Vec::with_capacity(pages.len());
    for (i, page) in pages.iter().enumerate() {
        let cleaned = clean_page(page);
        if !cleaned.is_empty() {
            out.push(cleaned);
            continue;
        }
        let ocr_text = match &doc_pages {
            Some((page_ids, doc)) => ocr_page(doc, page_ids, i)?,
            None => None,
        };
        out.push(ocr_text.map_or_else(String::new, |t| clean_page(&t)));
    }
    Ok(out)
}

/// Extract per-page text, transparently falling back to OCR when the user did
/// not pass `--ocr` but the document has no text layer at all.
///
/// Returns the per-page text and whether OCR was actually performed (which
/// drives the wording of the no-text error if OCR also finds nothing).
fn extract_pages_with_fallback(data: &[u8], ocr: bool) -> Result<(Vec<String>, bool)> {
    let pages = extract_pages_with_ocr(data, ocr)?;
    if ocr || pages.iter().any(|p| !p.is_empty()) {
        return Ok((pages, ocr));
    }
    // Every page is textless and OCR wasn't requested: this is a scan, so
    // retry with OCR. A textless-but-empty document (no images) costs only a
    // re-parse here; `ocr_page` finds nothing and reports it at the call site.
    let ocr_pages = extract_pages_with_ocr(data, true)?;
    Ok((ocr_pages, true))
}

/// OCR the embedded images of one page (0-based index into `page_ids`).
/// `None` when the page has no OCR-able images or no text was detected.
///
/// Page-order invariant: `pdf_extract` iterates the page tree in document
/// order and lopdf's `get_pages()` numbers pages from the same tree in
/// ascending order, so `page_ids[i]` corresponds to `pages[i]` produced by
/// [`extract_pages`].
fn ocr_page(
    doc: &lopdf::Document,
    page_ids: &[lopdf::ObjectId],
    page_index: usize,
) -> Result<Option<String>> {
    let Some(page_id) = page_ids.get(page_index) else {
        return Ok(None);
    };
    let Ok(images) = doc.get_page_images(*page_id) else {
        return Ok(None);
    };
    let mut texts = Vec::new();
    for img in crate::pdf_ocr::ocr_candidates(&images) {
        let Some(rgb) = crate::pdf_ocr::decode_pdf_image(img) else {
            continue;
        };
        if let Some(text) = crate::ocr::ocr_rgb_image(&rgb)? {
            texts.push(text);
        }
    }
    if texts.is_empty() {
        Ok(None)
    } else {
        Ok(Some(texts.join("\n")))
    }
}

fn no_text_error(ocr: bool) -> BatdocError {
    let message = if ocr {
        "PDF contains no extractable text (no text layer; OCR found nothing)"
    } else {
        "PDF contains no extractable text (may be scanned/image-only)"
    };
    BatdocError::Document(message.into())
}

/// Extract plain text from a PDF.
pub(crate) fn extract_plain(data: &[u8], opts: ExtractOptions) -> Result<String> {
    let mut out = String::new();
    extract_plain_to(data, opts, &mut out)?;
    Ok(out)
}

/// Stream plain text from a PDF into `sink`, one cleaned page at a time.
///
/// Non-empty pages are joined with `\n` (no trailing newline); if no page has
/// text, [`no_text_error`] is returned with the OCR-attempt wording.
pub(crate) fn extract_plain_to(
    data: &[u8],
    opts: ExtractOptions,
    sink: &mut impl ExtractSink,
) -> Result<()> {
    let (pages, ocr_attempted) = extract_pages_with_fallback(data, opts.ocr)?;

    let mut first = true;
    let mut wrote = false;
    for page in pages {
        if page.is_empty() {
            continue;
        }
        if !first {
            sink.write_str("\n")?;
        }
        sink.write_str(&page)?;
        first = false;
        wrote = true;
    }

    if !wrote {
        return Err(no_text_error(ocr_attempted));
    }
    Ok(())
}

/// Extract markdown from a PDF.
///
/// Positioned layout pipeline: pages are classified into headings and
/// paragraphs (see [`extract_markdown_to`]). Multiple emitted pages get a
/// `## Page N` heading; a single emitted page omits it as redundant.
pub(crate) fn extract_markdown(data: &[u8], opts: ExtractOptions) -> Result<String> {
    let mut out = String::new();
    extract_markdown_to(data, opts, &mut out)?;
    Ok(out)
}

/// Result of rendering all pages with a given set of document signals.
struct RenderedPages {
    emitted: Vec<(u32, String)>,
    ocr_attempted: bool,
    had_native_lines: bool,
}

/// Render every page into a markdown buffer using the supplied `signals`.
/// Returns the emitted pages, whether OCR was attempted, and whether any
/// page had non-empty native assembled lines (the safety-net trigger).
fn render_pages(
    doc: &lopdf::Document,
    page_ids: &[lopdf::ObjectId],
    page_count: u32,
    signals: &crate::pdf_layout::DocSignals,
    opts: ExtractOptions,
) -> Result<RenderedPages> {
    let mut emitted: Vec<(u32, String)> = Vec::new();
    let mut ocr_attempted = opts.ocr;
    let mut had_native_lines = false;
    for (page_num, page_id) in (1..=page_count).zip(page_ids) {
        let page = crate::pdf_text::extract_positioned_page(doc, page_num)?;
        let garbled = page_looks_garbled(&page);
        let mut native = crate::pdf_layout::assemble(&page);
        if !native.is_empty() {
            had_native_lines = true;
        }
        if garbled {
            native.clear(); // garbage native layer: OCR replaces it
        }
        // OCR when asked, when the page assembled no non-empty lines
        // (auto-fallback, mirroring `extract_pages_with_fallback`), or when
        // the native layer is garbage.
        let need_ocr = opts.ocr || garbled || native.iter().all(|l| l.text.trim().is_empty());
        let mut ocr_lines = Vec::new();
        let mut unplaced_text = String::new();
        if need_ocr {
            ocr_attempted = true;
            let placed = crate::pdf_geometry::placed_images(doc, *page_id);
            if let Ok(images) = doc.get_page_images(*page_id) {
                for img in crate::pdf_ocr::ocr_candidates(&images) {
                    let Some(rgb) = crate::pdf_ocr::decode_pdf_image(img) else {
                        continue;
                    };
                    let lines = crate::ocr::ocr_rgb_image_lines(&rgb)?;
                    if lines.is_empty() {
                        continue;
                    }
                    match placed.iter().find(|p| p.object_id == img.id) {
                        Some(p) => ocr_lines.extend(crate::pdf_ocr::map_ocr_lines(
                            p,
                            rgb.width(),
                            rgb.height(),
                            &lines,
                        )),
                        None => {
                            // Inline (BI/EI) or otherwise unplaceable image:
                            // page-end append (today's fallback, spec §4.1).
                            for l in &lines {
                                unplaced_text.push_str(&l.text);
                                unplaced_text.push('\n');
                            }
                        }
                    }
                }
            }
        }
        let merged = crate::pdf_ocr::merge(native, ocr_lines);
        // Tables are detected on the assembly-order stream (row fragments
        // are adjacent there), then regions join the xy-cut as opaque
        // full-width bands.
        let items = crate::pdf_layout::find_tables(&merged, signals);
        let ordered = crate::pdf_layout::reading_order_items(items);
        let blocks = crate::pdf_layout::classify_items(ordered, signals);
        let mut page_md = String::new();
        crate::pdf_layout::render(&blocks, &mut page_md)?;
        if !unplaced_text.trim().is_empty() {
            if !page_md.is_empty() {
                page_md.push('\n');
            }
            page_md.push_str(unplaced_text.trim_end());
            page_md.push('\n');
        }
        if !page_md.trim().is_empty() {
            emitted.push((page.page_num, page_md));
        }
    }
    Ok(RenderedPages {
        emitted,
        ocr_attempted,
        had_native_lines,
    })
}

/// Stream markdown from a PDF into `sink` via the two-pass positioned-layout
/// driver (spec §4.2):
///
/// - **Pass 1** collects document-wide signals (modal body size, repeated
///   first/last-line signatures); each page is dropped immediately.
/// - **Pass 2** keeps one positioned page in memory at a time: assemble →
///   garbled-page detection → OCR merge (region-aware when the image has a
///   placement, page-end append otherwise) → reading order → classify →
///   render into a page-local buffer.
///
/// Pages whose rendered text is empty/whitespace are skipped. A single
/// emitted page is written bare; multiple pages each get a
/// `## Page N\n\n` heading (N = ORIGINAL 1-based page number, not the
/// filtered index) with `\n` between pages. If nothing is emitted,
/// [`no_text_error`] is returned with the OCR-attempt wording.
pub(crate) fn extract_markdown_to(
    data: &[u8],
    opts: ExtractOptions,
    sink: &mut impl ExtractSink,
) -> Result<()> {
    let mut doc = match lopdf::Document::load_mem(data) {
        Ok(d) => d,
        Err(e) => {
            return Err(BatdocError::Document(format!("PDF extraction failed: {e}")));
        }
    };
    if doc.is_encrypted() {
        // Mirror pdf-extract's `maybe_decrypt`: empty-password attempt,
        // and the same error string the plain path produces on failure.
        if let Err(e) = doc.decrypt("") {
            use lopdf::encryption::DecryptionError;
            if matches!(
                e,
                lopdf::Error::Decryption(DecryptionError::IncorrectPassword)
            ) {
                return Err(BatdocError::Document(format!(
                    "PDF extraction failed: {}",
                    pdf_extract::OutputError::PdfError(e)
                )));
            }
            // Other decrypt failures: log-and-continue, per the plan ruling.
        }
    }
    let page_ids: Vec<lopdf::ObjectId> = doc.get_pages().into_values().collect();
    // Absurd-page-count guard: > u32::MAX pages cannot be numbered anyway.
    let page_count = u32::try_from(page_ids.len()).unwrap_or(u32::MAX);

    // Pass 1: doc-level signals (tiny aggregates; pages dropped immediately).
    let mut sig_builder = crate::pdf_layout::DocSignalsBuilder::new();
    for page_num in 1..=page_count {
        let page = crate::pdf_text::extract_positioned_page(&doc, page_num)?;
        sig_builder.add_lines(&crate::pdf_layout::assemble(&page));
    }
    let signals = sig_builder.finish(page_ids.len());

    // Pass 2: one positioned page in memory at a time. Rendered per-page
    // markdown is buffered (bounded by output size, same memory profile as
    // today's Vec<String> of all page text; the single-page heading rule
    // requires not emitting page 1 before knowing whether page 2 exists).
    let rendered = render_pages(&doc, &page_ids, page_count, &signals, opts)?;
    let mut emitted = rendered.emitted;

    // Safety net: if furniture stripping dropped every native line, re-run
    // pass 2 with empty header/footer sets (body_size preserved) so real
    // documents whose only text was misclassified as furniture still emit.
    if emitted.is_empty() && rendered.had_native_lines {
        let fallback_signals = crate::pdf_layout::DocSignals {
            body_size: signals.body_size,
            headers: std::collections::HashSet::new(),
            footers: std::collections::HashSet::new(),
        };
        let fallback = render_pages(&doc, &page_ids, page_count, &fallback_signals, opts)?;
        if !fallback.emitted.is_empty() {
            emitted = fallback.emitted;
        }
    }

    if emitted.is_empty() {
        return Err(no_text_error(rendered.ocr_attempted));
    }
    if emitted.len() == 1 {
        // Single page — no heading needed
        sink.write_str(&emitted[0].1)?;
    } else {
        for (i, (page_num, page_md)) in emitted.iter().enumerate() {
            if i > 0 {
                sink.write_str("\n")?;
            }
            sink.write_str(&format!("## Page {page_num}\n\n"))?;
            sink.write_str(page_md)?;
        }
    }
    Ok(())
}

/// ≥10% replacement/PUA chars means the native text layer is garbage
/// (matches the fork's recovery-flip threshold, spec §5).
/// Driver-level garble check: a page is garbled when ≥10% of its chars are
/// U+FFFD or PUA (U+E000–U+F8FF).
///
/// Deliberately has no 20-character minimum sample floor, unlike the fork's
/// font-level prescan (`FontRecovery::observe` flips only after ≥20 decoded
/// chars): a very short page of pure garbage must still trip the check and
/// fall back to OCR. The flip side — partial-page garble goes undetected —
/// is a documented v1 limitation (plan ruling #6).
fn page_looks_garbled(page: &crate::pdf_text::PositionedPage) -> bool {
    let total = page.chars.len();
    if total == 0 {
        return false;
    }
    let bad = page
        .chars
        .iter()
        .filter(|c| c.ch == '\u{FFFD}' || ('\u{E000}'..='\u{F8FF}').contains(&c.ch))
        .count();
    bad * 10 >= total
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
        let result = extract_plain(garbage, crate::ExtractOptions::default());
        assert!(result.is_err());
    }

    #[test]
    fn empty_pdf_header_returns_error() {
        // A minimal PDF header with no real content
        let data = b"%PDF-1.4\n%%EOF\n";
        let result = extract_plain(data, crate::ExtractOptions::default());
        assert!(result.is_err());
    }

    #[test]
    fn no_text_error_wording_depends_on_ocr_flag() {
        let plain = no_text_error(false).to_string();
        assert!(plain.contains("scanned/image-only"));
        assert!(!plain.contains("OCR"));
        let ocr_on = no_text_error(true).to_string();
        assert!(ocr_on.contains("OCR found nothing"));
    }

    #[test]
    fn extract_plain_textless_pdf_auto_ocrs_and_reports() {
        // A structurally valid PDF with no page tree (so text extraction succeeds
        // but yields zero pages); large enough for `%%EOF`/`startxref` detection.
        let data = textless_pdf();
        // The PDF falls back to OCR even without --ocr, so a textless-but-empty
        // document reports the OCR-failed message in both the explicit and the
        // implicit (auto-fallback) OCR cases.
        for opts in [
            crate::ExtractOptions::default(),
            crate::ExtractOptions {
                ocr: true,
                ..crate::ExtractOptions::default()
            },
        ] {
            let err = extract_plain(data, opts).unwrap_err().to_string();
            assert!(err.contains("OCR found nothing"), "got: {err}");
        }
    }

    /// Structurally valid PDF with an empty page tree: text extraction
    /// succeeds but yields zero pages, so it exercises the textless/fallback
    /// error paths without needing OCR models.
    fn textless_pdf() -> &'static [u8] {
        b"%PDF-1.4\n\
1 0 obj\n\
<< /Type /Catalog /Pages 2 0 R >>\n\
endobj\n\
2 0 obj\n\
<< /Type /Pages /Kids [] /Count 0 >>\n\
endobj\n\
xref\n\
0 3\n\
0000000000 65535 f \n\
0000000009 00000 n \n\
0000000058 00000 n \n\
trailer\n\
<< /Size 3 /Root 1 0 R >>\n\
startxref\n\
110\n\
%%EOF\n"
    }

    #[test]
    fn extract_plain_auto_ocr_false_does_not_ocr() {
        let err = extract_plain(
            textless_pdf(),
            crate::ExtractOptions {
                images: false,
                ocr: false,
                auto_ocr: false,
                max_output_bytes: None,
            },
        )
        .unwrap_err()
        .to_string();
        assert!(err.contains("may be scanned/image-only"), "got: {err}");
        assert!(!err.contains("OCR"), "got: {err}");
    }

    #[test]
    fn extract_plain_ocr_true_auto_ocr_false_still_ocrs() {
        let err = extract_plain(
            textless_pdf(),
            crate::ExtractOptions {
                images: false,
                ocr: true,
                auto_ocr: false,
                max_output_bytes: None,
            },
        )
        .unwrap_err()
        .to_string();
        assert!(err.contains("OCR found nothing"), "got: {err}");
    }

    #[test]
    fn extract_markdown_auto_ocr_false_does_not_ocr() {
        // Use a 1-page PDF with an empty text layer (not the zero-page
        // `textless_pdf()`): page_count 0 never enters `render_pages`, which
        // would let this pass without exercising the markdown `need_ocr` branch.
        let err = extract_markdown(
            &build_text_pdf(&[""]),
            crate::ExtractOptions {
                images: false,
                ocr: false,
                auto_ocr: false,
                max_output_bytes: None,
            },
        )
        .unwrap_err()
        .to_string();
        assert!(err.contains("no extractable text"), "got: {err}");
        assert!(!err.contains("OCR found nothing"), "got: {err}");
    }

    #[test]
    fn pdf_extract_pages_does_not_require_owned_copy() {
        // Behavior lock: empty/malformed PDF error wording unchanged.
        let err = crate::extract_plain(b"%PDF-not-really", crate::Format::Pdf)
            .unwrap_err()
            .to_string();
        assert!(
            err.contains("PDF")
                || err.contains("extract")
                || err.contains("malformed")
                || err.contains("no extractable"),
            "unexpected error wording: {err}"
        );
    }

    #[test]
    fn extract_plain_to_matches_extract_plain_error() {
        for data in [&b"not a pdf at all"[..], &b"%PDF-1.4\n%%EOF\n"[..]] {
            let expected = extract_plain(data, crate::ExtractOptions::default())
                .unwrap_err()
                .to_string();
            let mut out = String::new();
            let actual = extract_plain_to(data, crate::ExtractOptions::default(), &mut out)
                .unwrap_err()
                .to_string();
            assert_eq!(actual, expected);
            assert!(out.is_empty());
        }
    }

    #[test]
    fn extract_markdown_to_matches_extract_markdown_error() {
        for data in [&b"not a pdf at all"[..], &b"%PDF-1.4\n%%EOF\n"[..]] {
            let expected = extract_markdown(data, crate::ExtractOptions::default())
                .unwrap_err()
                .to_string();
            let mut out = String::new();
            let actual = extract_markdown_to(data, crate::ExtractOptions::default(), &mut out)
                .unwrap_err()
                .to_string();
            assert_eq!(actual, expected);
            assert!(out.is_empty());
        }
    }

    #[test]
    fn render_blocks_blank_line_separated() {
        let blocks = vec![
            crate::pdf_layout::Block::Heading {
                level: 2,
                text: "Title".into(),
            },
            crate::pdf_layout::Block::Paragraph("Body text.".into()),
        ];
        let mut out = String::new();
        crate::pdf_layout::render(&blocks, &mut out).unwrap();
        assert_eq!(out, "## Title\n\nBody text.\n");
    }

    #[test]
    fn page_looks_garbled_threshold() {
        use crate::pdf_text::{PositionedChar, PositionedPage};
        let mk = |bad: usize, total: usize| {
            let chars = (0..total)
                .map(|i| PositionedChar {
                    ch: if i < bad { '\u{FFFD}' } else { 'a' },
                    x: 0.0,
                    y: 0.0,
                    font_size: 12.0,
                    advance: 6.0,
                    rotation: 0,
                })
                .collect();
            PositionedPage {
                page_num: 1,
                media_box: (0.0, 0.0, 612.0, 792.0),
                chars,
            }
        };
        assert!(page_looks_garbled(&mk(2, 10))); // 20% bad ≥ 10%
        assert!(!page_looks_garbled(&mk(1, 20))); // 5% bad
        assert!(!page_looks_garbled(&mk(0, 0))); // empty page is not garbled
    }

    #[test]
    fn markdown_single_page_omits_page_heading() {
        let data = build_text_pdf(&["Hello World"]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert_eq!(md, "Hello World\n");
        assert!(!md.contains("## Page"));
    }

    #[test]
    fn markdown_multi_page_keeps_page_headings() {
        let data = build_text_pdf(&["PageOne", "", "PageThree"]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert!(
            md.contains("## Page 1") && md.contains("## Page 3") && !md.contains("## Page 2"),
            "unexpected markdown: {md:?}"
        );
        assert!(md.contains("PageOne") && md.contains("PageThree"));
    }

    #[test]
    fn markdown_textless_pdf_reports_no_text_error() {
        // For the zero-page fixture, the markdown path emits the no-text
        // error (not the OCR wording the plain path uses after auto-OCR).
        let data = textless_pdf();
        let err = extract_markdown(data, crate::ExtractOptions::default())
            .unwrap_err()
            .to_string();
        assert!(err.contains("no extractable text"), "got: {err}");
    }

    #[test]
    fn markdown_heading_from_font_size_e2e() {
        // 24pt title (short) then 12pt body (longer, so the modal size is
        // 12 and 24/12 = 2.0 → H1) — from a real PDF, end to end.
        let data = build_text_pdf_content(
            "BT /F1 24 Tf 72 700 Td (Big) Tj ET\nBT /F1 12 Tf 72 660 Td (some body text goes here) Tj ET",
        );
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert_eq!(md, "# Big\n\nsome body text goes here\n", "got: {md:?}");
    }

    #[test]
    fn markdown_two_column_reading_order_e2e() {
        let data = build_text_pdf_content(
            "BT /F1 12 Tf 72 700 Td (L1) Tj ET\nBT /F1 12 Tf 320 700 Td (R1) Tj ET\nBT /F1 12 Tf 72 680 Td (L2) Tj ET",
        );
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        let l1 = md.find("L1").unwrap();
        let l2 = md.find("L2").unwrap();
        let r1 = md.find("R1").unwrap();
        assert!(l1 < l2 && l2 < r1, "reading order wrong: {md:?}");
    }

    #[test]
    fn narrow_gutter_columns_read_in_column_order() {
        // The sharp version of the fused-gutter characterization: two
        // rows per column at a 3em-ish gutter. Assembly splits the rows
        // (gap > 2.5x font size) and xy-cut orders column by column —
        // "L1 R1 L2 R2" interleaving would mean the gutter fused.
        let data = build_text_pdf_content(
            "BT /F1 12 Tf 72 700 Td (L1) Tj ET\n\
             BT /F1 12 Tf 180 700 Td (R1) Tj ET\n\
             BT /F1 12 Tf 72 680 Td (L2) Tj ET\n\
             BT /F1 12 Tf 180 680 Td (R2) Tj ET",
        );
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        let (l1, l2, r1, r2) = (
            md.find("L1").unwrap(),
            md.find("L2").unwrap(),
            md.find("R1").unwrap(),
            md.find("R2").unwrap(),
        );
        assert!(l1 < l2 && l2 < r1 && r1 < r2, "column order wrong: {md:?}");
    }
    #[test]
    fn furniture_stripping_safety_net_keeps_body_text() {
        // Without the safety net, the repeated first line is a header and the
        // repeated last line is a footer, so both are stripped and the output
        // is empty ("no extractable text"). The net re-runs with empty
        // header/footer sets and emits the body text.
        let content =
            "BT /F1 12 Tf 72 720 Td (Report page) Tj ET\nBT /F1 12 Tf 72 700 Td (Body text here) Tj ET";
        let data = build_text_pdf_pages(&[content, content]);
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert!(md.contains("Body text here"), "safety net failed: {md:?}");
        assert!(md.contains("Report page"), "safety net failed: {md:?}");
    }

    /// Build a minimal one-or-more page PDF with a WinAnsi Helvetica text layer,
    /// so `pdf_extract` returns non-empty per-page text.
    fn build_text_pdf(page_texts: &[&str]) -> Vec<u8> {
        use lopdf::{dictionary, Document as LopdfDoc, Object, Stream};

        let mut doc = LopdfDoc::with_version("1.4");
        let pages_id = doc.new_object_id();

        let font_id = doc.add_object(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
            "Encoding" => "WinAnsiEncoding",
        });
        let resources_id = doc.add_object(dictionary! {
            "Font" => dictionary! { "F1" => font_id },
        });

        let mut kids = Vec::new();
        for text in page_texts {
            let content = format!("BT /F1 12 Tf 72 720 Td ({text}) Tj ET");
            let content_id =
                doc.add_object(Stream::new(lopdf::Dictionary::new(), content.into_bytes()));
            let page_id = doc.add_object(dictionary! {
                "Type" => "Page",
                "Parent" => pages_id,
                "Contents" => content_id,
                "Resources" => resources_id,
                "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
            });
            kids.push(Object::Reference(page_id));
        }

        let pages = dictionary! {
            "Type" => "Pages",
            "Kids" => kids,
            "Count" => page_texts.len() as i64,
        };
        doc.objects.insert(pages_id, Object::Dictionary(pages));
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);

        let mut buf = Vec::new();
        doc.save_to(&mut buf).unwrap();
        buf
    }

    #[test]
    fn extract_plain_to_matches_extract_plain_multipage() {
        let data = build_text_pdf(&["PageOne", "PageTwo"]);
        let expected = extract_plain(&data, crate::ExtractOptions::default()).unwrap();
        assert!(
            expected.contains("PageOne"),
            "no extracted text: {expected:?}"
        );
        let mut out = String::new();
        extract_plain_to(&data, crate::ExtractOptions::default(), &mut out).unwrap();
        assert_eq!(out, expected);
    }

    #[test]
    fn extract_markdown_to_matches_extract_markdown_multipage() {
        let data = build_text_pdf(&["PageOne", "", "PageThree"]);
        let expected = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert!(
            expected.contains("## Page 1") && expected.contains("## Page 3"),
            "unexpected markdown: {expected:?}"
        );
        let mut out = String::new();
        extract_markdown_to(&data, crate::ExtractOptions::default(), &mut out).unwrap();
        assert_eq!(out, expected);
    }

    /// Golden byte-stability corpus (spec D2): plain output for these
    /// synthetic clean documents must never change, across the pdf-extract
    /// fork and all later PDF work. Garbled-document output may change;
    /// clean-document output may not.
    #[test]
    fn golden_plain_output_clean_pdfs() {
        let cases: Vec<(Vec<u8>, &str)> = vec![
            (build_text_pdf(&["PageOne"]), "PageOne\n"),
            (
                build_text_pdf(&["PageOne", "PageTwo"]),
                "PageOne\n\nPageTwo\n",
            ),
            (
                build_text_pdf(&["Hello World", "", "PageThree"]),
                "Hello World\n\nPageThree\n",
            ),
            (
                build_text_pdf_content("BT /F1 12 Tf 72 720 Td (caf\\351) Tj ET"),
                "café\n",
            ),
            (
                build_text_pdf_content("BT /F1 12 Tf 72 720 Td [(A) -120 (B)] TJ ET"),
                "A B\n",
            ),
        ];
        for (data, expected) in cases {
            let actual = crate::extract_plain(&data, crate::Format::Pdf).unwrap();
            assert_eq!(actual, expected);
        }
    }

    /// Like `build_text_pdf`, but with a caller-supplied content stream
    /// (for escapes/TJ arrays the simple builder can't express).
    fn build_text_pdf_content(content: &str) -> Vec<u8> {
        use lopdf::{dictionary, Document as LopdfDoc, Object, Stream};
        let mut doc = LopdfDoc::with_version("1.4");
        let pages_id = doc.new_object_id();
        let font_id = doc.add_object(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
            "Encoding" => "WinAnsiEncoding",
        });
        let resources_id = doc.add_object(dictionary! {
            "Font" => dictionary! { "F1" => font_id },
        });
        let content_id = doc.add_object(Stream::new(
            lopdf::Dictionary::new(),
            content.as_bytes().to_vec(),
        ));
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => content_id,
            "Resources" => resources_id,
            "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
        });
        doc.objects.insert(
            pages_id,
            Object::Dictionary(dictionary! {
                "Type" => "Pages",
                "Kids" => vec![Object::Reference(page_id)],
                "Count" => 1_i64,
            }),
        );
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);
        let mut buf = Vec::new();
        doc.save_to(&mut buf).unwrap();
        buf
    }

    /// Like `build_text_pdf_content`, but one caller-supplied content stream
    /// per page (for multi-page fixtures with custom layout per page).
    fn build_text_pdf_pages(contents: &[&str]) -> Vec<u8> {
        use lopdf::{dictionary, Document as LopdfDoc, Object, Stream};
        let mut doc = LopdfDoc::with_version("1.4");
        let pages_id = doc.new_object_id();
        let font_id = doc.add_object(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
            "Encoding" => "WinAnsiEncoding",
        });
        let resources_id = doc.add_object(dictionary! {
            "Font" => dictionary! { "F1" => font_id },
        });
        let mut kids = Vec::new();
        for content in contents {
            let content_id = doc.add_object(Stream::new(
                lopdf::Dictionary::new(),
                content.as_bytes().to_vec(),
            ));
            let page_id = doc.add_object(dictionary! {
                "Type" => "Page",
                "Parent" => pages_id,
                "Contents" => content_id,
                "Resources" => resources_id,
                "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
            });
            kids.push(Object::Reference(page_id));
        }
        doc.objects.insert(
            pages_id,
            Object::Dictionary(dictionary! {
                "Type" => "Pages",
                "Kids" => kids,
                "Count" => contents.len() as i64,
            }),
        );
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);
        let mut buf = Vec::new();
        doc.save_to(&mut buf).unwrap();
        buf
    }

    #[test]
    fn markdown_table_e2e() {
        // Three aligned rows at fixed x positions → pipe table in output.
        let data = build_text_pdf_content(
            "BT /F1 12 Tf 72 700 Td (Name) Tj ET\n\
             BT /F1 12 Tf 200 700 Td (Age) Tj ET\n\
             BT /F1 12 Tf 72 680 Td (Alice) Tj ET\n\
             BT /F1 12 Tf 200 680 Td (30) Tj ET\n\
             BT /F1 12 Tf 72 660 Td (Bob) Tj ET\n\
             BT /F1 12 Tf 200 660 Td (25) Tj ET",
        );
        let md = extract_markdown(&data, crate::ExtractOptions::default()).unwrap();
        assert!(md.contains("| Name | Age |"), "got: {md:?}");
        assert!(md.contains("| Alice | 30 |"), "got: {md:?}");
        assert!(md.contains("| --- | --- |"), "got: {md:?}");
    }

    #[test]
    #[ignore = "benchmark fixture generator — writes /tmp/batdoc-bench.pdf"]
    fn make_benchmark_pdf() {
        use lopdf::{dictionary, Document as LopdfDoc, Object, Stream};
        // 100+ words; the first/last line on each page rotates through this
        // list so no single normalized signature repeats enough times to be
        // classified as a header/footer (digits would normalize away, so we
        // use letter words).
        let words: &[&str] = &[
            "alpha",
            "bravo",
            "charlie",
            "delta",
            "echo",
            "foxtrot",
            "golf",
            "hotel",
            "india",
            "juliett",
            "kilo",
            "lima",
            "mike",
            "november",
            "oscar",
            "papa",
            "quebec",
            "romeo",
            "sierra",
            "tango",
            "uniform",
            "victor",
            "whiskey",
            "xray",
            "yankee",
            "zulu",
            "apple",
            "banana",
            "cherry",
            "date",
            "elderberry",
            "fig",
            "grape",
            "honeydew",
            "kiwi",
            "lemon",
            "mango",
            "nectarine",
            "orange",
            "papaya",
            "quince",
            "raspberry",
            "strawberry",
            "tangerine",
            "ugli",
            "vanilla",
            "watermelon",
            "xigua",
            "yam",
            "zucchini",
            "amber",
            "blue",
            "crimson",
            "emerald",
            "fuchsia",
            "gold",
            "indigo",
            "jade",
            "khaki",
            "lavender",
            "magenta",
            "navy",
            "olive",
            "purple",
            "rose",
            "silver",
            "teal",
            "umber",
            "violet",
            "wheat",
            "azure",
            "beige",
            "coral",
            "denim",
            "ebony",
            "ivory",
            "jet",
            "lime",
            "maroon",
            "mustard",
            "ochre",
            "peach",
            "plum",
            "rust",
            "sage",
            "tan",
            "turquoise",
            "wine",
            "ash",
            "birch",
            "cedar",
            "dogwood",
            "elm",
            "fir",
            "ginkgo",
            "hickory",
            "ironwood",
            "juniper",
            "koa",
            "larch",
            "maple",
            "oak",
            "pine",
            "redwood",
            "spruce",
            "teak",
            "walnut",
            "yew",
            "zebrawood",
        ];
        assert!(
            words.len() >= 100,
            "word list too small to avoid header/footer folding"
        );
        let mut doc = LopdfDoc::with_version("1.4");
        let pages_id = doc.new_object_id();
        let font_id = doc.add_object(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
            "Encoding" => "WinAnsiEncoding",
        });
        let resources_id = doc.add_object(dictionary! {
            "Font" => dictionary! { "F1" => font_id },
        });
        let mut kids = Vec::new();
        for i in 0..200 {
            let word = words[i % words.len()];
            let body = "lorem ipsum dolor sit amet consectetur adipiscing elit ".repeat(3);
            let content = format!(
                "BT /F1 18 Tf 72 720 Td (Topic {word} overview) Tj ET\n\
                 BT /F1 11 Tf 72 690 Td ({body}) Tj ET\n\
                 BT /F1 11 Tf 72 660 Td (End {word} summary) Tj ET"
            );
            let content_id =
                doc.add_object(Stream::new(lopdf::Dictionary::new(), content.into_bytes()));
            let page_id = doc.add_object(dictionary! {
                "Type" => "Page",
                "Parent" => pages_id,
                "Contents" => content_id,
                "Resources" => resources_id,
                "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
            });
            kids.push(Object::Reference(page_id));
        }
        doc.objects.insert(
            pages_id,
            Object::Dictionary(dictionary! {
                "Type" => "Pages",
                "Kids" => kids,
                "Count" => 200_i64,
            }),
        );
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);
        let mut buf = Vec::new();
        doc.save_to(&mut buf).unwrap();
        std::fs::write("/tmp/batdoc-bench.pdf", buf).unwrap();
    }
}
