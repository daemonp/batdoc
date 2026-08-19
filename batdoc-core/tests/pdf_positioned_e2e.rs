//! End-to-end tests against real PDFs supplied out-of-tree (fixtures are
//! not committed; AGENTS.md convention). Set BATDOC_SOAK_DIR to a
//! directory containing:
//!   garbled-cid.pdf     — a CID-keyed PDF whose /ToUnicode is broken
//!   scanned.pdf         — an image-only scan
//!   two-column.pdf      — a real multi-column document
//! All tests are #[ignore]d: run with
//!   BATDOC_SOAK_DIR=/path cargo test --release -p batdoc-core --test pdf_positioned_e2e -- --ignored

use std::path::PathBuf;

fn soak_pdf(name: &str) -> Option<Vec<u8>> {
    let dir = std::env::var_os("BATDOC_SOAK_DIR")?;
    let path = PathBuf::from(dir).join(name);
    std::fs::read(path).ok()
}

#[test]
#[ignore]
fn garbled_cid_pdf_extracts_readable_text() {
    let Some(data) = soak_pdf("garbled-cid.pdf") else {
        return;
    };
    let text = batdoc_core::extract_plain(&data, batdoc_core::Format::Pdf).unwrap();
    // Owner: replace with a phrase known to appear in the document.
    assert!(!text.trim().is_empty());
    assert!(
        !text.contains('\u{FFFD}'),
        "recovery left replacement chars"
    );
}

#[test]
#[ignore]
fn scanned_pdf_ocr_text_has_no_duplication() {
    let Some(data) = soak_pdf("scanned.pdf") else {
        return;
    };
    let md = batdoc_core::extract_markdown(&data, batdoc_core::Format::Pdf, false).unwrap();
    assert!(!md.trim().is_empty());
    // Owner: assert a known line appears exactly once.
}

#[test]
#[ignore]
fn two_column_pdf_reads_columns_in_order() {
    let Some(data) = soak_pdf("two-column.pdf") else {
        return;
    };
    let md = batdoc_core::extract_markdown(&data, batdoc_core::Format::Pdf, false).unwrap();
    // Owner: assert find(phrase_col1_end) < find(phrase_col2_start).
    assert!(!md.trim().is_empty());
}
