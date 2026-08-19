//! Positioned per-character PDF text capture.
//!
//! Drives pdf-extract's public `OutputDev`/`output_doc_page` to collect one
//! [`PositionedPage`] at a time — every glyph with its page-space position,
//! font size, advance, and quantized rotation. Raw material for
//! `pdf_layout` (reading order, columns, headings) and `pdf_ocr`
//! (region-aware merge). Stock-API only (spec D1): the `catch_unwind`
//! wrapper stays as defense-in-depth for downstream consumers on
//! un-patched pdf-extract (spec D3).

use crate::error::{BatdocError, Result};
use pdf_extract::{MediaBox, OutputDev, Transform};
use std::panic::{self, AssertUnwindSafe};

/// One decoded glyph in top-down page coordinates (y = distance from the
/// top of the page, matching markdown reading intuition).
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct PositionedChar {
    /// The decoded character.
    pub ch: char,
    /// Page-point x of the glyph (left edge).
    pub x: f64,
    /// Page-point y, TOP-DOWN (distance from the page top).
    pub y: f64,
    /// Transformed (scaled) font size.
    pub font_size: f64,
    /// Horizontal advance: `width * font_size + spacing`, the same quantity
    /// pdf-extract's `PlainTextOutput` tracks for its space threshold.
    pub advance: f64,
    /// Glyph rotation quantized to 0/90/180/270 degrees. `u16`, not `u8`:
    /// 270° > `u8::MAX` and would overflow.
    pub rotation: u16,
}

/// All positioned characters of a single page. No all-pages `Vec` is ever
/// materialized (spec §4.1 / §11): one of these at a time, dropped after
/// emit.
#[derive(Debug)]
pub(crate) struct PositionedPage {
    /// 1-based, matching `doc.get_pages()` keys.
    pub page_num: u32,
    /// `MediaBox` as (llx, lly, urx, ury).
    pub media_box: (f64, f64, f64, f64),
    /// Every glyph pdf-extract decoded for the page, in output order.
    pub chars: Vec<PositionedChar>,
}

/// `OutputDev` implementation for one-page positioned capture.
///
/// y-flip math: pdf-extract's `PlainTextOutput` computes the top-down
/// position as `trm.post_transform(&flip_ctm)` where
/// `flip_ctm = row_major(1, 0, 0, -1, 0, ury - lly)`. With euclid 0.20's
/// `post_transform` definition this reduces element-wise to
/// `x = trm.m31`, `y = (ury - lly) - trm.m32` — see the plan's Task 6
/// notes (verified against euclid 0.20.14 source).
struct PositionedOutputDev {
    media_box: (f64, f64, f64, f64),
    chars: Vec<PositionedChar>,
}

impl OutputDev for PositionedOutputDev {
    fn begin_page(
        &mut self,
        _page_num: u32,
        media_box: &MediaBox,
        _art_box: Option<(f64, f64, f64, f64)>,
    ) -> std::result::Result<(), pdf_extract::OutputError> {
        self.media_box = (media_box.llx, media_box.lly, media_box.urx, media_box.ury);
        Ok(())
    }

    fn end_page(&mut self) -> std::result::Result<(), pdf_extract::OutputError> {
        Ok(())
    }

    fn output_character(
        &mut self,
        trm: &Transform,
        width: f64,
        spacing: f64,
        font_size: f64,
        ch: &str,
    ) -> std::result::Result<(), pdf_extract::OutputError> {
        // y-flip post-transform reduces to (m31, height - m32) — see the
        // struct-level note / plan for the euclid algebra.
        let page_height = self.media_box.3 - self.media_box.1;
        let x = trm.m31;
        let y = page_height - trm.m32;
        // Transformed font size: sqrt(|(m11+m21)·fs · (m12+m22)·fs|), the
        // same area-preserving measure PlainTextOutput uses. `.abs()` guards
        // against a reflected text matrix yielding a negative product (the
        // upstream code produces NaN there; we degrade to a size instead).
        let sx = (trm.m11 + trm.m21) * font_size;
        let sy = (trm.m12 + trm.m22) * font_size;
        let size = (sx * sy).abs().sqrt();
        let rotation = quantize_rotation(trm);
        for c in ch.chars() {
            self.chars.push(PositionedChar {
                ch: c,
                x,
                y,
                font_size: size,
                // `mul_add` is the fused equivalent of `width * size +
                // spacing` (the quantity PlainTextOutput tracks); more
                // accurate, same value within f64.
                advance: width.mul_add(size, spacing),
                rotation,
            });
        }
        Ok(())
    }

    fn begin_word(&mut self) -> std::result::Result<(), pdf_extract::OutputError> {
        Ok(())
    }
    fn end_word(&mut self) -> std::result::Result<(), pdf_extract::OutputError> {
        Ok(())
    }
    fn end_line(&mut self) -> std::result::Result<(), pdf_extract::OutputError> {
        Ok(())
    }
}

/// Snap the text matrix's rotation to 0/90/180/270.
fn quantize_rotation(trm: &Transform) -> u16 {
    let deg = trm.m12.atan2(trm.m11).to_degrees();
    // `.rem_euclid(4.0)` bounds the snapped quarter to [0, 4) before the
    // narrowing cast, so no real value can truncate or lose sign. The
    // multiplication is in u16: quarter=3 (270°) must not overflow.
    #[allow(clippy::cast_sign_loss, clippy::cast_possible_truncation)]
    let quarter = (deg / 90.0).round().rem_euclid(4.0) as u16;
    quarter * 90
}

/// Extract one page of positioned characters. `page_num` is 1-based and
/// matches `lopdf::Document::get_pages()` (the same page-tree order
/// `extract_pages` relies on for its `page_ids[i] ↔ pages[i]` invariant).
pub(crate) fn extract_positioned_page(
    doc: &lopdf::Document,
    page_num: u32,
) -> Result<PositionedPage> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut dev = PositionedOutputDev {
            media_box: (0.0, 0.0, 0.0, 0.0),
            chars: Vec::new(),
        };
        pdf_extract::output_doc_page(doc, &mut dev, page_num)?;
        Ok::<_, pdf_extract::OutputError>(dev)
    }));
    match result {
        Ok(Ok(dev)) => Ok(PositionedPage {
            page_num,
            media_box: dev.media_box,
            chars: dev.chars,
        }),
        Ok(Err(e)) => Err(BatdocError::Document(format!("PDF extraction failed: {e}"))),
        Err(_) => Err(BatdocError::Document(
            "PDF extraction panicked (malformed document)".into(),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use lopdf::dictionary;

    fn positioned(content: &str) -> PositionedPage {
        let data = build_text_pdf_content(content);
        let doc = lopdf::Document::load_mem(&data).unwrap();
        extract_positioned_page(&doc, 1).unwrap()
    }

    #[test]
    fn captures_positions_top_down() {
        // MediaBox 612x792; Tm places baseline origin at (100, 700) PDF
        // space → top-down y = 792 - 700 = 92.
        let page = positioned("BT /F1 12 Tf 1 0 0 1 100 700 Tm (AB) Tj ET");
        assert_eq!(page.page_num, 1);
        assert_eq!(page.media_box, (0.0, 0.0, 612.0, 792.0));
        assert_eq!(page.chars.len(), 2);
        let a = page.chars[0];
        assert_eq!(a.ch, 'A');
        assert!((a.x - 100.0).abs() < 0.01, "x={}", a.x);
        assert!((a.y - 92.0).abs() < 0.01, "y={}", a.y);
        assert!((a.font_size - 12.0).abs() < 0.01);
        assert!(a.advance > 0.0);
        assert_eq!(a.rotation, 0);
        let b = page.chars[1];
        assert_eq!(b.ch, 'B');
        assert!(b.x > a.x, "B right of A: {} vs {}", b.x, a.x);
        assert!((b.y - a.y).abs() < 0.01);
    }

    #[test]
    fn captures_rotation() {
        // Tm with a 90° rotation: 0 1 -1 0 x y.
        let page = positioned("BT /F1 12 Tf 0 1 -1 0 100 700 Tm (R) Tj ET");
        assert_eq!(page.chars[0].rotation, 90);
    }

    #[test]
    fn captures_rotation_270() {
        // Tm with a 270° rotation: 0 -1 1 0 x y. atan2(m12, m11) =
        // atan2(-1, 0) = -90° → quarter 3 → 270°. 270 does not fit a u8
        // (overflow panics in debug), so the field must be u16.
        let page = positioned("BT /F1 12 Tf 0 -1 1 0 100 700 Tm (R) Tj ET");
        assert_eq!(page.chars[0].rotation, 270);
    }

    #[test]
    fn missing_page_is_an_error_not_a_panic() {
        let data = build_text_pdf_content("BT ET");
        let doc = lopdf::Document::load_mem(&data).unwrap();
        assert!(extract_positioned_page(&doc, 99).is_err());
    }

    /// Build a one-page PDF with a WinAnsi Helvetica font and a
    /// caller-supplied content stream (duplicates the helper in pdf.rs's
    /// tests deliberately — do not refactor across files).
    fn build_text_pdf_content(content: &str) -> Vec<u8> {
        let mut doc = lopdf::Document::with_version("1.4");
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
        let content_id = doc.add_object(lopdf::Stream::new(
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
            lopdf::Object::Dictionary(dictionary! {
                "Type" => "Pages",
                "Kids" => vec![lopdf::Object::Reference(page_id)],
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
}
