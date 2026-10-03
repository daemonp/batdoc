//! PDF page rasterization for the OCR fallback.
//!
//! Not every textless PDF is a scan. A producer can emit glyphs as filled
//! vector paths instead of text or images — "Microsoft: Print To PDF" does
//! exactly this — leaving a page with no text layer *and* no embedded image
//! to OCR. [`PageRasterizer`] renders such a page to a bitmap with the
//! pure-Rust `hayro` rasterizer so the existing OCR engine can read it.
//!
//! `hayro` is pure Rust (no C toolchain, no GPU), which keeps the musl,
//! macOS, and wasm release targets building unchanged. It is compiled in
//! only with the `ocr` feature, so the lean `--no-default-features` builds
//! (including wasm) never pull it in.
#![cfg(feature = "ocr")]

use std::panic::{self, AssertUnwindSafe};

/// Target rasterization resolution, in dots per inch. 300 DPI is the usual
/// floor for reliable OCR of small print.
const TARGET_DPI: f32 = 300.0;
/// PDF user-space units per inch (PDF 32000 §8.3.2: 1 unit = 1/72 inch).
const PDF_UNITS_PER_INCH: f32 = 72.0;
/// Hard cap on either rendered dimension, matching
/// [`crate::ocr::MAX_OCR_IMAGE_DIM`]: an oversized page is downscaled to fit
/// rather than rendered past the OCR pixel budget.
const MAX_RASTER_DIM: u32 = 10_000;

/// A parsed PDF that can rasterize pages on demand.
///
/// Construction parses the document once; each [`Self::rasterize`] call
/// renders a single page, so a multi-page fallback does not re-parse.
pub(crate) struct PageRasterizer {
    pdf: hayro::hayro_syntax::Pdf,
}

impl PageRasterizer {
    /// Parse `data` for rasterization.
    ///
    /// Returns `None` when the document cannot be read by the renderer (or
    /// parsing panics), so callers fall through to the no-text error exactly
    /// as they did before rasterization existed.
    pub(crate) fn new(data: &[u8]) -> Option<Self> {
        let pdf = panic::catch_unwind(AssertUnwindSafe(|| {
            hayro::hayro_syntax::Pdf::new(data.to_vec()).ok()
        }))
        .ok()
        .flatten()?;
        Some(Self { pdf })
    }

    /// Render page `page_index` (0-based) to an RGB bitmap.
    ///
    /// `None` when the index is out of range, rendering panics, or the
    /// rendered dimensions are empty or exceed [`MAX_RASTER_DIM`].
    pub(crate) fn rasterize(&self, page_index: usize) -> Option<image::RgbImage> {
        let page = self.pdf.pages().get(page_index)?;
        let scale = render_scale(page.render_dimensions());
        let settings = hayro::RenderSettings {
            x_scale: scale,
            y_scale: scale,
            // `None` lets the renderer size the viewport from the page.
            width: None,
            height: None,
            // White paper: OCR wants dark glyphs on a light background, and
            // an opaque background keeps the un-premultiplied RGB meaningful.
            bg_color: hayro::vello_cpu::color::palette::css::WHITE,
        };
        let cache = hayro::RenderCache::new();
        let interpreter = hayro::hayro_interpret::InterpreterSettings::default();
        let pixmap = panic::catch_unwind(AssertUnwindSafe(|| {
            hayro::render(page, &cache, &interpreter, &settings)
        }))
        .ok()?;
        let (width, height) = (u32::from(pixmap.width()), u32::from(pixmap.height()));
        if width == 0 || height == 0 || width > MAX_RASTER_DIM || height > MAX_RASTER_DIM {
            return None;
        }
        // `hayro` renders premultiplied RGBA; on the opaque white background
        // the un-premultiplied form is the exact pixel color.
        let mut rgb = Vec::with_capacity(width as usize * height as usize * 3);
        for px in pixmap.take_unpremultiplied() {
            rgb.push(px.r);
            rgb.push(px.g);
            rgb.push(px.b);
        }
        image::RgbImage::from_raw(width, height, rgb)
    }
}

/// Scale factor that reaches [`TARGET_DPI`], reduced so the longest page
/// edge stays within [`MAX_RASTER_DIM`] pixels.
#[allow(clippy::cast_precision_loss)] // MAX_RASTER_DIM (10_000) is exact in f32
fn render_scale((width, height): (f32, f32)) -> f32 {
    let scale = TARGET_DPI / PDF_UNITS_PER_INCH;
    let longest = width.max(height);
    if longest > 0.0 {
        scale.min(MAX_RASTER_DIM as f32 / longest)
    } else {
        scale
    }
    // A degenerate page must still produce a non-zero scale so the
    // renderer does not divide by zero.
    .max(0.01)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[allow(clippy::suboptimal_flops)] // readable arithmetic beats mul_add here
    fn scale_is_300_dpi_for_letter() {
        // 612x792 pt letter at 300 DPI → 2550x3300 px, under the cap.
        let s = render_scale((612.0, 792.0));
        assert!((s - 300.0 / 72.0).abs() < 1e-6, "{s}");
        assert!((612.0 * s - 2550.0).abs() < 0.5, "{}", 612.0 * s);
    }

    #[test]
    #[allow(clippy::suboptimal_flops, clippy::cast_precision_loss)] // see render_scale
    fn scale_shrinks_for_oversized_page() {
        // A 5000x5000 pt page at 300 DPI would be ~20833 px; the cap must
        // bring the longest edge down to MAX_RASTER_DIM.
        let s = render_scale((5000.0, 5000.0));
        assert!((5000.0 * s - MAX_RASTER_DIM as f32).abs() < 1.0);
        assert!(s < 300.0 / 72.0);
    }

    #[test]
    fn scale_never_zero_for_degenerate_page() {
        assert!(render_scale((0.0, 0.0)) > 0.0);
    }

    #[test]
    fn rasterizer_rejects_garbage() {
        assert!(PageRasterizer::new(b"not a pdf").is_none());
    }

    #[test]
    fn rasterizes_vector_outline_page() {
        // The "Microsoft: Print To PDF" shape: a page whose only content is
        // filled path operators — no text layer, no embedded image. This is
        // the case the rasterize-then-OCR fallback exists for, and it must
        // yield a bitmap with the drawn mark.
        let data = outline_pdf();
        let rasterizer = PageRasterizer::new(&data).expect("outline PDF parses");
        let img = rasterizer.rasterize(0).expect("page renders");
        // Letter at 300 DPI is 2550x3300; allow a pixel of float rounding.
        assert!(img.width().abs_diff(2550) <= 2, "width {}", img.width());
        assert!(img.height().abs_diff(3300) <= 2, "height {}", img.height());
        let dark = img.pixels().filter(|p| p.0[0] < 128).count();
        let total = (img.width() * img.height()) as usize;
        assert!(dark > 0, "no dark pixels rendered");
        // A 200x100pt black bar is a small fraction of a letter page.
        assert!(dark < total / 4, "unexpectedly much ink: {dark}/{total}");
    }

    #[test]
    fn rasterize_out_of_range_page_is_none() {
        let data = outline_pdf();
        let rasterizer = PageRasterizer::new(&data).expect("outline PDF parses");
        assert!(rasterizer.rasterize(9).is_none());
    }

    /// One-page PDF drawn with path operators only (no fonts, no images):
    /// a filled 200x100pt rectangle.
    fn outline_pdf() -> Vec<u8> {
        use lopdf::{dictionary, Document, Object, Stream};
        let mut doc = Document::with_version("1.4");
        let pages_id = doc.new_object_id();
        let content = b"0 0 0 rg 100 100 m 300 100 l 300 200 l 100 200 l h f".to_vec();
        let content_id = doc.add_object(Stream::new(lopdf::Dictionary::new(), content));
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => content_id,
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
}
