//! Region-aware OCR merge for PDF embedded images.
//!
//! Moved from `pdf.rs`: image decoding and OCR candidate selection helpers.
//! New: map OCR pixel-space lines into page points, and merge them with
//! native PDF text lines while dropping overlaps.

use crate::ocr::MAX_OCR_IMAGE_DIM;
use crate::pdf_geometry::{PlacedImage, PtRect};
use crate::pdf_layout::{Line, LineSource};
use lopdf::xobject::PdfImage;
use std::io::Cursor;

/// Maximum number of embedded images OCR'd per page (largest first).
const MAX_OCR_IMAGES_PER_PAGE: usize = 4;
/// Maximum decoded pixels per embedded image (~100 MP ≈ 300 MB RGB).
/// Larger images are skipped so a single image cannot exhaust memory.
const MAX_OCR_IMAGE_PIXELS: u64 = 100_000_000;

/// Overlap tolerance in page points for OCR/native dedup (liteparse
/// `ocr_merge.rs:688` uses 2.0 pt).
#[allow(dead_code)] // consumed by merge; merge is not yet called by the driver.
const OVERLAP_TOLERANCE_PT: f64 = 2.0;

/// Select up to [`MAX_OCR_IMAGES_PER_PAGE`] OCR candidates, largest by
/// declared pixel area first, skipping images whose decoded size would
/// exceed [`MAX_OCR_IMAGE_PIXELS`]. The cap must bound peak memory, which
/// the declared dimensions determine (compressed images expand when decoded).
pub(crate) fn ocr_candidates<'a>(images: &'a [PdfImage<'a>]) -> Vec<&'a PdfImage<'a>> {
    let mut candidates: Vec<(&PdfImage, u64)> = images
        .iter()
        .map(|img| {
            (
                img,
                u64::try_from(img.width.max(0))
                    .unwrap_or(0)
                    .saturating_mul(u64::try_from(img.height.max(0)).unwrap_or(0)),
            )
        })
        .filter(|(_, area)| *area > 0 && *area <= MAX_OCR_IMAGE_PIXELS)
        .collect();
    candidates.sort_by_key(|(_, area)| std::cmp::Reverse(*area));
    candidates
        .into_iter()
        .take(MAX_OCR_IMAGES_PER_PAGE)
        .map(|(img, _)| img)
        .collect()
}

/// Decode JPEG bytes with strict dimension limits, so a stream whose real
/// dimensions exceed the PDF dictionary's claim cannot allocate beyond the
/// OCR pixel budget.
pub(crate) fn decode_jpeg_bounded(bytes: &[u8]) -> Option<image::RgbImage> {
    let mut reader = image::ImageReader::new(Cursor::new(bytes));
    reader.set_format(image::ImageFormat::Jpeg);
    let mut limits = image::Limits::default();
    limits.max_image_width = Some(MAX_OCR_IMAGE_DIM);
    limits.max_image_height = Some(MAX_OCR_IMAGE_DIM);
    reader.limits(limits);
    reader.decode().ok().map(image::DynamicImage::into_rgb8)
}

/// Decode a `lopdf::PdfImage` into an RGB image usable by the OCR engine.
///
/// Handles `DCTDecode` (JPEG bytes) and raw/`FlateDecode` 8-bit `DeviceGray`/`DeviceRGB`,
/// plus ICC-based spaces whose `ColorSpace` is an indirect reference (which lopdf
/// leaves as `None`) or `ICCBased` — component count is read from the decoded
/// buffer length.
/// Returns `None` for other encodings (`CCITT`, `JPX`, `JBIG2`, `Indexed`, `CMYK`).
pub(crate) fn decode_pdf_image(img: &PdfImage<'_>) -> Option<image::RgbImage> {
    let filters: Vec<&str> = img.filters.iter().flatten().map(String::as_str).collect();
    if filters.contains(&"DCTDecode") {
        return decode_jpeg_bounded(img.content);
    }
    if filters.iter().any(|f| *f != "FlateDecode") {
        return None; // unknown encoding
    }
    let width = u32::try_from(img.width).ok()?;
    let height = u32::try_from(img.height).ok()?;
    if width == 0 || height == 0 {
        return None;
    }
    let content = if filters.contains(&"FlateDecode") {
        let mut stream = lopdf::Stream::new(img.origin_dict.clone(), img.content.to_vec());
        stream.decompress().ok()?;
        stream.content
    } else {
        img.content.to_vec()
    };
    // Validate the exact buffer length so `ImageSource::from_bytes` never
    // receives a mis-sized buffer (it errors on non-exact lengths).
    let area = u64::from(width) * u64::from(height);
    let len = u64::try_from(content.len()).unwrap_or(0);
    let bpc = img.bits_per_component.unwrap_or(8);
    let channels: u64 = match (img.color_space.as_deref(), bpc) {
        (Some("DeviceRGB"), 8) => 3,
        (Some("DeviceGray"), 8) => 1,
        // Indirect / ICCBased color spaces surface as `None` (a `/ColorSpace N 0 R`
        // reference is never dereferenced by lopdf) or `"ICCBased"`; there is no
        // channel count on `PdfImage`, so the decoded buffer length disambiguates.
        (Some("ICCBased") | None, 8) if area > 0 && len != 0 && len % area == 0 => {
            match len / area {
                1 | 3 => len / area,
                _ => return None,
            }
        }
        _ => return None,
    };
    if len != area * channels {
        return None;
    }
    if channels == 3 {
        image::RgbImage::from_raw(width, height, content)
    } else {
        let gray = image::GrayImage::from_raw(width, height, content)?;
        Some(image::DynamicImage::ImageLuma8(gray).into_rgb8())
    }
}

/// Map OCR pixel-space lines for one embedded image into top-down page
/// points via its placement (spec §6: `page_pt_per_pixel` = placed / pixels).
#[allow(dead_code)] // consumed by the driver's OCR merge pass in a later task.
pub(crate) fn map_ocr_lines(
    placed: &PlacedImage,
    img_width_px: u32,
    img_height_px: u32,
    lines: &[crate::ocr::OcrTextLine],
) -> Vec<Line> {
    if img_width_px == 0 || img_height_px == 0 {
        return Vec::new();
    }
    let sx = (placed.rect.x1 - placed.rect.x0) / f64::from(img_width_px);
    let sy = (placed.rect.y1 - placed.rect.y0) / f64::from(img_height_px);
    lines
        .iter()
        .map(|line| {
            let x0 = f64::from(line.x).mul_add(sx, placed.rect.x0);
            let x1 = (f64::from(line.x) + f64::from(line.width)).mul_add(sx, placed.rect.x0);
            let y0 = f64::from(line.y).mul_add(sy, placed.rect.y0);
            let y1 = (f64::from(line.y) + f64::from(line.height)).mul_add(sy, placed.rect.y0);
            let rect = PtRect { x0, y0, x1, y1 };
            Line {
                text: line.text.clone(),
                words: vec![],
                rect,
                font_size: (y1 - y0) * 0.8,
                source: LineSource::Ocr,
            }
        })
        .collect()
}

/// Drop OCR lines that overlap native text (2.0 pt tolerance, spec §6) and
/// return the rest, concatenated after `native` (unordered — the driver's
/// reading-order pass sorts everything together).
#[allow(dead_code)] // consumed by the driver's OCR merge pass in a later task.
pub(crate) fn merge(native: Vec<Line>, ocr: Vec<Line>) -> Vec<Line> {
    let mut out = native;
    for ocr_line in ocr {
        let overlaps = out.iter().any(|line| {
            matches!(line.source, LineSource::Native)
                && line
                    .rect
                    .intersects_expanded(&ocr_line.rect, OVERLAP_TOLERANCE_PT)
        });
        if !overlaps {
            out.push(ocr_line);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ocr::OcrTextLine;
    use crate::pdf_layout::LineSource;

    fn ocr_line(text: &str, x: i32, y: i32, w: u32, h: u32) -> OcrTextLine {
        OcrTextLine {
            text: text.into(),
            x,
            y,
            width: w,
            height: h,
        }
    }

    fn native_line(text: &str, x0: f64, y0: f64, x1: f64, y1: f64) -> Line {
        Line {
            text: text.into(),
            words: vec![],
            rect: PtRect { x0, y0, x1, y1 },
            font_size: 12.0,
            source: LineSource::Native,
        }
    }

    #[test]
    fn maps_pixel_lines_to_page_points() {
        // 100x50 px image placed at page rect x 72..272 (200pt), y 142..192 (50pt)
        // → 2 pt/px in x, 1 pt/px in y.
        let placed = PlacedImage {
            object_id: (5, 0),
            rect: PtRect {
                x0: 72.0,
                y0: 142.0,
                x1: 272.0,
                y1: 192.0,
            },
        };
        let lines = vec![ocr_line("Hi", 10, 20, 20, 10)];
        let mapped = map_ocr_lines(&placed, 100, 50, &lines);
        assert_eq!(mapped.len(), 1);
        let r = &mapped[0].rect;
        assert!((r.x0 - 92.0).abs() < 0.01, "{r:?}"); // 72 + 10*2
        assert!((r.x1 - 132.0).abs() < 0.01, "{r:?}"); // 72 + 30*2
        assert!((r.y0 - 162.0).abs() < 0.01, "{r:?}"); // 142 + 20*1
        assert!((r.y1 - 172.0).abs() < 0.01, "{r:?}"); // 142 + 30*1
        assert_eq!(mapped[0].text, "Hi");
        assert!(matches!(mapped[0].source, LineSource::Ocr));
        assert!((mapped[0].font_size - 8.0).abs() < 0.01); // 10pt height * 0.8
    }

    #[test]
    fn merge_drops_overlapping_ocr() {
        let native = vec![native_line("Hello", 100.0, 90.0, 160.0, 104.0)];
        let ocr = vec![
            // Overlaps the native line within 2pt → dropped.
            Line {
                text: "Hello".into(),
                words: vec![],
                rect: PtRect {
                    x0: 101.0,
                    y0: 91.0,
                    x1: 159.0,
                    y1: 103.0,
                },
                font_size: 11.0,
                source: LineSource::Ocr,
            },
            // Far away → kept.
            Line {
                text: "Caption".into(),
                words: vec![],
                rect: PtRect {
                    x0: 100.0,
                    y0: 300.0,
                    x1: 200.0,
                    y1: 314.0,
                },
                font_size: 11.0,
                source: LineSource::Ocr,
            },
        ];
        let merged = merge(native, ocr);
        assert_eq!(merged.len(), 2);
        assert!(merged
            .iter()
            .any(|l| l.text == "Hello" && matches!(l.source, LineSource::Native)));
        assert!(merged.iter().any(|l| l.text == "Caption"));
        assert!(!merged
            .iter()
            .any(|l| matches!(l.source, LineSource::Ocr) && l.text == "Hello"));
    }

    #[test]
    fn merge_keeps_everything_when_no_overlap() {
        let merged = merge(
            vec![native_line("a", 0.0, 0.0, 10.0, 10.0)],
            vec![Line {
                text: "b".into(),
                words: vec![],
                rect: PtRect {
                    x0: 0.0,
                    y0: 50.0,
                    x1: 10.0,
                    y1: 60.0,
                },
                font_size: 10.0,
                source: LineSource::Ocr,
            }],
        );
        assert_eq!(merged.len(), 2);
    }

    #[test]
    fn merge_keeps_two_ocr_lines_that_only_overlap_each_other() {
        let native = vec![native_line("Native", 0.0, 0.0, 10.0, 10.0)];
        let ocr = vec![
            Line {
                text: "OcrOne".into(),
                words: vec![],
                rect: PtRect {
                    x0: 100.0,
                    y0: 100.0,
                    x1: 120.0,
                    y1: 114.0,
                },
                font_size: 11.0,
                source: LineSource::Ocr,
            },
            // Overlaps the first OCR line but not the native line → must be kept.
            Line {
                text: "OcrTwo".into(),
                words: vec![],
                rect: PtRect {
                    x0: 110.0,
                    y0: 105.0,
                    x1: 130.0,
                    y1: 119.0,
                },
                font_size: 11.0,
                source: LineSource::Ocr,
            },
        ];
        let merged = merge(native, ocr);
        assert_eq!(merged.len(), 3);
        assert!(merged
            .iter()
            .any(|l| l.text == "Native" && matches!(l.source, LineSource::Native)));
        assert!(merged
            .iter()
            .any(|l| l.text == "OcrOne" && matches!(l.source, LineSource::Ocr)));
        assert!(merged
            .iter()
            .any(|l| l.text == "OcrTwo" && matches!(l.source, LineSource::Ocr)));
    }

    #[test]
    fn ocr_candidates_caps_and_ranks() {
        let dict = lopdf::Dictionary::new();
        let mk = |width: i64, height: i64| PdfImage {
            id: (1, 0),
            width,
            height,
            color_space: Some("DeviceRGB".to_string()),
            filters: Some(Vec::new()),
            bits_per_component: Some(8),
            content: &[],
            origin_dict: &dict,
        };
        let images = vec![
            mk(100, 100),       // 10k px
            mk(20_000, 20_000), // 400 MP — over budget, excluded
            mk(1000, 1000),     // 1 MP — largest valid
            mk(500, 500),
            mk(200, 200),
            mk(50, 50),
            mk(0, 100), // zero area — excluded
        ];
        let picked = ocr_candidates(&images);
        // 5 valid images → capped at 4, largest first; oversized/zero excluded.
        assert_eq!(picked.len(), 4);
        assert_eq!((picked[0].width, picked[0].height), (1000, 1000));
        assert_eq!((picked[1].width, picked[1].height), (500, 500));
        assert_eq!((picked[2].width, picked[2].height), (200, 200));
        assert_eq!((picked[3].width, picked[3].height), (100, 100));
    }

    #[test]
    fn decode_pdf_image_raw_rgb() {
        let mut dict = lopdf::Dictionary::new();
        dict.set("Subtype", lopdf::Object::Name(b"Image".to_vec()));
        let img = PdfImage {
            id: (1, 0),
            width: 2,
            height: 2,
            color_space: Some("DeviceRGB".to_string()),
            filters: Some(Vec::new()),
            bits_per_component: Some(8),
            content: &[255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255],
            origin_dict: &dict,
        };
        let rgb = decode_pdf_image(&img).unwrap();
        assert_eq!(rgb.dimensions(), (2, 2));
        assert_eq!(rgb.get_pixel(0, 0), &image::Rgb([255, 0, 0]));
        assert_eq!(rgb.get_pixel(1, 1), &image::Rgb([255, 255, 255]));
    }

    #[test]
    fn decode_pdf_image_flate_gray() {
        let mut dict = lopdf::Dictionary::new();
        dict.set("Subtype", lopdf::Object::Name(b"Image".to_vec()));
        dict.set("Filter", lopdf::Object::Name(b"FlateDecode".to_vec()));
        // zlib-compressed `[0, 255, 128, 64]` (stored deflate block + adler-32).
        // Inlined because `Stream::compress()` will not compress a 4-byte payload.
        let compressed = [
            0x78, 0x9C, 0x01, 0x04, 0x00, 0xFB, 0xFF, 0x00, 0xFF, 0x80, 0x40, 0x04, 0x41, 0x01,
            0xC0,
        ];
        let mut origin = lopdf::Dictionary::new();
        origin.set("Width", 2);
        origin.set("Height", 2);
        origin.set("ColorSpace", lopdf::Object::Name(b"DeviceGray".to_vec()));
        origin.set("BitsPerComponent", 8);
        origin.set("Filter", lopdf::Object::Name(b"FlateDecode".to_vec()));
        let img = PdfImage {
            id: (1, 0),
            width: 2,
            height: 2,
            color_space: Some("DeviceGray".to_string()),
            filters: Some(vec!["FlateDecode".to_string()]),
            bits_per_component: Some(8),
            content: &compressed,
            origin_dict: &origin,
        };
        let rgb = decode_pdf_image(&img).unwrap();
        assert_eq!(rgb.dimensions(), (2, 2));
        assert_eq!(rgb.get_pixel(1, 0), &image::Rgb([255, 255, 255])); // Luma 255
        assert_eq!(rgb.get_pixel(0, 1), &image::Rgb([128, 128, 128])); // Luma 128
    }

    #[test]
    fn decode_pdf_image_indirect_iccbased_gray() {
        let mut origin = lopdf::Dictionary::new();
        origin.set("Width", 2);
        origin.set("Height", 2);
        origin.set("BitsPerComponent", 8);
        origin.set("Filter", lopdf::Object::Name(b"FlateDecode".to_vec()));
        // The `/ColorSpace` here is an indirect reference to an ICCBased array,
        // so lopdf reports `color_space = None` (reference never dereferenced).
        let compressed = [
            0x78, 0x9C, 0x01, 0x04, 0x00, 0xFB, 0xFF, 0x00, 0xFF, 0x80, 0x40, 0x04, 0x41, 0x01,
            0xC0,
        ];
        let img = PdfImage {
            id: (1, 0),
            width: 2,
            height: 2,
            color_space: None,
            filters: Some(vec!["FlateDecode".to_string()]),
            bits_per_component: Some(8),
            content: &compressed,
            origin_dict: &origin,
        };
        let rgb = decode_pdf_image(&img).unwrap();
        assert_eq!(rgb.dimensions(), (2, 2));
        assert_eq!(rgb.get_pixel(1, 0), &image::Rgb([255, 255, 255])); // Luma 255
        assert_eq!(rgb.get_pixel(0, 1), &image::Rgb([128, 128, 128])); // Luma 128
    }

    #[test]
    fn decode_pdf_image_no_color_space_non_byte_multiple_skipped() {
        let mut dict = lopdf::Dictionary::new();
        dict.set("Subtype", lopdf::Object::Name(b"Image".to_vec()));
        dict.set("Filter", lopdf::Object::Name(b"FlateDecode".to_vec()));
        let img = PdfImage {
            id: (1, 0),
            width: 2,
            height: 2,
            color_space: None,
            filters: Some(vec!["FlateDecode".to_string()]),
            bits_per_component: Some(1), // packed bits: length not a multiple of area
            content: &[0xFF],
            origin_dict: &dict,
        };
        assert!(decode_pdf_image(&img).is_none());
    }

    #[test]
    fn decode_pdf_image_overlong_buffer_skipped() {
        let mut dict = lopdf::Dictionary::new();
        dict.set("Subtype", lopdf::Object::Name(b"Image".to_vec()));
        let img = PdfImage {
            id: (1, 0),
            width: 2,
            height: 2,
            color_space: Some("DeviceRGB".to_string()),
            filters: Some(Vec::new()),
            bits_per_component: Some(8),
            content: &[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], // 15 bytes ≠ 12
            origin_dict: &dict,
        };
        assert!(decode_pdf_image(&img).is_none());
    }

    #[test]
    fn unknown_pdf_image_encoding_is_skipped() {
        let mut dict = lopdf::Dictionary::new();
        dict.set("Subtype", lopdf::Object::Name(b"Image".to_vec()));
        let img = PdfImage {
            id: (1, 0),
            width: 2,
            height: 2,
            color_space: Some("DeviceCMYK".to_string()),
            filters: Some(vec!["DCTDecode".to_string()]),
            bits_per_component: Some(8),
            content: &[0xFF, 0xD8, 0xFF, 0x00], // JPEG magic but not decodable
            origin_dict: &dict,
        };
        assert!(decode_pdf_image(&img).is_none());
    }

    #[test]
    fn decode_jpeg_bounded_enforces_dim_limits() {
        let encode = |w: u32, h: u32| {
            let img = image::RgbImage::from_pixel(w, h, image::Rgb([128, 128, 128]));
            let mut buf = std::io::Cursor::new(Vec::new());
            image::DynamicImage::ImageRgb8(img)
                .write_to(&mut buf, image::ImageFormat::Jpeg)
                .unwrap();
            buf.into_inner()
        };
        // Within limits: decodes.
        let ok = encode(64, 64);
        assert_eq!(decode_jpeg_bounded(&ok).unwrap().dimensions(), (64, 64));
        // Width beyond MAX_OCR_IMAGE_DIM: rejected by the strict dim limit.
        let wide = encode(MAX_OCR_IMAGE_DIM + 1, 2);
        assert!(decode_jpeg_bounded(&wide).is_none());
    }
}
