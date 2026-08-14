//! PDF text extraction.
//!
//! Uses [`pdf_extract`] to pull text from PDF files. Since `pdf_extract` can
//! panic on malformed input (rather than returning errors), all calls are
//! wrapped in [`std::panic::catch_unwind`] to convert panics into
//! [`BatdocError::Document`] errors.

use crate::error::{BatdocError, Result};
use crate::ocr::MAX_OCR_IMAGE_DIM;
use crate::ExtractOptions;
use lopdf::xobject::PdfImage;
use std::fmt::Write as _;
use std::io::Cursor;
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

/// Maximum number of embedded images OCR'd per page (largest first).
const MAX_OCR_IMAGES_PER_PAGE: usize = 4;
/// Maximum decoded pixels per embedded image (~100 MP ≈ 300 MB RGB).
/// Larger images are skipped so a single image cannot exhaust memory.
const MAX_OCR_IMAGE_PIXELS: u64 = 100_000_000;

/// Select up to [`MAX_OCR_IMAGES_PER_PAGE`] OCR candidates, largest by
/// declared pixel area first, skipping images whose decoded size would
/// exceed [`MAX_OCR_IMAGE_PIXELS`]. The cap must bound peak memory, which
/// the declared dimensions determine (compressed images expand when decoded).
fn ocr_candidates<'a>(images: &'a [PdfImage<'a>]) -> Vec<&'a PdfImage<'a>> {
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
fn decode_jpeg_bounded(bytes: &[u8]) -> Option<image::RgbImage> {
    let mut reader = image::ImageReader::new(Cursor::new(bytes));
    reader.set_format(image::ImageFormat::Jpeg);
    let mut limits = image::Limits::default();
    limits.max_image_width = Some(MAX_OCR_IMAGE_DIM);
    limits.max_image_height = Some(MAX_OCR_IMAGE_DIM);
    reader.limits(limits);
    reader.decode().ok().map(image::DynamicImage::into_rgb8)
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
    for img in ocr_candidates(&images) {
        let Some(rgb) = decode_pdf_image(img) else {
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

/// Decode a `lopdf::PdfImage` into an RGB image usable by the OCR engine.
///
/// Handles `DCTDecode` (JPEG bytes) and raw/`FlateDecode` 8-bit `DeviceGray`/`DeviceRGB`.
/// Returns `None` for other encodings (`CCITT`, `JPX`, `JBIG2`, `Indexed`, `CMYK`).
fn decode_pdf_image(img: &PdfImage<'_>) -> Option<image::RgbImage> {
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
    let channels: u64 = match (
        img.color_space.as_deref(),
        img.bits_per_component.unwrap_or(8),
    ) {
        (Some("DeviceRGB"), 8) => 3,
        (Some("DeviceGray"), 8) => 1,
        _ => return None,
    };
    let expected = u64::from(width) * u64::from(height) * channels;
    if u64::try_from(content.len()).ok() != Some(expected) {
        return None;
    }
    if channels == 3 {
        image::RgbImage::from_raw(width, height, content)
    } else {
        let gray = image::GrayImage::from_raw(width, height, content)?;
        Some(image::DynamicImage::ImageLuma8(gray).into_rgb8())
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
    let pages = extract_pages_with_ocr(data, opts.ocr)?;
    let nonempty: Vec<&str> = pages
        .iter()
        .map(String::as_str)
        .filter(|s| !s.is_empty())
        .collect();

    if nonempty.is_empty() {
        return Err(no_text_error(opts.ocr));
    }

    Ok(nonempty.join("\n"))
}

/// Extract markdown from a PDF.
///
/// Each page gets a `## Page N` heading. Single-page documents omit the
/// heading since it would be redundant.
pub(crate) fn extract_markdown(data: &[u8], opts: ExtractOptions) -> Result<String> {
    let pages = extract_pages_with_ocr(data, opts.ocr)?;
    let nonempty: Vec<(usize, &str)> = pages
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
        return Err(no_text_error(opts.ocr));
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
    fn extract_plain_ocr_flag_errors_differ() {
        // A structurally valid PDF with no page tree (so text extraction succeeds
        // but yields zero pages); large enough for `%%EOF`/`startxref` detection.
        let data = b"%PDF-1.4\n\
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
%%EOF\n";
        let plain_err = extract_plain(data, crate::ExtractOptions::default())
            .unwrap_err()
            .to_string();
        assert!(plain_err.contains("scanned/image-only"));
        assert!(!plain_err.contains("OCR"));
        let ocr_err = extract_plain(
            data,
            crate::ExtractOptions {
                ocr: true,
                ..crate::ExtractOptions::default()
            },
        )
        .unwrap_err()
        .to_string();
        assert!(ocr_err.contains("OCR found nothing"));
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
