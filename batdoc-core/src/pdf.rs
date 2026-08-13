//! PDF text extraction.
//!
//! Uses [`pdf_extract`] to pull text from PDF files. Since `pdf_extract` can
//! panic on malformed input (rather than returning errors), all calls are
//! wrapped in [`std::panic::catch_unwind`] to convert panics into
//! [`BatdocError::Document`] errors.

use crate::error::{BatdocError, Result};
use lopdf::xobject::PdfImage;
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

/// Produce the per-page text, OCR'ing empty pages when `ocr` is set.
fn extract_pages_with_ocr(data: &[u8], ocr: bool) -> Result<Vec<String>> {
    let pages = extract_pages(data)?;
    if !ocr {
        return Ok(pages.iter().map(|p| clean_page(p)).collect());
    }
    let mut out = Vec::with_capacity(pages.len());
    for (i, page) in pages.iter().enumerate() {
        let cleaned = clean_page(page);
        if !cleaned.is_empty() {
            out.push(cleaned);
        } else if let Some(ocr_text) = ocr_page(data, i)? {
            out.push(clean_page(&ocr_text));
        } else {
            out.push(String::new());
        }
    }
    Ok(out)
}

/// OCR the embedded images of one page (0-based index). `None` when the page
/// has no OCR-able images or no text was detected.
fn ocr_page(data: &[u8], page_index: usize) -> Result<Option<String>> {
    let Ok(doc) = lopdf::Document::load_mem(data) else {
        return Ok(None);
    };
    let pages = doc.get_pages();
    let Some(page_id) = pages.values().nth(page_index) else {
        return Ok(None);
    };
    let Ok(images) = doc.get_page_images(*page_id) else {
        return Ok(None);
    };

    let mut decoded: Vec<image::RgbImage> = Vec::new();
    for img in images {
        if let Some(rgb) = decode_pdf_image(&img) {
            decoded.push(rgb);
        }
    }
    // OCR the largest images first, capped at 4 per page.
    decoded.sort_by_key(|img| std::cmp::Reverse(u64::from(img.width()) * u64::from(img.height())));

    let mut texts = Vec::new();
    for img in decoded.into_iter().take(4) {
        if let Some(text) = crate::ocr::ocr_rgb_image(&img)? {
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
        return image::load_from_memory(img.content)
            .ok()
            .map(image::DynamicImage::into_rgb8);
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
pub(crate) fn extract_plain(data: &[u8], ocr: bool) -> Result<String> {
    let pages = extract_pages_with_ocr(data, ocr)?;
    let nonempty: Vec<&str> = pages
        .iter()
        .map(String::as_str)
        .filter(|s| !s.is_empty())
        .collect();

    if nonempty.is_empty() {
        return Err(no_text_error(ocr));
    }

    Ok(nonempty.join("\n"))
}

/// Extract markdown from a PDF.
///
/// Each page gets a `## Page N` heading. Single-page documents omit the
/// heading since it would be redundant.
pub(crate) fn extract_markdown(data: &[u8], ocr: bool) -> Result<String> {
    let pages = extract_pages_with_ocr(data, ocr)?;
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
        return Err(no_text_error(ocr));
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
        let result = extract_plain(garbage, false);
        assert!(result.is_err());
    }

    #[test]
    fn empty_pdf_header_returns_error() {
        // A minimal PDF header with no real content
        let data = b"%PDF-1.4\n%%EOF\n";
        let result = extract_plain(data, false);
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
        let plain_err = extract_plain(data, false).unwrap_err().to_string();
        assert!(plain_err.contains("scanned/image-only"));
        assert!(!plain_err.contains("OCR"));
        let ocr_err = extract_plain(data, true).unwrap_err().to_string();
        assert!(ocr_err.contains("OCR found nothing"));
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
}
