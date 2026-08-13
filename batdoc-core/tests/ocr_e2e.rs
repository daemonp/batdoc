//! End-to-end OCR tests. All require the OCR models; they are `#[ignore]`d
//! and run by CI (and locally via `cargo test --release -p batdoc-core --test ocr_e2e -- --ignored`).
//!
//! Fixtures are generated in-memory: a bitmap-font image (pure `image` code,
//! no fonts), a DOCX containing it, and a PDF page containing it as a raw
//! RGB `XObject`.
//!
//! Fixture note: ocrs' models (fixed 800×600 detection input; every detected
//! line is resized into one 64px-tall recognition input) read the 5×7 bitmap
//! font reliably only for short isolated lines. The spec's single long line
//! "BATDOC OCR 123" deterministically misreads the 5×7 `D` as `O` at every
//! scale ("BATOOC OCR 123"), and scaling up only adds noise. Two short lines
//! — "BATDOC" and "123", separated by 12 blank rows with the second line
//! offset 16 glyph columns to the right — are read deterministically and
//! exactly, so the fixture uses that layout and the assertions below.

use std::io::Write;

use batdoc_core::{ExtractOptions, Format};

/// 5x7 bitmap glyphs, MSB of each byte is the leftmost column.
const GLYPHS: &[(u8, [u8; 7])] = &[
    (
        b'B',
        [
            0b11110, 0b10001, 0b10001, 0b11110, 0b10001, 0b10001, 0b11110,
        ],
    ),
    (
        b'A',
        [
            0b01110, 0b10001, 0b10001, 0b11111, 0b10001, 0b10001, 0b10001,
        ],
    ),
    (
        b'T',
        [
            0b11111, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100,
        ],
    ),
    (
        b'D',
        [
            0b11110, 0b10001, 0b10001, 0b10001, 0b10001, 0b10001, 0b11110,
        ],
    ),
    (
        b'O',
        [
            0b01110, 0b10001, 0b10001, 0b10001, 0b10001, 0b10001, 0b01110,
        ],
    ),
    (
        b'C',
        [
            0b01111, 0b10000, 0b10000, 0b10000, 0b10000, 0b10000, 0b01111,
        ],
    ),
    (
        b'R',
        [
            0b11110, 0b10001, 0b10001, 0b11110, 0b10100, 0b10010, 0b10001,
        ],
    ),
    (
        b'1',
        [
            0b00100, 0b01100, 0b00100, 0b00100, 0b00100, 0b00100, 0b01110,
        ],
    ),
    (
        b'2',
        [
            0b01110, 0b10001, 0b00001, 0b00010, 0b00100, 0b01000, 0b11111,
        ],
    ),
    (
        b'3',
        [
            0b11110, 0b00001, 0b00001, 0b01110, 0b00001, 0b00001, 0b11110,
        ],
    ),
    (b' ', [0, 0, 0, 0, 0, 0, 0]),
];

/// Two short lines of bitmap text (see the fixture note at the top of this
/// file for why the single-long-line layout does not OCR cleanly).
///
/// `line2` is offset to the right so the detection model returns two separate
/// line boxes; lines overlapped horizontally get merged into one recognition
/// input and read poorly.
#[allow(clippy::cast_possible_truncation)] // enum indexes are tiny; u32 pixel math is never truncated here
fn render_test_image() -> image::RgbImage {
    const SCALE: u32 = 5;
    const GLYPH_W: u32 = 5;
    const GLYPH_H: u32 = 7;
    const GAP: u32 = 2;
    const LINE_GAP_ROWS: u32 = 12;
    const PAD_ROWS: u32 = 8;
    const LINE2_X_OFFSET: u32 = 16;
    const LINES: [&[u8]; 2] = [b"BATDOC", b"123"];

    let width = ((LINE2_X_OFFSET + LINES[1].len() as u32) * (GLYPH_W + GAP) + GAP) * SCALE;
    let height = (LINES.len() as u32 * (GLYPH_H + LINE_GAP_ROWS) + PAD_ROWS) * SCALE;
    let mut img = image::RgbImage::from_pixel(width, height, image::Rgb([255, 255, 255]));
    for (line_idx, line) in LINES.iter().enumerate() {
        for (char_idx, ch) in line.iter().enumerate() {
            let glyph = GLYPHS
                .iter()
                .find(|(c, _)| c == ch)
                .map_or([0u8; 7], |(_, rows)| *rows);
            let x_off = if line_idx == 1 { LINE2_X_OFFSET } else { 0 };
            let ox = ((x_off + char_idx as u32) * (GLYPH_W + GAP) + GAP) * SCALE;
            let oy = (line_idx as u32 * (GLYPH_H + LINE_GAP_ROWS) + PAD_ROWS / 2) * SCALE;
            for (row, bits) in glyph.iter().enumerate() {
                for col in 0..GLYPH_W {
                    if bits & (1 << (4 - col)) != 0 {
                        for dy in 0..SCALE {
                            for dx in 0..SCALE {
                                img.put_pixel(
                                    ox + col * SCALE + dx,
                                    oy + row as u32 * SCALE + dy,
                                    image::Rgb([0, 0, 0]),
                                );
                            }
                        }
                    }
                }
            }
        }
    }
    img
}

/// The test image as PNG bytes.
fn test_image_png() -> Vec<u8> {
    let mut buf = Vec::new();
    image::DynamicImage::ImageRgb8(render_test_image())
        .write_to(&mut std::io::Cursor::new(&mut buf), image::ImageFormat::Png)
        .unwrap();
    buf
}

/// Build a minimal DOCX whose second paragraph contains the image.
fn build_docx_with_image(png: &[u8]) -> Vec<u8> {
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    let mut writer = ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    writer.start_file("word/document.xml", opts).unwrap();
    writer
        .write_all(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
 <w:body>
  <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
  <w:p><w:r><w:drawing>
   <wp:inline><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>
  </w:drawing></w:r></w:p>
 </w:body>
</w:document>"#,
        )
        .unwrap();

    writer
        .start_file("word/_rels/document.xml.rels", opts)
        .unwrap();
    writer
        .write_all(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#,
        )
        .unwrap();

    writer.start_file("word/media/image1.png", opts).unwrap();
    writer.write_all(png).unwrap();

    writer.finish().unwrap().into_inner()
}

/// Build a minimal one-page PDF whose page contains the image as a raw RGB `XObject`.
fn build_pdf_with_image(rgb: &image::RgbImage) -> Vec<u8> {
    use lopdf::{dictionary, Document, Object, Stream};

    let mut doc = Document::with_version("1.5");
    let pages_obj = doc.new_object_id();

    let mut img_dict = lopdf::Dictionary::new();
    img_dict.set("Type", Object::Name(b"XObject".to_vec()));
    img_dict.set("Subtype", Object::Name(b"Image".to_vec()));
    img_dict.set("Width", i64::from(rgb.width()));
    img_dict.set("Height", i64::from(rgb.height()));
    img_dict.set("ColorSpace", Object::Name(b"DeviceRGB".to_vec()));
    img_dict.set("BitsPerComponent", 8);
    let img_id = doc.add_object(Stream::new(img_dict, rgb.as_raw().clone()));

    // Empty content stream so pdf-extract returns an empty text layer.
    let content_id = doc.add_object(Stream::new(lopdf::Dictionary::new(), Vec::new()));

    let resources_id = doc.add_object(dictionary! {
        "XObject" => dictionary! { "Im0" => img_id },
    });
    let page_id = doc.add_object(dictionary! {
        "Type" => "Page",
        "Parent" => pages_obj,
        "Contents" => content_id,
        "Resources" => resources_id,
        "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
    });
    let pages = dictionary! {
        "Type" => "Pages",
        "Kids" => vec![page_id.into()],
        "Count" => 1,
    };
    doc.objects.insert(pages_obj, Object::Dictionary(pages));
    let catalog_id = doc.add_object(dictionary! {
        "Type" => "Catalog",
        "Pages" => pages_obj,
    });
    doc.trailer.set("Root", catalog_id);

    let mut buf = Vec::new();
    doc.save_to(&mut buf).unwrap();
    buf
}

#[test]
#[ignore = "requires OCR models (downloaded on first use)"]
fn ocr_direct_image_input() {
    let png = test_image_png();
    let text =
        batdoc_core::extract_plain_with(&png, Format::Image, ExtractOptions::default()).unwrap();
    assert!(text.contains("123"), "OCR text missing digits: {text:?}");
    assert!(
        text.contains("BATDOC"),
        "OCR text missing letters: {text:?}"
    );
}

#[test]
#[ignore = "requires OCR models (downloaded on first use)"]
fn ocr_docx_embedded_image_markdown() {
    let docx = build_docx_with_image(&test_image_png());
    let md = batdoc_core::extract_markdown_with(
        &docx,
        Format::Docx,
        ExtractOptions {
            images: true,
            ocr: true,
        },
    )
    .unwrap();
    assert!(md.contains("Hello"), "got: {md}");
    assert!(md.contains("![][image1]"), "got: {md}");
    assert!(md.contains("> "), "expected blockquote in: {md}");
    assert!(md.contains("123"), "OCR text missing: {md}");
    // Without --ocr the same document is today's output
    let no_ocr = batdoc_core::extract_markdown_with(
        &docx,
        Format::Docx,
        ExtractOptions {
            images: true,
            ocr: false,
        },
    )
    .unwrap();
    assert!(!no_ocr.contains("> "));
    assert!(no_ocr.contains("![][image1]"));
}

#[test]
#[ignore = "requires OCR models (downloaded on first use)"]
fn ocr_pdf_embedded_image_page() {
    let pdf = build_pdf_with_image(&render_test_image());
    // Today's behavior: no --ocr → clean error.
    let err = batdoc_core::extract_plain_with(&pdf, Format::Pdf, ExtractOptions::default())
        .unwrap_err()
        .to_string();
    assert!(err.contains("scanned/image-only"), "got: {err}");
    // With --ocr → OCR'd text.
    let text = batdoc_core::extract_plain_with(
        &pdf,
        Format::Pdf,
        ExtractOptions {
            ocr: true,
            images: false,
        },
    )
    .unwrap();
    assert!(text.contains("123"), "OCR text missing: {text:?}");
}
