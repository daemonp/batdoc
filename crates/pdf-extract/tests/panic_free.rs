//! The fork must never panic on malformed fonts/pages — errors or
//! replacement characters instead (batdoc spec §5.1).

use lopdf::{dictionary, Document, Object, Stream};

/// One-page PDF whose simple font's Differences array references a glyph
/// name that exists in no table (`/not_a_real_glyph_name`).
fn unknown_glyph_name_pdf() -> Vec<u8> {
    let mut doc = Document::with_version("1.4");
    let pages_id = doc.new_object_id();
    let font_id = doc.add_object(dictionary! {
        "Type" => "Font",
        "Subtype" => "Type1",
        "BaseFont" => "Helvetica",
        "Encoding" => dictionary! {
            "Type" => "Encoding",
            "BaseEncoding" => "WinAnsiEncoding",
            "Differences" => vec![Object::Integer(65), Object::Name(b"not_a_real_glyph_name".to_vec())],
        },
    });
    let resources_id = doc.add_object(dictionary! {
        "Font" => dictionary! { "F1" => font_id },
    });
    let content_id = doc.add_object(Stream::new(
        lopdf::Dictionary::new(),
        b"BT /F1 12 Tf 72 720 Td (A) Tj ET".to_vec(),
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

#[test]
fn unknown_glyph_name_does_not_panic() {
    // Upstream 0.10 panics here ("invalid glyph name"); the fork must
    // return Ok with a replacement char or skip the glyph.
    let result = pdf_extract::extract_text_from_mem(&unknown_glyph_name_pdf());
    result.expect("extraction must not error/panic");
}

#[test]
fn page_with_missing_contents_does_not_panic() {
    // A page dict with no /Contents: output_doc_inner must return Err,
    // not unwrap-panic.
    let mut doc = Document::with_version("1.4");
    let pages_id = doc.new_object_id();
    let page_id = doc.add_object(dictionary! {
        "Type" => "Page",
        "Parent" => pages_id,
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
    let result = pdf_extract::extract_text_from_mem(&buf);
    // Either an Ok with empty text or an Err — anything but a panic.
    let _ = result;
}

/// A page whose /Contents is a stream of bytes that do not parse as a
/// content stream: Content::decode fails, extraction must not panic.
#[test]
fn undecodable_content_stream_does_not_panic() {
    let mut doc = Document::with_version("1.4");
    let pages_id = doc.new_object_id();
    // 0xFF and other non-token bytes defeat the content lexer.
    let content_id = doc.add_object(Stream::new(
        lopdf::Dictionary::new(),
        vec![0xFF, 0xD8, 0xFF, 0xE0, 0x00],
    ));
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
    let result = pdf_extract::extract_text_from_mem(&buf);
    // Ok (page skipped) or Err — anything but a panic.
    let _ = result;
}

/// A text-showing operator with a non-string operand must not panic.
#[test]
fn malformed_tj_operand_does_not_panic() {
    let mut doc = Document::with_version("1.4");
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
    // `42 Tj` — Tj with an integer operand.
    let content_id = doc.add_object(Stream::new(
        lopdf::Dictionary::new(),
        b"BT /F1 12 Tf 72 720 Td 42 Tj ET".to_vec(),
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
    let result = pdf_extract::extract_text_from_mem(&buf);
    let _ = result;
}

/// A Type0 font missing its /DescendantFonts must not panic.
#[test]
fn cid_font_without_descendants_does_not_panic() {
    let mut doc = Document::with_version("1.4");
    let pages_id = doc.new_object_id();
    let font_id = doc.add_object(dictionary! {
        "Type" => "Font",
        "Subtype" => "Type0",
        "BaseFont" => "Broken",
        "Encoding" => "Identity-H",
    });
    let resources_id = doc.add_object(dictionary! {
        "Font" => dictionary! { "F1" => font_id },
    });
    let content_id = doc.add_object(Stream::new(
        lopdf::Dictionary::new(),
        b"BT /F1 12 Tf 72 720 Td (AB) Tj ET".to_vec(),
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
    let result = pdf_extract::extract_text_from_mem(&buf);
    let _ = result;
}
