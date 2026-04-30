//! Shared XML utility functions for OOXML parsers.
//!
//! Used by `docx.rs`, `xlsx.rs`, and `pptx.rs` to extract attributes from
//! `quick_xml` elements, parse relationship files, and compute `_rels` paths
//! without duplicating the parsing logic.

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::HashMap;
use std::io::{Cursor, Read};
use zip::ZipArchive;

/// Relationship map: rId → target URL.
pub(crate) type Rels = HashMap<String, String>;

/// Get an attribute value from an XML element by name.
pub(crate) fn get_attr(e: &quick_xml::events::BytesStart, attr_name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == attr_name {
            return std::str::from_utf8(&attr.value).ok().map(String::from);
        }
    }
    None
}

/// Parse an OOXML relationships XML string into an rId → URL map.
///
/// Only includes relationships with `TargetMode="External"` (hyperlinks)
/// or those whose Type ends with `/hyperlink`.
pub(crate) fn parse_rels_xml(xml: &str) -> Rels {
    let mut rels = Rels::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let id = get_attr(e, b"Id").unwrap_or_default();
                let target = get_attr(e, b"Target").unwrap_or_default();
                let rel_type = get_attr(e, b"Type").unwrap_or_default();
                let target_mode = get_attr(e, b"TargetMode").unwrap_or_default();

                let is_external = target_mode.eq_ignore_ascii_case("External");
                let is_hyperlink = rel_type.ends_with("/hyperlink");
                if !id.is_empty() && !target.is_empty() && (is_external || is_hyperlink) {
                    rels.insert(id, target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    rels
}

/// Parse an OOXML relationships XML string into an rId → target path map
/// for image relationships.
///
/// Only includes relationships whose Type ends with `/image`.
pub(crate) fn parse_image_rels_xml(xml: &str) -> Rels {
    let mut rels = Rels::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let id = get_attr(e, b"Id").unwrap_or_default();
                let target = get_attr(e, b"Target").unwrap_or_default();
                let rel_type = get_attr(e, b"Type").unwrap_or_default();

                if !id.is_empty() && !target.is_empty() && rel_type.ends_with("/image") {
                    rels.insert(id, target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    rels
}

/// Load image relationships from a `.rels` file in a ZIP archive.
///
/// Returns an empty map if the file doesn't exist or can't be read.
pub(crate) fn load_image_rels(archive: &mut ZipArchive<Cursor<&[u8]>>, path: &str) -> Rels {
    let mut xml = String::new();
    match archive.by_name(path) {
        Ok(mut entry) => {
            if entry.read_to_string(&mut xml).is_err() {
                return Rels::new();
            }
        }
        Err(_) => return Rels::new(),
    }
    parse_image_rels_xml(&xml)
}

/// Load a relationships file from a ZIP archive and parse it into a `Rels` map.
///
/// Returns an empty map if the file doesn't exist or can't be read.
pub(crate) fn load_rels(archive: &mut ZipArchive<Cursor<&[u8]>>, path: &str) -> Rels {
    let mut xml = String::new();
    match archive.by_name(path) {
        Ok(mut entry) => {
            if entry.read_to_string(&mut xml).is_err() {
                return Rels::new();
            }
        }
        Err(_) => return Rels::new(),
    }
    parse_rels_xml(&xml)
}

/// Read image bytes from a ZIP archive given a relationship target and base directory.
///
/// The `target` is the value from a `.rels` file (e.g., `"media/image1.png"` or
/// `"../media/image1.png"`). The `base_dir` is the directory of the part that
/// owns the relationship (e.g., `"word"` for `word/document.xml`).
///
/// Returns `None` if the entry doesn't exist or can't be read.
pub(crate) fn read_image_from_zip(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    target: &str,
    base_dir: &str,
) -> Option<Vec<u8>> {
    // Resolve relative path: join base_dir + target, then normalize "../"
    let full_path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        let raw = if base_dir.is_empty() {
            target.to_string()
        } else {
            format!("{base_dir}/{target}")
        };
        normalize_zip_path(&raw)
    };

    let mut data = Vec::new();
    archive
        .by_name(&full_path)
        .ok()?
        .read_to_end(&mut data)
        .ok()?;
    Some(data)
}

/// Normalize a ZIP path by resolving `..` segments.
///
/// `"ppt/slides/../media/image1.png"` → `"ppt/media/image1.png"`
fn normalize_zip_path(path: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for segment in path.split('/') {
        if segment == ".." {
            parts.pop();
        } else if !segment.is_empty() && segment != "." {
            parts.push(segment);
        }
    }
    parts.join("/")
}

/// Compute the `_rels` file path for a given OOXML part path.
///
/// For `xl/worksheets/sheet1.xml`, returns `xl/worksheets/_rels/sheet1.xml.rels`.
/// For `ppt/slides/slide1.xml`, returns `ppt/slides/_rels/slide1.xml.rels`.
pub(crate) fn rels_path(part_path: &str) -> String {
    if let Some((dir, file)) = part_path.rsplit_once('/') {
        format!("{dir}/_rels/{file}.rels")
    } else {
        format!("_rels/{part_path}.rels")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── parse_rels_xml ───────────────────────────────────────────

    #[test]
    fn parse_rels_hyperlinks() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:test@example.com" TargetMode="External"/>
</Relationships>"#;
        let rels = parse_rels_xml(xml);
        assert_eq!(rels.get("rId4").unwrap(), "https://example.com");
        assert_eq!(rels.get("rId5").unwrap(), "mailto:test@example.com");
        assert!(!rels.contains_key("rId1")); // styles.xml is not a hyperlink
    }

    #[test]
    fn parse_rels_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;
        let rels = parse_rels_xml(xml);
        assert!(rels.is_empty());
    }

    #[test]
    fn parse_rels_external_not_hyperlink_type() {
        // External target without hyperlink type — should still be included
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/custom" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;
        let rels = parse_rels_xml(xml);
        assert_eq!(rels.get("rId1").unwrap(), "https://example.com");
    }

    // ── parse_image_rels_xml ─────────────────────────────────────

    #[test]
    fn parse_image_rels_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpeg"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;
        let rels = parse_image_rels_xml(xml);
        assert_eq!(rels.len(), 2);
        assert_eq!(rels.get("rId2").unwrap(), "media/image1.png");
        assert_eq!(rels.get("rId3").unwrap(), "media/image2.jpeg");
        assert!(!rels.contains_key("rId1")); // styles
        assert!(!rels.contains_key("rId4")); // hyperlink
    }

    #[test]
    fn parse_image_rels_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;
        let rels = parse_image_rels_xml(xml);
        assert!(rels.is_empty());
    }

    // ── rels_path ─────────────────────────────────────────────────

    #[test]
    fn rels_path_nested() {
        assert_eq!(
            rels_path("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
    }

    #[test]
    fn rels_path_slides() {
        assert_eq!(
            rels_path("ppt/slides/slide1.xml"),
            "ppt/slides/_rels/slide1.xml.rels"
        );
    }

    #[test]
    fn rels_path_no_dir() {
        assert_eq!(rels_path("sheet1.xml"), "_rels/sheet1.xml.rels");
    }

    // ── normalize_zip_path ────────────────────────────────────────

    #[test]
    fn normalize_simple() {
        assert_eq!(
            normalize_zip_path("word/media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn normalize_dotdot() {
        assert_eq!(
            normalize_zip_path("ppt/slides/../media/image1.png"),
            "ppt/media/image1.png"
        );
    }

    #[test]
    fn normalize_multiple_dotdot() {
        assert_eq!(
            normalize_zip_path("a/b/c/../../d/image.png"),
            "a/d/image.png"
        );
    }

    #[test]
    fn normalize_no_dotdot() {
        assert_eq!(
            normalize_zip_path("xl/media/image1.png"),
            "xl/media/image1.png"
        );
    }
}
