//! Image placement extraction from PDF content streams.
//!
//! Walks the page content stream (and any nested Form `XObjects`) to record
//! every image `XObject` painted on the page, with its page-space bounding
//! rectangle expressed in top-down points (same coordinate system as
//! [`pdf_text::PositionedChar`]). This is the join key for region-aware OCR
//! merge in later phases.
//!
//! This module's image-placement machinery is exercised only by the OCR
//! path; with the `ocr` feature off it is compiled but unreferenced
//! (`PtRect` and its helpers stay in use by the layout pipeline
//! unconditionally).
#![cfg_attr(not(feature = "ocr"), allow(dead_code))]
//! PDF coordinate rules followed here:
//! - Image `XObjects` paint the unit square `[0,0,1,1]` (PDF 32000 §8.9.5),
//!   transformed by the current `CTM`.
//! - `cm` concatenates as `ctm_new = mul(ctm_old, cm_matrix)` because the
//!   matrix maps *new* user space into *current* user space (PDF 32000 §8.3.4).
//! - Final rects are flipped to top-down page coordinates using the effective
//!   `MediaBox` height.
//!
//! Malformed streams are skipped; the function never panics on bad input.

use lopdf::content::Content;
use lopdf::{Dictionary, Object, ObjectId, Stream};

/// Axis-aligned rectangle in top-down page points (same space as
/// `pdf_text::PositionedChar`: y measured from the page top).
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct PtRect {
    pub x0: f64,
    pub y0: f64,
    pub x1: f64,
    pub y1: f64,
}

impl PtRect {
    /// Intersection test expanded by `tol` on all sides.
    pub(crate) fn intersects_expanded(&self, other: &Self, tol: f64) -> bool {
        self.x0 <= other.x1 + tol
            && self.x1 >= other.x0 - tol
            && self.y0 <= other.y1 + tol
            && self.y1 >= other.y0 - tol
    }
}

/// An image `XObject` painted on a page, keyed by the same object id
/// `lopdf::Document::get_page_images` reports (the OCR join key, spec §6).
#[allow(dead_code)] // consumed by pdf_ocr/pdf_layout in later phases
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct PlacedImage {
    pub object_id: lopdf::ObjectId,
    pub rect: PtRect,
}

/// Find every image `XObject` placement on `page_id`, including placements
/// inside nested `Form` `XObjects`.
///
/// Never panics on malformed content; parse failures are skipped and the
/// function returns the placements that could be decoded.
pub(crate) fn placed_images(doc: &lopdf::Document, page_id: lopdf::ObjectId) -> Vec<PlacedImage> {
    let Some(page_height) = page_height(doc, page_id) else {
        return Vec::new();
    };
    let Ok((direct, resource_ids)) = doc.get_page_resources(page_id) else {
        return Vec::new();
    };
    let mut merged = Dictionary::new();
    // Inheritance order: root → leaf, so leaf entries override ancestors.
    for id in resource_ids.iter().rev() {
        if let Ok(dict) = doc.get_dictionary(*id) {
            for (k, v) in dict {
                merged.set(k.clone(), v.clone());
            }
        }
    }
    if let Some(dict) = direct {
        for (k, v) in dict {
            merged.set(k.clone(), v.clone());
        }
    }
    if merged.is_empty() {
        return Vec::new();
    }
    let Ok(content) = doc.get_page_content(page_id) else {
        return Vec::new();
    };

    let mut walker = Walker {
        doc,
        page_height,
        out: Vec::new(),
    };
    walker.walk(&content, &merged, IDENTITY, 0);
    walker.out
}

const IDENTITY: [f64; 6] = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
const MAX_FORM_DEPTH: usize = 8;

struct Walker<'a> {
    doc: &'a lopdf::Document,
    page_height: f64,
    out: Vec<PlacedImage>,
}

impl Walker<'_> {
    fn walk(&mut self, content: &[u8], resources: &Dictionary, ctm: [f64; 6], depth: usize) {
        let Ok(content) = Content::decode(content) else {
            return;
        };

        let mut stack: Vec<[f64; 6]> = vec![ctm];
        for op in &content.operations {
            match op.operator.as_str() {
                "q" => stack.push(*stack.last().unwrap_or(&IDENTITY)),
                "Q" => {
                    stack.pop();
                    if stack.is_empty() {
                        stack.push(IDENTITY);
                    }
                }
                "cm" => {
                    if let Some(cur) = stack.last_mut() {
                        if let Some(m) = operands_to_matrix(&op.operands) {
                            *cur = mul(*cur, m);
                        }
                    }
                }
                "Do" => {
                    let cur = *stack.last().unwrap_or(&IDENTITY);
                    self.handle_do(resources, &op.operands, cur, depth);
                }
                _ => {}
            }
        }
    }

    fn handle_do(
        &mut self,
        resources: &Dictionary,
        operands: &[Object],
        ctm: [f64; 6],
        depth: usize,
    ) {
        let Some(Object::Name(name)) = operands.first() else {
            return;
        };
        let xobject_dict = match resources.get(b"XObject") {
            Ok(obj) => match self.doc.dereference(obj) {
                Ok((_, Object::Dictionary(dict))) => dict,
                _ => return,
            },
            Err(_) => return,
        };
        let obj_id = match xobject_dict.get(name) {
            Ok(Object::Reference(id)) => *id,
            Ok(obj) => match self.doc.dereference(obj) {
                Ok((_, Object::Reference(id))) => *id,
                _ => return,
            },
            Err(_) => return,
        };
        let Ok(Object::Stream(stream)) = self.doc.get_object(obj_id) else {
            return;
        };
        let subtype = stream
            .dict
            .get(b"Subtype")
            .ok()
            .and_then(|o| self.doc.dereference(o).ok())
            .and_then(|(_, o)| o.as_name().ok())
            .and_then(|name| std::str::from_utf8(name).ok())
            .unwrap_or("");
        match subtype {
            "Image" => {
                self.out.push(PlacedImage {
                    object_id: obj_id,
                    rect: unit_rect_in_page_space(ctm, self.page_height),
                });
            }
            "Form" => {
                if depth >= MAX_FORM_DEPTH {
                    return;
                }
                let form_matrix = matrix_from_stream(self.doc, stream);
                let ctm2 = mul(ctm, form_matrix);
                let form_resources = stream
                    .dict
                    .get(b"Resources")
                    .ok()
                    .and_then(|o| self.doc.dereference(o).ok())
                    .and_then(|(_, o)| o.as_dict().ok())
                    .unwrap_or(resources);
                let form_content = match stream.decompressed_content() {
                    Ok(bytes) => bytes,
                    Err(_) => {
                        // lopdf's `decompressed_content` returns Err when the
                        // stream dict has no /Filter (i.e. raw content). Treat
                        // missing Filter as identity and use the raw bytes.
                        if stream.dict.get(b"Filter").is_err() {
                            stream.content.clone()
                        } else {
                            return;
                        }
                    }
                };
                self.walk(&form_content, form_resources, ctm2, depth + 1);
            }
            _ => {}
        }
    }
}

/// Extract the form `/Matrix` from a Form stream dictionary, defaulting to
/// identity.
fn matrix_from_stream(doc: &lopdf::Document, stream: &Stream) -> [f64; 6] {
    let Ok(obj) = stream.dict.get(b"Matrix") else {
        return IDENTITY;
    };
    match doc.dereference(obj) {
        Ok((_, o)) => operands_to_matrix_from_object(o).unwrap_or(IDENTITY),
        Err(_) => operands_to_matrix_from_object(obj).unwrap_or(IDENTITY),
    }
}

/// Transform the unit square by `ctm` and convert to a top-down `PtRect`.
fn unit_rect_in_page_space(ctm: [f64; 6], page_height: f64) -> PtRect {
    let corners = [
        transform_point(ctm, (0.0, 0.0)),
        transform_point(ctm, (1.0, 0.0)),
        transform_point(ctm, (0.0, 1.0)),
        transform_point(ctm, (1.0, 1.0)),
    ];
    let mut min_x = f64::INFINITY;
    let mut max_x = f64::NEG_INFINITY;
    let mut min_y = f64::INFINITY;
    let mut max_y = f64::NEG_INFINITY;
    for (x, y) in corners {
        min_x = min_x.min(x);
        max_x = max_x.max(x);
        min_y = min_y.min(y);
        max_y = max_y.max(y);
    }
    PtRect {
        x0: min_x,
        y0: page_height - max_y,
        x1: max_x,
        y1: page_height - min_y,
    }
}

#[allow(clippy::suboptimal_flops)] // Brief specifies the exact affine formula; reordering is not desired.
fn transform_point(m: [f64; 6], (x, y): (f64, f64)) -> (f64, f64) {
    (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])
}

/// Convert six numeric operands into an affine matrix `[a b c d e f]`.
#[allow(clippy::cast_precision_loss)] // PDF integer operands are within f64 precision for coordinates.
fn operands_to_matrix(operands: &[Object]) -> Option<[f64; 6]> {
    if operands.len() < 6 {
        return None;
    }
    let mut out = [0.0; 6];
    for (i, obj) in operands.iter().enumerate().take(6) {
        out[i] = match obj {
            Object::Integer(n) => *n as f64,
            Object::Real(f) => f64::from(*f),
            Object::Reference(_) => {
                // Indirect operands are invalid in a content stream; skip.
                return None;
            }
            _ => return None,
        };
    }
    Some(out)
}

/// Convert six numeric objects (e.g. an array) into an affine matrix.
fn operands_to_matrix_from_object(obj: &Object) -> Option<[f64; 6]> {
    match obj {
        Object::Array(arr) => operands_to_matrix(arr),
        _ => None,
    }
}

/// Multiply two affine matrices: `mul(m1, m2)` means apply `m2` first, then `m1`.
#[allow(clippy::suboptimal_flops)] // Brief specifies the exact affine formula; reordering is not desired.
fn mul(m1: [f64; 6], m2: [f64; 6]) -> [f64; 6] {
    [
        m1[0] * m2[0] + m1[2] * m2[1],
        m1[1] * m2[0] + m1[3] * m2[1],
        m1[0] * m2[2] + m1[2] * m2[3],
        m1[1] * m2[2] + m1[3] * m2[3],
        m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
        m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
    ]
}

/// Effective `MediaBox` height for `page_id`, walking the `/Parent` chain.
fn page_height(doc: &lopdf::Document, page_id: ObjectId) -> Option<f64> {
    let dict = doc.get_dictionary(page_id).ok()?;
    let media_box = media_box_for_dict(doc, dict)?;
    Some(media_box[3] - media_box[1])
}

const MAX_MEDIABOX_PARENT_DEPTH: u32 = 32;

/// Resolve the effective `MediaBox` array for a dictionary, walking `/Parent`.
fn media_box_for_dict(doc: &lopdf::Document, dict: &Dictionary) -> Option<[f64; 4]> {
    media_box_for_dict_inner(doc, dict, 0)
}

fn media_box_for_dict_inner(
    doc: &lopdf::Document,
    dict: &Dictionary,
    depth: u32,
) -> Option<[f64; 4]> {
    if depth > MAX_MEDIABOX_PARENT_DEPTH {
        // Cyclic or absurdly deep parent chain: fall back to the default.
        return Some([0.0, 0.0, 612.0, 792.0]);
    }
    if let Ok(obj) = dict.get(b"MediaBox") {
        if let Ok(rect) = rect_array(doc, obj) {
            return Some(rect);
        }
    }
    if let Ok(parent) = dict.get(b"Parent") {
        let parent_id = parent.as_reference().ok()?;
        let parent_dict = doc.get_dictionary(parent_id).ok()?;
        return media_box_for_dict_inner(doc, parent_dict, depth + 1);
    }
    // PDF spec default: US Letter.
    Some([0.0, 0.0, 612.0, 792.0])
}

/// Convert an indirect or direct array into `[llx, lly, urx, ury]`.
#[allow(clippy::cast_precision_loss)] // PDF MediaBox coordinates are within f64 precision.
fn rect_array(doc: &lopdf::Document, obj: &Object) -> Result<[f64; 4], ()> {
    let (_, obj) = doc.dereference(obj).map_err(|_| ())?;
    let arr = obj.as_array().map_err(|_| ())?;
    if arr.len() < 4 {
        return Err(());
    }
    let mut out = [0.0; 4];
    for (i, item) in arr.iter().enumerate().take(4) {
        out[i] = match item {
            Object::Integer(n) => *n as f64,
            Object::Real(f) => f64::from(*f),
            _ => return Err(()),
        };
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use lopdf::{dictionary, Document, Object, Stream};

    #[test]
    fn placed_image_rect_from_cm_do() {
        // q 200 0 0 50 72 600 cm /Im0 Do Q  →  PDF-space rect
        // x: 72..272, y: 600..650; page height 792 → top-down y: 142..192.
        let (doc, page_id) = image_pdf("q 200 0 0 50 72 600 cm /Im0 Do Q");
        let placed = placed_images(&doc, page_id);
        assert_eq!(placed.len(), 1);
        let r = &placed[0].rect;
        assert!((r.x0 - 72.0).abs() < 0.01, "{r:?}");
        assert!((r.x1 - 272.0).abs() < 0.01, "{r:?}");
        assert!((r.y0 - 142.0).abs() < 0.01, "{r:?}");
        assert!((r.y1 - 192.0).abs() < 0.01, "{r:?}");
    }

    #[test]
    fn object_id_matches_get_page_images() {
        let (doc, page_id) = image_pdf("q 100 0 0 100 0 0 cm /Im0 Do Q");
        let placed = placed_images(&doc, page_id);
        let images = doc.get_page_images(page_id).unwrap();
        assert_eq!(placed.len(), 1);
        assert_eq!(placed[0].object_id, images[0].id);
    }

    #[test]
    fn nested_form_xobject_recursion() {
        // A Form XObject (own resources, /Matrix translate by 10,20)
        // containing "q 50 0 0 50 1 2 cm /Im1 Do Q"; page content does
        // "q 2 0 0 2 100 100 cm /Fm0 Do Q". Expected PDF-space unit-square
        // mapping: ctm = [2 0 0 2 100 100] ∘ [1 0 0 1 10 20] ∘ [50 0 0 50 1 2]
        //   → x = 2*(1*50x ... ) — compute in the test via the same corner
        //   math, but assert the final top-down rect numerically:
        //   inner point (0,0) → form: (1,2) → +matrix (11,22) → page ctm:
        //   (122, 144); inner (1,1) → form (51,52) → (61,72) → (222, 244).
        //   top-down (h=792): y0 = 792-244 = 548, y1 = 792-144 = 648.
        let (doc, page_id) = form_pdf();
        let placed = placed_images(&doc, page_id);
        assert_eq!(placed.len(), 1);
        let r = &placed[0].rect;
        assert!((r.x0 - 122.0).abs() < 0.01, "{r:?}");
        assert!((r.x1 - 222.0).abs() < 0.01, "{r:?}");
        assert!((r.y0 - 548.0).abs() < 0.01, "{r:?}");
        assert!((r.y1 - 648.0).abs() < 0.01, "{r:?}");
    }

    #[test]
    fn resources_inherited_from_pages_node() {
        // Page dict omits /Resources; the /Pages node carries them.
        // Same content as the first test → same rect.
        let (doc, page_id) = image_pdf_inherited_resources("q 200 0 0 50 72 600 cm /Im0 Do Q");
        let placed = placed_images(&doc, page_id);
        assert_eq!(placed.len(), 1);
        assert!((placed[0].rect.x0 - 72.0).abs() < 0.01);
    }

    #[test]
    fn rect_intersection_with_tolerance() {
        let a = PtRect {
            x0: 0.0,
            y0: 0.0,
            x1: 10.0,
            y1: 10.0,
        };
        let b = PtRect {
            x0: 11.0,
            y0: 0.0,
            x1: 20.0,
            y1: 10.0,
        };
        assert!(!a.intersects_expanded(&b, 0.5));
        assert!(a.intersects_expanded(&b, 2.0));
    }

    #[test]
    fn cyclic_parent_chain_does_not_stack_overflow() {
        // A cyclic /Parent chain must not hang or crash. get_page_resources
        // detects the cycle and returns Err, so placed_images returns empty.
        // The key property is that the call terminates without a stack overflow.
        let (doc, page_id) = image_pdf_cyclic_parent("q 200 0 0 50 72 600 cm /Im0 Do Q");
        let placed = placed_images(&doc, page_id);
        assert!(placed.is_empty());
    }

    #[test]
    fn media_box_for_dict_guards_against_cyclic_parent() {
        // Directly exercise media_box_for_dict with a Pages node whose Parent
        // is itself. Without the depth cap it would recurse until stack overflow.
        let mut doc = Document::with_version("1.4");
        let pages_id = doc.new_object_id();
        let pages_dict = dictionary! {
            "Type" => "Pages",
            "Kids" => Vec::<Object>::new(),
            "Count" => 0_i64,
            "Parent" => pages_id,
        };
        doc.objects.insert(pages_id, Object::Dictionary(pages_dict));
        let pages_dict = doc
            .get_dictionary(pages_id)
            .expect("pages dict we just inserted");
        let rect = media_box_for_dict(&doc, pages_dict);
        // Falls back to the PDF default US-Letter MediaBox.
        assert_eq!(rect, Some([0.0, 0.0, 612.0, 792.0]));
    }

    /// Build a one-page PDF with an image XObject in resources and a
    /// caller-supplied content stream.
    fn image_pdf(content: &str) -> (Document, ObjectId) {
        let mut doc = Document::with_version("1.4");
        let pages_id = doc.new_object_id();

        let image_id = doc.add_object(image_stream());
        let resources_id = doc.add_object(dictionary! {
            "XObject" => dictionary! { "Im0" => image_id },
        });
        let content_id =
            doc.add_object(Stream::new(Dictionary::new(), content.as_bytes().to_vec()));
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
        let loaded = Document::load_mem(&buf).unwrap();
        (loaded, page_id)
    }

    /// Build a one-page PDF where the /Pages node carries /Resources and the
    /// page dict omits it.
    fn image_pdf_inherited_resources(content: &str) -> (Document, ObjectId) {
        let mut doc = Document::with_version("1.4");
        let pages_id = doc.new_object_id();

        let image_id = doc.add_object(image_stream());
        let resources_id = doc.add_object(dictionary! {
            "XObject" => dictionary! { "Im0" => image_id },
        });
        let content_id =
            doc.add_object(Stream::new(Dictionary::new(), content.as_bytes().to_vec()));
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
                "Resources" => resources_id,
            }),
        );
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);

        let mut buf = Vec::new();
        doc.save_to(&mut buf).unwrap();
        let loaded = Document::load_mem(&buf).unwrap();
        (loaded, page_id)
    }

    /// Build a one-page PDF with a cyclic /Parent chain. The page has direct
    /// /Resources so resource lookup succeeds, but it has no /MediaBox and the
    /// Pages node's /Parent points to itself, so media_box_for_dict must guard
    /// against infinite recursion.
    fn image_pdf_cyclic_parent(content: &str) -> (Document, ObjectId) {
        let mut doc = Document::with_version("1.4");
        let pages_id = doc.new_object_id();

        let image_id = doc.add_object(image_stream());
        let resources_id = doc.add_object(dictionary! {
            "XObject" => dictionary! { "Im0" => image_id },
        });
        let content_id =
            doc.add_object(Stream::new(Dictionary::new(), content.as_bytes().to_vec()));
        // Page intentionally omits /MediaBox so the parent walk is exercised.
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => content_id,
            "Resources" => resources_id,
        });
        // Cyclic Pages node: its /Parent points to itself and it has no /MediaBox.
        doc.objects.insert(
            pages_id,
            Object::Dictionary(dictionary! {
                "Type" => "Pages",
                "Kids" => vec![Object::Reference(page_id)],
                "Count" => 1_i64,
                "Parent" => pages_id,
            }),
        );
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);

        let mut buf = Vec::new();
        doc.save_to(&mut buf).unwrap();
        let loaded = Document::load_mem(&buf).unwrap();
        (loaded, page_id)
    }

    /// Build a one-page PDF with a Form XObject that contains an image.
    fn form_pdf() -> (Document, ObjectId) {
        let mut doc = Document::with_version("1.4");
        let pages_id = doc.new_object_id();

        let image_id = doc.add_object(image_stream_named("Im1"));
        let form_resources_id = doc.add_object(dictionary! {
            "XObject" => dictionary! { "Im1" => image_id },
        });
        let form_content = b"q 50 0 0 50 1 2 cm /Im1 Do Q".to_vec();
        let form_id = doc.add_object(Stream::new(
            dictionary! {
                "Type" => "XObject",
                "Subtype" => "Form",
                "BBox" => vec![0.into(), 0.into(), 100.into(), 100.into()],
                "Matrix" => vec![1.into(), 0.into(), 0.into(), 1.into(), 10.into(), 20.into()],
                "Resources" => form_resources_id,
            },
            form_content,
        ));
        let page_resources_id = doc.add_object(dictionary! {
            "XObject" => dictionary! { "Fm0" => form_id },
        });
        let page_content = b"q 2 0 0 2 100 100 cm /Fm0 Do Q".to_vec();
        let content_id = doc.add_object(Stream::new(Dictionary::new(), page_content));
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => content_id,
            "Resources" => page_resources_id,
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
        let loaded = Document::load_mem(&buf).unwrap();
        (loaded, page_id)
    }

    fn image_stream() -> Stream {
        image_stream_named("Im0")
    }

    fn image_stream_named(_name: &str) -> Stream {
        // 2x2 DeviceRGB image, raw (no compression needed for tests).
        Stream::new(
            dictionary! {
                "Type" => "XObject",
                "Subtype" => "Image",
                "Width" => 2_i64,
                "Height" => 2_i64,
                "ColorSpace" => "DeviceRGB",
                "BitsPerComponent" => 8_i64,
            },
            vec![255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255],
        )
    }
}
