//! Watermark detection: reconstruct diagonal glyph runs and match them
//! against user needles or document-wide repetitions.
//!
//! PDF watermarks are frequently drawn as rotated text — a brand string at
//! 45°, for example. `pdf_layout` quantizes glyph rotation to the four
//! orthogonal directions, so such runs scatter into one glyph per line and
//! never appear as contiguous text in the output. This module recovers the
//! runs *before* layout assembly by grouping glyphs along their true
//! direction (`PositionedChar::angle_deg`) and offset perpendicular to it.

use crate::pdf_text::PositionedPage;
use std::collections::HashSet;

/// One reconstructed run of text written along a single, consistent
/// direction. `indices` index back into the source page's `chars`, so a
/// caller can drop exactly the glyphs that were matched.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct GlyphRun {
    /// Glyph characters in reading order (whitespace preserved as decoded).
    pub text: String,
    /// Indices into `page.chars`, in the same reading order as `text`.
    pub indices: Vec<usize>,
}

/// Glyphs more than this far off the nearest orthogonal direction are
/// treated as skewed (watermark candidates). 15° keeps genuine
/// near-horizontal/vertical text out while catching 45° watermarks.
const SKEW_TOLERANCE_DEG: f64 = 15.0;
/// Skewed glyphs within this many degrees of each other form one direction
/// group. Watermark glyphs share an exact angle; the slack absorbs producer
/// rounding.
const ANGLE_GROUP_TOL_DEG: f64 = 5.0;
/// Perpendicular-offset bucket width as a fraction of the group's median
/// glyph size, with an absolute floor. Two parallel watermark lines must
/// land in different buckets (the real corpus separates them by ~33pt at
/// 27pt type), while drift within one line must not.
const LINE_OFFSET_SIZE_RATIO: f64 = 0.5;
const LINE_OFFSET_FLOOR: f64 = 3.0;

/// Reconstruct every skewed text run on a page, in no particular order.
///
/// Orthogonal (0/90/180/270) text is ignored: it assembles correctly
/// through the normal layout path and must not be reprocessed here.
#[allow(clippy::cast_precision_loss, clippy::cast_possible_truncation)]
pub(crate) fn skewed_runs(page: &PositionedPage) -> Vec<GlyphRun> {
    let mut skewed: Vec<(usize, f64)> = page
        .chars
        .iter()
        .enumerate()
        .filter(|(_, c)| !c.ch.is_control() && skew_deg(c.angle_deg) > SKEW_TOLERANCE_DEG)
        .map(|(i, c)| (i, c.angle_deg.rem_euclid(360.0)))
        .collect();
    if skewed.is_empty() {
        return Vec::new();
    }

    // Direction groups: sort by angle and split wherever the gap exceeds
    // the tolerance. Skewed angles can only be near 45/135/225/315, so the
    // 0/360 seam never separates a genuine group.
    skewed.sort_by(|a, b| a.1.total_cmp(&b.1));
    let mut groups: Vec<Vec<(usize, f64)>> = Vec::new();
    for item in skewed {
        match groups.last_mut() {
            Some(g) if item.1 - g.last().map_or(0.0, |&(_, a)| a) <= ANGLE_GROUP_TOL_DEG => {
                g.push(item);
            }
            _ => groups.push(vec![item]),
        }
    }

    let mut runs = Vec::new();
    for group in groups {
        let angle = group.iter().map(|&(_, a)| a).sum::<f64>() / group.len() as f64;
        let rad = angle.to_radians();
        // Stored coordinates are top-down, so the PDF-space direction
        // (cos, sin) appears as (cos, -sin); a perpendicular is (sin, cos).
        let (dx, dy) = (rad.cos(), -rad.sin());
        let (px, py) = (rad.sin(), rad.cos());

        let mut sizes: Vec<f64> = group
            .iter()
            .map(|&(i, _)| page.chars[i].font_size)
            .collect();
        sizes.sort_by(f64::total_cmp);
        let median = sizes.get(sizes.len() / 2).copied().unwrap_or(0.0);
        let tol = (median * LINE_OFFSET_SIZE_RATIO).max(LINE_OFFSET_FLOOR);

        // Bucket by perpendicular offset (separates parallel lines), then
        // order along the direction (reading order within a line).
        let mut placed: Vec<(i64, f64, usize)> = group
            .iter()
            .map(|&(i, _)| {
                let c = &page.chars[i];
                let offset = c.x.mul_add(px, c.y * py);
                let along = c.x.mul_add(dx, c.y * dy);
                ((offset / tol).round() as i64, along, i)
            })
            .collect();
        placed.sort_by(|a, b| a.0.cmp(&b.0).then_with(|| a.1.total_cmp(&b.1)));

        let mut cur_bucket = i64::MIN;
        let mut current: Option<GlyphRun> = None;
        for (bucket, _, i) in placed {
            if bucket != cur_bucket {
                if let Some(run) = current.take() {
                    runs.push(run);
                }
                cur_bucket = bucket;
            }
            let run = current.get_or_insert_with(|| GlyphRun {
                text: String::new(),
                indices: Vec::new(),
            });
            run.text.push(page.chars[i].ch);
            run.indices.push(i);
        }
        if let Some(run) = current {
            runs.push(run);
        }
    }
    runs
}

/// Return `page` with the glyphs of every skewed run removed when the run
/// matches a needle in `needles` or its normalized signature is in
/// `watermarks`.
///
/// Takes ownership and edits `chars` in place: with no filters, or when no
/// run matches, the page is returned untouched and no allocation happens.
/// The default extraction path therefore pays nothing for this feature.
pub(crate) fn strip_runs(
    mut page: PositionedPage,
    needles: &[String],
    watermarks: &HashSet<String>,
) -> PositionedPage {
    if needles.is_empty() && watermarks.is_empty() {
        return page;
    }
    // Normalize the needles once rather than per run.
    let needles: Vec<String> = needles
        .iter()
        .map(|needle| normalize(needle))
        .filter(|needle| !needle.is_empty())
        .collect();
    let mut drop = HashSet::new();
    for run in skewed_runs(&page) {
        let hay = normalize(&run.text);
        if needle_match(&hay, &needles) || watermarks.contains(&hay) {
            drop.extend(run.indices);
        }
    }
    if drop.is_empty() {
        return page;
    }
    let mut i = 0;
    page.chars.retain(|_| {
        let keep = !drop.contains(&i);
        i += 1;
        keep
    });
    page
}

/// Distance in degrees from `angle` to the nearest orthogonal direction,
/// in `[0, 45]`.
fn skew_deg(angle: f64) -> f64 {
    let d = angle.rem_euclid(90.0);
    d.min(90.0 - d)
}

/// Case- and whitespace-insensitive form used for needle matching, so
/// `draftcopy` matches a run decoded as `Draft Copy`.
pub(crate) fn normalize(s: &str) -> String {
    s.chars()
        .filter(|c| !c.is_whitespace())
        .flat_map(char::to_lowercase)
        .collect()
}

/// True when the already-[`normalize`]d `hay` contains any already-normalized
/// needle.
fn needle_match(hay: &str, needles: &[String]) -> bool {
    needles.iter().any(|needle| hay.contains(needle))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pdf_text::PositionedChar;

    /// A char at `(x, y)` (top-down) advancing `adv` points along `angle`.
    fn pc(ch: char, x: f64, y: f64, size: f64, angle_deg: f64) -> PositionedChar {
        // Rotation is the orthogonal snap; the module must use angle_deg.
        #[allow(clippy::cast_sign_loss, clippy::cast_possible_truncation)]
        let rotation = ((angle_deg / 90.0).round().rem_euclid(4.0) as u16) * 90;
        PositionedChar {
            ch,
            x,
            y,
            font_size: size,
            advance: 7.0,
            rotation,
            angle_deg,
        }
    }

    /// Emit `text` as a single run starting at `(x0, y0)`, advancing along
    /// `angle_deg` so each glyph sits on its predecessor's baseline.
    fn emit(text: &str, x0: f64, y0: f64, angle_deg: f64, out: &mut Vec<PositionedChar>) {
        let (dx, dy) = angle_dir(angle_deg);
        for (i, ch) in text.chars().enumerate() {
            #[allow(clippy::cast_precision_loss)] // synthetic positions
            let t = i as f64 * 7.0;
            out.push(pc(ch, x0 + dx * t, y0 + dy * t, 27.0, angle_deg));
        }
    }

    /// Direction unit vector in the stored TOP-DOWN frame, matching
    /// `pdf_text`'s y-flip: PDF `(cos, sin)` becomes `(cos, -sin)`.
    fn angle_dir(angle_deg: f64) -> (f64, f64) {
        let r = angle_deg.to_radians();
        (r.cos(), -r.sin())
    }

    fn page(chars: Vec<PositionedChar>) -> PositionedPage {
        PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars,
        }
    }

    #[test]
    fn reconstructs_parallel_diagonal_runs() {
        let mut chars = Vec::new();
        emit("Draft ", 100.0, 500.0, 45.0, &mut chars);
        emit("Watermark", 100.0, 560.0, 45.0, &mut chars);
        let runs = skewed_runs(&page(chars));
        let mut texts: Vec<&str> = runs.iter().map(|r| r.text.as_str()).collect();
        texts.sort_unstable();
        assert_eq!(texts, vec!["Draft ", "Watermark"]);
    }

    #[test]
    fn ignored_orthogonal_text_produces_no_runs() {
        let mut chars = Vec::new();
        emit("horizontal", 100.0, 500.0, 0.0, &mut chars);
        emit("vertical", 300.0, 500.0, 90.0, &mut chars);
        assert!(skewed_runs(&page(chars)).is_empty());
    }

    #[test]
    fn run_indices_point_at_the_source_glyphs() {
        let mut chars = Vec::new();
        emit("ab", 100.0, 500.0, 45.0, &mut chars);
        let runs = skewed_runs(&page(chars));
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "ab");
        assert_eq!(runs[0].indices, vec![0, 1]);
    }

    #[test]
    fn strip_runs_drops_matched_glyphs_and_keeps_the_rest() {
        let mut chars = Vec::new();
        emit("body", 72.0, 700.0, 0.0, &mut chars);
        emit("XZQ", 100.0, 300.0, 45.0, &mut chars);
        let stripped = strip_runs(page(chars), &["xzq".into()], &HashSet::new());
        let text: String = stripped.chars.iter().map(|c| c.ch).collect();
        assert!(text.contains("body"), "body lost: {text:?}");
        assert!(
            !text.contains('X') && !text.contains('Z') && !text.contains('Q'),
            "watermark kept: {text:?}"
        );
    }

    #[test]
    fn strip_runs_drops_detected_watermark_signature() {
        let mut chars = Vec::new();
        emit("body", 72.0, 700.0, 0.0, &mut chars);
        emit("XZQ", 100.0, 300.0, 45.0, &mut chars);
        let sigs = HashSet::from([normalize("XZQ")]);
        let stripped = strip_runs(page(chars), &[], &sigs);
        let text: String = stripped.chars.iter().map(|c| c.ch).collect();
        assert!(
            !text.contains('X') && !text.contains('Z') && !text.contains('Q'),
            "{text:?}"
        );
    }

    #[test]
    fn strip_runs_with_no_needles_is_identity() {
        let mut chars = Vec::new();
        emit("XZQ", 100.0, 300.0, 45.0, &mut chars);
        let expected = chars.clone();
        let after = strip_runs(page(chars), &[], &HashSet::new());
        assert_eq!(after.chars, expected);
    }

    #[test]
    fn normalize_folds_case_and_whitespace() {
        assert_eq!(normalize("Draft Copy "), "draftcopy");
    }

    #[test]
    fn needle_match_uses_source_normalization() {
        let n = |s: &str| normalize(s);
        assert!(needle_match(&n("DraftCopy"), &[n("draftcopy")]));
        assert!(needle_match(
            &n("Sample Watermark"),
            &[n("sample watermark")]
        ));
        assert!(!needle_match(&n("Ordinary body text"), &[n("draftcopy")]));
    }

    #[test]
    fn reconstructs_runs_at_distinct_skew_angles() {
        // Opposite diagonals are separate direction groups, each recovered.
        let mut chars = Vec::new();
        emit("First", 100.0, 500.0, 45.0, &mut chars);
        emit("Second", 300.0, 300.0, 135.0, &mut chars);
        let runs = skewed_runs(&page(chars));
        let mut texts: Vec<&str> = runs.iter().map(|r| r.text.as_str()).collect();
        texts.sort_unstable();
        assert_eq!(texts, vec!["First", "Second"]);
    }
}
