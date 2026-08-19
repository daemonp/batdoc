//! Layout algorithms derived from run-llama/liteparse (Apache-2.0);
//! reimplemented for batdoc under MIT — no code or tables copied.
//!
//! Assembles the positioned per-char stream from [`crate::pdf_text`] into
//! words and lines: canonical rotation → rotation partition → baseline
//! clustering → advance-gap word breaks → horizontal (column) line split.
//! liteparse runs the same pipeline over whole text items with PDFium-sized
//! boxes; here it runs per char with `font_size`/`advance` from pdf-extract,
//! so the thresholds are restated in those units (see the constants below).
//!
//! Coordinate frames: chars arrive in top-down page space. [`assemble`]
//! canonicalizes the page rotation (majority text becomes horizontal),
//! assembles each rotation group in its own frame, and transforms every
//! emitted [`Line::rect`] back to top-down page space — the join key for
//! region-aware OCR merge. [`Word`] x-extents stay in the line's assembly
//! frame (identity unless the page was canonicalized).

use crate::pdf_geometry::PtRect;
use crate::pdf_text::{PositionedChar, PositionedPage};

/// Where a line's text came from. `Ocr` lines are synthesized by the
/// region-aware merge (later phase) rather than from native PDF glyphs.
///
/// Consumed by Tasks 10–13 — allow dead code until those land.
#[allow(dead_code)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum LineSource {
    Native,
    Ocr,
}

/// A word with its x-extent, rebuilt from per-char advance gaps (spec §7 —
/// NOT naive space splitting).
///
/// Consumed by Tasks 10–13 — allow dead code until those land.
#[allow(dead_code)]
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Word {
    pub text: String,
    pub x0: f64,
    pub x1: f64,
}

/// A baseline cluster of chars assembled into text, with the union of its
/// char boxes in top-down page space.
///
/// Consumed by Tasks 10–13 — allow dead code until those land.
#[allow(dead_code)]
#[derive(Clone, Debug)]
pub(crate) struct Line {
    pub text: String,
    pub words: Vec<Word>,
    pub rect: PtRect,
    /// Median transformed font size of the line's chars.
    pub font_size: f64,
    pub source: LineSource,
}

/// Word break: gap between consecutive chars exceeds a quarter of the line
/// font size, with a 1pt floor so tiny fonts don't shatter words.
const WORD_GAP_RATIO: f64 = 0.25;
const WORD_GAP_FLOOR: f64 = 1.0;
/// Horizontal line split: a same-baseline gap this many times the font size
/// means two columns, not two words (liteparse `form_lines` equivalent).
const LINE_SPLIT_RATIO: f64 = 1.5;
/// Baseline band: chars join a line when `[y - 0.8·size, y + 0.2·size]`
/// overlaps the line's running band (asymmetric: top-down y grows downward,
/// so the band reaches further up toward the previous baseline).
const BAND_UP: f64 = 0.8;
const BAND_DOWN: f64 = 0.2;
/// One rotation value must cover this share of the page's chars before the
/// whole page is canonicalized to it.
const CANONICAL_SHARE_PCT: usize = 70;

// Rotation math runs in `u16` (`PositionedChar.rotation` is u16 degrees),
// and the module must model the full 0/90/180/270 set (spec-pinned
// transforms).
const R90: u16 = 90;
const R180: u16 = 180;
const R270: u16 = 270;
const R360: u16 = 360;

/// Assemble one page of positioned chars into lines (words, rects, sizes).
///
/// Consumed by Tasks 10–13 — allow dead code until those land.
#[allow(dead_code)]
pub(crate) fn assemble(page: &PositionedPage) -> Vec<Line> {
    if page.chars.is_empty() {
        return Vec::new();
    }
    let page_w = page.media_box.2 - page.media_box.0;
    let page_h = page.media_box.3 - page.media_box.1;
    let (chars, page_rot) = canonicalize(&page.chars, page_w, page_h);
    let (frame_w, frame_h) = frame_dims(page_rot, page_w, page_h);

    // Partition by rotation relative to the canonical frame: the majority
    // group is rotation 0; minority rotations land in their own groups. The
    // relative rotation is kept in a local u16 (270 does not fit the char
    // field's u8) and never written back into a `PositionedChar`.
    let mut groups: [Vec<PositionedChar>; 4] = [Vec::new(), Vec::new(), Vec::new(), Vec::new()];
    for c in chars {
        let rel = (c.rotation + R360 - page_rot) % R360;
        groups[rot_index(rel)].push(c);
    }

    let back = (R360 - page_rot) % R360;
    let mut out = assemble_group(&groups[0]);
    for line in &mut out {
        line.rect = rotate_rect(line.rect, back, frame_w, frame_h);
    }
    // Minority groups (rotated side-labels on an otherwise-canonical page):
    // assemble each in its own unrotated frame so they are not silently
    // dropped, then map rects group frame → canonical frame → page space.
    for (i, group) in groups.iter().enumerate().skip(1) {
        if group.is_empty() {
            continue;
        }
        let g = u16::try_from(i * 90).unwrap_or_default(); // i ∈ 1..=3 → ≤ 270
        let unrotated: Vec<PositionedChar> = group
            .iter()
            .map(|c| {
                let (x, y) = unrotate_point(g, c.x, c.y, frame_w, frame_h);
                PositionedChar {
                    x,
                    y,
                    rotation: 0,
                    ..*c
                }
            })
            .collect();
        let (gw, gh) = frame_dims(g, frame_w, frame_h);
        let mut lines = assemble_group(&unrotated);
        for line in &mut lines {
            let canonical = rotate_rect(line.rect, (R360 - g) % R360, gw, gh);
            line.rect = rotate_rect(canonical, back, frame_w, frame_h);
        }
        out.extend(lines);
    }
    out
}

/// If one non-zero rotation covers ≥ 70% of chars, rotate every char into
/// that majority's reading frame and return `(chars, applied_rotation)`.
/// Char rotations are left untouched; the caller combines them with the
/// applied rotation (in u16) to partition minority groups. Font size and
/// advance are lengths and survive rotation unchanged.
fn canonicalize(chars: &[PositionedChar], w: f64, h: f64) -> (Vec<PositionedChar>, u16) {
    let mut counts = [0usize; 4];
    for c in chars {
        counts[rot_index(c.rotation)] += 1;
    }
    let mut best: u16 = 0;
    let mut best_count = 0;
    for (i, &n) in counts.iter().enumerate().skip(1) {
        if n > best_count {
            best_count = n;
            best = u16::try_from(i * 90).unwrap_or_default(); // i ∈ 1..=3 → ≤ 270
        }
    }
    if best_count * 100 < chars.len() * CANONICAL_SHARE_PCT {
        return (chars.to_vec(), 0);
    }
    let out = chars
        .iter()
        .map(|c| {
            let (x, y) = unrotate_point(best, c.x, c.y, w, h);
            PositionedChar { x, y, ..*c }
        })
        .collect();
    (out, best)
}

/// Cluster chars into baseline groups: sorted by y, a char joins the current
/// cluster while its band overlaps the cluster's running band.
fn cluster_lines(chars: &[PositionedChar]) -> Vec<Vec<PositionedChar>> {
    let mut sorted = chars.to_vec();
    sorted.sort_by(|a, b| a.y.total_cmp(&b.y).then_with(|| a.x.total_cmp(&b.x)));
    let mut lines: Vec<Vec<PositionedChar>> = Vec::new();
    let mut band = (0.0, 0.0);
    for c in sorted {
        let lo = (-BAND_UP).mul_add(c.font_size, c.y);
        let hi = BAND_DOWN.mul_add(c.font_size, c.y);
        if let Some(cur) = lines.last_mut() {
            if lo <= band.1 && hi >= band.0 {
                cur.push(c);
                band.0 = band.0.min(lo);
                band.1 = band.1.max(hi);
                continue;
            }
        }
        band = (lo, hi);
        lines.push(vec![c]);
    }
    lines
}

/// Assemble one rotation group (all rotation 0 in its own frame) into
/// lines: baseline clusters, then split each cluster horizontally where a
/// gap exceeds the column threshold.
fn assemble_group(chars: &[PositionedChar]) -> Vec<Line> {
    let mut out = Vec::new();
    for cluster in cluster_lines(chars) {
        let mut by_x = cluster;
        by_x.sort_by(|a, b| a.x.total_cmp(&b.x));
        let split_gap = LINE_SPLIT_RATIO * median_size(&by_x);
        let mut cur: Vec<PositionedChar> = Vec::new();
        for c in by_x {
            if let Some(prev) = cur.last() {
                if c.x - (prev.x + prev.advance) > split_gap {
                    out.push(build_line(&cur));
                    cur.clear();
                }
            }
            cur.push(c);
        }
        if !cur.is_empty() {
            out.push(build_line(&cur));
        }
    }
    out
}

/// Build one line from its x-sorted chars: word breaks on explicit space
/// glyphs and on advance gaps, rect = union of char boxes.
fn build_line(chars: &[PositionedChar]) -> Line {
    let size = median_size(chars);
    let word_gap = (WORD_GAP_RATIO * size).max(WORD_GAP_FLOOR);
    let mut words: Vec<Word> = Vec::new();
    let mut text = String::new();
    let (mut wx0, mut wx1) = (0.0, 0.0);
    let mut rect = PtRect {
        x0: f64::INFINITY,
        y0: f64::INFINITY,
        x1: f64::NEG_INFINITY,
        y1: f64::NEG_INFINITY,
    };
    let mut prev: Option<&PositionedChar> = None;
    for c in chars {
        rect.x0 = rect.x0.min(c.x);
        rect.y0 = rect.y0.min(c.y - c.font_size);
        rect.x1 = rect.x1.max(c.x + c.advance);
        rect.y1 = rect.y1.max(c.y);
        // A real space glyph always breaks the word (and is not itself
        // emitted); so does an advance gap past the threshold.
        let break_before =
            prev.is_some_and(|p| p.ch == ' ' || c.ch == ' ' || c.x - (p.x + p.advance) > word_gap);
        if break_before && !text.is_empty() {
            words.push(Word {
                text: std::mem::take(&mut text),
                x0: wx0,
                x1: wx1,
            });
        }
        if c.ch != ' ' {
            if text.is_empty() {
                wx0 = c.x;
                wx1 = c.x + c.advance;
            } else {
                wx1 = wx1.max(c.x + c.advance);
            }
            text.push(c.ch);
        }
        prev = Some(c);
    }
    if !text.is_empty() {
        words.push(Word {
            text,
            x0: wx0,
            x1: wx1,
        });
    }
    let text = words
        .iter()
        .map(|w| w.text.as_str())
        .collect::<Vec<_>>()
        .join(" ");
    Line {
        text,
        words,
        rect,
        font_size: size,
        source: LineSource::Native,
    }
}

/// Median char size of a non-empty slice (mean of the two middles when even).
fn median_size(chars: &[PositionedChar]) -> f64 {
    let mut sizes: Vec<f64> = chars.iter().map(|c| c.font_size).collect();
    sizes.sort_by(f64::total_cmp);
    let n = sizes.len();
    f64::midpoint(sizes[(n - 1) / 2], sizes[n / 2])
}

/// Rotation bucket 0..=3 for a quantized rotation value.
#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)] // rot < 360 → rot/90 ∈ 0..=3
const fn rot_index(rot: u16) -> usize {
    (rot / 90 % 4) as usize
}

/// Inverse-rotate a point by `rot` (0/90/180/270) within a `w`×`h` frame so
/// that text written at `rot` reads horizontally. Pinned by
/// `rotated_page_is_canonicalized`: 90° text reads bottom-to-top, so the
/// first glyph (largest top-down y) must map to the smallest x'.
fn unrotate_point(rot: u16, x: f64, y: f64, w: f64, h: f64) -> (f64, f64) {
    match rot {
        R90 => (h - y, x),
        R180 => (w - x, h - y),
        R270 => (y, w - x),
        _ => (x, y),
    }
}

/// Frame dimensions after `unrotate_point`: quarter turns swap the axes.
const fn frame_dims(rot: u16, w: f64, h: f64) -> (f64, f64) {
    if rot == R90 || rot == R270 {
        (h, w)
    } else {
        (w, h)
    }
}

/// Rotate an axis-aligned rect by `rot` within a `w`×`h` frame (the inverse
/// of [`unrotate_point`] when `rot = 360 - applied`), re-normalizing corners.
fn rotate_rect(rect: PtRect, rot: u16, w: f64, h: f64) -> PtRect {
    if rot == 0 {
        return rect;
    }
    let (ax, ay) = unrotate_point(rot, rect.x0, rect.y0, w, h);
    let (bx, by) = unrotate_point(rot, rect.x1, rect.y1, w, h);
    PtRect {
        x0: ax.min(bx),
        y0: ay.min(by),
        x1: ax.max(bx),
        y1: ay.max(by),
    }
}

// ---------------------------------------------------------------------------
// Reading order: recursive XY-cut (liteparse `xy_cut`, reimplemented).
//
// liteparse builds a bucketed density projection per axis and searches for
// valleys; here lines are whole baseline clusters, so the same idea collapses
// to interval merging: merge each axis's coverage intervals, and the widest
// surviving gap between merged clusters is a cut no line crosses. The axis
// with the LARGER absolute gap wins (classic xy-cut: a column gutter dwarfs
// inter-line y-gaps, while a full-width spanning line blocks the vertical
// cut entirely, letting the horizontal band break fire first).
// ---------------------------------------------------------------------------

/// Column gutter: a gap in x-coverage this many times the median font size.
const XCUT_COLUMN_RATIO: f64 = 2.0;
/// Band break: a gap in y-coverage this many times the median font size
/// (half the column ratio — stacked paragraphs sit closer than columns).
const XCUT_BAND_RATIO: f64 = 1.0;
/// Recursion backstop: adversarial geometry peeling one line per cut must
/// not blow the stack (repo rule: no panics on malformed input).
const XCUT_MAX_DEPTH: u32 = 32;

/// A qualifying cut: `gap` is its absolute width (the larger-gap axis wins),
/// `at` the gap midpoint to compare line centers against.
struct Cut {
    gap: f64,
    at: f64,
    vertical: bool,
}

/// Order `lines` for reading: recursive XY-cut, columns left→right inside
/// bands top→bottom. Consumed by the Task 13 driver on the merged
/// (native + OCR) line set — allow dead code until it lands.
#[allow(dead_code)]
pub(crate) fn reading_order(lines: Vec<Line>) -> Vec<Line> {
    let mut out = Vec::with_capacity(lines.len());
    let median = median_font_size(&lines);
    xy_cut(lines, median, 0, &mut out);
    out
}

/// Median `font_size` of the line set (mean of the two middles when even).
fn median_font_size(lines: &[Line]) -> f64 {
    let mut sizes: Vec<f64> = lines.iter().map(|l| l.font_size).collect();
    sizes.sort_by(f64::total_cmp);
    if sizes.is_empty() {
        return 0.0;
    }
    let n = sizes.len();
    f64::midpoint(sizes[(n - 1) / 2], sizes[n / 2])
}

/// Append `lines` to `out` in reading order. The median is computed once at
/// the top level and passed down so nested regions use page-consistent
/// thresholds.
fn xy_cut(lines: Vec<Line>, median: f64, depth: u32, out: &mut Vec<Line>) {
    if lines.len() <= 1 || depth >= XCUT_MAX_DEPTH {
        return push_sorted(lines, out);
    }
    let Some(cut) = find_cut(&lines, median) else {
        return push_sorted(lines, out);
    };
    // Split every line by which side of the gap midpoint its rect center
    // falls on. No line can straddle the gap (a straddler would have merged
    // the coverage intervals), so both sides are non-empty and recursion
    // strictly shrinks. First = top for horizontal cuts, left for vertical.
    let (mut first, mut second) = (Vec::new(), Vec::new());
    for l in lines {
        let center = if cut.vertical {
            f64::midpoint(l.rect.x0, l.rect.x1)
        } else {
            f64::midpoint(l.rect.y0, l.rect.y1)
        };
        if center < cut.at {
            first.push(l);
        } else {
            second.push(l);
        }
    }
    xy_cut(first, median, depth + 1, out);
    xy_cut(second, median, depth + 1, out);
}

/// No qualifying cut: a leaf sorts top→bottom, then left→right.
fn push_sorted(mut lines: Vec<Line>, out: &mut Vec<Line>) {
    lines.sort_by(|a, b| {
        a.rect
            .y0
            .total_cmp(&b.rect.y0)
            .then_with(|| a.rect.x0.total_cmp(&b.rect.x0))
    });
    out.extend(lines);
}

/// Pick the cut on the axis with the larger absolute qualifying gap; `None`
/// when neither axis has one.
fn find_cut(lines: &[Line], median: f64) -> Option<Cut> {
    let v = widest_gap(lines, XCUT_COLUMN_RATIO * median, true);
    let h = widest_gap(lines, XCUT_BAND_RATIO * median, false);
    match (v, h) {
        (Some(v), Some(h)) => Some(if v.gap > h.gap { v } else { h }),
        (Some(v), None) => Some(v),
        (None, h) => h,
    }
}

/// Widest gap between consecutive merged coverage intervals along one axis,
/// if it clears `threshold`. Merging first is what makes a spanning line
/// block the cut: it overlaps both neighbors, so no gap survives between
/// them (liteparse `xy_find_best_cut`, restated as interval merging).
fn widest_gap(lines: &[Line], threshold: f64, vertical: bool) -> Option<Cut> {
    let mut ivals: Vec<(f64, f64)> = lines
        .iter()
        .map(|l| {
            if vertical {
                (l.rect.x0, l.rect.x1)
            } else {
                (l.rect.y0, l.rect.y1)
            }
        })
        .collect();
    ivals.sort_by(|a, b| a.0.total_cmp(&b.0));
    // Sole caller (`xy_cut`) guarantees ≥ 2 lines, so `ivals` is non-empty.
    let mut best: Option<Cut> = None;
    let mut end = ivals[0].1;
    for &(n0, n1) in &ivals[1..] {
        if n0 <= end {
            // Overlapping/adjacent: same coverage cluster, extend its end.
            end = end.max(n1);
            continue;
        }
        let gap = n0 - end;
        if gap >= threshold && best.as_ref().is_none_or(|b| gap > b.gap) {
            best = Some(Cut {
                gap,
                at: f64::midpoint(end, n0),
                vertical,
            });
        }
        end = n1;
    }
    best
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pdf_text::{PositionedChar, PositionedPage};

    fn pc(ch: char, x: f64, y: f64, size: f64, adv: f64) -> PositionedChar {
        PositionedChar {
            ch,
            x,
            y,
            font_size: size,
            advance: adv,
            rotation: 0,
        }
    }

    /// "Hi" at 12pt: H at x=100 (adv 8), i at x=108 (adv 3).
    fn hi_page() -> PositionedPage {
        PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars: vec![
                pc('H', 100.0, 92.0, 12.0, 8.0),
                pc('i', 108.0, 92.0, 12.0, 3.0),
            ],
        }
    }

    #[test]
    fn assembles_one_line() {
        let lines = assemble(&hi_page());
        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].text, "Hi");
        assert_eq!(lines[0].words.len(), 1);
        assert!((lines[0].rect.x0 - 100.0).abs() < 0.01);
        assert!((lines[0].font_size - 12.0).abs() < 0.01);
    }

    #[test]
    fn word_break_on_advance_gap() {
        // "Hi there": gap of 6pt (> 0.25*12=3) after "Hi".
        let mut chars = hi_page().chars;
        #[allow(clippy::cast_precision_loss)] // tiny synthetic index; exact in f64
        let mut rest = "there"
            .chars()
            .enumerate()
            .map(|(i, c)| pc(c, 117.0 + i as f64 * 6.0, 92.0, 12.0, 5.5))
            .collect::<Vec<_>>();
        chars.append(&mut rest);
        let lines = assemble(&PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars,
        });
        assert_eq!(lines[0].text, "Hi there");
        assert_eq!(lines[0].words.len(), 2);
    }

    #[test]
    fn explicit_space_glyph_does_not_double() {
        let chars = vec![
            pc('a', 100.0, 92.0, 12.0, 6.0),
            pc(' ', 106.0, 92.0, 12.0, 3.0),
            pc('b', 109.0, 92.0, 12.0, 6.0),
        ];
        let lines = assemble(&PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars,
        });
        assert_eq!(lines[0].text, "a b");
    }

    #[test]
    fn two_baselines_make_two_lines() {
        let mut chars = hi_page().chars;
        chars.push(pc('B', 100.0, 106.0, 12.0, 7.0));
        chars.push(pc('y', 107.0, 106.0, 12.0, 6.0));
        chars.push(pc('e', 113.0, 106.0, 12.0, 6.0));
        let lines = assemble(&PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars,
        });
        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].text, "Hi");
        assert_eq!(lines[1].text, "Bye");
    }

    #[test]
    fn rotated_page_is_canonicalized() {
        // 4 of 5 chars at rotation 90 → page canonicalized; the line reads
        // in ascending original-y order. "abc" written bottom-to-top:
        // a at (50, 300), b at (50, 290), c at (50, 280), rotation 90.
        let chars = vec![
            PositionedChar {
                ch: 'a',
                x: 50.0,
                y: 300.0,
                font_size: 12.0,
                advance: 7.0,
                rotation: 90,
            },
            PositionedChar {
                ch: 'b',
                x: 50.0,
                y: 290.0,
                font_size: 12.0,
                advance: 7.0,
                rotation: 90,
            },
            PositionedChar {
                ch: 'c',
                x: 50.0,
                y: 280.0,
                font_size: 12.0,
                advance: 7.0,
                rotation: 90,
            },
        ];
        let lines = assemble(&PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars,
        });
        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].text, "abc");
    }

    #[test]
    fn wide_horizontal_gap_splits_line() {
        // "L1" at x=72 and "R1" at x=320 on the same baseline → two lines
        // (gap 234pt > 1.5*12), which is what lets xy-cut see columns.
        let chars = vec![
            pc('L', 72.0, 92.0, 12.0, 8.0),
            pc('1', 80.0, 92.0, 12.0, 6.0),
            pc('R', 320.0, 92.0, 12.0, 8.0),
            pc('1', 328.0, 92.0, 12.0, 6.0),
        ];
        let lines = assemble(&PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars,
        });
        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].text, "L1");
        assert_eq!(lines[1].text, "R1");
    }

    fn line(text: &str, x0: f64, y0: f64, x1: f64) -> Line {
        Line {
            text: text.into(),
            words: vec![],
            rect: PtRect {
                x0,
                y0,
                x1,
                y1: y0 + 12.0,
            },
            font_size: 12.0,
            source: LineSource::Native,
        }
    }

    fn texts(lines: &[Line]) -> Vec<&str> {
        lines.iter().map(|l| l.text.as_str()).collect()
    }

    #[test]
    fn single_column_sorts_by_y() {
        let lines = vec![
            line("second", 72.0, 100.0, 200.0),
            line("first", 72.0, 80.0, 200.0),
        ];
        assert_eq!(texts(&reading_order(lines)), ["first", "second"]);
    }

    #[test]
    fn two_columns_left_then_right() {
        // Left column x 72..250, right column x 320..500 (gap 70pt >= 2*12).
        let lines = vec![
            line("R1", 320.0, 80.0, 500.0),
            line("L2", 72.0, 100.0, 250.0),
            line("L1", 72.0, 80.0, 250.0),
            line("R2", 320.0, 100.0, 500.0),
        ];
        assert_eq!(texts(&reading_order(lines)), ["L1", "L2", "R1", "R2"]);
    }

    #[test]
    fn full_width_title_then_columns() {
        let lines = vec![
            line("R1", 320.0, 100.0, 500.0),
            line("TITLE", 72.0, 60.0, 500.0), // spans both columns
            line("L1", 72.0, 100.0, 250.0),
        ];
        assert_eq!(texts(&reading_order(lines)), ["TITLE", "L1", "R1"]);
    }

    #[test]
    fn narrow_gap_is_not_a_column_cut() {
        // 15pt gap < 2*12=24pt threshold → one column, y-sorted.
        let lines = vec![line("b", 200.0, 90.0, 280.0), line("a", 72.0, 80.0, 185.0)];
        assert_eq!(texts(&reading_order(lines)), ["a", "b"]);
    }

    #[test]
    fn empty_page_assembles_empty() {
        let lines = assemble(&PositionedPage {
            page_num: 1,
            media_box: (0.0, 0.0, 612.0, 792.0),
            chars: vec![],
        });
        assert!(lines.is_empty());
    }
}
