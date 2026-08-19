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
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum LineSource {
    Native,
    Ocr,
}

/// A word with its x-extent, rebuilt from per-char advance gaps (spec §7 —
/// NOT naive space splitting).
///
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Word {
    pub text: String,
    pub x0: f64,
    pub x1: f64,
}

/// A baseline cluster of chars assembled into text, with the union of its
/// char boxes in top-down page space.
///
#[derive(Clone, Debug)]
pub(crate) struct Line {
    pub text: String,
    /// Per-word x-extents in the line's ASSEMBLY frame (see the module
    /// docs): the table detector (Task 14) clusters word x0s across lines
    /// of a run — all canonical-group lines share one frame, so that is
    /// safe — but a word's x0/x1 must never be compared against its own
    /// line's page-space `rect`.
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
/// Raised from 1.5 to 10 so that wide table gutters stay on one line:
/// a 12 pt table with 96 pt between columns must not shatter into one-word
/// lines, otherwise table detection sees no multi-word candidates.
const LINE_SPLIT_RATIO: f64 = 10.0;
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
/// bands top→bottom. Runs on the merged (native + OCR) line set.
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

// ---------------------------------------------------------------------------
// Document-level signals + heading/paragraph classification
// (liteparse `markdown_layout/{headings,repetition,classify}`, minimal port).
//
// Pass 1 of the Task 13 driver feeds each page's assembled lines to
// [`DocSignalsBuilder`] and drops the page; pass 2 runs [`classify`] over the
// merged, reading-ordered line set. liteparse gates headings on `size > body
// + 0.5pt` and derives levels from a doc-wide rank of distinct heading sizes
// (its headings.rs:578/:644) — the plan deliberately trimmed that to a pure
// size-ratio scheme so signals stay three small fields (spec §7).
// ---------------------------------------------------------------------------

/// Heading gate: a line at least this much larger than the body size.
const HEADING_MIN_RATIO: f64 = 1.2;
/// Size-ratio bands → heading level: ≥1.8 → H1, ≥1.5 → H2, ≥1.3 → H3, and
/// anything still over [`HEADING_MIN_RATIO`] → H4.
const HEADING_H1_RATIO: f64 = 1.8;
const HEADING_H2_RATIO: f64 = 1.5;
const HEADING_H3_RATIO: f64 = 1.3;
/// Paragraph split: a top-to-top vertical gap larger than this many times
/// the previous line's font size starts a new paragraph.
const PARAGRAPH_GAP_RATIO: f64 = 1.5;
/// Body size when the document yields no measurable text (image-only PDF).
const DEFAULT_BODY_SIZE: f64 = 12.0;
/// A normalized first/last-line signature must repeat on at least this many
/// pages to count as a header/footer…
const HEADER_FOOTER_MIN_PAGES: usize = 2;
/// …or on one page in `HEADER_FOOTER_PAGE_DIV`, whichever is larger.
const HEADER_FOOTER_PAGE_DIV: usize = 3;

/// Document-wide aggregates collected in driver pass 1 (tiny — no pages
/// retained, spec §4.2/§11).
#[derive(Debug)]
pub(crate) struct DocSignals {
    /// Modal font size across the document (quarter-point buckets) — the
    /// body text size.
    pub body_size: f64,
    /// Normalized signatures of lines repeated as a page's first line.
    pub headers: std::collections::HashSet<String>,
    /// Same for last lines (footers).
    pub footers: std::collections::HashSet<String>,
}

/// Pass-1 accumulator for [`DocSignals`].
#[derive(Default)]
pub(crate) struct DocSignalsBuilder {
    /// Quarter-point bucket key → summed text length at that size.
    size_weights: std::collections::HashMap<u64, usize>,
    /// Normalized first-line signature → pages seen on.
    first_counts: std::collections::HashMap<String, usize>,
    /// Normalized last-line signature → pages seen on.
    last_counts: std::collections::HashMap<String, usize>,
}

/// A classified block of body content. The List variant arrives in
/// Tasks 15–16.
#[derive(Debug)]
pub(crate) enum Block {
    Heading {
        level: u8,
        text: String,
    },
    Paragraph(String),
    /// Borderless table (Task 14): rows of cells, empty string = no word
    /// landed in that column for that row.
    Table(Vec<Vec<String>>),
}

impl DocSignalsBuilder {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    /// Fold one page's assembled lines into the aggregates. Lines are
    /// unordered here: the page's "first line" is the one with min
    /// `rect.y0`, its "last line" the one with max `rect.y0`.
    /// Whitespace-only lines carry no signal and are skipped entirely.
    pub(crate) fn add_lines(&mut self, lines: &[Line]) {
        let mut first: Option<&Line> = None;
        let mut last: Option<&Line> = None;
        for l in lines {
            let text = l.text.trim();
            if text.is_empty() {
                continue;
            }
            if l.font_size > 0.0 {
                *self
                    .size_weights
                    .entry(size_bucket(l.font_size))
                    .or_insert(0) += text.chars().count();
            }
            if first.is_none_or(|f| l.rect.y0 < f.rect.y0) {
                first = Some(l);
            }
            if last.is_none_or(|f| l.rect.y0 > f.rect.y0) {
                last = Some(l);
            }
        }
        if let Some(f) = first {
            *self
                .first_counts
                .entry(normalize_signature(&f.text))
                .or_insert(0) += 1;
        }
        if let Some(l) = last {
            *self
                .last_counts
                .entry(normalize_signature(&l.text))
                .or_insert(0) += 1;
        }
    }

    /// Emit the aggregates. `body_size` is the bucket with the most text
    /// (ties go to the larger size — the body is never smaller than its own
    /// footnotes); signatures under the repetition threshold are dropped.
    pub(crate) fn finish(self, page_count: usize) -> DocSignals {
        let body_size = self
            .size_weights
            .iter()
            .max_by(|a, b| a.1.cmp(b.1).then_with(|| a.0.cmp(b.0)))
            .map_or(DEFAULT_BODY_SIZE, |(&key, _)| bucket_size(key));
        let threshold = HEADER_FOOTER_MIN_PAGES.max(page_count / HEADER_FOOTER_PAGE_DIV);
        let collect = |counts: std::collections::HashMap<String, usize>| {
            counts
                .into_iter()
                .filter(|(_, n)| *n >= threshold)
                .map(|(sig, _)| sig)
                .collect()
        };
        DocSignals {
            body_size,
            headers: collect(self.first_counts),
            footers: collect(self.last_counts),
        }
    }
}

/// Quarter-point histogram bucket for a font size. Positive sizes only —
/// callers skip `<= 0.0` (a malformed-PDF guard, not a real case).
#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)] // size > 0 and ≤ ~1e3pt
fn size_bucket(size: f64) -> u64 {
    (size * 4.0).round() as u64
}

/// Inverse of [`size_bucket`].
#[allow(clippy::cast_precision_loss)] // bucket keys are small integers
fn bucket_size(key: u64) -> f64 {
    key as f64 / 4.0
}

/// Normalize a line for header/footer matching: lowercase, every RUN of
/// ASCII digits collapses to one `#`, whitespace runs collapse to one space
/// (liteparse `normalize_for_repetition`). Run-collapse — not per-digit —
/// so "Page 3 of 9" and "Page 12 of 9" normalize together.
fn normalize_signature(text: &str) -> String {
    let mut out = String::new();
    let mut in_digits = false;
    let mut pending_space = false;
    for c in text.trim().chars().flat_map(char::to_lowercase) {
        if c.is_whitespace() {
            pending_space = !out.is_empty();
            in_digits = false;
            continue;
        }
        if pending_space {
            out.push(' ');
            pending_space = false;
        }
        if c.is_ascii_digit() {
            if !in_digits {
                out.push('#');
                in_digits = true;
            }
        } else {
            out.push(c);
            in_digits = false;
        }
    }
    out
}

/// Open paragraph under construction: text plus the last line's font
/// size and top edge (the gap rule needs both).
struct Para {
    text: String,
    font_size: f64,
    y0: f64,
}

/// Incremental heading/paragraph classifier: [`detect_tables`] feeds it
/// the non-table lines interleaved with detected table blocks, so the
/// paragraph state must survive across calls. [`classify`] is the thin
/// all-lines wrapper; the heading/paragraph rules live here exactly once.
struct Classifier<'a> {
    signals: &'a DocSignals,
    blocks: Vec<Block>,
    para: Option<Para>,
}

impl Classifier<'_> {
    const fn new(signals: &DocSignals) -> Classifier<'_> {
        Classifier {
            signals,
            blocks: Vec::new(),
            para: None,
        }
    }

    fn flush(&mut self) {
        if let Some(p) = self.para.take() {
            self.blocks.push(Block::Paragraph(p.text));
        }
    }

    /// Fold one line into the block stream: header/footer drop, heading
    /// gate, then paragraph merge/split — see [`classify`] for the rules.
    fn feed(&mut self, line: &Line) {
        let text = line.text.trim();
        if text.is_empty() {
            return;
        }
        let sig = normalize_signature(text);
        if self.signals.headers.contains(&sig) || self.signals.footers.contains(&sig) {
            self.flush();
            return;
        }
        let ratio = if self.signals.body_size > 0.0 {
            line.font_size / self.signals.body_size
        } else {
            0.0
        };
        if ratio >= HEADING_MIN_RATIO {
            self.flush();
            let level = if ratio >= HEADING_H1_RATIO {
                1
            } else if ratio >= HEADING_H2_RATIO {
                2
            } else if ratio >= HEADING_H3_RATIO {
                3
            } else {
                4
            };
            self.blocks.push(Block::Heading {
                level,
                text: text.to_string(),
            });
            return;
        }
        match self.para.as_mut() {
            Some(p) if line.rect.y0 - p.y0 <= PARAGRAPH_GAP_RATIO * p.font_size => {
                if p.text.ends_with('-') {
                    p.text.pop();
                } else {
                    p.text.push(' ');
                }
                p.text.push_str(text);
                p.font_size = line.font_size;
                p.y0 = line.rect.y0;
            }
            _ => {
                self.flush();
                self.para = Some(Para {
                    text: text.to_string(),
                    font_size: line.font_size,
                    y0: line.rect.y0,
                });
            }
        }
    }

    fn finish(mut self) -> Vec<Block> {
        self.flush();
        self.blocks
    }
}

/// Classify reading-ordered `lines` (native + OCR merged — OCR lines
/// participate exactly like native ones) into blocks:
///
/// 1. Lines whose normalized text is a known header/footer are dropped
///    (and break any open paragraph — chrome never glues text together).
/// 2. A line at ≥ [`HEADING_MIN_RATIO`]× the body size is a heading; the
///    ratio band picks the level.
/// 3. Remaining consecutive lines merge into paragraphs, joined with a
///    space — except a trailing `-` dehyphenates ("exam-" + "ple" →
///    "example"). A top-to-top gap > [`PARAGRAPH_GAP_RATIO`]× the previous
///    line's font size splits the paragraph.
///
/// Whitespace-only lines (all-space clusters) are dropped, never emitted
/// as empty paragraphs.
// Pass-by-value is the established Task 13 signature (mirrors the driver,
// which hands over the owned line vec); the lines are only borrowed now
// that classification is incremental.
#[allow(dead_code, clippy::needless_pass_by_value)]
pub(crate) fn classify(lines: Vec<Line>, signals: &DocSignals) -> Vec<Block> {
    let mut c = Classifier::new(signals);
    for line in &lines {
        c.feed(line);
    }
    c.finish()
}

// ---------------------------------------------------------------------------
// Borderless table detection (spec §7 v1.5; liteparse `try_detect_table*`,
// adapted: no ruled-line graphics, words rebuilt from advance gaps, word x0s
// clustered across the run instead of PDFium span pieces).
// ---------------------------------------------------------------------------

/// A table needs at least this many consecutive candidate lines.
const TABLE_MIN_ROWS: usize = 3;
/// …and each candidate line needs at least this many words (single-word
/// lines break runs — they carry no column signal).
const TABLE_MIN_WORDS: usize = 2;
/// Candidate lines sit within this fraction of the body size (heading-size
/// lines break runs).
const TABLE_BODY_SIZE_TOLERANCE: f64 = 0.2;
/// Run continuity: the bottom-to-top vertical gap between consecutive
/// candidate lines stays within this many times the previous line's font
/// size (liteparse `table_rows_adjacent`, restated in `Line` geometry).
const TABLE_ROW_GAP_RATIO: f64 = 2.5;
/// Column anchors: word x0s within this many times the run's font size
/// belong to the same column.
const TABLE_COL_TOLERANCE_RATIO: f64 = 0.5;
/// Gutter check: between consecutive anchor columns a line occupies, the
/// whitespace gap must reach this many times the line's font size. Prose
/// word spacing (~0.25em) fails; table gutters pass.
const TABLE_GUTTER_RATIO: f64 = 0.75;
/// Occupancy (liteparse `count_text_table_runs` criterion): ≥60% of a
/// line's words start at an anchor, and each anchor column is occupied by
/// ≥60% of the run's lines. Exact 3/5 comparisons, so a 2/3 sparse row
/// (67%) passes.
const TABLE_OCCUPANCY_NUM: usize = 3;
const TABLE_OCCUPANCY_DEN: usize = 5;
/// Fewer than two columns is a definition list, not a table.
const TABLE_MIN_COLUMNS: usize = 2;

/// Detect borderless tables in reading-ordered `lines` and classify the
/// rest: a run of [`TABLE_MIN_ROWS`]+ consecutive body-size lines whose
/// word x0s cluster into ≥ [`TABLE_MIN_COLUMNS`] anchor columns — with a
/// real gutter between occupied column pairs and 60% line/column
/// occupancy — becomes [`Block::Table`]; every other line goes through the
/// same heading/paragraph rules as [`classify`]. Runs classification
/// internally and replaces the bare `classify` call in the driver
/// (Task 15 wires it in).
///
/// OCR-source lines carry no word geometry (`words` is empty), so they
/// fail the word-count gate and never join a run.
// Pass-by-value is the binding Task 14 signature (mirrors `classify`).
#[allow(clippy::needless_pass_by_value)]
pub(crate) fn detect_tables(lines: Vec<Line>, signals: &DocSignals) -> Vec<Block> {
    let mut c = Classifier::new(signals);
    let mut i = 0;
    while i < lines.len() {
        if let Some((rows, end)) = try_table_run(&lines, i, signals) {
            c.flush();
            c.blocks.push(Block::Table(rows));
            i = end;
        } else {
            c.feed(&lines[i]);
            i += 1;
        }
    }
    c.finish()
}

/// If a table run starts at `lines[start]`, return its rows of cells and
/// the index one past the run's last line; `None` falls through to normal
/// classification of that one line (the walk retries at `start + 1`, so a
/// table whose run begins deeper inside a candidate band is still found).
fn try_table_run(
    lines: &[Line],
    start: usize,
    signals: &DocSignals,
) -> Option<(Vec<Vec<String>>, usize)> {
    if !table_candidate(&lines[start], signals) {
        return None;
    }
    let mut end = start + 1;
    while end < lines.len()
        && table_candidate(&lines[end], signals)
        && lines[end].rect.y0 - lines[end - 1].rect.y1
            <= TABLE_ROW_GAP_RATIO * lines[end - 1].font_size
    {
        end += 1;
    }
    let run = &lines[start..end];
    if run.len() < TABLE_MIN_ROWS {
        return None;
    }
    // All run lines are within 20% of the body size, so the first line's
    // size is a representative scale for the anchor tolerance.
    let tol = TABLE_COL_TOLERANCE_RATIO * run[0].font_size;
    let anchors = column_anchors(run, tol);
    if anchors.len() < TABLE_MIN_COLUMNS {
        return None;
    }
    // Word → column assignment: nearest anchor within tolerance; words
    // landing between columns (unassigned) count against line occupancy.
    let assign = |line: &Line| -> Vec<Option<usize>> {
        line.words
            .iter()
            .map(|w| {
                anchors
                    .iter()
                    .enumerate()
                    .filter(|(_, a)| (w.x0 - *a).abs() <= tol)
                    .min_by(|(ai, a), (bi, b)| {
                        (w.x0 - *a)
                            .abs()
                            .total_cmp(&(w.x0 - *b).abs())
                            .then_with(|| ai.cmp(bi))
                    })
                    .map(|(i, _)| i)
            })
            .collect()
    };
    let assigned: Vec<Vec<Option<usize>>> = run.iter().map(assign).collect();
    // Gutter check: for every pair of consecutive anchor columns a line
    // occupies, the whitespace between the right edge of the left cell's
    // last word and the left edge of the right cell's first word must be
    // a real gutter, not a prose word space.
    for (line, cols) in run.iter().zip(&assigned) {
        let gutter = TABLE_GUTTER_RATIO * line.font_size;
        for pair in 0..anchors.len() - 1 {
            let left = line
                .words
                .iter()
                .zip(cols)
                .filter(|(_, c)| **c == Some(pair))
                .map(|(w, _)| w.x1)
                .fold(f64::NEG_INFINITY, f64::max);
            let right = line
                .words
                .iter()
                .zip(cols)
                .filter(|(_, c)| **c == Some(pair + 1))
                .map(|(w, _)| w.x0)
                .fold(f64::INFINITY, f64::min);
            if left.is_finite() && right.is_finite() && right - left < gutter {
                return None;
            }
        }
    }
    // Line occupancy: ≥60% of the line's words start at an anchor.
    for (line, cols) in run.iter().zip(&assigned) {
        let on_anchor = cols.iter().filter(|c| c.is_some()).count();
        if on_anchor * TABLE_OCCUPANCY_DEN < line.words.len() * TABLE_OCCUPANCY_NUM {
            return None;
        }
    }
    // Column occupancy: each anchor is used by ≥60% of the run's lines.
    for col in 0..anchors.len() {
        let occupied = assigned
            .iter()
            .filter(|cols| cols.contains(&Some(col)))
            .count();
        if occupied * TABLE_OCCUPANCY_DEN < run.len() * TABLE_OCCUPANCY_NUM {
            return None;
        }
    }
    // Cell text: words assigned to a column join with spaces; a column
    // with no word in a line yields an empty cell.
    let rows = run
        .iter()
        .zip(&assigned)
        .map(|(line, cols)| {
            let mut cells = vec![String::new(); anchors.len()];
            for (w, c) in line.words.iter().zip(cols) {
                if let Some(col) = c {
                    if !cells[*col].is_empty() {
                        cells[*col].push(' ');
                    }
                    cells[*col].push_str(&w.text);
                }
            }
            cells
        })
        .collect();
    Some((rows, end))
}

/// A line that can join a table run: enough words to carry column signal
/// (OCR lines have none and are excluded here) and a font size within
/// [`TABLE_BODY_SIZE_TOLERANCE`] of the body size.
fn table_candidate(line: &Line, signals: &DocSignals) -> bool {
    if line.words.len() < TABLE_MIN_WORDS {
        return false;
    }
    signals.body_size > 0.0
        && (line.font_size - signals.body_size).abs()
            <= TABLE_BODY_SIZE_TOLERANCE * signals.body_size
}

/// Cluster the run's word x0 positions into sorted column anchors:
/// positions within `tol` of a cluster's running mean merge into it (the
/// mean keeps the anchor centered as left/right-aligned cells drift in).
#[allow(clippy::cast_precision_loss)] // cluster sizes are tiny; exact in f64
fn column_anchors(run: &[Line], tol: f64) -> Vec<f64> {
    let mut xs: Vec<f64> = run
        .iter()
        .flat_map(|l| l.words.iter().map(|w| w.x0))
        .collect();
    xs.sort_by(f64::total_cmp);
    // (sum, count) per open cluster; anchors are the final means.
    let mut clusters: Vec<(f64, usize)> = Vec::new();
    for x in xs {
        match clusters.last_mut() {
            Some((sum, n)) if (x - *sum / *n as f64).abs() <= tol => {
                *sum += x;
                *n += 1;
            }
            _ => clusters.push((x, 1)),
        }
    }
    clusters.iter().map(|(sum, n)| sum / *n as f64).collect()
}

// ---------------------------------------------------------------------------
// Markdown render (Task 13 driver output).
// ---------------------------------------------------------------------------

/// Render classified blocks to markdown: blocks separated by exactly one
/// blank line, document ends with a single trailing newline.
pub(crate) fn render(
    blocks: &[Block],
    sink: &mut impl crate::ExtractSink,
) -> crate::error::Result<()> {
    for (i, block) in blocks.iter().enumerate() {
        if i > 0 {
            sink.write_str("\n")?;
        }
        match block {
            Block::Heading { level, text } => {
                sink.write_str(&"#".repeat(usize::from(*level)))?;
                sink.write_str(" ")?;
                sink.write_str(text)?;
                sink.write_str("\n")?;
            }
            Block::Paragraph(text) => {
                sink.write_str(text)?;
                sink.write_str("\n")?;
            }
            Block::Table(rows) => {
                render_table(rows, sink)?;
            }
        }
    }
    Ok(())
}

/// Escape a table cell for pipe-table markdown: backslashes first, then
/// pipes, so we don't double-escape the separator we just inserted.
fn escape_cell(s: &str) -> String {
    s.replace('\\', "\\\\").replace('|', "\\|")
}

/// Render a pipe table: header row, `| --- |` separator per column, then
/// body rows. The last row ends with a newline; the surrounding render loop
/// supplies the blank-line separation between blocks.
fn render_table(
    rows: &[Vec<String>],
    sink: &mut impl crate::ExtractSink,
) -> crate::error::Result<()> {
    let Some(header) = rows.first() else {
        return Ok(());
    };
    render_table_row(header, sink)?;
    for _ in header {
        sink.write_str("| --- ")?;
    }
    sink.write_str("|\n")?;
    for row in rows.iter().skip(1) {
        render_table_row(row, sink)?;
    }
    Ok(())
}

fn render_table_row(
    row: &[String],
    sink: &mut impl crate::ExtractSink,
) -> crate::error::Result<()> {
    for cell in row {
        sink.write_str("| ")?;
        sink.write_str(&escape_cell(cell))?;
        sink.write_str(" ")?;
    }
    sink.write_str("|\n")?;
    Ok(())
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

    fn sig_lines(pages: &[Vec<Line>]) -> DocSignals {
        let mut b = DocSignalsBuilder::new();
        for p in pages {
            b.add_lines(p);
        }
        b.finish(pages.len())
    }

    #[test]
    fn body_size_is_modal() {
        let pages = vec![vec![
            line("body text here", 72.0, 100.0, 300.0), // 12pt (helper default)
            Line {
                font_size: 18.0,
                ..line("H", 72.0, 60.0, 100.0)
            },
        ]];
        assert!((sig_lines(&pages).body_size - 12.0).abs() < 0.01);
    }

    #[test]
    fn repeated_first_line_becomes_header() {
        // The trailing "common closer" line is deliberate: without it the
        // digit-normalized body line is every page's LAST line and would
        // itself qualify as a footer (first/last-line signatures have no
        // page-band gating — see task-12 report). It doubles as a footer
        // fixture: classify must drop it too.
        let mk = |n: usize| {
            vec![
                line("CONFIDENTIAL", 72.0, 30.0, 200.0),
                line(&format!("body {n}"), 72.0, 100.0, 300.0),
                line("common closer", 72.0, 700.0, 300.0),
            ]
        };
        let pages: Vec<Vec<Line>> = (0..4).map(mk).collect();
        let sig = sig_lines(&pages);
        assert!(sig.headers.contains("confidential"));
        let blocks = classify(pages[0].clone(), &sig);
        assert_eq!(blocks.len(), 1);
        assert!(matches!(&blocks[0], Block::Paragraph(t) if t == "body 0"));
    }

    #[test]
    fn page_numbers_in_footers_normalize_together() {
        let mk = |n: usize| {
            vec![
                line("body", 72.0, 100.0, 300.0),
                line(&format!("Page {n} of 9"), 250.0, 750.0, 350.0),
            ]
        };
        let pages: Vec<Vec<Line>> = (1..=3).map(mk).collect();
        let sig = sig_lines(&pages);
        assert!(sig.footers.contains("page # of #"), "{:?}", sig.footers);
    }

    #[test]
    fn heading_levels_from_size_ratio() {
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let lines = vec![
            Line {
                font_size: 24.0,
                ..line("Big", 72.0, 50.0, 200.0)
            }, // 2.0x → H1
            Line {
                font_size: 18.0,
                ..line("Mid", 72.0, 90.0, 200.0)
            }, // 1.5x → H2
            Line {
                font_size: 14.5,
                ..line("Small", 72.0, 120.0, 200.0)
            }, // 1.21x → H4
            line("body", 72.0, 150.0, 300.0),
        ];
        let blocks = classify(lines, &sig);
        assert!(matches!(&blocks[0], Block::Heading { level: 1, text } if text == "Big"));
        assert!(matches!(&blocks[1], Block::Heading { level: 2, .. }));
        assert!(matches!(&blocks[2], Block::Heading { level: 4, .. }));
        assert!(matches!(&blocks[3], Block::Paragraph(t) if t == "body"));
    }

    #[test]
    fn paragraph_join_and_dehyphenation() {
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let lines = vec![
            line("This is an exam-", 72.0, 100.0, 300.0),
            line("ple of joining.", 72.0, 114.0, 300.0),
            line("New para.", 72.0, 140.0, 300.0), // 26pt gap > 1.5*12 → split
        ];
        let blocks = classify(lines, &sig);
        assert_eq!(blocks.len(), 2);
        assert!(matches!(&blocks[0], Block::Paragraph(t) if t == "This is an example of joining."));
        assert!(matches!(&blocks[1], Block::Paragraph(t) if t == "New para."));
    }

    #[test]
    fn whitespace_only_lines_are_dropped() {
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let lines = vec![
            line("real text", 72.0, 100.0, 300.0),
            Line {
                text: "   ".into(),
                ..line("", 72.0, 114.0, 300.0)
            },
        ];
        let blocks = classify(lines, &sig);
        assert_eq!(blocks.len(), 1);
        assert!(matches!(&blocks[0], Block::Paragraph(t) if t == "real text"));
    }

    /// Build a word-ful Line (detect_tables reads words, not just text).
    #[allow(clippy::cast_precision_loss)] // tiny synthetic word lengths; exact in f64
    fn wline(words: &[&str], x0s: &[f64], y0: f64) -> Line {
        assert_eq!(words.len(), x0s.len());
        let words_v: Vec<Word> = words
            .iter()
            .zip(x0s)
            .map(|(t, x)| Word {
                text: t.to_string(),
                x0: *x,
                x1: *x + t.len() as f64 * 6.0,
            })
            .collect();
        let text = words.join(" ");
        Line {
            text,
            rect: PtRect {
                x0: x0s[0],
                y0,
                x1: 400.0,
                y1: y0 + 12.0,
            },
            font_size: 12.0,
            source: LineSource::Native,
            words: words_v,
        }
    }

    #[test]
    fn detects_simple_grid() {
        let lines = vec![
            wline(&["Name", "Age", "City"], &[72.0, 200.0, 300.0], 100.0),
            wline(&["Alice", "30", "NYC"], &[72.0, 200.0, 300.0], 114.0),
            wline(&["Bob", "25", "LA"], &[72.0, 200.0, 300.0], 128.0),
        ];
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let blocks = detect_tables(lines, &sig);
        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            Block::Table(rows) => {
                assert_eq!(rows.len(), 3);
                assert_eq!(rows[0], vec!["Name", "Age", "City"]);
                assert_eq!(rows[2], vec!["Bob", "25", "LA"]);
            }
            other => panic!("expected table, got {other:?}"),
        }
    }

    #[test]
    fn prose_is_not_a_table() {
        // Realistic prose spacing (6pt/char, ~3pt word gaps at 12pt):
        // aligned-looking x0s but gutters of 3pt < 0.75*12=9pt → rejected
        // by the gutter check.
        let lines = vec![
            wline(&["The", "quick", "brown"], &[72.0, 93.0, 126.0], 100.0),
            wline(&["fox", "jumps", "over"], &[72.0, 93.0, 126.0], 114.0),
            wline(&["the", "lazy", "dog"], &[72.0, 93.0, 126.0], 128.0),
        ];
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let blocks = detect_tables(lines, &sig);
        assert!(blocks.iter().all(|b| matches!(b, Block::Paragraph(_))));
    }

    #[test]
    fn short_run_is_not_a_table() {
        // Only 2 aligned rows — below the 3-line minimum.
        let lines = vec![
            wline(&["a", "b"], &[72.0, 200.0], 100.0),
            wline(&["c", "d"], &[72.0, 200.0], 114.0),
        ];
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let blocks = detect_tables(lines, &sig);
        assert!(blocks.iter().all(|b| matches!(b, Block::Paragraph(_))));
    }

    #[test]
    fn missing_cell_yields_empty_string() {
        let lines = vec![
            wline(&["Name", "Age", "City"], &[72.0, 200.0, 300.0], 100.0),
            wline(&["Alice", "NYC"], &[72.0, 300.0], 114.0), // no Age cell
            wline(&["Bob", "25", "LA"], &[72.0, 200.0, 300.0], 128.0),
        ];
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let blocks = detect_tables(lines, &sig);
        match &blocks[0] {
            Block::Table(rows) => assert_eq!(rows[1], vec!["Alice", "", "NYC"]),
            other => panic!("expected table, got {other:?}"),
        }
    }

    #[test]
    fn ocr_lines_join_paragraphs() {
        let sig = DocSignals {
            body_size: 12.0,
            headers: Default::default(),
            footers: Default::default(),
        };
        let ocr = Line {
            source: LineSource::Ocr,
            ..line("from ocr", 72.0, 114.0, 300.0)
        };
        let lines = vec![line("native", 72.0, 100.0, 300.0), ocr];
        let blocks = classify(lines, &sig);
        assert_eq!(blocks.len(), 1);
        assert!(matches!(&blocks[0], Block::Paragraph(t) if t == "native from ocr"));
    }

    #[test]
    fn render_table_block() {
        let blocks = vec![
            Block::Paragraph("Before.".into()),
            Block::Table(vec![
                vec!["Name".into(), "Age".into()],
                vec!["Alice".into(), "30".into()],
            ]),
            Block::Paragraph("After.".into()),
        ];
        let mut out = String::new();
        render(&blocks, &mut out).unwrap();
        assert_eq!(
            out,
            "Before.\n\n| Name | Age |\n| --- | --- |\n| Alice | 30 |\n\nAfter.\n"
        );
    }

    #[test]
    fn render_table_escapes_pipes() {
        let blocks = vec![Block::Table(vec![vec!["a|b".into()], vec!["c".into()]])];
        let mut out = String::new();
        render(&blocks, &mut out).unwrap();
        assert_eq!(out, "| a\\|b |\n| --- |\n| c |\n");
    }
}
