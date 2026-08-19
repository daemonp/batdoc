//! Reverse cmap: glyph-id -> unicode, parsed from an embedded font
//! program's `cmap` table via ttf-parser. Used to recover text from
//! CID-keyed fonts whose /ToUnicode is missing or garbage.
//!
//! Algorithm derived from run-llama/liteparse `font_cmap.rs`
//! (Apache-2.0); reimplemented — no code copied.
//!
//! Preference order (liteparse): UCS-4 subtables (format 12, platform 0/4
//! or 3/10) over BMP (format 4, platform 0/3 or 3/1); later mappings for a
//! glyph do not overwrite earlier ones; Private Use Area codepoints
//! (U+E000..=U+F8FF) are rejected as they are never real text.

use std::collections::HashMap;

#[derive(Clone)]
pub(crate) struct ReverseCmap {
    map: HashMap<u32, char>, // glyph id -> unicode
}

impl ReverseCmap {
    /// Parse `data` as an sfnt and invert its unicode cmap subtables.
    /// `None` if the data isn't a parseable font or has no usable cmap.
    pub(crate) fn from_font_program(data: &[u8]) -> Option<ReverseCmap> {
        let face = ttf_parser::Face::parse(data, 0).ok()?;
        let cmap = face.tables().cmap?;
        // Unicode subtables, best-first: format 12 (`SegmentedCoverage`,
        // full UCS-4) before BMP formats (4/0/6).
        let mut subs: Vec<_> = cmap
            .subtables
            .into_iter()
            .filter(|s| s.is_unicode())
            .collect();
        subs.sort_by_key(|s| {
            u8::from(!matches!(
                s.format,
                ttf_parser::cmap::Format::SegmentedCoverage(_)
            ))
        });
        let mut map = HashMap::new();
        for sub in subs {
            sub.codepoints(|cp| {
                // PUA codepoints are never real text; skip.
                if (0xE000..=0xF8FF).contains(&cp) {
                    return;
                }
                let Some(ch) = char::from_u32(cp) else { return };
                if let Some(glyph) = sub.glyph_index(cp) {
                    // First mapping for a glyph wins (UCS-4 subtable ran
                    // first); glyph 0 already surfaces as None.
                    map.entry(u32::from(glyph.0)).or_insert(ch);
                }
            });
        }
        if map.is_empty() {
            None
        } else {
            Some(ReverseCmap { map })
        }
    }

    pub(crate) fn lookup(&self, glyph_id: u32) -> Option<char> {
        self.map.get(&glyph_id).copied()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal synthetic sfnt: a `cmap` table with one format-4
    /// subtable mapping A(0x41)->glyph 1, B(0x42)->glyph 2 via an idDelta
    /// segment [0x41..=0x42], delta = -0x40, plus the required 0xFFFF
    /// terminator segment. ttf-parser's `Face::parse` hard-requires
    /// `head`/`hhea`/`maxp`, so those are included with minimal valid
    /// contents. Kept synthetic: no checked-in binary fixtures.
    fn tiny_sfnt() -> Vec<u8> {
        // cmap format 4 layout: https://learn.microsoft.com/en-us/typography/opentype/spec/cmap#format-4
        let mut sub = Vec::new();
        let seg_count = 2_u16; // [0x41..0x42], [0xFFFF]
        let seg_x2 = seg_count * 2;
        sub.extend_from_slice(&4_u16.to_be_bytes()); // format
        sub.extend_from_slice(&0_u16.to_be_bytes()); // length (patched below)
        sub.extend_from_slice(&0_u16.to_be_bytes()); // language
        sub.extend_from_slice(&(seg_x2).to_be_bytes());
        // searchRange = 2 * 2^floor(log2(segCount)) = 4
        sub.extend_from_slice(&4_u16.to_be_bytes());
        // entrySelector = floor(log2(segCount)) = 1
        sub.extend_from_slice(&1_u16.to_be_bytes());
        // rangeShift = 2*segCount - searchRange = 0
        sub.extend_from_slice(&0_u16.to_be_bytes());
        sub.extend_from_slice(&0x42_u16.to_be_bytes()); // endCode[0]
        sub.extend_from_slice(&0xFFFF_u16.to_be_bytes()); // endCode[1]
        sub.extend_from_slice(&0_u16.to_be_bytes()); // reservedPad
        sub.extend_from_slice(&0x41_u16.to_be_bytes()); // startCode[0]
        sub.extend_from_slice(&0xFFFF_u16.to_be_bytes()); // startCode[1]
        sub.extend_from_slice(&(-0x40_i16).to_be_bytes()); // idDelta[0]: 0x41->1
        sub.extend_from_slice(&1_i16.to_be_bytes()); // idDelta[1]: 0xFFFF->0
        sub.extend_from_slice(&0_u16.to_be_bytes()); // idRangeOffset[0]
        sub.extend_from_slice(&0_u16.to_be_bytes()); // idRangeOffset[1]
        let sub_len = sub.len() as u16;
        sub[2..4].copy_from_slice(&sub_len.to_be_bytes());

        let mut cmap = Vec::new();
        cmap.extend_from_slice(&0_u16.to_be_bytes()); // version
        cmap.extend_from_slice(&1_u16.to_be_bytes()); // numTables
        cmap.extend_from_slice(&3_u16.to_be_bytes()); // platform: Windows
        cmap.extend_from_slice(&1_u16.to_be_bytes()); // encoding: Unicode BMP
        cmap.extend_from_slice(&12_u32.to_be_bytes()); // subtable offset
        cmap.extend_from_slice(&sub);

        // head: 54 bytes; only unitsPerEm (offset 18) is validated
        // (16..=16384).
        let mut head = Vec::new();
        head.extend_from_slice(&0x0001_0000_u32.to_be_bytes()); // version
        head.extend_from_slice(&0_u32.to_be_bytes()); // fontRevision
        head.extend_from_slice(&0_u32.to_be_bytes()); // checkSumAdjustment
        head.extend_from_slice(&0x5F0F_3CF5_u32.to_be_bytes()); // magicNumber
        head.extend_from_slice(&0_u16.to_be_bytes()); // flags
        head.extend_from_slice(&1000_u16.to_be_bytes()); // unitsPerEm
        head.extend_from_slice(&[0; 16]); // created + modified
        head.extend_from_slice(&[0; 8]); // xMin..yMax
        head.extend_from_slice(&[0; 10]); // macStyle..glyphDataFormat
        assert_eq!(head.len(), 54);

        // hhea: 36 bytes; numberOfMetrics (u16) at offset 34.
        let mut hhea = Vec::new();
        hhea.extend_from_slice(&0x0001_0000_u32.to_be_bytes()); // version
        hhea.extend_from_slice(&[0; 30]);
        hhea.extend_from_slice(&3_u16.to_be_bytes()); // numberOfMetrics
        assert_eq!(hhea.len(), 36);

        // maxp: version 1.0; numGlyphs (u16, nonzero) at offset 4.
        let mut maxp = Vec::new();
        maxp.extend_from_slice(&0x0001_0000_u32.to_be_bytes()); // version
        maxp.extend_from_slice(&3_u16.to_be_bytes()); // numGlyphs
        maxp.extend_from_slice(&[0; 26]); // remaining v1.0 fields
        assert_eq!(maxp.len(), 32);

        // Table directory records must be sorted by tag (ttf-parser
        // binary-searches them): cmap < head < hhea < maxp.
        let tables: [(&[u8; 4], &[u8]); 4] = [
            (b"cmap", &cmap),
            (b"head", &head),
            (b"hhea", &hhea),
            (b"maxp", &maxp),
        ];
        let num_tables = tables.len() as u16;
        let dir_len = 12 + tables.len() * 16;
        let mut sfnt = Vec::new();
        sfnt.extend_from_slice(&0x0001_0000_u32.to_be_bytes()); // sfnt version
        sfnt.extend_from_slice(&num_tables.to_be_bytes());
        sfnt.extend_from_slice(&64_u16.to_be_bytes()); // searchRange = 16*2^2
        sfnt.extend_from_slice(&2_u16.to_be_bytes()); // entrySelector
        sfnt.extend_from_slice(&0_u16.to_be_bytes()); // rangeShift
        let mut offset = dir_len as u32;
        for (tag, data) in &tables {
            sfnt.extend_from_slice(tag.as_slice());
            sfnt.extend_from_slice(&0_u32.to_be_bytes()); // checksum (unused)
            sfnt.extend_from_slice(&offset.to_be_bytes());
            sfnt.extend_from_slice(&(data.len() as u32).to_be_bytes());
            offset += data.len() as u32;
        }
        for (_, data) in &tables {
            sfnt.extend_from_slice(data);
        }
        sfnt
    }

    #[test]
    fn reverse_cmap_inverts_format4() {
        let rc = ReverseCmap::from_font_program(&tiny_sfnt()).unwrap();
        assert_eq!(rc.lookup(1), Some('A'));
        assert_eq!(rc.lookup(2), Some('B'));
        assert_eq!(rc.lookup(3), None);
    }

    #[test]
    fn reverse_cmap_rejects_garbage() {
        assert!(ReverseCmap::from_font_program(b"not a font").is_none());
        assert!(ReverseCmap::from_font_program(&[]).is_none());
    }
}
