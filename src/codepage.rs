//! Windows codepage to `encoding_rs` encoding mapping.
//!
//! Provides a function to decode 8-bit text using a Windows codepage ID,
//! shared by the `.doc` and `.xls` parsers. Falls back to cp1252 (Western
//! European) for unknown or unsupported codepages.

use encoding_rs::Encoding;

/// Decode a single byte using the given Windows codepage, returning its
/// Unicode code point. Used for per-character decoding in the `.doc` parser.
///
/// For ASCII bytes (< 0x80), returns the byte value directly (all Windows
/// codepages are ASCII-compatible). For high bytes, decodes through
/// `encoding_rs`.
pub(crate) fn decode_byte(byte: u8, codepage: u16) -> char {
    if byte < 0x80 {
        return char::from(byte);
    }
    let encoding = codepage_to_encoding(codepage);
    let buf = [byte];
    let (cow, _encoding_used, _had_errors) = encoding.decode(&buf);
    cow.chars().next().unwrap_or('\u{FFFD}')
}

/// Map a Windows codepage ID to an `encoding_rs` encoding.
///
/// Covers the codepages most commonly encountered in Office documents.
/// Unknown codepages fall back to Windows-1252 (Western European), which
/// is the most common encoding in legacy Office files.
fn codepage_to_encoding(codepage: u16) -> &'static Encoding {
    match codepage {
        437 => encoding_rs::IBM866, // DOS US — closest available; not perfect
        874 => encoding_rs::WINDOWS_874,
        932 => encoding_rs::SHIFT_JIS,
        936 => encoding_rs::GBK,
        949 => encoding_rs::EUC_KR,
        950 => encoding_rs::BIG5,
        1250 => encoding_rs::WINDOWS_1250,
        1251 => encoding_rs::WINDOWS_1251,
        1253 => encoding_rs::WINDOWS_1253,
        1254 => encoding_rs::WINDOWS_1254,
        1255 => encoding_rs::WINDOWS_1255,
        1256 => encoding_rs::WINDOWS_1256,
        1257 => encoding_rs::WINDOWS_1257,
        1258 => encoding_rs::WINDOWS_1258,
        10000 => encoding_rs::MACINTOSH,
        20866 => encoding_rs::KOI8_R,
        21866 => encoding_rs::KOI8_U,
        28592 => encoding_rs::ISO_8859_2,
        28595 => encoding_rs::ISO_8859_5,
        28597 => encoding_rs::ISO_8859_7,
        28598 => encoding_rs::ISO_8859_8,
        65001 => encoding_rs::UTF_8,
        _ => encoding_rs::WINDOWS_1252, // cp1252 / ISO-8859-1 / default
    }
}

/// Map a Word `lidFE` (locale ID) to a likely Windows codepage.
///
/// The `lidFE` field in the FIB indicates the language used for text
/// in the document. We map it to the Windows codepage that locale
/// typically uses. This is a heuristic — Word documents may override
/// this per-run in the piece table, but for the 8-bit fallback case
/// this is the best we can do without parsing the full piece table.
pub(crate) const fn lid_to_codepage(lid: u16) -> u16 {
    // Strip sublanguage bits for primary language matching
    let primary = lid & 0x03FF;
    match primary {
        0x0004 => 936,                                               // Chinese (Simplified)
        0x0011 => 932,                                               // Japanese
        0x0012 => 949,                                               // Korean
        0x0019 | 0x0022 | 0x0023 | 0x0402 => 1251, // Russian, Ukrainian, Belarusian, Bulgarian
        0x001A | 0x0005 | 0x000E | 0x0015 | 0x001B | 0x0024 => 1250, // Central European
        0x0025..=0x0027 => 1257,                   // Baltic (Estonian, Latvian, Lithuanian)
        0x0008 => 1253,                            // Greek
        0x001F => 1254,                            // Turkish
        0x000D => 1255,                            // Hebrew
        0x0001 | 0x0029 => 1256,                   // Arabic, Persian/Farsi
        0x002A => 1258,                            // Vietnamese
        0x001E => 874,                             // Thai
        0x0404 => 950,                             // Chinese (Traditional)
        _ => 1252,                                 // Western European (default)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decode_byte_ascii() {
        assert_eq!(decode_byte(b'A', 1252), 'A');
    }

    #[test]
    fn decode_byte_cp1251_high() {
        // 0xC0 in cp1251 = А
        assert_eq!(decode_byte(0xC0, 1251), '\u{0410}');
    }

    #[test]
    fn lid_russian() {
        assert_eq!(lid_to_codepage(0x0419), 1251); // Russian (Russia)
    }

    #[test]
    fn lid_japanese() {
        assert_eq!(lid_to_codepage(0x0411), 932); // Japanese
    }

    #[test]
    fn lid_chinese_simplified() {
        assert_eq!(lid_to_codepage(0x0804), 936); // Chinese (PRC)
    }

    #[test]
    fn lid_english_default() {
        assert_eq!(lid_to_codepage(0x0409), 1252); // English (US)
    }

    #[test]
    fn lid_polish() {
        assert_eq!(lid_to_codepage(0x0415), 1250); // Polish
    }
}
