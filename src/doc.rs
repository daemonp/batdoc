//! OLE2 `.doc` (Word 97+) binary format parser.
//!
//! Reads the `WordDocument` stream from the OLE2 compound file, parses
//! the FIB (File Information Block) header for text boundaries and flags,
//! and extracts the text using the 256-byte block Unicode/8-bit heuristic
//! from the original C `catdoc` project. Field codes (HYPERLINK, TOC, etc.)
//! are suppressed.

use cfb::CompoundFile;
use std::io::{Cursor, Read};

use crate::codepage;
use crate::error::BatdocError;
use crate::heuristic;

// FIB flag bits
const F_ENCRYPTED: u16 = 0x0100;
const F_EXT_CHAR: u16 = 0x1000;

/// Extract markdown-formatted text from an OLE2 .doc file.
///
/// Since .doc binary format doesn't carry style information through the text
/// stream, we apply heuristics to infer headings and tables from the plain text:
///   - Numbered lines like "1. Foo" or "1.2 Bar" that are short → headings
///   - Short standalone lines (< 80 chars, no sentence-ending punctuation) → bold
///   - Tab-separated lines with consistent columns → markdown tables
pub(crate) fn extract_markdown(data: &[u8]) -> crate::error::Result<String> {
    let plain = extract_plain(data)?;
    Ok(heuristic::plain_to_markdown(&plain))
}

/// Extract plain text from an OLE2 .doc file.
/// Returns the document text as a String with paragraph separation.
pub(crate) fn extract_plain(data: &[u8]) -> crate::error::Result<String> {
    let cursor = Cursor::new(data);
    let mut cfb = CompoundFile::open(cursor)?;

    let stream_path = "/WordDocument";
    if !cfb.exists(stream_path) {
        return Err(BatdocError::Document(
            "not a Word document (no WordDocument stream)".into(),
        ));
    }

    let mut stream = cfb.open_stream(stream_path)?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf)?;

    if buf.len() < 32 {
        return Err(BatdocError::Document(
            "WordDocument stream too short".into(),
        ));
    }

    let flags = u16::from_le_bytes([buf[10], buf[11]]);

    if flags & F_ENCRYPTED != 0 {
        return Err(BatdocError::Document("document is encrypted".into()));
    }

    // FIB `lid` (install language) at offset 6-7, used to infer codepage
    // for 8-bit text blocks when no piece table is available.
    let lid = u16::from_le_bytes([buf[6], buf[7]]);
    let cp = codepage::lid_to_codepage(lid);

    let text_start = u32::from_le_bytes([buf[24], buf[25], buf[26], buf[27]]) as usize; // u32 → usize: lossless on 32+ bit
    let text_end = u32::from_le_bytes([buf[28], buf[29], buf[30], buf[31]]) as usize;

    if text_start >= buf.len() || text_end > buf.len() || text_start >= text_end {
        return Err(BatdocError::Document(
            "invalid text boundaries in FIB".into(),
        ));
    }

    let text_data = &buf[text_start..text_end];
    let is_unicode = flags & F_EXT_CHAR != 0;

    let chars = if is_unicode {
        extract_word8_text(text_data, cp)
    } else {
        extract_8bit_text(text_data, cp)
    };

    Ok(chars_to_text(&chars))
}

/// Extract text from Word 97+ format using the 256-byte block heuristic.
fn extract_word8_text(data: &[u8], codepage: u16) -> Vec<u16> {
    let mut result = Vec::new();
    let mut offset = 0;

    while offset < data.len() {
        let block_end = (offset + 256).min(data.len());
        let block = &data[offset..block_end];

        if detect_unicode_block(block) {
            for pair in block.chunks_exact(2) {
                result.push(u16::from_le_bytes([pair[0], pair[1]]));
            }
        } else {
            for &b in block {
                let ch = codepage::decode_byte(b, codepage);
                // Convert char to u16 for the existing pipeline.
                // BMP characters fit in u16; supplementary plane chars
                // (unlikely in 8-bit codepages) get two u16 surrogates.
                let mut buf = [0u16; 2];
                let encoded = ch.encode_utf16(&mut buf);
                result.extend_from_slice(encoded);
            }
        }

        offset = block_end;
    }

    result
}

/// Detect if a 256-byte block is UTF-16LE encoded.
fn detect_unicode_block(block: &[u8]) -> bool {
    block.chunks_exact(2).any(|pair| {
        let c = pair[0];
        (c == 0x20 || c == 0x0D || c.is_ascii_punctuation()) && pair[1] == 0x00
    })
}

/// Extract text from pre-Word97 8-bit encoded stream.
fn extract_8bit_text(data: &[u8], codepage: u16) -> Vec<u16> {
    let mut result = Vec::new();
    for &b in data {
        let ch = codepage::decode_byte(b, codepage);
        let mut buf = [0u16; 2];
        let encoded = ch.encode_utf16(&mut buf);
        result.extend_from_slice(encoded);
    }
    result
}

/// Flush a paragraph buffer into the output string if non-empty.
fn flush_paragraph(paragraph: &mut String, output: &mut String, first: &mut bool) {
    let text = paragraph.trim_end();
    if !text.is_empty() {
        if !*first {
            output.push('\n');
        }
        output.push_str(text);
        output.push('\n');
        *first = false;
    }
    paragraph.clear();
}

/// Tracks what we're collecting during a field: instruction or display text.
#[derive(Clone)]
enum FieldState {
    /// Collecting field instruction text (between 0x0013 and 0x0014).
    Instruction(String),
    /// Collecting display text (between 0x0014 and 0x0015), with the URL
    /// if this is a HYPERLINK field.
    Display { url: Option<String>, text: String },
}

/// Process the u16 character stream into paragraphs, extracting hyperlinks
/// from field codes and suppressing other field codes.
///
/// Word field codes use three markers:
/// - `0x0013` — field begin (instruction text follows)
/// - `0x0014` — field separator (display text follows)
/// - `0x0015` — field end
///
/// For `HYPERLINK` fields, we capture the URL from the instruction and
/// emit `[display text](url)` inline. Other field types (TOC, PAGE, etc.)
/// are suppressed as before.
///
/// Handles UTF-16 surrogate pairs: a high surrogate (0xD800-0xDBFF) followed
/// by a low surrogate (0xDC00-0xDFFF) is decoded into the correct supplementary
/// plane character. Unpaired surrogates are replaced with U+FFFD.
fn chars_to_text(chars: &[u16]) -> String {
    let mut output = String::new();
    let mut paragraph = String::new();
    let mut first = true;
    let mut field_depth: i32 = 0;
    let mut field_stack: Vec<FieldState> = Vec::new();
    let mut pending_high_surrogate: Option<u16> = None;

    for &c in chars {
        // Handle surrogate pair completion
        if let Some(hi) = pending_high_surrogate.take() {
            if (0xDC00..=0xDFFF).contains(&c) {
                // Valid surrogate pair → supplementary plane character
                let code = 0x10000 + ((u32::from(hi) - 0xD800) << 10) + (u32::from(c) - 0xDC00);
                if let Some(ch) = char::from_u32(code) {
                    push_char_to_field_or_para(ch, &mut field_stack, &mut paragraph);
                }
                continue;
            }
            // Unpaired high surrogate — emit replacement and process `c` normally
            push_char_to_field_or_para('\u{FFFD}', &mut field_stack, &mut paragraph);
        }

        // Buffer high surrogates for the next iteration
        if (0xD800..=0xDBFF).contains(&c) {
            pending_high_surrogate = Some(c);
            continue;
        }

        match c {
            0x0013 => {
                // Field begin — start capturing instruction text
                field_depth += 1;
                field_stack.push(FieldState::Instruction(String::new()));
            }
            0x0014 => {
                // Field separator — switch from instruction to display text
                if let Some(state) = field_stack.last_mut() {
                    let url = if let FieldState::Instruction(ref instr) = state {
                        extract_hyperlink_url(instr)
                    } else {
                        None
                    };
                    *state = FieldState::Display {
                        url,
                        text: String::new(),
                    };
                }
            }
            0x0015 => {
                emit_field_end(&mut field_stack, &mut paragraph);
                if field_depth > 0 {
                    field_depth -= 1;
                }
            }
            _ if field_depth > 0 => {
                // Inside a field — accumulate into the appropriate buffer
                if c == 0x000D || c == 0x000B {
                    // Paragraph break inside field — clear current field text
                    if let Some(state) = field_stack.last_mut() {
                        match state {
                            FieldState::Instruction(ref mut s) => s.clear(),
                            FieldState::Display { ref mut text, .. } => text.clear(),
                        }
                    }
                } else if let Some(ch) = char::from_u32(u32::from(c)) {
                    if ch >= ' ' {
                        if let Some(state) = field_stack.last_mut() {
                            match state {
                                FieldState::Instruction(ref mut s) => s.push(ch),
                                FieldState::Display { ref mut text, .. } => text.push(ch),
                            }
                        }
                    }
                }
            }
            0x000B..=0x000D => {
                flush_paragraph(&mut paragraph, &mut output, &mut first);
            }
            0x0007 | 0x0009 => {
                paragraph.push('\t');
            }
            0x001E => {
                paragraph.push('-');
            }
            0x001F | 0x0002 | 0xFEFF => {}
            // Lone low surrogate (not preceded by high) — replace
            c if (0xDC00..=0xDFFF).contains(&c) => {
                paragraph.push('\u{FFFD}');
            }
            c if c < 0x0020 => {}
            c => {
                if let Some(ch) = char::from_u32(u32::from(c)) {
                    paragraph.push(ch);
                }
            }
        }
    }

    // Flush any trailing unpaired high surrogate
    if pending_high_surrogate.is_some() {
        paragraph.push('\u{FFFD}');
    }

    flush_paragraph(&mut paragraph, &mut output, &mut first);
    output
}

/// Process a field-end marker (0x0015): pop the field state and emit
/// the result into the paragraph.
fn emit_field_end(field_stack: &mut Vec<FieldState>, paragraph: &mut String) {
    if let Some(state) = field_stack.pop() {
        match state {
            FieldState::Display {
                url: Some(url),
                text,
            } => {
                // Emit markdown-style link
                paragraph.push('[');
                paragraph.push_str(text.trim());
                paragraph.push_str("](");
                paragraph.push_str(&url);
                paragraph.push(')');
            }
            FieldState::Display { url: None, text } => {
                // Non-hyperlink field with display text — emit the display text
                paragraph.push_str(&text);
            }
            FieldState::Instruction(_) => {
                // No separator seen — field is fully suppressed
            }
        }
    }
}

/// Push a character into the innermost field buffer, or into the paragraph
/// if we're not inside any field.
fn push_char_to_field_or_para(ch: char, field_stack: &mut [FieldState], paragraph: &mut String) {
    if let Some(state) = field_stack.last_mut() {
        match state {
            FieldState::Instruction(ref mut s) => s.push(ch),
            FieldState::Display { ref mut text, .. } => text.push(ch),
        }
    } else {
        paragraph.push(ch);
    }
}

/// Extract a URL from a HYPERLINK field instruction string.
///
/// Field instruction format: `HYPERLINK "http://example.com" \l "bookmark"`
/// or `HYPERLINK http://example.com`. We extract the URL, handling both
/// quoted and unquoted forms.
fn extract_hyperlink_url(instruction: &str) -> Option<String> {
    let trimmed = instruction.trim();

    // Must start with "HYPERLINK" (case-insensitive)
    let rest = if let Some(r) = trimmed.strip_prefix("HYPERLINK") {
        r
    } else if let Some(r) = trimmed.strip_prefix("hyperlink") {
        r
    } else {
        let lower = trimmed.to_lowercase();
        if let Some(idx) = lower.find("hyperlink") {
            &trimmed[idx + 9..]
        } else {
            return None;
        }
    };

    let rest = rest.trim_start();
    if rest.is_empty() {
        return None;
    }

    // Extract URL: may be quoted or unquoted
    let url = rest.strip_prefix('"').map_or_else(
        || {
            // Unquoted: take until whitespace or backslash (switch start)
            let end = rest
                .find(|c: char| c.is_whitespace() || c == '\\')
                .unwrap_or(rest.len());
            &rest[..end]
        },
        |inner| {
            // Quoted: find closing quote
            let end = inner.find('"').unwrap_or(inner.len());
            &inner[..end]
        },
    );

    if url.is_empty() {
        None
    } else {
        Some(url.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── detect_unicode_block ─────────────────────────────────────

    #[test]
    fn detect_unicode_space_zero() {
        let block = [0x20, 0x00, 0x41, 0x00]; // " A" in UTF-16LE
        assert!(detect_unicode_block(&block));
    }

    #[test]
    fn detect_unicode_cr_zero() {
        let block = [0x0D, 0x00, 0x00, 0x00];
        assert!(detect_unicode_block(&block));
    }

    #[test]
    fn detect_8bit_block() {
        let block = [0x48, 0x65, 0x6C, 0x6C]; // "Hell" in ASCII
        assert!(!detect_unicode_block(&block));
    }

    #[test]
    fn detect_empty_block() {
        assert!(!detect_unicode_block(&[]));
    }

    #[test]
    fn detect_single_byte() {
        assert!(!detect_unicode_block(&[0x20]));
    }

    // ── chars_to_text ────────────────────────────────────────────

    #[test]
    fn simple_paragraph() {
        let chars: Vec<u16> = "Hello world".encode_utf16().collect();
        let mut chars_with_cr = chars;
        chars_with_cr.push(0x000D);
        assert_eq!(chars_to_text(&chars_with_cr), "Hello world\n");
    }

    #[test]
    fn two_paragraphs() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("First".encode_utf16());
        chars.push(0x000D);
        chars.extend("Second".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "First\n\nSecond\n");
    }

    #[test]
    fn hyperlink_field_emits_link() {
        // HYPERLINK fields now emit markdown-style [text](url)
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Before ".encode_utf16());
        chars.push(0x0013); // field begin
        chars.extend("HYPERLINK \"http://example.com\"".encode_utf16());
        chars.push(0x0014); // field separator
        chars.extend("visible text".encode_utf16());
        chars.push(0x0015); // field end
        chars.push(0x000D);
        assert_eq!(
            chars_to_text(&chars),
            "Before [visible text](http://example.com)\n"
        );
    }

    #[test]
    fn hyperlink_field_unquoted() {
        let mut chars: Vec<u16> = Vec::new();
        chars.push(0x0013);
        chars.extend("HYPERLINK http://example.com".encode_utf16());
        chars.push(0x0014);
        chars.extend("click".encode_utf16());
        chars.push(0x0015);
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "[click](http://example.com)\n");
    }

    #[test]
    fn non_hyperlink_field_display_text_shown() {
        // Non-HYPERLINK fields with display text show the display text
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Page ".encode_utf16());
        chars.push(0x0013); // field begin
        chars.extend("PAGE".encode_utf16());
        chars.push(0x0014); // separator
        chars.extend("42".encode_utf16());
        chars.push(0x0015); // field end
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "Page 42\n");
    }

    #[test]
    fn field_codes_fully_suppressed() {
        // When 0x0015 comes without 0x0014, the whole field is hidden.
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Before ".encode_utf16());
        chars.push(0x0013); // field begin
        chars.extend("TOC hidden".encode_utf16());
        chars.push(0x0015); // field end (no separator)
        chars.extend(" After".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "Before  After\n");
    }

    #[test]
    fn extract_hyperlink_url_quoted() {
        assert_eq!(
            extract_hyperlink_url(r#" HYPERLINK "https://example.com" \l "top""#),
            Some("https://example.com".to_string())
        );
    }

    #[test]
    fn extract_hyperlink_url_unquoted() {
        assert_eq!(
            extract_hyperlink_url("HYPERLINK http://example.com"),
            Some("http://example.com".to_string())
        );
    }

    #[test]
    fn extract_hyperlink_url_not_hyperlink() {
        assert_eq!(extract_hyperlink_url("TOC \\o \\h"), None);
    }

    #[test]
    fn tab_characters() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("A".encode_utf16());
        chars.push(0x0009); // tab
        chars.extend("B".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "A\tB\n");
    }

    #[test]
    fn cell_marker_becomes_tab() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Cell1".encode_utf16());
        chars.push(0x0007); // cell marker
        chars.extend("Cell2".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "Cell1\tCell2\n");
    }

    #[test]
    fn non_breaking_hyphen() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("well".encode_utf16());
        chars.push(0x001E); // non-breaking hyphen
        chars.extend("known".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "well-known\n");
    }

    #[test]
    fn trailing_whitespace_trimmed() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Hello   ".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "Hello\n");
    }

    #[test]
    fn empty_paragraphs_skipped() {
        let chars: Vec<u16> = vec![0x000D, 0x000D, 0x000D];
        assert_eq!(chars_to_text(&chars), "");
    }

    #[test]
    fn bom_skipped() {
        let mut chars: Vec<u16> = vec![0xFEFF];
        chars.extend("Hello".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "Hello\n");
    }

    #[test]
    fn page_break_flushes() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Page1".encode_utf16());
        chars.push(0x000C); // page break
        chars.extend("Page2".encode_utf16());
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "Page1\n\nPage2\n");
    }

    #[test]
    fn text_without_trailing_cr() {
        let chars: Vec<u16> = "No newline".encode_utf16().collect();
        assert_eq!(chars_to_text(&chars), "No newline\n");
    }

    // ── surrogate pair handling ─────────────────────────────────

    #[test]
    fn surrogate_pair_emoji() {
        // U+1F600 (😀) = D83D DE00 in UTF-16
        let mut chars: Vec<u16> = vec![0xD83D, 0xDE00];
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "\u{1F600}\n");
    }

    #[test]
    fn unpaired_high_surrogate() {
        let mut chars: Vec<u16> = vec![0xD83D]; // high surrogate alone
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "\u{FFFD}\n");
    }

    #[test]
    fn unpaired_low_surrogate() {
        let mut chars: Vec<u16> = vec![0xDE00]; // low surrogate alone
        chars.push(0x000D);
        assert_eq!(chars_to_text(&chars), "\u{FFFD}\n");
    }

    // ── extract_8bit_text ────────────────────────────────────────

    #[test]
    fn extract_8bit_ascii() {
        let data = b"ABC";
        let result = extract_8bit_text(data, 1252);
        assert_eq!(result, vec![0x41, 0x42, 0x43]);
    }

    #[test]
    fn extract_8bit_special() {
        let data = [0x80]; // Euro sign in cp1252
        let result = extract_8bit_text(&data, 1252);
        assert_eq!(result, vec![0x20AC]);
    }

    #[test]
    fn extract_8bit_cyrillic() {
        // 0xC0 in cp1251 = А (U+0410)
        let data = [0xC0];
        let result = extract_8bit_text(&data, 1251);
        assert_eq!(result, vec![0x0410]);
    }
}
