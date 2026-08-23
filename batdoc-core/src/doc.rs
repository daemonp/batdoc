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
use crate::ExtractSink;

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
    let mut out = String::new();
    extract_plain_to(data, &mut out)?;
    Ok(out)
}

/// Stream plain text from an OLE2 .doc file into `sink`.
///
/// The `WordDocument` stream is still read fully into memory (the compound
/// file stream API is not chunked), but each 256-byte block is decoded into
/// the sink as it is scanned, instead of building a full `Vec<u16>` then a
/// full output `String`.
pub(crate) fn extract_plain_to(
    data: &[u8],
    sink: &mut impl ExtractSink,
) -> crate::error::Result<()> {
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

    let mut writer = DocWriter::new(sink);
    if is_unicode {
        extract_word8_text_to(text_data, cp, &mut writer)?;
    } else {
        extract_8bit_text_to(text_data, cp, &mut writer)?;
    }
    writer.finish()
}

/// Stream Word 97+ text out block by block using the 256-byte heuristic.
fn extract_word8_text_to<S: ExtractSink>(
    data: &[u8],
    codepage: u16,
    writer: &mut DocWriter<'_, S>,
) -> crate::error::Result<()> {
    let mut offset = 0;

    while offset < data.len() {
        let block_end = (offset + 256).min(data.len());
        let block = &data[offset..block_end];

        if detect_unicode_block(block) {
            for pair in block.as_chunks::<2>().0 {
                writer.push_char(u16::from_le_bytes(*pair))?;
            }
        } else {
            for &b in block {
                let ch = codepage::decode_byte(b, codepage);
                // Convert char to u16 for the shared pipeline.
                // BMP characters fit in u16; supplementary plane chars
                // (unlikely in 8-bit codepages) get two u16 surrogates.
                let mut buf = [0u16; 2];
                let encoded = ch.encode_utf16(&mut buf);
                for &cu in encoded.iter() {
                    writer.push_char(cu)?;
                }
            }
        }

        offset = block_end;
    }

    Ok(())
}

/// Stream pre-Word97 8-bit text out byte by byte.
fn extract_8bit_text_to<S: ExtractSink>(
    data: &[u8],
    codepage: u16,
    writer: &mut DocWriter<'_, S>,
) -> crate::error::Result<()> {
    for &b in data {
        let ch = codepage::decode_byte(b, codepage);
        let mut buf = [0u16; 2];
        let encoded = ch.encode_utf16(&mut buf);
        for &cu in encoded.iter() {
            writer.push_char(cu)?;
        }
    }
    Ok(())
}

/// Detect if a 256-byte block is UTF-16LE encoded.
fn detect_unicode_block(block: &[u8]) -> bool {
    block.as_chunks::<2>().0.iter().any(|pair| {
        let c = pair[0];
        (c == 0x20 || c == 0x0D || c.is_ascii_punctuation()) && pair[1] == 0x00
    })
}

/// Extract text from pre-Word97 8-bit encoded stream into a `Vec<u16>`.
///
/// Only used by tests: the streaming path decodes into the sink via
/// [`extract_8bit_text_to`].
#[cfg(test)]
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

/// Incremental writer that turns decoded u16 code units into paragraph text
/// and emits completed paragraphs to a sink.
///
/// Shared by the streaming [`extract_plain_to`] path and, via a `String`
/// sink, by the buffered `chars_to_text` used in tests.
struct DocWriter<'a, S: ExtractSink> {
    sink: &'a mut S,
    paragraph: String,
    first: bool,
    field_depth: i32,
    field_stack: Vec<FieldState>,
    pending_high_surrogate: Option<u16>,
}

impl<'a, S: ExtractSink> DocWriter<'a, S> {
    const fn new(sink: &'a mut S) -> Self {
        Self {
            sink,
            paragraph: String::new(),
            first: true,
            field_depth: 0,
            field_stack: Vec::new(),
            pending_high_surrogate: None,
        }
    }

    /// Feed a single u16 code unit through the field/surrogate/paragraph
    /// state machine, emitting completed paragraphs to the sink.
    fn push_char(&mut self, c: u16) -> crate::error::Result<()> {
        // Handle surrogate pair completion
        if let Some(hi) = self.pending_high_surrogate.take() {
            if (0xDC00..=0xDFFF).contains(&c) {
                // Valid surrogate pair → supplementary plane character
                let code = 0x10000 + ((u32::from(hi) - 0xD800) << 10) + (u32::from(c) - 0xDC00);
                if let Some(ch) = char::from_u32(code) {
                    push_char_to_field_or_para(ch, &mut self.field_stack, &mut self.paragraph);
                }
                return Ok(());
            }
            // Unpaired high surrogate — emit replacement and process `c` normally
            push_char_to_field_or_para('\u{FFFD}', &mut self.field_stack, &mut self.paragraph);
        }

        // Buffer high surrogates for the next iteration
        if (0xD800..=0xDBFF).contains(&c) {
            self.pending_high_surrogate = Some(c);
            return Ok(());
        }

        match c {
            0x0013 => {
                // Field begin — start capturing instruction text
                self.field_depth += 1;
                self.field_stack
                    .push(FieldState::Instruction(String::new()));
            }
            0x0014 => {
                // Field separator — switch from instruction to display text
                if let Some(state) = self.field_stack.last_mut() {
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
                emit_field_end(&mut self.field_stack, &mut self.paragraph);
                if self.field_depth > 0 {
                    self.field_depth -= 1;
                }
            }
            _ if self.field_depth > 0 => {
                // Inside a field — accumulate into the appropriate buffer
                if c == 0x000D || c == 0x000B {
                    // Paragraph break inside field — clear current field text
                    if let Some(state) = self.field_stack.last_mut() {
                        match state {
                            FieldState::Instruction(ref mut s) => s.clear(),
                            FieldState::Display { ref mut text, .. } => text.clear(),
                        }
                    }
                } else if let Some(ch) = char::from_u32(u32::from(c)) {
                    if ch >= ' ' {
                        if let Some(state) = self.field_stack.last_mut() {
                            match state {
                                FieldState::Instruction(ref mut s) => s.push(ch),
                                FieldState::Display { ref mut text, .. } => text.push(ch),
                            }
                        }
                    }
                }
            }
            0x000B..=0x000D => {
                self.flush_paragraph()?;
            }
            0x0007 | 0x0009 => {
                self.paragraph.push('\t');
            }
            0x001E => {
                self.paragraph.push('-');
            }
            0x001F | 0x0002 | 0xFEFF => {}
            // Lone low surrogate (not preceded by high) — replace
            c if (0xDC00..=0xDFFF).contains(&c) => {
                self.paragraph.push('\u{FFFD}');
            }
            c if c < 0x0020 => {}
            c => {
                if let Some(ch) = char::from_u32(u32::from(c)) {
                    self.paragraph.push(ch);
                }
            }
        }
        Ok(())
    }

    /// Flush the current paragraph buffer into the sink if non-empty.
    fn flush_paragraph(&mut self) -> crate::error::Result<()> {
        let text = self.paragraph.trim_end();
        if !text.is_empty() {
            if !self.first {
                self.sink.write_str("\n")?;
            }
            self.sink.write_str(text)?;
            self.sink.write_str("\n")?;
            self.first = false;
        }
        self.paragraph.clear();
        Ok(())
    }

    /// Flush any trailing unpaired surrogate and the final paragraph.
    fn finish(&mut self) -> crate::error::Result<()> {
        if self.pending_high_surrogate.is_some() {
            self.paragraph.push('\u{FFFD}');
        }
        self.flush_paragraph()
    }
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
#[cfg(test)]
fn chars_to_text(chars: &[u16]) -> String {
    let mut out = String::new();
    {
        let mut writer = DocWriter::new(&mut out);
        for &c in chars {
            writer.push_char(c).expect("String sink cannot fail");
        }
        writer.finish().expect("String sink cannot fail");
    }
    out
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
        {
            let idx = lower.find("hyperlink")?;
            &trimmed[idx + 9..]
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

    // ── extract_plain_to parity ──────────────────────────────────

    /// Build a minimal OLE2 `.doc` whose `WordDocument` stream is a 32-byte
    /// FIB header (lid=0x0409, `flags`, text boundaries) followed by `text_data`.
    fn build_doc(flags: u16, text_data: &[u8]) -> Vec<u8> {
        use std::io::Write;

        let header_len = 32usize;
        let mut doc = vec![0u8; header_len];
        doc[6..8].copy_from_slice(&0x0409u16.to_le_bytes()); // lid (US English)
        doc[10..12].copy_from_slice(&flags.to_le_bytes());
        let text_start = header_len as u32;
        let text_end = text_start + u32::try_from(text_data.len()).unwrap();
        doc[24..28].copy_from_slice(&text_start.to_le_bytes());
        doc[28..32].copy_from_slice(&text_end.to_le_bytes());
        doc.extend_from_slice(text_data);

        let mut cfb = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        {
            let mut stream = cfb.create_stream("/WordDocument").unwrap();
            stream.write_all(&doc).unwrap();
        }
        cfb.flush().unwrap();
        cfb.into_inner().into_inner()
    }

    #[test]
    fn extract_plain_to_matches_extract_plain_unicode() {
        let mut chars: Vec<u16> = Vec::new();
        chars.extend("Hello".encode_utf16());
        chars.push(0x0013);
        chars.extend("HYPERLINK \"http://example.com\"".encode_utf16());
        chars.push(0x0014);
        chars.extend("there".encode_utf16());
        chars.push(0x0015);
        chars.push(0x000D);
        chars.extend("Second".encode_utf16());
        chars.push(0x000D);
        chars.extend("\u{1F600}".encode_utf16()); // surrogate pair, no trailing CR
        let bytes: Vec<u8> = chars.iter().flat_map(|c| c.to_le_bytes()).collect();

        let data = build_doc(F_EXT_CHAR, &bytes);
        let expected = extract_plain(&data).unwrap();
        let mut out = String::new();
        extract_plain_to(&data, &mut out).unwrap();
        assert_eq!(out, expected);
    }

    #[test]
    fn extract_plain_to_matches_extract_plain_8bit() {
        let text = b"Hello world\rSecond line\r";
        let data = build_doc(0, text);
        let expected = extract_plain(&data).unwrap();
        let mut out = String::new();
        extract_plain_to(&data, &mut out).unwrap();
        assert_eq!(out, expected);
    }
}
