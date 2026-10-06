#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
// Denied for the reason `classify.rs` gives at its top: this module reads attacker-chosen
// offsets out of a CFB before anything has authenticated it. Production code here has
// none of the three (measured 2026-09-05). `clippy::indexing_slicing` is left off for
// the reason given there too, though this module has no hits of it at all.
//! Recognise a legacy binary Office document, say whether it is encrypted, and read the
//! structures the decrypt paths walk to find out.
//!
//! Word 97-2003, Excel 97-2003 and PowerPoint 97-2003 are CFB containers like an
//! encrypted OOXML file, but they carry no `EncryptionInfo` stream: the encryption is
//! woven into the format's own records. Without this module [`crate::classify()`] sees a CFB
//! it cannot read and reports `Unknown`, which is honest but tells a caller nothing —
//! not even that the file *is* encrypted.
//!
//! That gap was worth closing before the decrypt paths existed (GH #4), because the
//! failure it prevents is the one this crate has already made once. GH #11 reported
//! `WrongPassword` for a file whose *algorithm* was unimplemented, and LibreOffice makes a
//! worse version of the same mistake on exactly these files: its PowerPoint filter never
//! checks for encryption, parses ciphertext as records, and tells the user
//! *"Incorrect file version"* — the codec is present in `filter/source/msfilter/mscodec`
//! and wired into Writer and Calc, just never into Impress. A detection step in front of
//! the parser is what avoids that, and it costs no cryptography at all.
//!
//! # One reader, two callers
//!
//! The FIB, the BIFF record walk and PowerPoint's persist directory are read here, once,
//! by functions that return `Option` and never fail loudly. [`probe`] maps `None` to
//! "could not tell"; the decrypt modules (`word97`, `excel97`, `powerpoint97`, behind the
//! `legacy-binary` feature) map the same `None` to an error naming the field. Two walkers
//! that could disagree about where a record starts would be the crossed-parameter bug
//! this crate has already shipped once, in a different disguise.
//!
//! # What each format is asked
//!
//! | format | container marker | encryption marker |
//! | --- | --- | --- |
//! | Word | `WordDocument` stream | FIB `fEncrypted` / `fObfuscated` flags, then the `EncryptionHeader` at the start of the table stream |
//! | Excel | `Workbook` or `Book` stream | the `FILEPASS` record (`0x002F`) in the BIFF stream |
//! | PowerPoint | `PowerPoint Document` + `Current User` | the `UserEditAtom` at `offsetToCurrentEdit` is `0x20` bytes when encrypted; its `encryptSessionPersistIdRef` then leads through the persist directory to the `CryptSession10Container` |
//!
//! PowerPoint's is the non-obvious one. `CurrentUserAtom.headerToken` looks like the
//! answer and is not — a real encrypted file written by PowerPoint 16 still carries the
//! *unencrypted* token, so reading it reports the wrong verdict. `msoffcrypto-tool`
//! (`format/ppt97.py:812-837`) follows `offsetToCurrentEdit` into the document stream and
//! tests the `UserEditAtom`'s length instead, which is what this module does.
//!
//! # Hostile input
//!
//! Everything read here is attacker-chosen. Every offset is bounds-checked before use,
//! every scan is bounded by [`crate::limits`], and no path can panic or loop unboundedly:
//! `probe` returns `None` for anything it cannot make sense of, exactly as `classify`
//! returns `Unknown`. The PowerPoint walk seeks rather than reading a prefix, so an
//! `offsetToCurrentEdit` near the end of a 50 MB deck costs a few small reads and not
//! 50 MB of memory.

use crate::limits::{
    BIFF_SCAN_CAP, BINARY_HEADER_READ_CAP, CURRENT_USER_READ_CAP, ENCRYPTION_HEADER_STRUCTURE_MAX,
    PPT_PERSIST_DIRECTORY_READ_CAP, PPT_PERSIST_OBJECTS_MAX,
};
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Seek, SeekFrom};
use std::ops::Range;

/// Which 97-2003 application format a CFB container holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BinaryFormat {
    Word,
    Excel,
    PowerPoint,
}

/// What [`probe`] found in a legacy binary container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct BinaryVerdict {
    pub(crate) format: BinaryFormat,
    /// `Some(false)` for a document carrying no password-to-open, `None` when the
    /// format was recognised but its encryption marker could not be reached.
    ///
    /// The distinction is the point of this module: reporting a file as unprotected
    /// because we failed to read it is worse than admitting we do not know.
    pub(crate) encrypted: Option<bool>,
    /// `EncryptionVersionInfo` where the format exposes one. `None` for an unencrypted
    /// document, and for XOR obfuscation, which predates that header entirely.
    pub(crate) version: Option<(u16, u16)>,
    /// `EncryptionHeader.KeySize`, in bits, where one was readable.
    pub(crate) key_bits: Option<u32>,
    /// XOR obfuscation rather than an RC4 family. A different scheme, not a weaker one.
    pub(crate) xor_obfuscated: bool,
}

// ---- stream names --------------------------------------------------------------------

pub(crate) const WORD_DOCUMENT: &str = "/WordDocument";
pub(crate) const TABLE_0: &str = "/0Table";
pub(crate) const TABLE_1: &str = "/1Table";
#[cfg(feature = "legacy-binary")]
pub(crate) const DATA: &str = "/Data";
pub(crate) const WORKBOOK: &str = "/Workbook";
/// The BIFF5 (Excel 5.0/95) name for the workbook stream. Recognised so that such a file
/// is named rather than `Unknown`; its record layout is not the one this crate reads.
pub(crate) const BOOK: &str = "/Book";
pub(crate) const POWERPOINT_DOCUMENT: &str = "/PowerPoint Document";
pub(crate) const CURRENT_USER: &str = "/Current User";

/// Which format a container's stream names announce, or `None` for a CFB that carries
/// none of the three markers.
///
/// Order matters only in that the markers are mutually exclusive in practice.
pub(crate) fn format_of<F: Read + Seek>(cfb: &cfb::CompoundFile<F>) -> Option<BinaryFormat> {
    if cfb.exists(WORD_DOCUMENT) {
        return Some(BinaryFormat::Word);
    }
    if cfb.exists(WORKBOOK) || cfb.exists(BOOK) {
        return Some(BinaryFormat::Excel);
    }
    if cfb.exists(POWERPOINT_DOCUMENT) && cfb.exists(CURRENT_USER) {
        return Some(BinaryFormat::PowerPoint);
    }
    None
}

/// The workbook stream a container carries: BIFF8's `Workbook`, else BIFF5's `Book`.
pub(crate) fn workbook_stream_name<F: Read + Seek>(
    cfb: &cfb::CompoundFile<F>,
) -> Option<&'static str> {
    if cfb.exists(WORKBOOK) {
        Some(WORKBOOK)
    } else if cfb.exists(BOOK) {
        Some(BOOK)
    } else {
        None
    }
}

/// Recognise a legacy binary Office container, or return `None`.
///
/// `None` means "not a legacy binary Office document, or unreadable as one" — never an
/// error, because a caller of `classify` gets a verdict for every input.
pub(crate) fn probe(data: &[u8]) -> Option<BinaryVerdict> {
    let mut cfb = cfb::CompoundFile::open(Cursor::new(data)).ok()?;
    match format_of(&cfb)? {
        BinaryFormat::Word => probe_word(&mut cfb),
        BinaryFormat::Excel => probe_excel(&mut cfb),
        BinaryFormat::PowerPoint => probe_powerpoint(&mut cfb),
    }
}

/// Read at most `cap` bytes of a stream. A stream shorter than `cap` is not an error.
fn read_capped<F: Read + Seek>(
    cfb: &mut cfb::CompoundFile<F>,
    name: &str,
    cap: usize,
) -> Option<Vec<u8>> {
    let mut buf = Vec::new();
    cfb.open_stream(name)
        .ok()?
        .take(cap as u64)
        .read_to_end(&mut buf)
        .ok()?;
    Some(buf)
}

pub(crate) fn le16(b: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes([
        *b.get(at)?,
        *b.get(at.checked_add(1)?)?,
    ]))
}

pub(crate) fn le32(b: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes([
        *b.get(at)?,
        *b.get(at.checked_add(1)?)?,
        *b.get(at.checked_add(2)?)?,
        *b.get(at.checked_add(3)?)?,
    ]))
}

/// Parse the `EncryptionVersionInfo` + `EncryptionHeader` prefix shared by the binary
/// formats' encryption blobs, for reporting only.
///
/// Layout per [MS-OFFCRYPTO] §2.3.5.1 / §2.3.2: `vMajor`(2) `vMinor`(2) `Flags`(4)
/// `HeaderSize`(4), then the header itself — `Flags`(4) `SizeExtra`(4) `AlgID`(4)
/// `AlgIDHash`(4) `KeySize`(4) … — so `KeySize` sits at offset 28. Returns
/// `(version, key_bits)`. The strict parse the decrypt path runs is
/// `rc4_cryptoapi::parse`; this one reports what a file *says* and refuses nothing.
pub(crate) fn parse_encryption_header(blob: &[u8]) -> (Option<(u16, u16)>, Option<u32>) {
    let version = match (le16(blob, 0), le16(blob, 2)) {
        (Some(major), Some(minor)) => Some((major, minor)),
        _ => None,
    };
    // Only the CryptoAPI-era header carries a KeySize; the 1.1 RC4 header does not lay
    // its fields out this way, so do not invent one for it.
    let key_bits = match version {
        Some((2..=4, 2)) => le32(blob, 28).filter(|k| *k > 0 && *k <= 4096),
        _ => None,
    };
    (version, key_bits)
}

// ---- Word ----------------------------------------------------------------------------

/// The bytes of the `WordDocument` stream that are never encrypted: the FIB's first 68
/// ([MS-DOC] §2.2.6.2 and §2.2.6.3, "beyond the initial 68 bytes").
#[cfg(feature = "legacy-binary")]
pub(crate) const FIB_CLEAR_LEN: usize = 0x44;

/// `FibBase.wIdent` — [MS-DOC] §2.5.2, "MUST be 0xA5EC".
const FIB_W_IDENT: u16 = 0xA5EC;
/// Offset of the 16 flag bits `fDot` … `fObfuscated` within the FIB.
const FIB_FLAGS_AT: usize = 0x0A;
/// Offset of `FibBase.lKey`.
const FIB_L_KEY_AT: usize = 0x0E;
/// `fEncrypted` (bit F), `fWhichTblStm` (bit G) and `fObfuscated` (bit M) of the flag
/// word — [MS-DOC] §2.5.2.
pub(crate) const FIB_F_ENCRYPTED: u16 = 0x0100;
pub(crate) const FIB_F_WHICH_TBL_STM: u16 = 0x0200;
pub(crate) const FIB_F_OBFUSCATED: u16 = 0x8000;

/// The three `FibBase` fields the encryption paths act on — [MS-DOC] §2.5.2.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct FibBase {
    /// The flag word at offset 0x0A.
    pub(crate) flags: u16,
    /// `lKey`: the size of the `EncryptionHeader` at the start of the table stream when
    /// `fEncrypted` is set and `fObfuscated` clear; the XOR password verifier when both
    /// are set; otherwise zero.
    pub(crate) l_key: u32,
}

impl FibBase {
    /// The FIB at the start of a `WordDocument` stream, or `None` when the stream is too
    /// short to hold one or does not open with `wIdent`. Anything else is a CFB that
    /// merely happens to carry a stream with this name.
    pub(crate) fn parse(word_document: &[u8]) -> Option<Self> {
        if le16(word_document, 0)? != FIB_W_IDENT {
            return None;
        }
        Some(Self {
            flags: le16(word_document, FIB_FLAGS_AT)?,
            l_key: le32(word_document, FIB_L_KEY_AT)?,
        })
    }

    pub(crate) fn encrypted(self) -> bool {
        self.flags & FIB_F_ENCRYPTED != 0
    }

    pub(crate) fn obfuscated(self) -> bool {
        self.flags & FIB_F_OBFUSCATED != 0
    }

    /// The table stream the FIB refers to — [MS-DOC] §2.5.2 `fWhichTblStm`.
    pub(crate) fn table_stream(self) -> &'static str {
        if self.flags & FIB_F_WHICH_TBL_STM != 0 {
            TABLE_1
        } else {
            TABLE_0
        }
    }
}

/// Offsets of the two FIB fields a decrypt clears: the flag word and `lKey`.
#[cfg(feature = "legacy-binary")]
pub(crate) const FIB_FLAGS_OFFSET: usize = FIB_FLAGS_AT;
#[cfg(feature = "legacy-binary")]
pub(crate) const FIB_L_KEY_OFFSET: usize = FIB_L_KEY_AT;

fn probe_word<F: Read + Seek>(cfb: &mut cfb::CompoundFile<F>) -> Option<BinaryVerdict> {
    let word = read_capped(cfb, WORD_DOCUMENT, BINARY_HEADER_READ_CAP)?;
    let fib = FibBase::parse(&word)?;

    if !fib.encrypted() {
        return Some(BinaryVerdict {
            format: BinaryFormat::Word,
            encrypted: Some(false),
            version: None,
            key_bits: None,
            xor_obfuscated: false,
        });
    }
    if fib.obfuscated() {
        // XOR obfuscation predates the EncryptionHeader; there is nothing further to read.
        return Some(BinaryVerdict {
            format: BinaryFormat::Word,
            encrypted: Some(true),
            version: None,
            key_bits: None,
            xor_obfuscated: true,
        });
    }

    // The encryption header lives at the start of whichever table stream the FIB names.
    let (version, key_bits) = read_capped(cfb, fib.table_stream(), BINARY_HEADER_READ_CAP)
        .map(|b| parse_encryption_header(&b))
        .unwrap_or((None, None));

    Some(BinaryVerdict {
        format: BinaryFormat::Word,
        encrypted: Some(true),
        version,
        key_bits,
        xor_obfuscated: false,
    })
}

// ---- Excel ---------------------------------------------------------------------------

/// `BOF` -- the record every BIFF stream must open with ([MS-XLS] §2.1.7.20).
pub(crate) const BIFF_BOF: u16 = 0x0809;
/// `FILEPASS` -- the encryption marker ([MS-XLS] §2.4.117).
pub(crate) const BIFF_FILEPASS: u16 = 0x002F;
/// `BoundSheet8` -- whose `lbPlyPos` is the one field of an encrypted record that must
/// stay in the clear ([MS-XLS] §2.2.10, §2.4.28).
#[cfg(feature = "legacy-binary")]
pub(crate) const BIFF_BOUNDSHEET8: u16 = 0x0085;

/// The records [MS-XLS] permits **in the clear ahead of `FILEPASS`**, and the records a
/// writer never encrypts at all.
///
/// §2.2.10 makes every record body encrypted and names the exceptions that must never
/// be: `BOF` (§2.4.21), `FilePass` (§2.4.117), `UsrExcl` (§2.4.339), `FileLock`
/// (§2.4.116), `InterfaceHdr` (§2.4.146), `RRDInfo` (§2.4.227) and `RRDHead` (§2.4.226)
/// -- the same set `msoffcrypto-tool` leaves unciphered when it strips a workbook
/// (`format/xls97.py:573-583`). Only these can sit between `BOF` and `FILEPASS`.
///
/// That is what turns "no `FILEPASS` yet" into a *proof* of absence. The moment the walk
/// meets a record outside this set it is reading, in the clear, a record the format
/// requires to be encrypted -- so no `FILEPASS` precedes it, and none can follow. A plain
/// workbook is decided at its third record (`BOF`, `INTERFACEHDR`, then `MMS` at offset
/// 0x1A in one LibreOffice writes) and never by reaching the end of a stream that can run
/// to megabytes of shared strings. That is also why `BIFF_SCAN_CAP` stays a bound on
/// hostile input rather than on how large a legitimate workbook may be.
pub(crate) const BIFF_NEVER_ENCRYPTED: [u16; 7] = [
    BIFF_BOF,
    BIFF_FILEPASS,
    0x00E1,
    0x0138,
    0x0194,
    0x0195,
    0x0196,
];

/// One BIFF record: its type and where its body lies in the stream. The 4-byte header
/// sits immediately before `body`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BiffRecord {
    pub(crate) id: u16,
    pub(crate) body: Range<usize>,
}

/// Where a BIFF walk stopped making sense: a header that cannot be read in full, or a
/// declared length that runs past the end of the stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg(feature = "legacy-binary")]
pub(crate) struct BiffWalkError {
    pub(crate) at: usize,
}

/// The records of a BIFF stream in order — [MS-XLS] §2.1.4: a 2-byte type, a 2-byte
/// length, then that many bytes of body.
///
/// Ends cleanly only at exactly the end of the stream. Bounded: every record advances
/// the position by at least its 4-byte header.
#[cfg(feature = "legacy-binary")]
pub(crate) struct BiffRecords<'a> {
    book: &'a [u8],
    pos: usize,
    failed: bool,
}

#[cfg(feature = "legacy-binary")]
pub(crate) fn biff_records(book: &[u8]) -> BiffRecords<'_> {
    BiffRecords {
        book,
        pos: 0,
        failed: false,
    }
}

#[cfg(feature = "legacy-binary")]
impl Iterator for BiffRecords<'_> {
    type Item = Result<BiffRecord, BiffWalkError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.failed || self.pos >= self.book.len() {
            return None;
        }
        let at = self.pos;
        let (Some(id), Some(len)) = (le16(self.book, at), le16(self.book, at + 2)) else {
            self.failed = true;
            return Some(Err(BiffWalkError { at }));
        };
        let start = at + 4;
        let end = start + usize::from(len);
        if end > self.book.len() {
            self.failed = true;
            return Some(Err(BiffWalkError { at }));
        }
        self.pos = end;
        Some(Ok(BiffRecord {
            id,
            body: start..end,
        }))
    }
}

/// What a walk from `BOF` towards `FILEPASS` decided.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum FilePassScan {
    /// `FILEPASS` reached; here it is.
    Found(BiffRecord),
    /// A record that must be encrypted was met in the clear before any `FILEPASS`, so
    /// there is none: a proof, not an absence of evidence.
    ProvenAbsent,
    /// The stream does not open with `BOF`, a header could not be read, a declared
    /// length overshoots the buffer, or the buffer ended before either verdict.
    Unreadable,
}

/// Walk the record headers from the top of a workbook stream until one of the three
/// verdicts is reached.
///
/// Every way out is one of three, and only one of them is "not encrypted":
///
///   * FILEPASS reached                                    -> `Found`
///   * a record that must be encrypted, met in the clear   -> `ProvenAbsent`
///   * anything else -- a header that cannot be read, a declared length overshooting
///     the buffer, a stream truncated or larger than the cap before either of the
///     above                                               -> `Unreadable`
///
/// The loop this replaced collapsed the last two. Running off the end of the buffer for
/// any reason fell through to "not encrypted", so a 20-byte `/Workbook` -- BOF and
/// nothing after it -- classified as a document needing no password, which is the guess
/// in the attacker's favour this module's doc says it exists to refuse.
///
/// One leniency the strict [`biff_records`] walk does not have, and the reason this is
/// its own loop over the same headers: a `FILEPASS` whose body is cut short -- by
/// truncation or by the read cap -- is still `Found`, with the body clamped to what is
/// there. The marker's meaning does not depend on its body, so that is "encrypted,
/// scheme unread" rather than unknown. A decrypt has no such tolerance: it needs the
/// body, and refuses a record whose declared length overshoots the stream.
pub(crate) fn scan_for_filepass(book: &[u8]) -> FilePassScan {
    // A stream that does not open with BOF is a CFB that merely carries the name --
    // msoffcrypto asserts the same at `format/xls97.py:484`. Its records mean nothing,
    // and neither does the absence of a FILEPASS among them.
    if le16(book, 0) != Some(BIFF_BOF) {
        return FilePassScan::Unreadable;
    }
    // Bounded: every pass either returns or advances `pos` by at least the 4-byte
    // header, and `book` is at most the caller's cap long.
    let mut pos = 0usize;
    loop {
        let (Some(id), Some(len)) = (le16(book, pos), le16(book, pos + 2)) else {
            return FilePassScan::Unreadable;
        };
        let start = pos + 4;
        let end = start + usize::from(len);
        if id == BIFF_FILEPASS {
            return FilePassScan::Found(BiffRecord {
                id,
                body: start..end.min(book.len()),
            });
        }
        if !BIFF_NEVER_ENCRYPTED.contains(&id) {
            return FilePassScan::ProvenAbsent;
        }
        if end > book.len() {
            return FilePassScan::Unreadable;
        }
        pos = end;
    }
}

fn probe_excel<F: Read + Seek>(cfb: &mut cfb::CompoundFile<F>) -> Option<BinaryVerdict> {
    let stream = workbook_stream_name(cfb)?;
    let book = read_capped(cfb, stream, BIFF_SCAN_CAP)?;

    match scan_for_filepass(&book) {
        FilePassScan::Unreadable => Some(unreadable_excel()),
        FilePassScan::ProvenAbsent => Some(BinaryVerdict {
            format: BinaryFormat::Excel,
            encrypted: Some(false),
            version: None,
            key_bits: None,
            xor_obfuscated: false,
        }),
        // FILEPASS is the marker; its body only says which scheme. A body cut short by
        // the cap or by truncation leaves the marker's meaning intact, so that is
        // "encrypted, scheme unread" rather than unknown. wEncryptionType 0 is XOR
        // obfuscation; 1 is an RC4 family described by the EncryptionHeader after it.
        FilePassScan::Found(record) => {
            let body = book.get(record.body).unwrap_or(&[]);
            Some(match le16(body, 0) {
                Some(0) => BinaryVerdict {
                    format: BinaryFormat::Excel,
                    encrypted: Some(true),
                    version: None,
                    key_bits: None,
                    xor_obfuscated: true,
                },
                Some(_) => {
                    let (version, key_bits) = body
                        .get(2..)
                        .map(parse_encryption_header)
                        .unwrap_or((None, None));
                    BinaryVerdict {
                        format: BinaryFormat::Excel,
                        encrypted: Some(true),
                        version,
                        key_bits,
                        xor_obfuscated: false,
                    }
                }
                None => BinaryVerdict {
                    format: BinaryFormat::Excel,
                    encrypted: Some(true),
                    version: None,
                    key_bits: None,
                    xor_obfuscated: false,
                },
            })
        }
    }
}

/// An Excel container recognised as such but whose BIFF stream could not be walked to a
/// verdict: no `BOF`, a header that cannot be read, a length past the buffer, or a stream
/// that ends -- by truncation or by the cap -- before any record decides it.
///
/// `encrypted: None`, never `Some(false)`. The name is still reported: "this is an Excel
/// workbook we could not read" is more useful to a caller than `Unknown`, and it costs
/// nothing to say.
fn unreadable_excel() -> BinaryVerdict {
    BinaryVerdict {
        format: BinaryFormat::Excel,
        encrypted: None,
        version: None,
        key_bits: None,
        xor_obfuscated: false,
    }
}

// ---- PowerPoint ----------------------------------------------------------------------

/// Positioned reads over a stream, so the PowerPoint walk can follow offsets in either
/// a `cfb` stream (detection: a handful of small reads into a deck of any size) or a
/// buffer already in memory (decryption, which rewrites the whole stream anyway).
///
/// `read_at` returns exactly `len` bytes or `None`; a short read is unreadable, not a
/// prefix. Every `len` a caller passes is a record header, a fixed atom, or a length
/// already checked against a [`crate::limits`] cap.
pub(crate) trait ReadAt {
    fn read_at(&mut self, at: u64, len: usize) -> Option<Vec<u8>>;
}

impl ReadAt for &[u8] {
    fn read_at(&mut self, at: u64, len: usize) -> Option<Vec<u8>> {
        let at = usize::try_from(at).ok()?;
        self.get(at..at.checked_add(len)?).map(<[u8]>::to_vec)
    }
}

impl<F: Read + Seek> ReadAt for cfb::Stream<F> {
    fn read_at(&mut self, at: u64, len: usize) -> Option<Vec<u8>> {
        self.seek(SeekFrom::Start(at)).ok()?;
        let mut buf = vec![0u8; len];
        self.read_exact(&mut buf).ok()?;
        Some(buf)
    }
}

/// `RecordType` values — [MS-PPT] §2.13.24. (`RT_CurrentUserAtom`, `0x0FF6`, is not
/// checked: the `Current User` stream holds exactly one record by definition and the
/// atom's offset field is validated by what it points at.)
pub(crate) const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
pub(crate) const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
pub(crate) const RT_CRYPT_SESSION10_CONTAINER: u16 = 0x2F14;

/// `UserEditAtom.rh.recLen` — [MS-PPT] §2.3.3: `0x1C`, or `0x20` when the optional
/// `encryptSessionPersistIdRef` is present, which it must be in an encrypted document.
pub(crate) const USER_EDIT_ATOM_LEN_PLAIN: u32 = 0x1C;
pub(crate) const USER_EDIT_ATOM_LEN_ENCRYPTED: u32 = 0x20;

/// A `RecordHeader` — [MS-PPT] §2.3.1: `recVer`/`recInstance` packed in 2 bytes,
/// `recType`(2), `recLen`(4).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RecordHeader {
    pub(crate) ver_instance: u16,
    pub(crate) rec_type: u16,
    pub(crate) rec_len: u32,
}

impl RecordHeader {
    pub(crate) const LEN: usize = 8;

    pub(crate) fn rec_ver(self) -> u16 {
        self.ver_instance & 0x000F
    }
}

pub(crate) fn record_header(src: &mut impl ReadAt, at: u64) -> Option<RecordHeader> {
    let b = src.read_at(at, RecordHeader::LEN)?;
    Some(RecordHeader {
        ver_instance: le16(&b, 0)?,
        rec_type: le16(&b, 2)?,
        rec_len: le32(&b, 4)?,
    })
}

/// Where `CurrentUserAtom.headerToken` sits — [MS-PPT] §2.3.2, whose layout is
/// `rh`(8) `size`(4) `headerToken`(4) `offsetToCurrentEdit`(4), so the third field is at
/// offset 12.
///
/// `headerToken` is NOT the encryption indicator -- see the module docs.
#[cfg(feature = "legacy-binary")]
pub(crate) const CURRENT_USER_HEADER_TOKEN_AT: usize = 12;

/// `CurrentUserAtom.offsetToCurrentEdit` — the fourth field of the same layout, at
/// offset 16, and the offset the `UserEditAtom` walk starts from.
///
/// Ungated: `classify` follows it in the detection build. The value is attacker-chosen,
/// which is why every reader of it checks the header it lands on rather than the offset
/// itself.
pub(crate) fn current_edit_offset(current_user: &[u8]) -> Option<u32> {
    le32(current_user, 16)
}

/// The `UserEditAtom` fields the encryption paths act on — [MS-PPT] §2.3.3.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct UserEditAtom {
    pub(crate) rec_len: u32,
    /// Offset of the previous user edit's atom, `0` when there is none.
    pub(crate) offset_last_edit: u32,
    pub(crate) offset_persist_directory: u32,
    /// Present only when `rec_len` is `0x20`.
    pub(crate) encrypt_session_persist_id_ref: Option<u32>,
}

#[cfg(feature = "legacy-binary")]
impl UserEditAtom {
    /// Where `rh.recLen` sits, relative to the atom's start — [MS-PPT] §2.3.3.
    pub(crate) const REC_LEN_AT: usize = 4;
    /// Where the optional `encryptSessionPersistIdRef` sits, relative to the same start:
    /// past the 8-byte record header and the atom's first 0x1C bytes. Present only when
    /// `recLen` is `0x20`, which is what the field's own `Option` records.
    pub(crate) const ENCRYPT_SESSION_REF_AT: usize = RecordHeader::LEN + 0x1C;
}

/// Read the `UserEditAtom` at `at`, or `None` unless its header is one.
///
/// The offset that led here is attacker-chosen, so the header is checked before its
/// length is believed: `recVer`/`recInstance` must be `0x0000` and `recType` `0x0FF5`,
/// which is what msoffcrypto asserts before it reads `recLen`
/// (`format/ppt97.py:229-231`). A `recLen` read from an arbitrary byte pair is not
/// evidence of anything, and the check this replaced -- `rec_len == 0x20`, otherwise
/// "not encrypted" -- reported "needs no password" for every offset pointing at anything
/// but an encrypted atom. A length that is neither shape is likewise a header we do not
/// understand.
pub(crate) fn user_edit_atom(src: &mut impl ReadAt, at: u64) -> Option<UserEditAtom> {
    let rh = record_header(src, at)?;
    if rh.ver_instance != 0x0000 || rh.rec_type != RT_USER_EDIT_ATOM {
        return None;
    }
    if rh.rec_len != USER_EDIT_ATOM_LEN_PLAIN && rh.rec_len != USER_EDIT_ATOM_LEN_ENCRYPTED {
        return None;
    }
    // lastSlideIdRef(4) version(2) minorVersion(1) majorVersion(1) offsetLastEdit(4)
    // offsetPersistDirectory(4) docPersistIdRef(4) persistIdSeed(4) lastView(2)
    // unused(2) [encryptSessionPersistIdRef(4)]
    let body = src.read_at(
        at.checked_add(RecordHeader::LEN as u64)?,
        usize::try_from(rh.rec_len).ok()?,
    )?;
    Some(UserEditAtom {
        rec_len: rh.rec_len,
        offset_last_edit: le32(&body, 8)?,
        offset_persist_directory: le32(&body, 12)?,
        encrypt_session_persist_id_ref: if rh.rec_len == USER_EDIT_ATOM_LEN_ENCRYPTED {
            Some(le32(&body, 0x1C)?)
        } else {
            None
        },
    })
}

/// The persist object directory a `PersistDirectoryAtom` at `at` declares, as
/// `persistId -> stream offset` — [MS-PPT] §2.3.4, §2.3.5.
///
/// Each entry packs a 20-bit starting `persistId` and a 12-bit `cPersist` into one
/// `u32`, followed by `cPersist` 4-byte offsets for consecutive identifiers. A later
/// entry naming an identifier already seen replaces its offset (§2.1.2 part 1, step 8c).
///
/// Bounded twice: the atom's `recLen` by [`PPT_PERSIST_DIRECTORY_READ_CAP`], and the
/// number of offsets by [`PPT_PERSIST_OBJECTS_MAX`], the 20-bit identifier space.
pub(crate) fn persist_directory(src: &mut impl ReadAt, at: u64) -> Option<BTreeMap<u32, u32>> {
    let rh = record_header(src, at)?;
    if rh.ver_instance != 0x0000 || rh.rec_type != RT_PERSIST_DIRECTORY_ATOM {
        return None;
    }
    let len = usize::try_from(rh.rec_len).ok()?;
    if len > PPT_PERSIST_DIRECTORY_READ_CAP {
        return None;
    }
    let entries = src.read_at(at.checked_add(RecordHeader::LEN as u64)?, len)?;

    let mut directory = BTreeMap::new();
    let mut objects = 0usize;
    let mut pos = 0usize;
    while pos < entries.len() {
        let word = le32(&entries, pos)?;
        let first_id = word & 0x000F_FFFF;
        let count = usize::try_from(word >> 20).ok()?;
        // [MS-PPT] §2.3.5: cPersist MUST be at least 1. A zero count would also make
        // this loop spin on the same 4 bytes.
        if count == 0 {
            return None;
        }
        objects = objects.checked_add(count)?;
        if objects > PPT_PERSIST_OBJECTS_MAX {
            return None;
        }
        pos = pos.checked_add(4)?;
        for i in 0..count {
            let offset = le32(&entries, pos)?;
            let id = first_id.checked_add(u32::try_from(i).ok()?)?;
            directory.insert(id, offset);
            pos = pos.checked_add(4)?;
        }
    }
    Some(directory)
}

/// The `data` of the `CryptSession10Container` at `at` — the RC4 CryptoAPI encryption
/// header structure ([MS-PPT] §2.3.7, [MS-OFFCRYPTO] §2.3.5.1) — or `None` unless the
/// record header is one.
///
/// `recVer` must be `0xF`; `recInstance` is not checked, because msoffcrypto found real
/// files failing the spec's `0x000` (`format/ppt97.py:439-442`) and PowerPoint opens
/// them. The length is bounded by [`ENCRYPTION_HEADER_STRUCTURE_MAX`] before it is read.
pub(crate) fn crypt_session_container(src: &mut impl ReadAt, at: u64) -> Option<Vec<u8>> {
    let rh = record_header(src, at)?;
    if rh.rec_ver() != 0xF || rh.rec_type != RT_CRYPT_SESSION10_CONTAINER {
        return None;
    }
    let len = usize::try_from(rh.rec_len).ok()?;
    if len > ENCRYPTION_HEADER_STRUCTURE_MAX {
        return None;
    }
    src.read_at(at.checked_add(RecordHeader::LEN as u64)?, len)
}

/// Follow an encrypted presentation's `UserEditAtom` to its encryption header structure.
///
/// The route [MS-PPT] §2.3.7 lays out: `encryptSessionPersistIdRef` is looked up in the
/// persist directory the atom names, and the persist object there is the
/// `CryptSession10Container` whose `data` is the header.
pub(crate) fn presentation_encryption_header(
    src: &mut impl ReadAt,
    atom: &UserEditAtom,
) -> Option<Vec<u8>> {
    let id = atom.encrypt_session_persist_id_ref?;
    let directory = persist_directory(src, u64::from(atom.offset_persist_directory))?;
    let offset = *directory.get(&id)?;
    crypt_session_container(src, u64::from(offset))
}

fn probe_powerpoint<F: Read + Seek>(cfb: &mut cfb::CompoundFile<F>) -> Option<BinaryVerdict> {
    let cu = read_capped(cfb, CURRENT_USER, CURRENT_USER_READ_CAP)?;
    let offset = current_edit_offset(&cu)?;

    // The offset is attacker-controlled and indexes the document stream. Seek to it and
    // read one atom; nothing here materialises the stream.
    let mut doc = cfb.open_stream(POWERPOINT_DOCUMENT).ok()?;
    let Some(atom) = user_edit_atom(&mut doc, u64::from(offset)) else {
        return Some(unreadable_powerpoint());
    };

    // An encrypted presentation's atom carries an extra encryptSessionPersistIdRef
    // field, taking recLen to 0x20 where a plain one is 0x1C (msoffcrypto-tool
    // format/ppt97.py:834). `user_edit_atom` admits no third length.
    if atom.rec_len == USER_EDIT_ATOM_LEN_PLAIN {
        return Some(BinaryVerdict {
            format: BinaryFormat::PowerPoint,
            encrypted: Some(false),
            version: None,
            key_bits: None,
            xor_obfuscated: false,
        });
    }

    // Encrypted. The header sits inside a CryptSession10Container reached through the
    // persist directory rather than at a fixed offset; reading it is best effort, and
    // a walk that fails leaves the verdict "encrypted, scheme unread".
    let (version, key_bits) = presentation_encryption_header(&mut doc, &atom)
        .map(|h| parse_encryption_header(&h))
        .unwrap_or((None, None));

    Some(BinaryVerdict {
        format: BinaryFormat::PowerPoint,
        encrypted: Some(true),
        version,
        key_bits,
        xor_obfuscated: false,
    })
}

/// A PowerPoint container recognised as such but whose edit atom could not be reached.
///
/// `encrypted: None`, not `Some(false)`. Claiming a file is unprotected because we failed
/// to read it is the error this module exists to avoid, and an attacker choosing the
/// `offsetToCurrentEdit` is exactly who would want that answer.
fn unreadable_powerpoint() -> BinaryVerdict {
    BinaryVerdict {
        format: BinaryFormat::PowerPoint,
        encrypted: None,
        version: None,
        key_bits: None,
        xor_obfuscated: false,
    }
}
