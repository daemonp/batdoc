//! Every number a `.doc`, `.xls` or `.ppt` declares about itself must produce an error —
//! never a panic, never a hang — and every refusal must name what it refused.
//!
//! The binary formats are worse than the OOXML ones in this respect: their offsets and
//! lengths index the streams the decrypt rewrites, and an unbounded one is a slice panic
//! or, for PowerPoint's persist objects, a decrypt of the whole stream repeated once per
//! directory entry. Each test here poisons one field, asserts the variant, and carries a
//! control differing only in that field, so a green result cannot come from the synthetic
//! container being refused for an unrelated reason (CLAUDE.md § Testing Rules).
//!
//! Two kinds of input. Tampered fixtures: the real Office-written files with one field
//! rewritten in place through `cfb`, which is the cheapest way to reach a parser with
//! everything else valid. Synthetic streams: built from record headers where the shape
//! under test is not one a fixture has. Neither is committed, for the reason
//! `malformed_input.rs` gives.

use crate::{decrypt_binary_office, Error};
use std::io::{Cursor, Read, Seek, SeekFrom, Write};

const PASSWORD: &str = "testpass";

fn fixture(name: &str) -> Vec<u8> {
    let path = format!("{}/tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"));
    std::fs::read(&path).unwrap_or_else(|e| panic!("fixture {name} must be present: {e}"))
}

/// `data` with `bytes` written at `at` inside stream `name`, nothing else changed.
fn patch_stream(data: &[u8], name: &str, at: u64, bytes: &[u8]) -> Vec<u8> {
    let mut cursor = Cursor::new(data.to_vec());
    {
        let mut cfb = cfb::CompoundFile::open(&mut cursor).expect("a CFB container");
        let mut stream = cfb.open_stream(name).expect("the stream exists");
        stream.seek(SeekFrom::Start(at)).unwrap();
        stream.write_all(bytes).unwrap();
        stream.flush().unwrap();
    }
    cursor.into_inner()
}

/// One stream of a container, whole.
fn read_stream(data: &[u8], name: &str) -> Vec<u8> {
    let mut cfb = cfb::CompoundFile::open(Cursor::new(data)).expect("a CFB container");
    let mut out = Vec::new();
    cfb.open_stream(name)
        .expect("the stream exists")
        .read_to_end(&mut out)
        .unwrap();
    out
}

/// A CFB container holding the named streams, and nothing else.
fn cfb_with(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut container = cfb::CompoundFile::create(&mut cursor).unwrap();
        for (name, bytes) in streams {
            let mut stream = container.create_stream(name).unwrap();
            stream.write_all(bytes).unwrap();
            stream.flush().unwrap();
        }
        container.flush().unwrap();
    }
    cursor.into_inner()
}

/// A BIFF stream: each record is a 2-byte id, a 2-byte body length, then the body.
fn biff(records: &[(u16, &[u8])]) -> Vec<u8> {
    let mut out = Vec::new();
    for (id, body) in records {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(body.len() as u16).to_le_bytes());
        out.extend_from_slice(body);
    }
    out
}

fn bad_parameters_naming(got: &Result<Vec<u8>, Error>, needle: &str) -> bool {
    matches!(got, Err(Error::BadParameters(m)) if m.contains(needle))
}

// ---- Word ------------------------------------------------------------------------------
//
// word97_password.doc, measured with olefile: FIB flags 0x13F0 at offset 0x0A, lKey 0xC6
// at 0x0E; the 1Table stream is 9 754 bytes and opens with the 198-byte encryption
// header structure, whose EncryptionHeader.KeySize sits at 12 + 16 = 28.

const DOC: &str = "word97_password.doc";
const FIB_FLAGS_AT: u64 = 0x0A;
const FIB_L_KEY_AT: u64 = 0x0E;
const DOC_TABLE_LEN: u32 = 9_754;
const TABLE_KEY_SIZE_AT: u64 = 28;

/// `FibBase.lKey` is a length the file declares over the table stream. Past the stream,
/// past the structure cap, and too short to hold a version are each refused by name;
/// the untouched value is the control.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn word_l_key_is_bounded_before_it_slices_the_table_stream() {
    let doc = fixture(DOC);
    for (what, l_key, needle) in [
        ("u32::MAX", u32::MAX, "over the"),
        (
            "over the structure cap",
            crate::limits::ENCRYPTION_HEADER_STRUCTURE_MAX as u32 + 1,
            "over the",
        ),
        ("too short for a version", 3, "too small"),
    ] {
        let tampered = patch_stream(&doc, "/WordDocument", FIB_L_KEY_AT, &l_key.to_le_bytes());
        let got = decrypt_binary_office(&tampered, PASSWORD);
        assert!(bad_parameters_naming(&got, needle), "{what}: {got:?}");
    }
    // Under the cap but past the stream: the fixture's own lKey over a table stream cut
    // to 100 bytes. The cap cannot catch this one; only the check against the stream can.
    let word = read_stream(&doc, "/WordDocument");
    let short_table = read_stream(&doc, "/1Table")[..100].to_vec();
    let got = decrypt_binary_office(
        &cfb_with(&[("/WordDocument", &word), ("/1Table", &short_table)]),
        PASSWORD,
    );
    assert!(
        bad_parameters_naming(&got, "table stream holds 100"),
        "{got:?}"
    );
    // Control: the fixture decrypts as it is, so the refusals above are the field's.
    assert!(decrypt_binary_office(&doc, PASSWORD).is_ok());
    let _ = DOC_TABLE_LEN;
    // And an lKey that is merely *larger* than the structure but within the stream is
    // accepted: the parser reads the structure by its own size fields.
    let generous = patch_stream(&doc, "/WordDocument", FIB_L_KEY_AT, &400u32.to_le_bytes());
    assert!(decrypt_binary_office(&generous, PASSWORD).is_ok());
}

/// The two FIB bits the format decides encryption by, each flipped alone.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn word_fib_flags_decide_not_encrypted_and_xor_by_name() {
    let doc = fixture(DOC);
    let flags = 0x13F0u16;
    // fEncrypted clear: nothing to decrypt, whatever the table stream holds.
    let plain = patch_stream(
        &doc,
        "/WordDocument",
        FIB_FLAGS_AT,
        &(flags & !0x0100).to_le_bytes(),
    );
    assert!(matches!(
        decrypt_binary_office(&plain, PASSWORD),
        Err(Error::NotEncrypted)
    ));
    // fObfuscated set: Word's XOR method, refused by name and never as a wrong password.
    let xor = patch_stream(
        &doc,
        "/WordDocument",
        FIB_FLAGS_AT,
        &(flags | 0x8000).to_le_bytes(),
    );
    assert!(matches!(
        decrypt_binary_office(&xor, PASSWORD),
        Err(Error::UnsupportedAlgorithm {
            what: "FibBase.fObfuscated",
            ..
        })
    ));
}

/// `EncryptionHeader.KeySize` is authenticated by the verifier and bounded before it
/// derives a key: 0 reads as 40 and then fails the password check (the file was written
/// at 128), a value off the grid is refused by name.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn word_key_size_is_bounded_and_verified() {
    let doc = fixture(DOC);
    let forty = patch_stream(&doc, "/1Table", TABLE_KEY_SIZE_AT, &0u32.to_le_bytes());
    assert!(matches!(
        decrypt_binary_office(&forty, PASSWORD),
        Err(Error::WrongPassword)
    ));
    for bits in [8u32, 136, 1024] {
        let tampered = patch_stream(&doc, "/1Table", TABLE_KEY_SIZE_AT, &bits.to_le_bytes());
        let got = decrypt_binary_office(&tampered, PASSWORD);
        assert!(
            bad_parameters_naming(&got, "KeySize"),
            "KeySize {bits}: {got:?}"
        );
    }
    // The version pair, refused by name.
    let versioned = patch_stream(&doc, "/1Table", 0, &[5, 0, 2, 0]);
    assert!(matches!(
        decrypt_binary_office(&versioned, PASSWORD),
        Err(Error::UnsupportedEncryptionVersion(5, 2))
    ));
}

/// A `WordDocument` stream shorter than the FIB, and one that is not a FIB at all.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn word_document_shorter_than_the_fib_is_refused() {
    let table = read_stream(&fixture(DOC), "/1Table");
    let mut short = read_stream(&fixture(DOC), "/WordDocument");
    short.truncate(0x30);
    let got = decrypt_binary_office(
        &cfb_with(&[("/WordDocument", &short), ("/1Table", &table)]),
        PASSWORD,
    );
    assert!(
        matches!(got, Err(Error::MissingStream(m)) if m.contains("68-byte")),
        "{got:?}"
    );

    let not_a_fib = vec![0u8; 0x44];
    let got = decrypt_binary_office(
        &cfb_with(&[("/WordDocument", &not_a_fib), ("/1Table", &table)]),
        PASSWORD,
    );
    assert!(bad_parameters_naming(&got, "wIdent"), "{got:?}");
}

/// Office 97/2000 RC4 end to end, on a container built with a test-side writer over the
/// production key schedule.
///
/// No fixture exists (see `rc4_office97`'s docs for why); the schedule's KDF and
/// verifier are pinned against msoffcrypto's vector there, so what this adds is the
/// wiring: the `1.1` dispatch in `word97`, the 512-byte block loop over three streams,
/// the clear FIB put back with its flags and `lKey` cleared. Self-agreement only for
/// the container layout, and labelled so.
#[test]
fn word_office97_rc4_container_decrypts_and_reports_a_wrong_password() {
    use crate::rc4::{decrypt_in_blocks, BlockKeySchedule, Keystream};
    use crate::rc4_office97::Office97KeySchedule;
    use md5::{Digest, Md5};

    let salt = [0x11u8; 16];
    let schedule = Office97KeySchedule::new(PASSWORD, &salt);

    // Verifier and its MD5, both under the block-0 keystream, continued.
    let verifier = [0x5Au8; 16];
    let mut keystream = Keystream::new(&schedule.block_key(0)).unwrap();
    let mut encrypted_verifier = verifier;
    keystream.apply(&mut encrypted_verifier);
    let mut encrypted_hash: [u8; 16] = Md5::digest(verifier).into();
    keystream.apply(&mut encrypted_hash);

    let mut header = vec![1u8, 0, 1, 0];
    header.extend_from_slice(&salt);
    header.extend_from_slice(&encrypted_verifier);
    header.extend_from_slice(&encrypted_hash);
    assert_eq!(header.len(), 52);

    // A FIB with fEncrypted | fWhichTblStm and lKey = 52, then 1 000 bytes of "document".
    let mut fib = vec![0u8; 0x44];
    fib[..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib[0x0A..0x0C].copy_from_slice(&0x0300u16.to_le_bytes());
    fib[0x0E..0x12].copy_from_slice(&52u32.to_le_bytes());
    let body: Vec<u8> = (0..1000u32).map(|i| (i % 251) as u8).collect();

    let mut word = fib.clone();
    word.extend_from_slice(&body);
    let mut table = header.clone();
    table.extend_from_slice(&body);
    let mut data = body.clone();
    // Encrypt whole (RC4 is symmetric), then restore the clear FIB and header.
    for stream in [&mut word, &mut table, &mut data] {
        decrypt_in_blocks(&schedule, stream, 0x200).unwrap();
    }
    word[..0x44].copy_from_slice(&fib);
    table[..52].copy_from_slice(&header);

    let container = cfb_with(&[
        ("/WordDocument", &word),
        ("/1Table", &table),
        ("/Data", &data),
    ]);
    let plain = decrypt_binary_office(&container, PASSWORD).expect("must decrypt");

    let word_plain = read_stream(&plain, "/WordDocument");
    assert_eq!(&word_plain[0x44..], &body[..], "the document body decrypts");
    assert_eq!(
        &word_plain[..2],
        &0xA5ECu16.to_le_bytes(),
        "the FIB survives"
    );
    assert_eq!(
        &word_plain[0x0A..0x0C],
        &0x0200u16.to_le_bytes(),
        "fEncrypted cleared, fWhichTblStm kept"
    );
    assert_eq!(&word_plain[0x0E..0x12], &[0, 0, 0, 0], "lKey cleared");
    assert_eq!(
        &read_stream(&plain, "/1Table")[52..],
        &body[..],
        "the table stream decrypts past the header"
    );
    assert_eq!(
        read_stream(&plain, "/Data"),
        body,
        "the Data stream decrypts whole"
    );

    assert!(matches!(
        decrypt_binary_office(&container, "not-the-password"),
        Err(Error::WrongPassword)
    ));
}

// ---- Excel -----------------------------------------------------------------------------
//
// excel97_password.xls: Workbook opens BOF (16-byte body), then FILEPASS at offset 20 with
// a 200-byte body: wEncryptionType 1, then the same 198-byte structure the .doc carries.
// excel97_xor.xls: FILEPASS at 20 with a 6-byte body: type 0, key 0xA6CE, verifier 0x9727.

const XLS: &str = "excel97_password.xls";
const XLS_XOR: &str = "excel97_xor.xls";
const BOF: u16 = 0x0809;
const FILEPASS: u16 = 0x002F;
const INTERFACEHDR: u16 = 0x00E1;
const MMS: u16 = 0x00C1;
const BOF_BODY: [u8; 16] = [
    0x00, 0x06, 0x05, 0x00, 0x5A, 0x4F, 0xCD, 0x07, 0xC1, 0x00, 0x02, 0x00, 0x06, 0x08, 0x00, 0x00,
];

/// The fixture's own FILEPASS body: public bytes, and the only way to build a synthetic
/// workbook that passes the verifier under `testpass`.
fn rc4_filepass_body() -> Vec<u8> {
    let book = read_stream(&fixture(XLS), "/Workbook");
    assert_eq!(&book[20..22], &FILEPASS.to_le_bytes());
    let len = u16::from_le_bytes([book[22], book[23]]) as usize;
    book[24..24 + len].to_vec()
}

/// A record whose declared length overshoots the stream, and a stream cut inside a
/// header: both refused by name, before any decrypt. The control is the same records
/// with the length honest.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn excel_record_length_past_the_stream_is_refused() {
    let filepass = rc4_filepass_body();
    let honest = biff(&[(BOF, &BOF_BODY), (FILEPASS, &filepass), (MMS, &[0, 0])]);
    assert!(decrypt_binary_office(&cfb_with(&[("/Workbook", &honest)]), PASSWORD).is_ok());

    let mut overshoot = biff(&[(BOF, &BOF_BODY), (FILEPASS, &filepass)]);
    overshoot.extend_from_slice(&MMS.to_le_bytes());
    overshoot.extend_from_slice(&0x4000u16.to_le_bytes()); // claims 16 KiB, has none
    let got = decrypt_binary_office(&cfb_with(&[("/Workbook", &overshoot)]), PASSWORD);
    assert!(
        bad_parameters_naming(&got, "past the end of the stream"),
        "{got:?}"
    );

    let mut cut = biff(&[(BOF, &BOF_BODY), (FILEPASS, &filepass)]);
    cut.extend_from_slice(&[0xC1]); // one byte of the next record id
    let got = decrypt_binary_office(&cfb_with(&[("/Workbook", &cut)]), PASSWORD);
    assert!(bad_parameters_naming(&got, "no readable header"), "{got:?}");

    // A FILEPASS whose own body is cut short: detection still says "encrypted", but a
    // decrypt needs the body and refuses.
    let mut short = biff(&[(BOF, &BOF_BODY)]);
    short.extend_from_slice(&FILEPASS.to_le_bytes());
    short.extend_from_slice(&200u16.to_le_bytes());
    let got = decrypt_binary_office(&cfb_with(&[("/Workbook", &short)]), PASSWORD);
    assert!(
        bad_parameters_naming(&got, "past the end of the stream"),
        "{got:?}"
    );
}

/// FILEPASS after a record the format requires to be encrypted is neither an encrypted
/// workbook nor a plain one, and is refused by name; a workbook with no FILEPASS at all
/// is not encrypted. The control is FILEPASS after the records the format permits ahead
/// of it, which decrypts.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn excel_filepass_out_of_order_is_refused_and_absence_is_not_encrypted() {
    let filepass = rc4_filepass_body();
    let in_order = biff(&[
        (BOF, &BOF_BODY),
        (INTERFACEHDR, &[0xB0, 0x04]),
        (FILEPASS, &filepass),
        (MMS, &[0, 0]),
    ]);
    assert!(decrypt_binary_office(&cfb_with(&[("/Workbook", &in_order)]), PASSWORD).is_ok());
    assert!(matches!(
        decrypt_binary_office(&cfb_with(&[("/Workbook", &in_order)]), "wrongpass"),
        Err(Error::WrongPassword)
    ));

    let out_of_order = biff(&[(BOF, &BOF_BODY), (MMS, &[0, 0]), (FILEPASS, &filepass)]);
    let got = decrypt_binary_office(&cfb_with(&[("/Workbook", &out_of_order)]), PASSWORD);
    assert!(
        bad_parameters_naming(&got, "FILEPASS follows record 0x00c1"),
        "{got:?}"
    );

    let plain = biff(&[
        (BOF, &BOF_BODY),
        (INTERFACEHDR, &[0xB0, 0x04]),
        (MMS, &[0, 0]),
    ]);
    assert!(matches!(
        decrypt_binary_office(&cfb_with(&[("/Workbook", &plain)]), PASSWORD),
        Err(Error::NotEncrypted)
    ));

    let not_bof = biff(&[(MMS, &[0, 0]), (FILEPASS, &filepass)]);
    let got = decrypt_binary_office(&cfb_with(&[("/Workbook", &not_bof)]), PASSWORD);
    assert!(bad_parameters_naming(&got, "open with BOF"), "{got:?}");
}

/// The FILEPASS body's own discriminators: an `wEncryptionType` that is neither scheme,
/// an XOR body too short for its two words, an RC4 body with no version, a version pair
/// this crate does not read, and a BIFF5 `Book` stream — each refused by name.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn excel_filepass_shapes_are_refused_by_name() {
    let mut rc4 = rc4_filepass_body();
    let cases: Vec<(&str, Vec<u8>)> = vec![
        ("wEncryptionType 2", {
            let mut b = rc4.clone();
            b[0] = 2;
            b
        }),
        ("XOR with one word", vec![0, 0, 0xCE, 0xA6]),
        ("RC4 with no version", vec![1, 0]),
    ];
    for (what, body) in cases {
        let got = decrypt_binary_office(
            &cfb_with(&[("/Workbook", &biff(&[(BOF, &BOF_BODY), (FILEPASS, &body)]))]),
            PASSWORD,
        );
        assert!(
            matches!(got, Err(Error::BadParameters(_))),
            "{what}: {got:?}"
        );
    }
    rc4[2] = 5; // vMajor 5
    let got = decrypt_binary_office(
        &cfb_with(&[("/Workbook", &biff(&[(BOF, &BOF_BODY), (FILEPASS, &rc4)]))]),
        PASSWORD,
    );
    assert!(
        matches!(got, Err(Error::UnsupportedEncryptionVersion(5, 2))),
        "{got:?}"
    );

    let book = read_stream(&fixture(XLS), "/Workbook");
    let got = decrypt_binary_office(&cfb_with(&[("/Book", &book)]), PASSWORD);
    assert!(
        matches!(
            got,
            Err(Error::UnsupportedAlgorithm {
                what: "Book stream",
                ..
            })
        ),
        "{got:?}"
    );
}

/// XOR's verifier is two words, and both are checked: a key that does not match the
/// password is refused even beside a verifier that does, and vice versa.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn excel_xor_key_and_verifier_are_both_checked() {
    let xls = fixture(XLS_XOR);
    let book = read_stream(&xls, "/Workbook");
    assert_eq!(
        &book[24..30],
        &[0, 0, 0xCE, 0xA6, 0x27, 0x97],
        "the fixture's FILEPASS"
    );
    assert!(decrypt_binary_office(&xls, PASSWORD).is_ok());

    let bad_key = patch_stream(&xls, "/Workbook", 26, &0xA6CFu16.to_le_bytes());
    assert!(matches!(
        decrypt_binary_office(&bad_key, PASSWORD),
        Err(Error::WrongPassword)
    ));
    let bad_verifier = patch_stream(&xls, "/Workbook", 28, &0x9726u16.to_le_bytes());
    assert!(matches!(
        decrypt_binary_office(&bad_verifier, PASSWORD),
        Err(Error::WrongPassword)
    ));
}

// ---- PowerPoint ------------------------------------------------------------------------

const PPT: &str = "powerpoint97_password.ppt";
const PPT_DOC: &str = "/PowerPoint Document";
const PPT_CURRENT_USER: &str = "/Current User";

/// Where the structures these tests tamper with actually are, **read out of the fixture**
/// rather than written down.
///
/// These were five hardcoded constants measured with olefile, and they were right until
/// the fixture was regenerated to clear its author metadata: the document stream went from
/// 202 461 bytes to 36 090, and every one of them pointed past the end. The tests failed
/// with `Cannot seek to 202179 bytes from start, because stream length is only 36090`,
/// which is a fixture-shaped failure wearing a bug's clothes.
///
/// They are file-derived numbers, so the file is what should derive them. The walk is the
/// one [MS-PPT] specifies and the one `binary_office::persist_directory` makes in
/// production: `CurrentUserAtom.offsetToCurrentEdit` (§2.3.2) locates the `UserEditAtom`
/// (§2.3.3), whose `offsetPersistDirectory` locates the `PersistDirectoryAtom` (§2.3.4),
/// whose entries map a persist id to a stream offset; the container this file encrypts
/// with is the one `encryptSessionPersistIdRef` names.
struct PptOffsets {
    /// The whole `PowerPoint Document` stream — one past the last byte a test may patch.
    doc_len: u32,
    /// The `UserEditAtom`. Its record header is 8 bytes, so its fields start at `+8`.
    edit_at: u64,
    /// The `PersistDirectoryAtom`. Record header 8 bytes, then one entry word.
    directory_at: u64,
    /// The first `rgPersistOffset` in that entry: header 8 + the entry word 4.
    first_offset_at: u64,
    /// `EncryptionHeader.KeySize` inside the `CryptSession10Container`: the container's
    /// record header is 8 bytes, `EncryptionVersionInfo` 4, `Flags` 4 and
    /// `EncryptionHeaderSize` 4, then the header itself, whose `KeySize` is 16 in —
    /// container + 36, measured 128 on this fixture.
    key_size_at: u64,
}

fn le32(bytes: &[u8], at: usize) -> u32 {
    let mut b = [0u8; 4];
    b.copy_from_slice(&bytes[at..at + 4]);
    u32::from_le_bytes(b)
}

fn ppt_offsets(ppt: &[u8]) -> PptOffsets {
    let current_user = read_stream(ppt, PPT_CURRENT_USER);
    let doc = read_stream(ppt, PPT_DOC);

    // CurrentUserAtom: record header 8, size 4, headerToken 4, then offsetToCurrentEdit.
    let edit_at = le32(&current_user, 16) as usize;
    assert_eq!(
        u16::from_le_bytes([doc[edit_at + 2], doc[edit_at + 3]]),
        0x0FF5,
        "offsetToCurrentEdit must land on a UserEditAtom"
    );
    // UserEditAtom fields, past its 8-byte header: lastSlideIdRef 4, version 2, minor 1,
    // major 1, offsetLastEdit 4, then offsetPersistDirectory.
    let directory_at = le32(&doc, edit_at + 8 + 12) as usize;
    assert_eq!(
        u16::from_le_bytes([doc[directory_at + 2], doc[directory_at + 3]]),
        0x1772,
        "offsetPersistDirectory must land on a PersistDirectoryAtom"
    );
    // ... and, 0x1C past the header, the persist id of the encryption container.
    let encrypt_ref = le32(&doc, edit_at + 8 + 0x1C);

    // Walk the directory for that id. Each entry is a packed word — persistId in the low
    // 20 bits, a run length in the high 12 — followed by that many 4-byte offsets.
    let dir_end = directory_at + 8 + le32(&doc, directory_at + 4) as usize;
    let mut at = directory_at + 8;
    let mut container = None;
    while at + 4 <= dir_end {
        let word = le32(&doc, at);
        at += 4;
        let (first_id, run) = (word & 0xF_FFFF, word >> 20);
        for i in 0..run {
            if at + 4 > dir_end {
                break;
            }
            if first_id + i == encrypt_ref {
                container = Some(le32(&doc, at) as usize);
            }
            at += 4;
        }
    }
    let container = container.expect("the persist directory names the encryption container");
    assert_eq!(
        u16::from_le_bytes([doc[container + 2], doc[container + 3]]),
        0x2F14,
        "encryptSessionPersistIdRef must land on a CryptSession10Container"
    );

    PptOffsets {
        doc_len: doc.len() as u32,
        edit_at: edit_at as u64,
        directory_at: directory_at as u64,
        first_offset_at: directory_at as u64 + 12,
        key_size_at: container as u64 + 36,
    }
}

#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn powerpoint_offsets_are_bounded_before_they_are_followed() {
    let ppt = fixture(PPT);
    let o = ppt_offsets(&ppt);
    assert!(decrypt_binary_office(&ppt, PASSWORD).is_ok(), "control");

    // offsetToCurrentEdit past the stream, and at a byte that is not a UserEditAtom.
    for (what, offset) in [
        ("past the stream", o.doc_len),
        ("u32::MAX", u32::MAX),
        ("mid-stream garbage", 1000),
    ] {
        let tampered = patch_stream(&ppt, "/Current User", 16, &offset.to_le_bytes());
        let got = decrypt_binary_office(&tampered, PASSWORD);
        assert!(
            bad_parameters_naming(&got, "no UserEditAtom"),
            "{what}: {got:?}"
        );
    }

    // offsetPersistDirectory at something that is not a PersistDirectoryAtom.
    let tampered = patch_stream(&ppt, PPT_DOC, o.edit_at + 8 + 12, &0u32.to_le_bytes());
    let got = decrypt_binary_office(&tampered, PASSWORD);
    assert!(
        bad_parameters_naming(&got, "no readable PersistDirectoryAtom"),
        "{got:?}"
    );

    // A PersistDirectoryAtom whose recLen exceeds the read cap.
    let tampered = patch_stream(&ppt, PPT_DOC, o.directory_at + 4, &u32::MAX.to_le_bytes());
    let got = decrypt_binary_office(&tampered, PASSWORD);
    assert!(
        bad_parameters_naming(&got, "no readable PersistDirectoryAtom"),
        "{got:?}"
    );

    // An entry with cPersist = 0, which [MS-PPT] §2.3.5 forbids: the atom is unreadable,
    // not a directory of nothing whose offsets are then parsed as further entries.
    let tampered = patch_stream(&ppt, PPT_DOC, o.directory_at + 8, &1u32.to_le_bytes());
    let got = decrypt_binary_office(&tampered, PASSWORD);
    assert!(
        bad_parameters_naming(&got, "no readable PersistDirectoryAtom"),
        "{got:?}"
    );

    // A persist object offset at or past the directory — the bound that keeps every
    // object's extent inside the stream and disjoint from the atoms after it.
    for offset in [o.directory_at as u32, o.doc_len, u32::MAX] {
        let tampered = patch_stream(&ppt, PPT_DOC, o.first_offset_at, &offset.to_le_bytes());
        let got = decrypt_binary_office(&tampered, PASSWORD);
        assert!(
            bad_parameters_naming(&got, "at or beyond the persist directory"),
            "offset {offset}: {got:?}"
        );
    }

    // encryptSessionPersistIdRef naming an object the directory does not have.
    let tampered = patch_stream(&ppt, PPT_DOC, o.edit_at + 8 + 0x1C, &999u32.to_le_bytes());
    let got = decrypt_binary_office(&tampered, PASSWORD);
    assert!(
        bad_parameters_naming(&got, "not in the persist directory"),
        "{got:?}"
    );

    // ...or one that is not a CryptSession10Container (persist object 1 is the
    // DocumentContainer, encrypted, so its header is noise).
    let tampered = patch_stream(&ppt, PPT_DOC, o.edit_at + 8 + 0x1C, &1u32.to_le_bytes());
    let got = decrypt_binary_office(&tampered, PASSWORD);
    assert!(
        bad_parameters_naming(&got, "no CryptSession10Container"),
        "{got:?}"
    );
}

/// [MS-PPT] §2.3.7: an encrypted presentation contains exactly one user edit. A second
/// is refused rather than half-decrypted; the plain atom shape is "not encrypted".
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn powerpoint_user_edit_shapes_decide_by_name() {
    let ppt = fixture(PPT);
    let o = ppt_offsets(&ppt);
    let chained = patch_stream(&ppt, PPT_DOC, o.edit_at + 8 + 8, &1u32.to_le_bytes());
    let got = decrypt_binary_office(&chained, PASSWORD);
    assert!(
        bad_parameters_naming(&got, "exactly one user edit"),
        "{got:?}"
    );

    let plain = patch_stream(&ppt, PPT_DOC, o.edit_at + 4, &0x1Cu32.to_le_bytes());
    assert!(matches!(
        decrypt_binary_office(&plain, PASSWORD),
        Err(Error::NotEncrypted)
    ));

    let odd = patch_stream(&ppt, PPT_DOC, o.edit_at + 4, &0x30u32.to_le_bytes());
    let got = decrypt_binary_office(&odd, PASSWORD);
    assert!(bad_parameters_naming(&got, "no UserEditAtom"), "{got:?}");
}

/// The header inside the container is parsed with the same bounds as Word's: KeySize 0
/// reads as 40 and fails the verifier, a value off the grid is refused by name.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn powerpoint_key_size_is_bounded_and_verified() {
    let ppt = fixture(PPT);
    let o = ppt_offsets(&ppt);
    let forty = patch_stream(&ppt, PPT_DOC, o.key_size_at, &0u32.to_le_bytes());
    assert!(matches!(
        decrypt_binary_office(&forty, PASSWORD),
        Err(Error::WrongPassword)
    ));
    let off_grid = patch_stream(&ppt, PPT_DOC, o.key_size_at, &200u32.to_le_bytes());
    let got = decrypt_binary_office(&off_grid, PASSWORD);
    assert!(bad_parameters_naming(&got, "KeySize"), "{got:?}");
}

/// A container that is a CFB but none of the three formats, and one that is not a CFB.
#[test]
fn something_that_is_not_a_binary_document_is_named() {
    assert!(matches!(
        decrypt_binary_office(&fixture("agile_encrypted.docx"), PASSWORD),
        Err(Error::MissingStream(_))
    ));
    assert!(matches!(
        decrypt_binary_office(&fixture("plain.docx"), PASSWORD),
        Err(Error::NotACfbFile)
    ));
    assert!(matches!(
        decrypt_binary_office(b"", PASSWORD),
        Err(Error::NotACfbFile)
    ));
}
