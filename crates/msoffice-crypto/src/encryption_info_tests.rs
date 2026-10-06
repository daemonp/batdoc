//! The writer against Microsoft's own output.
//!
//! The central test here is not a round-trip. Round-tripping through this crate's own
//! parser proves the two halves agree, which they would even if both were wrong about the
//! format — the failure mode CLAUDE.md § *Testing Rules* names, and the one a crate that is
//! both writer and reader is most exposed to.
//!
//! Instead: take the values out of a stream **Word 16 wrote**, hand them to [`write`], and
//! require the result to be the original stream byte for byte. That is a statement about
//! Microsoft's format rather than about this crate's self-consistency, and it is available
//! only because the repository already carries real-Office fixtures.

use super::*;
use crate::error::{EncryptParam, EncryptParamProblem};
use crate::limits;
use std::io::Read;

/// The three real-Office fixtures, all of which carry a 1 289-byte `EncryptionInfo`.
const OFFICE_FIXTURES: [&str; 3] = [
    "word16_agile.docx",
    "excel16_agile.xlsx",
    "powerpoint16_agile.pptx",
];

/// One fixture's `/EncryptionInfo` stream, whole.
///
/// A hard failure, never a skip — a test that silently passes when its fixture is missing
/// tests nothing, which is a bug this crate has already shipped once.
fn encryption_info_of(fixture: &str) -> Vec<u8> {
    let path = format!("{}/tests/fixtures/{fixture}", env!("CARGO_MANIFEST_DIR"));
    let data = std::fs::read(&path).unwrap_or_else(|e| panic!("fixture {path} must exist: {e}"));
    let mut container = cfb::CompoundFile::open(std::io::Cursor::new(data))
        .unwrap_or_else(|e| panic!("{fixture} is a CFB container: {e}"));
    let mut bytes = Vec::new();
    container
        .open_stream("/EncryptionInfo")
        .unwrap_or_else(|e| panic!("{fixture} has an EncryptionInfo stream: {e}"))
        .read_to_end(&mut bytes)
        .expect("in-memory read");
    bytes
}

/// Pull one attribute's value out of the XML, by name.
///
/// Deliberately **not** `agile::parse_encryption_info`, and deliberately not `quick-xml`:
/// the point of the byte-identity test is that the expectation comes from somewhere other
/// than this crate's own reading of the document. A four-line scan for `name="` is the
/// least this crate can bring to it.
///
/// Every attribute read here appears exactly once in the document, so the first hit is the
/// only hit — asserted below rather than assumed.
fn attr<'a>(xml: &'a str, name: &str) -> &'a str {
    let needle = format!("{name}=\"");
    let occurrences = xml.matches(&needle).count();
    assert_eq!(occurrences, 1, "{name} must appear exactly once");
    let start = xml.find(&needle).expect("checked above") + needle.len();
    let rest = &xml[start..];
    &rest[..rest.find('"').expect("an attribute value is quoted")]
}

/// Both `saltValue` attributes, in document order: `<keyData>`'s then `<p:encryptedKey>`'s.
///
/// `saltValue` is the one name that appears twice, on two elements, about two different
/// salts. Crossing them is the silent bug `AgileParams` is shaped to prevent, so the test
/// helper has to be careful about it too.
fn salt_values(xml: &str) -> (&str, &str) {
    let mut it = xml.match_indices("saltValue=\"").map(|(i, m)| {
        let rest = &xml[i + m.len()..];
        &rest[..rest.find('"').expect("quoted")]
    });
    let first = it.next().expect("keyData/@saltValue");
    let second = it.next().expect("p:encryptedKey/@saltValue");
    assert!(it.next().is_none(), "exactly two saltValue attributes");
    (first, second)
}

fn b64(value: &str) -> Vec<u8> {
    BASE64.decode(value).expect("Office writes valid base64")
}

/// **Fed Word's own values, this writer produces Word's own bytes.**
///
/// The strongest statement available about the writer, and the reason it is worth reading
/// the layout off a fixture rather than composing one: it covers the `\r\n` after the
/// declaration, the absence of a space before each `/>`, the attribute *order* on all three
/// elements, the `xmlns:c` namespace Word always declares and herumi declares only under
/// `isOffice2013`, and the 8-byte header — every one of which is a place a plausible
/// choice produces a plausible document that is not the one Office writes.
///
/// All three applications are checked because all three write the identical 1 289-byte
/// stream for this tuple; if that ever stops being true, this is where it surfaces.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn the_writer_reproduces_real_offices_stream_byte_for_byte() {
    for fixture in OFFICE_FIXTURES {
        let expected = encryption_info_of(fixture);
        let xml = std::str::from_utf8(&expected[8..]).expect("the document is UTF-8");
        let (key_data_salt, password_salt) = salt_values(xml);

        let got = write(&EncryptionInfoParams {
            // Office 16's tuple *is* `EncryptParams::default` — that default was read
            // off these same three fixtures — so the only value taken from the document
            // here is the spin count, and this test is therefore also the proof that the
            // default path still writes the bytes Word writes.
            params: EncryptParams {
                spin_count: attr(xml, "spinCount").parse().expect("an integer"),
                ..Default::default()
            },
            key_data_salt: &b64(key_data_salt),
            encrypted_hmac_key: &b64(attr(xml, "encryptedHmacKey")),
            encrypted_hmac_value: &b64(attr(xml, "encryptedHmacValue")),
            password_salt: &b64(password_salt),
            encrypted_verifier_hash_input: &b64(attr(xml, "encryptedVerifierHashInput")),
            encrypted_verifier_hash_value: &b64(attr(xml, "encryptedVerifierHashValue")),
            encrypted_key_value: &b64(attr(xml, "encryptedKeyValue")),
        })
        .expect("Office's own parameters must be writable");

        // Compare the text before the bytes: on a failure the assertion below prints two
        // 1 289-byte arrays, and the one above prints the diff a human can read.
        assert_eq!(
            std::str::from_utf8(&got[8..]).expect("we write UTF-8"),
            xml,
            "{fixture}: the XML differs from what Office wrote"
        );
        assert_eq!(got, expected, "{fixture}: the stream differs");
        assert_eq!(got.len(), 1289, "{fixture}: Office writes 1 289 bytes");
    }
}

/// The 8-byte header: agile 4.4 and `Reserved = 0x40`, little-endian.
///
/// Separated from the byte-identity test above because it is the half a reader checks
/// *before* the XML — `decrypt_ooxml` refuses a wrong Reserved word outright, and Word
/// refuses it with `0x800A141F` (measured, GH #5).
#[test]
fn the_header_is_agile_four_four_with_the_reserved_word_office_requires() {
    let stream = write(&params(100_000)).unwrap();
    assert_eq!(&stream[..2], &4u16.to_le_bytes(), "vMajor");
    assert_eq!(&stream[2..4], &4u16.to_le_bytes(), "vMinor");
    assert_eq!(&stream[4..8], &0x0000_0040u32.to_le_bytes(), "Reserved");
    assert_eq!(&stream[4..8], &[0x40, 0, 0, 0], "0x40, not zero");
}

/// A synthetic parameter set — distinct byte patterns so a crossed field is visible.
///
/// The tuple is [`EncryptParams::default`] with the spin count moved, so every blob
/// length below is the default path's: 16-byte salts, a 16-byte verifier input,
/// 64-byte SHA-512 blobs and a 32-byte wrapped AES-256 key.
fn params(spin_count: u32) -> EncryptionInfoParams<'static> {
    const KEY_DATA_SALT: [u8; 16] = [0x11; 16];
    const PASSWORD_SALT: [u8; 16] = [0x22; 16];
    const HMAC_KEY: [u8; 64] = [0x33; 64];
    const HMAC_VALUE: [u8; 64] = [0x44; 64];
    const VERIFIER_INPUT: [u8; 16] = [0x55; 16];
    const VERIFIER_VALUE: [u8; 64] = [0x66; 64];
    const KEY_VALUE: [u8; 32] = [0x77; 32];
    EncryptionInfoParams {
        params: EncryptParams {
            spin_count,
            ..Default::default()
        },
        key_data_salt: &KEY_DATA_SALT,
        encrypted_hmac_key: &HMAC_KEY,
        encrypted_hmac_value: &HMAC_VALUE,
        password_salt: &PASSWORD_SALT,
        encrypted_verifier_hash_input: &VERIFIER_INPUT,
        encrypted_verifier_hash_value: &VERIFIER_VALUE,
        encrypted_key_value: &KEY_VALUE,
    }
}

/// This crate's own parser accepts what this crate writes, for values no fixture supplies.
///
/// Secondary to the byte-identity test and not a substitute for it — but it covers the
/// case that one cannot: parameters chosen here rather than by Office, which is what step 4
/// will actually generate. A base64 or length rule that happened to work for Word's values
/// and not for others would show up here.
#[test]
fn the_parser_accepts_what_the_writer_produces() {
    let stream = write(&params(OFFICE_SPIN_COUNT)).unwrap();
    crate::agile::parse_encryption_info(&stream[8..])
        .expect("the writer must not emit a document this crate's own parser refuses");

    // The control: the parser is not simply accepting anything. Corrupting one base64
    // value in the body must be refused, so the acceptance above is a fact about the
    // document rather than about the parser being permissive.
    let xml = String::from_utf8(stream[8..].to_vec()).unwrap();
    let broken = xml.replace("keyBits=\"256\"", "keyBits=\"257\"");
    assert_ne!(broken, xml, "the substitution must have applied");
    assert!(
        crate::agile::parse_encryption_info(broken.as_bytes()).is_err(),
        "a keyBits the crate does not implement must still be refused"
    );
}

/// Every attribute value is base64 or digits, so no XML escaping is reachable.
///
/// The property the module header relies on to say there is no injection to defend
/// against. It holds only while every interpolated value stays base64 — a later free-text
/// attribute would break it silently, and this is what would notice.
#[test]
fn no_written_value_can_need_xml_escaping() {
    let stream = write(&params(OFFICE_SPIN_COUNT)).unwrap();
    let xml = std::str::from_utf8(&stream[8..]).unwrap();

    // Split on the quote rather than scanning for `="`: base64 values end in `=` padding,
    // so `saltValue="…Lyw=="` contains `="` *inside* the value and a naive scan walks off
    // the end of it. Since no value may contain a quote — which is most of what this test
    // asserts — the quotes strictly alternate, so every odd-indexed piece is a value. That
    // makes the split self-checking: a value that did contain a quote would misalign the
    // parity and land raw markup in the charset check below, which is exactly a failure.
    let pieces: Vec<&str> = xml.split('"').collect();
    assert!(
        pieces.len() % 2 == 1,
        "unbalanced quotes: the document is not well-formed"
    );
    let values: Vec<&str> = pieces.iter().skip(1).step_by(2).copied().collect();
    // 3 declaration + 3 namespaces + 8 on <keyData> + 2 on <dataIntegrity> + 1 uri
    // + 12 on <p:encryptedKey>. Pinned so an attribute cannot be dropped silently: the
    // charset loop below passes vacuously on a document that lost half its values.
    assert_eq!(
        values.len(),
        29,
        "every attribute value must be accounted for"
    );

    for value in &values {
        assert!(
            value
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || "+/=:.-_".contains(c)),
            "an attribute value is not base64, a number or a URI: {value:?}"
        );
    }
    // And no metacharacter survives anywhere a value was interpolated.
    assert!(!xml.contains('&'), "an escape would be needed for &");
}

/// A blob of the wrong length is refused before it becomes a file this crate cannot read.
///
/// One case per checked field, each differing from the good set only in that field, so a
/// refusal is attributable. The writer's job here is to fail rather than to emit a document
/// whose `saltSize` and `saltValue` disagree — which the parser would then reject, one
/// layer too late to say anything useful about why.
#[test]
fn a_blob_of_the_wrong_length_is_refused_by_the_writer() {
    const LONG: [u8; 128] = [0u8; 128];
    const SHORT: [u8; 1] = [0u8; 1];

    // The control first: the unmodified set is accepted, so every refusal below is
    // attributable to the one field it changes.
    assert!(write(&params(OFFICE_SPIN_COUNT)).is_ok());

    for (what, mutate) in [
        (
            "keyData/@saltValue",
            &(|p: &mut EncryptionInfoParams| p.key_data_salt = &SHORT)
                as &dyn Fn(&mut EncryptionInfoParams),
        ),
        ("p:encryptedKey/@saltValue", &|p| p.password_salt = &LONG),
        ("dataIntegrity/@encryptedHmacKey", &|p| {
            p.encrypted_hmac_key = &SHORT
        }),
        ("dataIntegrity/@encryptedHmacValue", &|p| {
            p.encrypted_hmac_value = &LONG
        }),
        ("p:encryptedKey/@encryptedVerifierHashInput", &|p| {
            p.encrypted_verifier_hash_input = &LONG
        }),
        ("p:encryptedKey/@encryptedVerifierHashValue", &|p| {
            p.encrypted_verifier_hash_value = &SHORT
        }),
        ("p:encryptedKey/@encryptedKeyValue", &|p| {
            p.encrypted_key_value = &SHORT
        }),
    ] {
        let mut p = params(OFFICE_SPIN_COUNT);
        mutate(&mut p);
        let got = write(&p).map(|s| s.len());
        assert!(
            matches!(&got, Err(Error::BadParameters(msg)) if msg.contains(what)),
            "a wrong-length {what} must be refused by name, got: {got:?}"
        );
    }
}

/// `encryptedKeyValue` is `roundUp(keyData/@keyBits / 8, blockSize)`, not `keyBits / 8`.
///
/// The direct guard on the one length that was wrong: the deleted `lengths` module fixed
/// this blob at the 32-byte session key, which is right for AES-256 by coincidence and
/// wrong for AES-192. [MS-OFFCRYPTO] §2.3.4.13 step 1 sizes the package key from
/// `Encryptor.KeyData.keyBits` and every blob in the document is padded to `blockSize`,
/// so a 24-byte AES-192 key is wrapped into **32** bytes — which is what
/// `agile::parse_encryption_info` requires on the way back in and what real Word 16
/// wrote in `tests/fixtures/agile_aes192_sha384.docx`.
///
/// Both directions are asserted, because either alone is weak: the refusal of 24 would
/// pass on a writer that refused every AES-192 file, and the acceptance of 32 would pass
/// on a writer that had stopped checking this blob at all.
///
/// The per-element wiring is checked in the same test. `keyBits` is the attribute most
/// easily crossed — it appears on both elements and §2.3.4.10 does **not** require the
/// two to be equal — so a writer that sized the blob from `p:encryptedKey/@keyBits`
/// would be measuring 256 here and would accept the 32-byte value for the wrong reason.
/// Parting the two in the parameters is what makes that visible.
#[test]
fn the_wrapped_key_is_sized_from_key_datas_key_bits_rounded_up_to_the_block() {
    const KEY_VALUE_24: [u8; 24] = [0x77; 24];
    const KEY_VALUE_32: [u8; 32] = [0x77; 32];

    let aes192 = EncryptParams {
        // AES-192 package key under an AES-256 password key: the spec joins neither to
        // the other, and the asymmetry is what makes the wiring testable.
        key_data_key_bits: 192,
        ..Default::default()
    };
    aes192
        .validate()
        .expect("AES-192 keyData under an AES-256 SHA-512 KEK is a conforming tuple");

    let mut short = params(OFFICE_SPIN_COUNT);
    short.params = aes192;
    short.encrypted_key_value = &KEY_VALUE_24;
    let got = write(&short).map(|s| s.len());
    assert!(
        matches!(
            &got,
            Err(Error::BadParameters(msg))
                if msg.contains("p:encryptedKey/@encryptedKeyValue")
                    && msg.contains("is 24 bytes")
                    && msg.contains("at 32")
        ),
        "a bare 24-byte AES-192 key must be refused: the blob is roundUp(24, 16) = 32, \
         got: {got:?}"
    );

    let mut padded = params(OFFICE_SPIN_COUNT);
    padded.params = aes192;
    padded.encrypted_key_value = &KEY_VALUE_32;
    let stream = write(&padded).expect("the padded 32-byte blob is the conforming length");
    let xml = std::str::from_utf8(&stream[8..]).unwrap();
    let (key_data, password) = xml
        .split_once("<p:encryptedKey ")
        .expect("the document has both elements");
    assert!(
        key_data.contains("keyBits=\"192\""),
        "keyData must carry the keyData keyBits: {key_data}"
    );
    assert!(
        password.contains("keyBits=\"256\""),
        "p:encryptedKey must carry the password keyBits: {password}"
    );
    crate::agile::parse_encryption_info(&stream[8..])
        .expect("this crate's own reader must accept the AES-192 document it writes");
}

/// A spin count past this crate's own ceiling is refused rather than written.
///
/// The ceiling exists to bound a hostile *file*, and the failure this prevents is the
/// crate writing one it would then refuse to open. The control is the value one below,
/// which is accepted — without it this would pass on a writer that refused every spin
/// count.
///
/// **The refusal now comes from `EncryptParams::validate`, not from `write`'s own
/// check**, and the whole payload is asserted rather than a substring of the message:
/// `write` used to test `SPIN_COUNT_MAX` itself and report `BadParameters`, which was one
/// fact checked in two places. The single check is the caller-facing one, so what this
/// test pins is the typed verdict a consumer would render — a message-substring
/// assertion would have passed unchanged through that swap of variants.
#[test]
fn a_spin_count_past_the_ceiling_is_refused_rather_than_written() {
    let got = write(&params(limits::SPIN_COUNT_MAX + 1)).map(|s| s.len());
    assert!(
        matches!(
            &got,
            Err(Error::EncryptParams {
                param: EncryptParam::SpinCount,
                problem: EncryptParamProblem::OutsideSpecRange,
                got: value,
                min: 0,
                max,
            }) if *value == limits::SPIN_COUNT_MAX + 1 && *max == limits::SPIN_COUNT_MAX
        ),
        "a spin count over the ceiling must be refused by the parameter validator with \
         the whole payload, got: {got:?}"
    );

    assert!(
        write(&params(limits::SPIN_COUNT_MAX)).is_ok(),
        "the ceiling itself must be writable -- the bound is inclusive"
    );
    assert!(write(&params(0)).is_ok(), "there is no floor on spinCount");
}

/// What Office writes, pinned so a later edit cannot quietly change it.
#[test]
fn the_office_spin_count_is_the_measured_one() {
    assert_eq!(OFFICE_SPIN_COUNT, 100_000);
    let stream = write(&params(OFFICE_SPIN_COUNT)).unwrap();
    let xml = std::str::from_utf8(&stream[8..]).unwrap();
    assert_eq!(attr(xml, "spinCount"), "100000");
}
