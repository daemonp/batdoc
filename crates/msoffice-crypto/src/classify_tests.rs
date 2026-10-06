#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
// `classify.rs` denies these for its production code; this file is that module's test
// body and uses all three the way a test should -- a fixture that fails to load or an
// assertion that fails is the test failing, which is the point.
//! Tests for [`crate::classify`].
//!
//! Two jobs, and the second is the important one:
//!
//! 1. Every fixture in the corpus classifies to the right family, the right algorithm
//!    tuple and the right `dataIntegrity` verdict.
//! 2. **Nothing makes it panic.** `classify` is the first thing a caller runs on a file
//!    from outside, so a panic there escapes the `Result` contract entirely — a caller's
//!    `map_err` never sees an unwind. The truncation and mutation sweeps below are a
//!    cheap stand-in for a fuzzer: every prefix of every fixture, every single-byte
//!    poisoning of the header region, and a set of hand-built degenerate inputs.
//!
//! These tests compile in **every** configuration, including
//! `cargo test --no-default-features` — they touch no cipher.

use super::*;

fn fixture(name: &str) -> Vec<u8> {
    let path = format!(
        concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/{}"),
        name
    );
    std::fs::read(&path).unwrap_or_else(|e| {
        panic!("fixture {path} must be present -- these tests are not optional: {e}")
    })
}

// ---- the corpus ----------------------------------------------------------------------

/// The agile fixture, attribute by attribute, against its own `EncryptionInfo` XML:
///
/// ```xml
/// <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
///          cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" .../>
/// <dataIntegrity encryptedHmacKey="..." encryptedHmacValue="..."/>
/// <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" .../>
/// ```
#[test]
fn agile_fixture_classifies_with_both_parameter_sets_and_a_declared_tag() {
    let class = classify(&fixture("agile_encrypted.docx"));

    assert_eq!(class.container, Container::Cfb);
    assert_eq!(class.version, Some((4, 4)));
    assert_eq!(class.family, Family::Agile);
    assert_eq!(class.data_integrity, IntegrityDeclaration::Declared);
    assert!(class.is_encrypted());
    assert!(class.is_supported());

    let key_data = class.key_data.expect("<keyData> is present");
    assert_eq!(
        key_data,
        AlgorithmParams {
            cipher: Some(CipherAlgorithm::Aes),
            hash: Some(HashAlgorithm::Sha512),
            key_bits: Some(256),
            block_size: Some(16),
            salt_size: Some(16),
            // <keyData> carries no spinCount -- that attribute belongs to the password
            // key encryptor alone. Asserting the None is the point: it is what proves
            // the two parameter sets are read separately rather than merged.
            spin_count: None,
        }
    );

    let password = class.password_key.expect("<p:encryptedKey> is present");
    assert_eq!(
        password,
        AlgorithmParams {
            cipher: Some(CipherAlgorithm::Aes),
            hash: Some(HashAlgorithm::Sha512),
            key_bits: Some(256),
            block_size: Some(16),
            salt_size: Some(16),
            spin_count: Some(100_000),
        }
    );
}

/// The two non-SHA-512 fixtures, in the build that has no cipher crate at all.
///
/// `classify` already reported every hash correctly before issue #11 — the bug was
/// entirely in the decrypt path — so this costs the detection suite nothing and is worth
/// having anyway: it is the assertion that would fail first if the reporting table and
/// the operational one ever drifted apart, which is precisely why there is only one
/// `HashAlgorithm`.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn the_non_sha512_agile_fixtures_report_their_own_hash_in_both_parameter_sets() {
    for (name, hash, key_bits) in [
        ("agile_aes256_sha384.docx", HashAlgorithm::Sha384, 256),
        ("agile_aes256_sha256.docx", HashAlgorithm::Sha256, 256),
        ("agile_aes128_sha1.docx", HashAlgorithm::Sha1, 128),
        ("agile_aes128_sha384.docx", HashAlgorithm::Sha384, 128),
        ("agile_aes192_sha384.docx", HashAlgorithm::Sha384, 192),
    ] {
        let class = classify(&fixture(name));
        assert_eq!(class.family, Family::Agile, "{name}");
        assert_eq!(
            class.data_integrity,
            IntegrityDeclaration::Declared,
            "{name}"
        );

        let expected = AlgorithmParams {
            cipher: Some(CipherAlgorithm::Aes),
            hash: Some(hash),
            key_bits: Some(key_bits),
            block_size: Some(16),
            salt_size: Some(16),
            spin_count: None,
        };
        assert_eq!(class.key_data.expect("<keyData>"), expected, "{name}");
        assert_eq!(
            class.password_key.expect("<p:encryptedKey>"),
            AlgorithmParams {
                spin_count: Some(100_000),
                ..expected
            },
            "{name}"
        );

        // ... and the SHA-512 fixture beside them still reports SHA-512, so the two
        // above are not passing because everything now reports the same thing.
        assert_eq!(
            classify(&fixture("agile_encrypted.docx"))
                .password_key
                .and_then(|p| p.hash),
            Some(HashAlgorithm::Sha512)
        );
    }
}

/// The standard fixture's `EncryptionHeader`, field by field:
/// `Flags = 0x36` (fAES set), `AlgID = 0x00006801`, `AlgIDHash = 0x00008004` (SHA-1),
/// `KeySize = 128`, `EncryptionVerifier.SaltSize = 16`.
///
/// `AlgID = 0x6801` is RC4's identifier, and the file sets `fAES` alongside it — a
/// combination [MS-OFFCRYPTO] §2.3.2's table does not list, and one §2.3.1 rules out in
/// prose ("If the fAES encryption bit is set, a block cipher that supports ECB mode MUST
/// be used"; RC4 is a stream cipher). It is AES-128 in fact: `decrypt_ooxml` decrypts it
/// as such. fAES wins here for exactly that reason, and this test is the record of the
/// file that pins the precedence.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn standard_fixture_classifies_as_aes_128_with_no_integrity_element() {
    let class = classify(&fixture("standard_encrypted.docx"));

    assert_eq!(class.container, Container::Cfb);
    assert_eq!(class.version, Some((4, 2)));
    assert_eq!(class.family, Family::Standard);
    assert_eq!(class.data_integrity, IntegrityDeclaration::NotApplicable);
    assert!(class.is_encrypted());
    assert!(class.is_supported());

    assert_eq!(
        class.key_data.expect("the EncryptionHeader is present"),
        AlgorithmParams {
            cipher: Some(CipherAlgorithm::Aes),
            hash: Some(HashAlgorithm::Sha1),
            key_bits: Some(128),
            // AES-ECB: no chaining block to declare, and the spin count is fixed at
            // 50 000 by [MS-OFFCRYPTO] §2.3.4.7 rather than carried in the file.
            block_size: None,
            salt_size: Some(16),
            spin_count: None,
        }
    );
    assert!(
        class.password_key.is_none(),
        "standard encryption has a single parameter set, not two"
    );
}

#[test]
fn plain_package_classifies_as_an_unencrypted_zip() {
    let class = classify(&fixture("plain.docx"));

    assert_eq!(class.container, Container::Zip);
    assert_eq!(class.version, None);
    assert_eq!(class.family, Family::Unencrypted);
    assert_eq!(class.data_integrity, IntegrityDeclaration::NotApplicable);
    assert_eq!(class.key_data, None);
    assert_eq!(class.password_key, None);
    assert!(!class.is_encrypted());
    assert!(!class.is_supported());
}

/// The fourth fixture is a 30-byte text file — the control for "not an Office container
/// at all". It must come back `Unknown`, **not** `Unencrypted`: this crate cannot say a
/// file it does not recognise is unencrypted, only that it could not tell.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn plain_text_fixture_is_unknown_not_unencrypted() {
    let class = classify(&fixture("plain_content.txt"));

    assert_eq!(class.container, Container::Unknown);
    assert_eq!(class.family, Family::Unknown);
    assert_eq!(class.data_integrity, IntegrityDeclaration::Unknown);
    assert_eq!(class.version, None);
    assert!(!class.is_encrypted());
}

// ---- the dataIntegrity verdict (design commitment D1) --------------------------------

/// Delete the whole `<dataIntegrity .../>` element from the fixture's `EncryptionInfo`.
///
/// That XML is stored in the clear, so this touches no cryptographic parameter and needs
/// no foreign writer: the result is a well-formed agile document that simply declares no
/// tag — which is precisely the edit an attacker makes to disable tamper detection, and
/// why the default policy refuses the result (GH #12). The base64
/// alphabet contains `/` but not `>`, so the first `/>` after the element name is its own
/// terminator. (The same helper exists in `lib.rs`'s decrypt tests; this copy is here so
/// the classify suite still runs with `--no-default-features`.)
fn agile_fixture_with_data_integrity_edited(edit: impl Fn(&mut Vec<u8>, usize, usize)) -> Vec<u8> {
    use std::io::{Read, Seek, SeekFrom, Write};

    let mut cursor = std::io::Cursor::new(fixture("agile_encrypted.docx"));
    {
        let mut container = cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB");
        let mut info = Vec::new();
        container
            .open_stream("/EncryptionInfo")
            .unwrap()
            .read_to_end(&mut info)
            .unwrap();

        let start = info
            .windows(14)
            .position(|w| w == b"<dataIntegrity")
            .expect("the fixture declares a dataIntegrity tag");
        let end = start + info[start..].windows(2).position(|w| w == b"/>").unwrap() + 2;
        edit(&mut info, start, end);

        let mut stream = container.open_stream("/EncryptionInfo").unwrap();
        stream.set_len(0).unwrap();
        stream.seek(SeekFrom::Start(0)).unwrap();
        stream.write_all(&info).unwrap();
        stream.flush().unwrap();
    }
    cursor.into_inner()
}

#[test]
fn an_agile_file_without_the_element_reports_absent() {
    let data = agile_fixture_with_data_integrity_edited(|info, start, end| {
        info.drain(start..end);
    });
    let class = classify(&data);

    assert_eq!(class.family, Family::Agile);
    assert_eq!(class.data_integrity, IntegrityDeclaration::Absent);
    // The control: everything else about the file is unchanged, so the verdict above
    // comes from the deleted element and not from the rewrite.
    assert_eq!(class.key_data.and_then(|p| p.key_bits), Some(256));
    assert_eq!(
        classify(&fixture("agile_encrypted.docx")).data_integrity,
        IntegrityDeclaration::Declared
    );
}

/// A half-written element is malformed, not absent — and `decrypt` refuses it with
/// `BadParameters`, at parse time, under **every** policy including `Skip`. `Absent` is a
/// different verdict with a different remedy: since GH #12 it too is refused by default,
/// but the explicitly chosen `VerifyIfPresent` / `Skip` opt-outs still decrypt it.
/// Collapsing the two would promise a caller an escape hatch that does not exist.
#[test]
fn an_element_missing_a_blob_reports_incomplete() {
    let data = agile_fixture_with_data_integrity_edited(|info, start, end| {
        let replacement = b"<dataIntegrity encryptedHmacKey=\"AAAA\"/>".to_vec();
        info.splice(start..end, replacement);
    });

    assert_eq!(
        classify(&data).data_integrity,
        IntegrityDeclaration::Incomplete
    );
}

// ---- version dispatch ----------------------------------------------------------------

/// Rewrite the two `EncryptionVersionInfo` u16s in place. Everything after them is left
/// alone, so each case differs from the fixture in exactly four bytes.
fn agile_fixture_with_version(v_major: u16, v_minor: u16) -> Vec<u8> {
    use std::io::{Read, Seek, SeekFrom, Write};

    let mut cursor = std::io::Cursor::new(fixture("agile_encrypted.docx"));
    {
        let mut container = cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB");
        let mut info = Vec::new();
        container
            .open_stream("/EncryptionInfo")
            .unwrap()
            .read_to_end(&mut info)
            .unwrap();
        info[0..2].copy_from_slice(&v_major.to_le_bytes());
        info[2..4].copy_from_slice(&v_minor.to_le_bytes());

        let mut stream = container.open_stream("/EncryptionInfo").unwrap();
        stream.seek(SeekFrom::Start(0)).unwrap();
        stream.write_all(&info).unwrap();
        stream.flush().unwrap();
    }
    cursor.into_inner()
}

#[test]
fn version_pairs_map_to_the_families_they_name() {
    // Extensible encryption and every undefined pair: encrypted, but not a family this
    // crate implements. `is_encrypted` is still true -- the file has an EncryptionInfo.
    for (v_major, v_minor) in [(3, 3), (4, 3), (1, 1), (2, 1), (4, 0), (9, 9), (0, 0)] {
        let class = classify(&agile_fixture_with_version(v_major, v_minor));
        assert_eq!(
            class.family,
            Family::Unsupported,
            "version {v_major}.{v_minor}"
        );
        assert_eq!(class.version, Some((v_major, v_minor)));
        assert!(class.is_encrypted() && !class.is_supported());
        assert_eq!(class.data_integrity, IntegrityDeclaration::Unknown);
    }

    // 2.2 and 3.2 are standard encryption exactly as 4.2 is. Here they are applied to an
    // agile body, so the binary header parse reads XML bytes as an EncryptionHeader --
    // the point is only that the version pair routes to the standard branch and that
    // doing so on nonsense produces a verdict rather than a panic.
    for (v_major, v_minor) in [(2, 2), (3, 2)] {
        let class = classify(&agile_fixture_with_version(v_major, v_minor));
        assert_eq!(class.version, Some((v_major, v_minor)));
        assert_eq!(class.data_integrity, IntegrityDeclaration::NotApplicable);
    }
}

/// `fAES`, not the version pair, separates ECMA-376 standard encryption from RC4
/// CryptoAPI: both are written with `vMinor = 2`. Clearing the bit in the real fixture
/// is a one-byte edit and must flip the family.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn clearing_f_aes_turns_the_standard_fixture_into_rc4_cryptoapi() {
    use std::io::{Read, Seek, SeekFrom, Write};

    let mut cursor = std::io::Cursor::new(fixture("standard_encrypted.docx"));
    {
        let mut container = cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB");
        let mut info = Vec::new();
        container
            .open_stream("/EncryptionInfo")
            .unwrap()
            .read_to_end(&mut info)
            .unwrap();
        // stream[8..12] = EncryptionHeaderSize, stream[12..] = EncryptionHeader,
        // whose first field is Flags.
        let flags = u32::from_le_bytes(info[12..16].try_into().unwrap());
        assert_eq!(flags & 0x20, 0x20, "the fixture declares fAES");
        info[12..16].copy_from_slice(&(flags & !0x20u32).to_le_bytes());

        let mut stream = container.open_stream("/EncryptionInfo").unwrap();
        stream.seek(SeekFrom::Start(0)).unwrap();
        stream.write_all(&info).unwrap();
        stream.flush().unwrap();
    }
    let class = classify(&cursor.into_inner());

    assert_eq!(class.family, Family::Rc4CryptoApi);
    assert_eq!(
        class.key_data.and_then(|p| p.cipher),
        Some(CipherAlgorithm::Rc4),
        "AlgID 0x6801 names RC4 once fAES is clear"
    );
    assert!(class.is_encrypted());
    assert!(
        !class.is_supported(),
        "RC4 CryptoAPI is slice S3, not today"
    );
}

/// `classify` reports **all three** of [MS-OFFCRYPTO] §2.3.2's AES `AlgID`s —
/// `0x0000660E`, `0x0000660F`, `0x00006610` — and the `KeySize` beside each, and it
/// always did.
///
/// Written because the claim was in dispute while the decrypt path was being taught to
/// read AES-192 and AES-256: `classify_standard` matches all three onto
/// `CipherAlgorithm::Aes` and passes `KeySize` straight through, so no change was needed
/// there and none was made. This is that check, run rather than asserted in prose.
///
/// It also records the asymmetry the decrypt work closed. Before 2026-09-20 `classify`
/// reported an AES-256 standard file as supported AES-256 and `decrypt_ooxml` then
/// refused it by name — the detector and the decryptor disagreeing about the same bytes,
/// which is the failure the `fAES`-before-`AlgID` precedence exists to prevent.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn classify_reports_all_three_aes_alg_ids_and_their_key_sizes() {
    use std::io::{Read, Seek, SeekFrom, Write};

    // stream[8..12] = EncryptionHeaderSize, stream[12..] = EncryptionHeader:
    // Flags at 12, SizeExtra at 16, AlgID at 20, AlgIDHash at 24, KeySize at 28.
    for (alg_id, key_bits) in [
        (0x0000_660Eu32, 128u32),
        (0x0000_660F, 192),
        (0x0000_6610, 256),
    ] {
        let mut cursor = std::io::Cursor::new(fixture("standard_encrypted.docx"));
        {
            let mut container = cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB");
            let mut info = Vec::new();
            container
                .open_stream("/EncryptionInfo")
                .unwrap()
                .read_to_end(&mut info)
                .unwrap();
            info[20..24].copy_from_slice(&alg_id.to_le_bytes());
            info[28..32].copy_from_slice(&key_bits.to_le_bytes());

            let mut stream = container.open_stream("/EncryptionInfo").unwrap();
            stream.seek(SeekFrom::Start(0)).unwrap();
            stream.write_all(&info).unwrap();
            stream.flush().unwrap();
        }
        let class = classify(&cursor.into_inner());

        assert_eq!(class.family, Family::Standard, "AlgID {alg_id:#010x}");
        let params = class.key_data.expect("the EncryptionHeader is present");
        assert_eq!(
            params.cipher,
            Some(CipherAlgorithm::Aes),
            "AlgID {alg_id:#010x}"
        );
        assert_eq!(
            params.key_bits,
            Some(key_bits),
            "AlgID {alg_id:#010x}: KeySize is reported as the file declares it"
        );
        assert!(class.is_supported(), "AlgID {alg_id:#010x}");
    }
}

// ---- it must never panic -------------------------------------------------------------

/// Every prefix of every fixture. This is the fuzz-shaped half of the suite: a truncated
/// upload is the single most likely malformed input a caller passes in, and it exercises
/// the CFB header, the FAT, the directory tree, the `EncryptionInfo` stream, the version
/// pair and both parsers at every depth.
///
/// Small prefixes are taken one byte at a time and larger ones on a stride, which keeps
/// the whole sweep at a few thousand `classify` calls rather than 120 000.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn every_truncation_of_every_fixture_returns_a_verdict() {
    for name in [
        "agile_encrypted.docx",
        "agile_aes256_sha384.docx",
        "agile_aes256_sha256.docx",
        "standard_encrypted.docx",
        "plain.docx",
        "plain_content.txt",
    ] {
        let data = fixture(name);
        let offsets = (0..=data.len().min(600))
            .chain((0..data.len()).step_by(97))
            .chain([data.len()]);
        for cut in offsets {
            let class = classify(&data[..cut]);
            // The only claim: it returned. Anything stronger would be asserting the
            // shape of a corrupt file, which is not a property worth pinning.
            let _ = class.is_encrypted();
        }
    }
}

/// Single-byte poisoning across the whole header region of both encrypted fixtures.
/// Truncation only ever produces *short* input; this produces input that is the right
/// length and internally inconsistent, which is what reaches the parsers' arithmetic.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn single_byte_corruption_of_the_header_region_returns_a_verdict() {
    for name in ["agile_encrypted.docx", "standard_encrypted.docx"] {
        let original = fixture(name);
        for offset in 0..original.len().min(4096) {
            for xor in [0x01u8, 0xFF, 0x80] {
                let mut data = original.clone();
                data[offset] ^= xor;
                let _ = classify(&data).is_encrypted();
            }
        }
    }
}

/// Degenerate and hand-built inputs: the shapes a sweep over a real file never produces.
#[test]
fn degenerate_inputs_return_a_verdict() {
    const CFB: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    let mut cases: Vec<Vec<u8>> = vec![
        Vec::new(),
        vec![0u8],
        vec![0xFFu8; 7],
        b"PK".to_vec(),
        b"PK\x03".to_vec(),
        b"PK\x03\x04".to_vec(),
        b"%PDF-1.4".to_vec(),
        CFB.to_vec(),
        // CFB magic with nothing behind it, and with a plausible-length garbage body.
        [CFB.as_slice(), &[0u8; 512]].concat(),
        [CFB.as_slice(), &[0xFFu8; 4096]].concat(),
        vec![0u8; 8192],
        vec![0xFFu8; 8192],
    ];
    // A CFB whose every byte after the magic is a repeating pattern -- enough structure
    // to get past the magic check and into the container parser.
    for pattern in [0x01u8, 0x7F, 0xAA] {
        cases.push([CFB.as_slice(), &vec![pattern; 3000]].concat());
    }

    for case in cases {
        let class = classify(&case);
        assert!(
            !class.is_supported() || class.container == Container::Cfb,
            "a supported family can only come from a CFB container"
        );
    }
}

/// The two ways an `EncryptionInfo` can be present and unreadable: shorter than the
/// 8-byte version header, and XML that does not parse. Both must classify, not fail.
#[test]
fn an_unreadable_encryption_info_is_unknown_rather_than_a_failure() {
    use std::io::{Seek, SeekFrom, Write};

    let rewrite_info = |body: &[u8]| {
        let mut cursor = std::io::Cursor::new(fixture("agile_encrypted.docx"));
        {
            let mut container = cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB");
            let mut stream = container.open_stream("/EncryptionInfo").unwrap();
            stream.set_len(0).unwrap();
            stream.seek(SeekFrom::Start(0)).unwrap();
            stream.write_all(body).unwrap();
            stream.flush().unwrap();
        }
        cursor.into_inner()
    };

    // Fewer than 8 bytes: no version pair to read.
    let class = classify(&rewrite_info(&[4, 0, 4]));
    assert_eq!(class.container, Container::Cfb);
    assert_eq!(class.family, Family::Unknown);
    assert_eq!(class.version, None);
    assert_eq!(class.data_integrity, IntegrityDeclaration::Unknown);

    // A well-formed 4.4 header over XML that is not XML. The version pair still reads,
    // so the family is Agile; the parse yields nothing, so no tuple and no tag.
    let class = classify(&rewrite_info(
        b"\x04\x00\x04\x00\x00\x00\x00\x00<encryption <<<",
    ));
    assert_eq!(class.family, Family::Agile);
    assert_eq!(class.version, Some((4, 4)));
    assert_eq!(class.key_data, None);
    assert_eq!(class.password_key, None);
    assert_eq!(class.data_integrity, IntegrityDeclaration::Absent);
}

/// A hostile `spinCount` / `keyBits` is *reported*, not refused. Bounding them belongs to
/// `decrypt` (`crate::limits`); a classifier that hid the value would leave its caller
/// unable to see what it was about to refuse.
#[test]
fn hostile_numbers_are_reported_rather_than_rejected() {
    let xml = br#"<encryption>
        <keyData saltSize="16" blockSize="16" keyBits="4294967295" hashSize="64"
                 cipherAlgorithm="AES" hashAlgorithm="SHA512" saltValue="AAAA"/>
        <p:encryptedKey spinCount="4294967295" keyBits="7" saltSize="0" blockSize="99"
                        hashAlgorithm="MD5" cipherAlgorithm="RC2"/>
        </encryption>"#;
    let class = super::classify_agile(xml, Some((4, 4)));

    let key_data = class.key_data.unwrap();
    assert_eq!(key_data.key_bits, Some(u32::MAX));
    let password = class.password_key.unwrap();
    assert_eq!(password.spin_count, Some(u32::MAX));
    assert_eq!(password.key_bits, Some(7));
    assert_eq!(password.block_size, Some(99));
    assert_eq!(password.salt_size, Some(0));
    // Unrecognised names are `None`, never a guess.
    assert_eq!(password.hash, None);
    assert_eq!(password.cipher, None);
}

/// Attribute values that are not numbers at all — a classifier must not choke on them
/// and must not silently report a wrong number.
#[test]
fn non_numeric_attributes_become_none() {
    let xml = br#"<encryption>
        <keyData keyBits="two hundred and fifty six" blockSize="" saltSize="-16"
                 spinCount="0x100" hashAlgorithm="SHA512"/>
        </encryption>"#;
    let params = super::classify_agile(xml, Some((4, 4)))
        .key_data
        .expect("<keyData> was seen");

    assert_eq!(params.key_bits, None);
    assert_eq!(params.block_size, None);
    assert_eq!(params.salt_size, None, "u32 does not parse a negative");
    assert_eq!(params.spin_count, None, "u32 does not parse a hex literal");
    // ...and the one well-formed attribute still comes through.
    assert_eq!(params.hash, Some(HashAlgorithm::Sha512));
}

// ---- legacy binary containers ---------------------------------------------------------
//
// These live here rather than in `tests/real_office_fixtures.rs` because that file is
// `crypto-ops`-gated, and this is the property worth pinning in the DETECTION build:
// recognising an encrypted `.doc` costs a CFB open and some integer reads, no cipher at
// all. `cargo tree --no-default-features` shows no cipher crate, and CI asserts it.

#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn legacy_binary_containers_are_named_without_any_cipher_dependency() {
    let cases = [
        (
            "word97_password.doc",
            Document::WordBinary,
            Family::Rc4CryptoApi,
            Some((4u16, 2u16)),
        ),
        (
            "excel97_password.xls",
            Document::ExcelBinary,
            Family::Rc4CryptoApi,
            Some((4, 2)),
        ),
        // PowerPoint's header sits inside a CryptSession10Container reached through the
        // persist directory. Until GH #4 detection stopped at the UserEditAtom and
        // reported `Unsupported`; now it follows the reference, in a handful of seeks.
        (
            "powerpoint97_password.ppt",
            Document::PowerPointBinary,
            Family::Rc4CryptoApi,
            Some((4, 2)),
        ),
    ];
    for (name, document, family, version) in cases {
        let c = classify(&fixture(name));
        assert_eq!(c.container, Container::Cfb, "{name}");
        assert_eq!(c.document, document, "{name}");
        assert_eq!(c.family, family, "{name}");
        assert_eq!(c.version, version, "{name}");
        assert!(c.is_encrypted(), "{name}");
    }
}

#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn an_ooxml_container_is_still_reported_as_an_ooxml_package() {
    // Guards against the binary probe capturing files it should not: everything reached
    // through an `EncryptionInfo` stream is ECMA-376 by construction.
    for name in [
        "agile_encrypted.docx",
        "standard_encrypted.docx",
        "word16_agile.docx",
        "excel16_agile.xlsx",
        "powerpoint16_agile.pptx",
    ] {
        assert_eq!(
            classify(&fixture(name)).document,
            Document::OoxmlPackage,
            "{name}: an EncryptionInfo container is an OOXML package, never a binary one"
        );
    }
    // ...and the other half of the same rule: a *plain* archive is not one of those, and
    // must not claim to be. `is_zip` reads four bytes and opens nothing, so `OoxmlPackage`
    // here would be an affirmative claim about a format nothing looked at.
    assert_eq!(
        classify(&fixture("plain.docx")).document,
        Document::ZipArchive,
        "a plain zip is reported as the archive it is, not as the package it might be"
    );
}

/// The contract [`Document::ZipArchive`] states, asserted rather than described: at this
/// layer a real `.docx` and a zip that merely starts with `PK\x03\x04` are *the same
/// verdict*, because the test that produced it is the same four bytes in both cases.
///
/// Delete `Document::ZipArchive` and route the ZIP branch back to `OoxmlPackage` and this
/// still passes -- it is about the two being equal, not about which variant they are. So
/// the assertion below on `plain.docx` is what makes the pair meaningful, and the two
/// tests are written to fail for different reasons on purpose.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn a_zip_is_not_inspected_so_a_package_and_an_impostor_are_indistinguishable() {
    // A zip signature followed by bytes belonging to no package at all -- the shape of a
    // look-alike that carries Office-ish structure without being an Office file.
    let impostor = b"PK\x03\x04not an OOXML package, and nothing here was read";
    let real = classify(&fixture("plain.docx"));
    let fake = classify(impostor);

    assert_eq!(real.document, fake.document);
    assert_eq!(real.container, fake.container);
    assert_eq!(real.family, fake.family);
    assert_eq!(
        real.document,
        Document::ZipArchive,
        "and the shared verdict is the one that claims nothing about the contents"
    );
    // Nothing was opened, so nothing failed to open -- which is a different fact from
    // the `Unreadable` a truncated CFB gets.
    assert_eq!(fake.container_read, ContainerRead::NotAttempted);
}

/// A prefix of a real file and a genuinely unrecognisable one are **not** the same
/// verdict, and before `container_read` existed they were.
///
/// This is the test the feature is for. Every other field agrees across the two inputs
/// below -- both are `Family::Unknown`, both `Document::Unknown`,
/// `IntegrityDeclaration::Unknown` -- so a caller dispatching on any of them takes the
/// same branch for "I was handed too few bytes" as for "this is not an Office file". One
/// of those is a missing answer and the other is a wrong one.
///
/// **Proven by removal**: route the `Err(Error::NotACfbFile)` arm in `classify_cfb` back
/// to `classify_binary(data)` and the first assertion fails with
/// `Opened`/`NotAttempted` -- because `binary_office::probe` opens the container with the
/// same call that already failed, so the walk reports "looked, found nothing".
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn a_prefix_is_reported_as_unreadable_rather_than_as_unknown() {
    let whole = fixture("agile_encrypted.docx");

    // 128 bytes is what a header sniff typically reads; 4096 is a whole CFB sector, and
    // is included because "peek a bigger header" is the obvious wrong fix.
    for cut in [128usize, 512, 4096] {
        let part = classify(&whole[..cut]);
        assert_eq!(
            part.container,
            Container::Cfb,
            "{cut}: the signature is in the first eight bytes and is still read"
        );
        assert_eq!(
            part.container_read,
            ContainerRead::Unreadable,
            "{cut}: the directory is past the end of these bytes"
        );
        // The fields a caller would otherwise dispatch on, all uninformative.
        assert_eq!(part.family, Family::Unknown, "{cut}");
        assert_eq!(part.document, Document::Unknown, "{cut}");
    }

    // The control: the same `Family::Unknown`, reached the other way. Without this the
    // test above could pass by reporting `Unreadable` for everything.
    let junk = classify(b"not an office file at all");
    assert_eq!(junk.family, Family::Unknown);
    assert_eq!(junk.container_read, ContainerRead::NotAttempted);

    // And the whole file, so the test cannot pass by never reporting `Opened`.
    assert_eq!(classify(&whole).container_read, ContainerRead::Opened);
    assert_eq!(classify(&whole).family, Family::Agile);
}

/// A CFB this crate opens but recognises nothing in reports `Opened`, not `Unreadable`.
///
/// The distinction the enum exists for, from the other side: `excel97_plain.xls` opens
/// fine and is a format the probe knows, so anything reporting `Unreadable` there would
/// be claiming the bytes were short when they were complete.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn a_container_that_opens_reports_opened_whatever_was_found_inside() {
    for name in [
        "agile_encrypted.docx",
        "standard_encrypted.docx",
        "word97_plain.doc",
        "excel97_plain.xls",
        "excel97_xor.xls",
        "powerpoint97_password.ppt",
    ] {
        assert_eq!(
            classify(&fixture(name)).container_read,
            ContainerRead::Opened,
            "{name}: the container opened, so the verdict describes bytes that were read"
        );
    }
}

#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn a_truncated_binary_container_never_panics_and_never_claims_to_be_plain() {
    // The offsets this reads -- Word's table-stream selector, Excel's FILEPASS walk, and
    // PowerPoint's attacker-chosen `offsetToCurrentEdit` -- are all file-controlled.
    // Truncation is the cheapest way to reach each of them with nothing behind it.
    for name in [
        "word97_password.doc",
        "excel97_password.xls",
        "powerpoint97_password.ppt",
    ] {
        let data = fixture(name);
        for cut in [
            0usize,
            1,
            8,
            64,
            512,
            4096,
            16384,
            data.len() / 2,
            data.len() - 1,
        ] {
            let c = classify(&data[..cut.min(data.len())]);
            // The only verdict dangerous to get wrong here is "not encrypted". Reporting
            // Unknown for a truncated file is correct; reporting Unencrypted would tell a
            // caller the file needs no password.
            assert!(
                c.family != Family::Unencrypted,
                "{name} truncated to {cut}: reported Unencrypted, which is a guess in the                  attacker's favour"
            );
        }
    }
}

// ---- the binary probes, on inputs built to reach every exit --------------------------
//
// Every legacy fixture in the corpus is password-protected, so until these landed no test
// ever asserted `Family::Unencrypted` for a CFB, and the code paths that tell a caller
// "this file needs no password" were never executed. They are the paths worth exercising
// on hostile input, because that is the verdict an attacker wants: a file reported as
// plain is a file the caller will not ask a password for. Each input below is built at
// runtime from record headers rather than committed, for the same reason
// `malformed_input.rs` builds its containers -- a deliberately broken `.xls` in the tree
// is indistinguishable from a corrupt one, and nothing in the bytes would say which
// record was poisoned.

/// A CFB container holding the named streams, and nothing else.
fn cfb_with(streams: &[(&str, &[u8])]) -> Vec<u8> {
    use std::io::Write;
    let mut cursor = std::io::Cursor::new(Vec::new());
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

/// A CFB whose `/Workbook` is exactly `stream`.
fn xls(stream: &[u8]) -> Vec<u8> {
    cfb_with(&[("/Workbook", stream)])
}

// The records these tests are built from. BOF, INTERFACEHDR and the other four in
// `BIFF_NEVER_ENCRYPTED` are the only ones [MS-XLS] lets precede FILEPASS in the
// clear; MMS is the first record a real workbook carries that is not among them.
const BOF: u16 = 0x0809;
const FILEPASS: u16 = 0x002F;
const INTERFACEHDR: u16 = 0x00E1;
const MMS: u16 = 0x00C1;
const EOF_RECORD: u16 = 0x000A;
const BOF_BODY: [u8; 16] = [
    0x00, 0x06, 0x05, 0x00, 0xE7, 0x29, 0xC1, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00,
];

/// The walk ran out before any record decided the verdict: `Unknown`, never
/// `Unencrypted`.
///
/// This is the bug. The old loop returned `Some(false)` from every exit that was not
/// "found FILEPASS", so each input here -- a workbook that is BOF and nothing else, one
/// whose second record claims a body the stream does not have, one cut off in the middle
/// of a header -- classified as an Excel document needing no password. The format is
/// still named: recognising the container is a separate fact from reading it.
#[test]
fn a_workbook_whose_walk_runs_out_is_unknown_not_unencrypted() {
    let bof_only = biff(&[(BOF, &BOF_BODY)]);
    let mut overshoot = biff(&[(BOF, &BOF_BODY)]);
    overshoot.extend_from_slice(&INTERFACEHDR.to_le_bytes());
    overshoot.extend_from_slice(&0x4000u16.to_le_bytes()); // claims 16 KiB, has none
    let mut cut_in_header = biff(&[(BOF, &BOF_BODY), (INTERFACEHDR, &[0xB0, 0x04])]);
    cut_in_header.extend_from_slice(&[0xC1]); // one byte of the next record id

    for (what, stream) in [
        ("BOF and nothing after it", bof_only),
        ("a record whose length overshoots the stream", overshoot),
        ("a stream cut inside a record header", cut_in_header),
    ] {
        let c = classify(&xls(&stream));
        assert_eq!(
            c.document,
            Document::ExcelBinary,
            "{what}: the format is still named"
        );
        assert_eq!(
            c.family,
            Family::Unknown,
            "{what}: an undecided walk must be Unknown, and was reported as {:?}",
            c.family
        );
        assert!(!c.is_encrypted(), "{what}");
    }
}

/// `Unencrypted` is a proof, not a default: it is reported only once a record that the
/// format requires to be encrypted has been read in the clear.
///
/// The negative control for the test above, and the part that keeps the fix from being
/// "everything is Unknown now". A real plain workbook decides at its third record, so
/// the first case is exactly the shape LibreOffice writes; the second is the degenerate
/// legal workbook, BOF then EOF. The third is the boundary: every record that MAY precede
/// FILEPASS, and then nothing -- the permitted set on its own must never decide.
#[test]
fn a_workbook_is_unencrypted_only_once_a_record_that_must_be_encrypted_is_in_the_clear() {
    for (what, records) in [
        (
            "BOF, INTERFACEHDR, then MMS -- the shape a real plain workbook has",
            vec![
                (BOF, &BOF_BODY[..]),
                (INTERFACEHDR, &[0xB0, 0x04][..]),
                (MMS, &[0, 0][..]),
            ],
        ),
        (
            "BOF then EOF",
            vec![(BOF, &BOF_BODY[..]), (EOF_RECORD, &[][..])],
        ),
    ] {
        let c = classify(&xls(&biff(&records)));
        assert_eq!(c.document, Document::ExcelBinary, "{what}");
        assert_eq!(c.family, Family::Unencrypted, "{what}");
        assert!(!c.is_encrypted(), "{what}");
    }

    // Every record permitted ahead of FILEPASS, in the clear, and then the stream ends.
    // None of them is evidence either way, so this must stay Unknown.
    let permitted_only = biff(&[
        (BOF, &BOF_BODY[..]),
        (0x0194, &[0, 0][..]), // USREXCL
        (0x0195, &[0, 0][..]), // FILELOCK
        (INTERFACEHDR, &[0xB0, 0x04][..]),
        (0x0196, &[0, 0][..]), // RRDINFO
        (0x0138, &[0, 0][..]), // RRDHEAD
    ]);
    let c = classify(&xls(&permitted_only));
    assert_eq!(c.document, Document::ExcelBinary);
    assert_eq!(
        c.family,
        Family::Unknown,
        "the permitted-in-the-clear set alone must never decide a verdict"
    );

    // And those same records do not hide a FILEPASS behind them: the walk continues
    // through them and finds it.
    let mut filepass_body = vec![0x01, 0x00]; // wEncryptionType = RC4 CryptoAPI
    filepass_body.extend_from_slice(&[4, 0, 2, 0]); // vMajor 4, vMinor 2
    filepass_body.resize(200, 0);
    let late = biff(&[
        (BOF, &BOF_BODY[..]),
        (0x0194, &[0, 0][..]),
        (INTERFACEHDR, &[0xB0, 0x04][..]),
        (FILEPASS, &filepass_body[..]),
    ]);
    let c = classify(&xls(&late));
    assert_eq!(c.document, Document::ExcelBinary);
    assert_eq!(
        c.family,
        Family::Rc4CryptoApi,
        "FILEPASS after permitted records is found"
    );
    assert_eq!(c.version, Some((4, 2)));
    assert!(c.is_encrypted());
}

/// A `/Workbook` that does not open with BOF is not a workbook, and its records prove
/// nothing -- including the absence of a FILEPASS among them.
///
/// Before the fix this was the cheapest false "plain" of all: any stream in which the
/// walk simply never met 0x002F fell through to `Some(false)`, and a stream of arbitrary
/// bytes usually does not contain that record.
#[test]
fn a_workbook_that_does_not_open_with_bof_is_unknown() {
    let not_bof = biff(&[(MMS, &[0, 0]), (INTERFACEHDR, &[0xB0, 0x04])]);
    let c = classify(&xls(&not_bof));
    assert_eq!(
        c.document,
        Document::ExcelBinary,
        "the container is still an Excel one"
    );
    assert_eq!(c.family, Family::Unknown);
    assert!(!c.is_encrypted());
}

/// FILEPASS is the marker; a body it does not have does not un-mark it.
///
/// `Unsupported` here means "encrypted by a scheme this crate could not name" -- the
/// mapping every encrypted binary verdict without a version pair lands on -- and the two
/// verdicts it must not be are the two an attacker would prefer: `Unencrypted`, or an
/// `Unknown` that a lenient caller might treat the same way.
#[test]
fn a_filepass_record_cut_short_is_still_encrypted() {
    let mut stream = biff(&[(BOF, &BOF_BODY)]);
    stream.extend_from_slice(&FILEPASS.to_le_bytes());
    stream.extend_from_slice(&200u16.to_le_bytes()); // declares a body, then ends
    let c = classify(&xls(&stream));
    assert_eq!(c.document, Document::ExcelBinary);
    assert_eq!(c.family, Family::Unsupported, "encrypted, scheme unread");
    assert!(c.is_encrypted());
    assert_eq!(c.version, None);
}

/// A PowerPoint verdict is believed only when the bytes at `offsetToCurrentEdit` are a
/// `UserEditAtom` header: `recVer`/`recInstance` 0x0000, `recType` 0x0FF5, and a length
/// that is one of the two shapes the atom has.
///
/// The offset is attacker-chosen. The old check read four bytes there and reported
/// "not encrypted" whenever they were not 0x20 -- so any offset pointing at anything but
/// an encrypted atom was a plain presentation. The first two cases are the controls: the
/// two legitimate atom shapes still decide.
#[test]
fn a_powerpoint_edit_atom_is_believed_only_when_its_header_is_a_user_edit_atom() {
    // CurrentUserAtom: rh(8) size(4) headerToken(4) offsetToCurrentEdit(4) = 0.
    let mut current_user = vec![0u8; 16];
    current_user.extend_from_slice(&0u32.to_le_bytes());

    let atom = |ver_instance: u16, rec_type: u16, rec_len: u32| -> Vec<u8> {
        let mut a = Vec::new();
        a.extend_from_slice(&ver_instance.to_le_bytes());
        a.extend_from_slice(&rec_type.to_le_bytes());
        a.extend_from_slice(&rec_len.to_le_bytes());
        a.resize(64, 0);
        a
    };

    for (what, doc, family) in [
        (
            "a plain atom, recLen 0x1C",
            atom(0x0000, 0x0FF5, 0x1C),
            Family::Unencrypted,
        ),
        (
            "an encrypted atom, recLen 0x20",
            atom(0x0000, 0x0FF5, 0x20),
            Family::Unsupported,
        ),
        (
            "the right length on the wrong record type",
            atom(0x0000, 0x0FF6, 0x1C),
            Family::Unknown,
        ),
        (
            "a UserEditAtom with a length it never has",
            atom(0x0000, 0x0FF5, 0x30),
            Family::Unknown,
        ),
        (
            "bytes that are not a record header at all",
            atom(0xFFFF, 0xFFFF, 0x00),
            Family::Unknown,
        ),
    ] {
        let data = cfb_with(&[
            ("/Current User", &current_user),
            ("/PowerPoint Document", &doc),
        ]);
        let c = classify(&data);
        assert_eq!(c.document, Document::PowerPointBinary, "{what}");
        assert_eq!(c.family, family, "{what}: reported {:?}", c.family);
        assert_eq!(c.is_encrypted(), family == Family::Unsupported, "{what}");
    }
}

/// Three documents Microsoft Office wrote with **no** password, which must come back
/// `Unencrypted`: the negative control for
/// `legacy_binary_containers_are_named_without_any_cipher_dependency` and for every
/// "an undecided walk is Unknown" test above.
///
/// Until these landed, every legacy fixture in the corpus was password-protected, so a
/// `probe` that returned `encrypted: Some(true)` unconditionally -- or, after the
/// hardening above, `None` unconditionally -- would have passed every legacy assertion
/// in the suite. Each is the twin of a `*_password` fixture saved by the same application
/// on the same Microsoft 365 build (`tools/gen_plain_binary_fixtures.ps1`), so within a
/// pair the password is the only variable. That is why they are Office-written rather
/// than LibreOffice-written: a control from a different writer leaves "or is that a
/// LibreOffice quirk?" open, which is the ambiguity a control exists to remove.
///
/// Confirmed independently of this crate before committing (olefile): the `.doc` FIB has
/// `fEncrypted` and `fObfuscated` clear; the `.xls` reads `BOF`, `INTERFACEHDR`, `MMS`,
/// so the walk decides at its third record without reaching EOF; the `.ppt`'s
/// `UserEditAtom` header is `0x0000 / 0x0FF5 / 0x1C`. And `msoffcrypto-tool`'s own
/// `is_encrypted()` answers `False` for all three.
///
/// Deliberately **not** added to
/// `a_truncated_binary_container_never_panics_and_never_claims_to_be_plain`: that sweep
/// asserts that a truncated *encrypted* file is never called plain. A plain `.doc` cut at
/// its tail still carries an intact FIB with `fEncrypted` clear and is, correctly, still
/// plain -- measured: `word97_plain.doc` truncated by one byte classifies `Unencrypted`.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn an_unprotected_legacy_binary_document_reports_unencrypted() {
    for (plain, twin, document) in [
        (
            "word97_plain.doc",
            "word97_password.doc",
            Document::WordBinary,
        ),
        (
            "excel97_plain.xls",
            "excel97_password.xls",
            Document::ExcelBinary,
        ),
        (
            "powerpoint97_plain.ppt",
            "powerpoint97_password.ppt",
            Document::PowerPointBinary,
        ),
    ] {
        let c = classify(&fixture(plain));
        assert_eq!(c.container, Container::Cfb, "{plain}");
        assert_eq!(c.document, document, "{plain}");
        assert_eq!(
            c.family,
            Family::Unencrypted,
            "{plain}: an Office-written document with no password reported {:?}",
            c.family
        );
        assert!(!c.is_encrypted(), "{plain}");
        assert_eq!(
            c.version, None,
            "{plain}: there is no EncryptionHeader to report"
        );

        // The pair differs in the password and in nothing the classifier reports on.
        let t = classify(&fixture(twin));
        assert_eq!(
            t.document, document,
            "{twin}: same application, same format"
        );
        assert!(t.is_encrypted(), "{twin}");
    }
}
