//! The standard writer, put through the real read path and the external oracles.
//!
//! The load-bearing tests are the ones that do not consult this module's own idea of the
//! format: `standard::parse_encryption_info` → `derive_standard_key` → `verify_password`
//! on what `write_encryption_info` wrote, the full `decrypt_ooxml` round trip, a diff of
//! the header against the committed fixture, and `office-crypto` reading the container.
//! Real Word 16 and `msoffcrypto-tool` are run on the artifact the last test writes, and
//! their verdicts are recorded in `CHANGELOG.md` against the commit that produced it.

use super::*;
use crate::standard;
use rand::SeedableRng;
use std::io::Read;

const PASSWORD: &str = "testpass";

/// The three key sizes [MS-OFFCRYPTO] §2.3.4.5 defines for this header, with the `AlgID`
/// §2.3.2 pairs with each — **typed literally from the specification**, never read back
/// from `standard_encrypt::AES_ROWS`. A table computed from the implementation proves
/// only that the implementation agrees with itself; these are the numbers the PDF gives,
/// so a table that changed would fail here rather than move silently.
const SPEC_ROWS: [(u32, u32); 3] = [(128, 0x0000_660E), (192, 0x0000_660F), (256, 0x0000_6610)];

/// The seed every golden below is measured under. All-zero, so it is obviously arbitrary
/// and obviously not chosen to make an assertion pass.
const SEED: [u8; 32] = [0u8; 32];

fn seeded() -> chacha20::ChaCha12Rng {
    chacha20::ChaCha12Rng::from_seed(SEED)
}

/// The key size every test that is not about the parameter uses — the default, so that
/// those tests keep measuring the path `encrypt_ooxml_standard` takes.
fn aes128() -> AesKeySize {
    AesKeySize::new(DEFAULT_KEY_BITS).expect("the default must be a size this writer accepts")
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn fixture(name: &str) -> Vec<u8> {
    let path = format!("{}/tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"));
    std::fs::read(&path)
        .unwrap_or_else(|e| panic!("fixture {name} must be present, not optional: {e}"))
}

fn plain_docx() -> Vec<u8> {
    fixture("plain.docx")
}

/// Read one stream out of a CFB container held in memory.
fn stream_of(data: &[u8], path: &str) -> Vec<u8> {
    let mut container =
        cfb::CompoundFile::open(std::io::Cursor::new(data)).expect("input is a CFB container");
    let mut bytes = Vec::new();
    container
        .open_stream(path)
        .unwrap_or_else(|e| panic!("container has no {path}: {e}"))
        .read_to_end(&mut bytes)
        .expect("in-memory read");
    bytes
}

/// Byte-exact material under a seeded RNG — plan D3's payoff at the key-schedule level.
///
/// Pins the draw order (salt, then verifier — nothing in the format enforces it, and
/// swapping the two moves every value below) and the derivation: any change to the KDF,
/// the verifier construction or the ECB calls moves these bytes. Measured, not computed
/// by hand; if a deliberate change moves them, re-measure with the same seed and say in
/// the commit what moved and why.
#[test]
fn the_material_is_byte_exact_under_a_seeded_rng() {
    let m = generate(PASSWORD, aes128(), &mut seeded()).unwrap();
    // One tuple, so a failure reports every value at once rather than the first.
    assert_eq!(
        (
            hex(&m.salt),
            m.key.with_secret(|k| hex(k)),
            hex(&m.encrypted_verifier),
            hex(&m.encrypted_verifier_hash),
        ),
        (
            // The salt is the seed's first 16 bytes -- the same bytes agile's password
            // salt draws under the same seed, which is what "same RNG, same draw" means.
            "9bf49a6a0755f953811fce125f2683d5".to_string(),
            "4fc16bef730947b705f403c5c0cc1e29".to_string(),
            "f27051fb20ffee956d61b12c6d6f767f".to_string(),
            "1f00f5d638f4cd5c3a2a59eaac8daeb146963fe038681093f6f066e95fbe8e8c".to_string(),
        ),
        "(salt, key, EncryptedVerifier, EncryptedVerifierHash)"
    );
}

/// **Write, parse, derive, verify — the four functions a decrypt actually runs.**
///
/// `generate` produces the encryptor, `write_encryption_info` serialises it,
/// `standard::parse_encryption_info` reads it back from bytes, and
/// `standard::verify_password` — the function `decrypt` calls, unmodified — accepts the
/// password against a key re-derived from the parsed salt. Nothing is re-derived here
/// for the test's benefit, so a writer that disagreed with the reader about the salt,
/// the KDF, the verifier hash or its padding fails.
///
/// The wrong-password case is the control. Without it this would pass on a
/// `verify_password` that accepted everything.
#[test]
fn the_generated_encryptor_verifies_through_the_real_path() {
    let m = generate(PASSWORD, aes128(), &mut seeded()).unwrap();
    let stream = write_encryption_info(
        m.key_size,
        &m.salt,
        &m.encrypted_verifier,
        &m.encrypted_verifier_hash,
    );

    let params = standard::parse_encryption_info(&stream[8..])
        .expect("the writer's output must parse -- that is the writer's contract");
    assert_eq!(params.salt, &m.salt);
    assert_eq!(params.encrypted_verifier, &m.encrypted_verifier);
    assert_eq!(params.encrypted_verifier_hash, &m.encrypted_verifier_hash);
    assert_eq!(params.key_size_bytes, AES128_KEY_LEN);

    let key = derive_standard_key(PASSWORD, params.salt, params.key_size_bytes).unwrap();
    standard::verify_password(&key, &params)
        .expect("the password used to generate the encryptor must verify against it");

    // The control.
    let wrong =
        derive_standard_key("not the password", params.salt, params.key_size_bytes).unwrap();
    assert!(
        matches!(
            standard::verify_password(&wrong, &params),
            Err(Error::WrongPassword)
        ),
        "a wrong password must be refused, and by name"
    );
}

/// The decrypt-side view of the two blobs: `EncryptedVerifierHash` unwraps to
/// `SHA1(verifier)` followed by **twelve zero bytes** — the whole 32-byte blob, not just
/// its first 20.
///
/// `verify_password` cannot see the tail (it compares the digest's 20 bytes), so a
/// writer padding with `0x36` would pass every test above and produce a file Word reads
/// as a wrong password — the GH #13 finding, on this format. This is the assertion that
/// catches it.
#[test]
fn the_verifier_hash_blob_is_sha1_of_the_verifier_zero_padded_to_the_whole_blob() {
    let m = generate(PASSWORD, aes128(), &mut seeded()).unwrap();
    let verifier = m
        .key
        .with_secret(|k| standard::aes_ecb_decrypt(k, &m.encrypted_verifier))
        .unwrap();
    let hash_blob = m
        .key
        .with_secret(|k| standard::aes_ecb_decrypt(k, &m.encrypted_verifier_hash))
        .unwrap();

    let mut expected = Sha1::digest(&verifier).to_vec();
    expected.resize(ENCRYPTED_VERIFIER_HASH_LEN, 0);
    assert_eq!(
        hash_blob, expected,
        "SHA1(verifier) || 0x00 * 12, the whole blob"
    );
    assert_eq!(
        &hash_blob[SHA1_LEN..],
        &[0u8; 12],
        "the tail is zero, as Word requires"
    );
    assert_eq!(verifier.len(), VERIFIER_LEN);
}

/// The header is the fixture's layout, byte for byte, with exactly three fields
/// replaced — the three where the fixture is non-conforming — plus the random-derived
/// ones.
///
/// `standard_encrypted.docx` declares `Flags 0x36` and RC4's `AlgID 0x6801` under
/// `fAES`, a pair [MS-OFFCRYPTO] §2.3.2 forbids, and a zero `Flags` copy where the spec
/// says "a copy". The writer emits the conforming `0x24` / `0x660E` / `0x24`. Everything
/// else — the version pair, `EncryptionHeaderSize`, `SizeExtra`, `AlgIDHash`, `KeySize`,
/// `ProviderType`, both reserved words, the CSP name and its terminator, `SaltSize`,
/// `VerifierHashSize`, the 224-byte total — must match the fixture exactly. The
/// fixture's original values are asserted first so the deviation is documented rather
/// than assumed.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn the_encryption_info_is_the_fixtures_layout_with_the_spec_values_where_it_deviates() {
    let m = generate(PASSWORD, aes128(), &mut seeded()).unwrap();
    let ours = write_encryption_info(
        m.key_size,
        &m.salt,
        &m.encrypted_verifier,
        &m.encrypted_verifier_hash,
    );

    let mut expected = stream_of(&fixture("standard_encrypted.docx"), "/EncryptionInfo");
    assert_eq!(expected.len(), 224, "the fixture's stream length");

    let u32_at = |b: &[u8], at: usize| u32::from_le_bytes(b[at..at + 4].try_into().unwrap());
    assert_eq!(u32_at(&expected, 4), 0, "the fixture's Flags copy is zero");
    assert_eq!(u32_at(&expected, 8), 140, "EncryptionHeaderSize");
    assert_eq!(u32_at(&expected, 12), 0x36, "the fixture's Flags");
    assert_eq!(
        u32_at(&expected, 20),
        0x6801,
        "the fixture's AlgID is RC4's"
    );

    // The three non-conforming fields, replaced with the spec's values.
    expected[4..8].copy_from_slice(&FLAGS.to_le_bytes());
    expected[12..16].copy_from_slice(&FLAGS.to_le_bytes());
    expected[20..24].copy_from_slice(&ALG_ID_AES_128.to_le_bytes());

    // The verifier's three random-derived fields. 12 + 140 = the verifier's offset.
    let v = 12 + 140;
    expected[v + 4..v + 20].copy_from_slice(&m.salt);
    expected[v + 20..v + 36].copy_from_slice(&m.encrypted_verifier);
    expected[v + 40..v + 72].copy_from_slice(&m.encrypted_verifier_hash);

    assert_eq!(ours.len(), expected.len());
    assert_eq!(
        ours, expected,
        "the header differs from the fixture somewhere it must not"
    );
}

/// The `EncryptedPackage` stream: the prefix is the plaintext length (not the padded
/// length), the body is the plaintext rounded up to a block, and ECB means the body is
/// the same whether it was chunked or not.
#[test]
fn the_encrypted_package_is_the_length_prefix_and_a_block_padded_body() {
    let key = DerivedKey::new(vec![0x5Au8; 16]);
    for len in [0usize, 1, 15, 16, 17, 4095, 4096, 4097, 10_000] {
        let plain: Vec<u8> = (0..len).map(|i| (i % 251) as u8).collect();
        let stream = encrypt_package(&plain, &key).unwrap();
        let declared = u64::from_le_bytes(stream[..8].try_into().unwrap());
        assert_eq!(declared as usize, len, "the prefix is the plaintext length");
        assert_eq!(
            stream.len() - 8,
            len.div_ceil(16) * 16,
            "the body is the plaintext rounded up to a block"
        );

        // Chunk-independence: one ECB pass over the zero-padded whole is the same bytes.
        let mut padded = plain.clone();
        padded.resize(len.div_ceil(16) * 16, 0);
        let whole = key.with_secret(|k| aes_ecb_encrypt(k, &padded)).unwrap();
        assert_eq!(&stream[8..], &whole[..], "len={len}");
    }
}

/// What this crate writes, this crate reads back — byte for byte, reporting the
/// integrity outcome the format allows — and a wrong password is a wrong password.
#[test]
fn encrypt_ooxml_standard_round_trips_and_reports_no_integrity_element() {
    let plain = plain_docx();
    let container = crate::encrypt_ooxml_standard(&plain, PASSWORD).expect("encrypt");
    assert!(crate::is_cfb_office(&container));

    let crate::Decrypted {
        package: back,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(&container, PASSWORD, crate::IntegrityPolicy::default())
        .expect("what this crate wrote, it must read");
    assert_eq!(outcome, crate::IntegrityOutcome::NotApplicable);
    assert!(
        !outcome.is_authenticated(),
        "the format cannot authenticate"
    );
    assert_eq!(
        back, plain,
        "the package must survive the round trip byte for byte"
    );
    assert_eq!(crate::decrypt_ooxml(&container, PASSWORD).unwrap(), plain);

    assert!(matches!(
        crate::decrypt_ooxml(&container, "not the password"),
        Err(Error::WrongPassword)
    ));
    // A caller demanding a guarantee the format cannot give is refused by name.
    assert!(matches!(
        crate::decrypt_ooxml_with_policy(&container, PASSWORD, crate::IntegrityPolicy::Require),
        Err(Error::IntegrityUnavailable(_))
    ));
}

/// **The negative control against the agile suite.** A flipped ciphertext byte in a
/// standard file is not detected — the format defines nothing that could detect it —
/// so the same tamper that is `IntegrityCheckFailed` for `encrypt_ooxml`'s output
/// *decrypts*, to the wrong bytes, with `NotApplicable`. Pinned so the difference between
/// the two entry points is a test rather than a sentence in a doc comment.
#[test]
fn a_tampered_standard_file_decrypts_to_the_wrong_bytes_because_the_format_cannot_tell() {
    use std::io::{Seek, SeekFrom, Write};
    let plain = plain_docx();
    let mut cursor = std::io::Cursor::new(crate::encrypt_ooxml_standard(&plain, PASSWORD).unwrap());
    {
        let mut cfb = cfb::CompoundFile::open(&mut cursor).expect("we wrote a CFB");
        let mut stream = cfb
            .open_stream("/EncryptedPackage")
            .expect("we wrote the stream");
        // Deep in the ciphertext body, past the prefix and the ZIP local header.
        let at = 8 + 4096 + 100;
        stream.seek(SeekFrom::Start(at)).unwrap();
        let mut b = [0u8; 1];
        stream.read_exact(&mut b).unwrap();
        stream.seek(SeekFrom::Start(at)).unwrap();
        stream.write_all(&[b[0] ^ 0x01]).unwrap();
        stream.flush().unwrap();
        cfb.flush().unwrap();
    }
    let tampered = cursor.into_inner();

    let crate::Decrypted {
        package: wrong,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(&tampered, PASSWORD, crate::IntegrityPolicy::default())
        .expect("standard encryption has no tag to fail");
    assert_eq!(outcome, crate::IntegrityOutcome::NotApplicable);
    assert_eq!(wrong.len(), plain.len());
    assert_ne!(wrong, plain, "the flipped byte must change the plaintext");
    // ECB: the damage is exactly one 16-byte block, and every other byte survives.
    let differing = wrong.iter().zip(&plain).filter(|(a, b)| a != b).count();
    assert!(
        (1..=16).contains(&differing),
        "{differing} bytes differ; ECB confines it to one block"
    );
}

/// `classify` sees exactly the file Office 2007 would have written: standard, AES-128,
/// SHA-1, `4.2`, 16-byte salt, no integrity element to declare. The detection half of the
/// crate and the encrypt half agree about what a file is.
#[test]
fn encrypt_ooxml_standard_output_classifies_as_office_2007_aes_128() {
    let c = crate::classify(&crate::encrypt_ooxml_standard(&plain_docx(), PASSWORD).unwrap());
    assert_eq!(c.container, crate::Container::Cfb);
    assert_eq!(c.document, crate::Document::OoxmlPackage);
    assert_eq!(c.version, Some(STANDARD_VERSION));
    assert_eq!(c.family, crate::Family::Standard);
    assert_eq!(c.data_integrity, crate::IntegrityDeclaration::NotApplicable);
    assert!(c.is_encrypted() && c.is_supported());
    assert_eq!(
        c.key_data.expect("the EncryptionHeader is present"),
        crate::AlgorithmParams {
            cipher: Some(crate::CipherAlgorithm::Aes),
            hash: Some(crate::HashAlgorithm::Sha1),
            key_bits: Some(128),
            block_size: None,
            salt_size: Some(16),
            spin_count: None,
        }
    );
    assert!(c.password_key.is_none());
}

/// The degenerate package: an 8-byte `EncryptedPackage` of nothing but its prefix, and
/// it round-trips to an empty `Vec`.
///
/// Driven through the seeded core rather than [`crate::encrypt_ooxml_standard`], because
/// the public entry point now refuses an input that is not a plain package and `&[]` is
/// not one — that refusal is `bytes_that_are_no_container_are_refused` in `lib.rs`. What
/// is checked here is the writer's handling of a degenerate payload, which is a separate
/// fact from whether the guard lets one through.
#[test]
fn an_empty_package_round_trips() {
    let container = encrypt(&[], PASSWORD, DEFAULT_KEY_BITS, &mut seeded()).unwrap();
    assert_eq!(stream_of(&container, "/EncryptedPackage").len(), 8);
    let crate::Decrypted {
        package: back,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(&container, PASSWORD, crate::IntegrityPolicy::default())
        .unwrap();
    assert_eq!(outcome, crate::IntegrityOutcome::NotApplicable);
    assert!(back.is_empty());
}

/// The whole write path under a seeded RNG is byte-stable, and this is its golden —
/// `EncryptionInfo`, `EncryptedPackage`, the four `DataSpaces` blobs and the CFB
/// directory `cfb` builds around them, timestamps zeroed. A change to any step moves
/// this digest.
///
/// Measured, not computed. If a deliberate change moves it, re-measure with the same
/// seed and say in the commit what moved and why. The plaintext is `plain.docx`, the
/// file the committed standard fixture decrypts to, so the fixture tests and this golden
/// are about the same bytes.
#[test]
fn the_whole_container_is_byte_exact_under_a_seeded_rng() {
    use sha2::{Digest as _, Sha256};
    let plain = plain_docx();
    let container = encrypt(&plain, PASSWORD, DEFAULT_KEY_BITS, &mut seeded()).unwrap();
    let digest = hex(&Sha256::digest(&container));
    assert_eq!(
        (container.len(), digest.as_str()),
        (GOLDEN_LEN, GOLDEN_SHA256),
        "the container's (length, SHA-256) under the seeded RNG"
    );

    // And the same seed reproduces it exactly; a golden that only held once is a fluke.
    let again = encrypt(&plain, PASSWORD, DEFAULT_KEY_BITS, &mut seeded()).unwrap();
    assert_eq!(container, again);
}

/// Measured on 2026-09-05 under `SEED`, `PASSWORD` and `plain.docx`.
const GOLDEN_LEN: usize = 40960;
const GOLDEN_SHA256: &str = "491298746ce1e46aed2c98b7bc9c061d97ccedf578062a901aa32f2249bc7e42";

/// The same seed reproduces everything; a different seed reproduces nothing — so the
/// generator is consuming the RNG rather than deriving from the password alone, which
/// would hand every document written with one password the same salt and key.
#[test]
fn the_rng_decides_the_output_and_the_seed_decides_the_rng() {
    let a = generate(PASSWORD, aes128(), &mut seeded()).unwrap();
    let b = generate(PASSWORD, aes128(), &mut seeded()).unwrap();
    assert_eq!(a.salt, b.salt);
    assert_eq!(a.encrypted_verifier, b.encrypted_verifier);
    assert_eq!(a.encrypted_verifier_hash, b.encrypted_verifier_hash);
    assert!(a.key.with_secret(|x| b.key.with_secret(|y| x == y)));

    let c = generate(
        PASSWORD,
        aes128(),
        &mut chacha20::ChaCha12Rng::from_seed([1u8; 32]),
    )
    .unwrap();
    assert_ne!(a.salt, c.salt);
    assert_ne!(a.encrypted_verifier, c.encrypted_verifier);
    assert_ne!(a.encrypted_verifier_hash, c.encrypted_verifier_hash);
    assert!(
        a.key.with_secret(|x| c.key.with_secret(|y| x != y)),
        "a different salt derives a different key from the same password"
    );
}

/// The public entry point works and is not accidentally deterministic: two calls must
/// differ, or the system RNG is not being consumed.
#[test]
fn the_public_entry_point_produces_a_fresh_file_each_call() {
    let a = crate::encrypt_ooxml_standard(b"PK\x03\x04 not really a zip", PASSWORD).unwrap();
    let b = crate::encrypt_ooxml_standard(b"PK\x03\x04 not really a zip", PASSWORD).unwrap();
    assert_ne!(
        a, b,
        "two encryptions of the same input must not share a salt"
    );
    assert_eq!(a.len(), b.len(), "but they are the same shape");
}

/// **An independent implementation reads a file this crate generated** — header,
/// verifier, package, container, all ours. `office-crypto` is a separate MIT crate with
/// its own CFB reader and its own parse of [MS-OFFCRYPTO] §2.3.4.5, so agreeing with it
/// is evidence about the file. The other external readers are run on the artifact the
/// last test writes, and recorded.
#[test]
fn an_independent_implementation_reads_what_encrypt_ooxml_standard_wrote() {
    let plain = plain_docx();
    let container = crate::encrypt_ooxml_standard(&plain, PASSWORD).unwrap();
    let theirs = office_crypto::decrypt_from_bytes(container, PASSWORD)
        .expect("office-crypto must open a file this crate wrote");
    assert_eq!(theirs, plain);
}

/// A package over the ceiling is refused before any key is derived or byte encrypted.
/// The allocation is lazy on every platform this runs on, so the test costs an
/// inequality, not a gigabyte.
///
/// The input must be a **plain ZIP** to reach the ceiling at all: the shape guard runs
/// first, so a gigabyte of zeros is now `UnknownContainer` rather than a size refusal.
/// That ordering has its own test, `the_shape_guard_runs_before_the_payload_ceiling`,
/// and this one is its negative control.
#[test]
fn a_package_over_the_ceiling_is_refused_before_any_work() {
    // `vec![0u8; N]` is `alloc_zeroed`, so these are lazy zero pages. Writing the magic
    // in place touches one of them; `resize` from a 4-byte `Vec` would memset a
    // gigabyte and make the comment above false.
    let mut big = vec![0u8; crate::limits::PAYLOAD_CEILING + 1];
    big[..4].copy_from_slice(b"PK\x03\x04");
    let got = crate::encrypt_ooxml_standard(&big, PASSWORD).map(|c| c.len());
    assert!(
        matches!(&got, Err(Error::BadParameters(msg)) if msg.contains("PAYLOAD_CEILING")),
        "over the ceiling must be refused by name, got: {got:?}"
    );
}

/// The shape guard runs before the payload ceiling, and the order is the claim. The
/// agile mirror of this test carries the full argument; reverse the two checks in
/// `crate::encrypt_ooxml_standard` and this fails with the ceiling message.
#[test]
fn the_shape_guard_runs_before_the_payload_ceiling() {
    let big = vec![0u8; crate::limits::PAYLOAD_CEILING + 1];
    let got = crate::encrypt_ooxml_standard(&big, PASSWORD).map(|c| c.len());
    assert!(
        matches!(&got, Err(Error::UnknownContainer)),
        "oversized junk is not a package first and oversized second, got: {got:?}"
    );
}

/// Writes the artifact the external readers are run against — real Word over COM,
/// LibreOffice over UNO, `msoffcrypto-tool` and `office-crypto` — and prints where.
/// `cargo test` cannot drive the first two; `tools/acceptance_gate.py` drives all four
/// and their verdicts are recorded in `CHANGELOG.md` against this exact file.
///
/// The directory is `MSOFFICE_CRYPTO_ARTIFACT_DIR` when set, else the system temp
/// directory — the same rule as the agile artifact, and for the same reason. **This test
/// honoured neither until the secure-gate rc.12 work**: it hardcoded the system temp
/// directory, so the one artifact whose writer was being changed was the one a
/// concurrent session could substitute. The mitigation was built for #8, documented on
/// the agile test, and applied to one of the two files.
///
/// The SHA-256 below is the one the gate prints on its first line: if they differ, the
/// gate did not read this run's file. That check is why the evidence in `CHANGELOG.md`
/// is recorded against a digest rather than a path, and this test could not support it
/// before, because it printed no digest at all.
#[test]
fn encrypt_ooxml_standard_writes_the_artifact_the_external_readers_are_run_on() {
    let dir = std::env::var_os("MSOFFICE_CRYPTO_ARTIFACT_DIR")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(std::env::temp_dir);
    let path = dir.join("msoffice_crypto_encrypt_ooxml_standard.docx");
    std::fs::write(
        &path,
        crate::encrypt_ooxml_standard(&plain_docx(), PASSWORD).unwrap(),
    )
    .unwrap();
    // Digested from the file rather than from the buffer that was written, so the number
    // below describes the bytes the gate will actually hash. A digest of the in-memory
    // value would agree with the gate in every case except the one worth catching -- a
    // short or transformed write -- and would claim to have verified it.
    let written = std::fs::read(&path).unwrap();
    let digest = hex(&crate::hash::HashAlgorithm::Sha256.digest(&written));
    println!(
        "encrypt_ooxml_standard artifact written to {} ({} bytes, SHA-256 {digest})",
        path.display(),
        written.len()
    );
}

/// **The default has not moved, and this is the assertion that says so by name.**
///
/// The golden above is byte-for-byte evidence for one key size; it would still hold if
/// `encrypt_ooxml_standard` stopped going through the parameterised entry point. This
/// pins the two facts that make the golden mean what it claims: the default key size is
/// AES-128, and the public no-parameter function writes exactly what the parameterised
/// one writes when handed it.
#[test]
fn the_default_key_size_is_aes_128_and_both_entry_points_write_it() {
    assert_eq!(
        DEFAULT_KEY_BITS, 128,
        "Office 2007 wrote AES-128; every acceptance-gate verdict and the committed \
         golden were measured on it"
    );
    assert_eq!(
        DEFAULT_KEY_BITS, SPEC_ROWS[0].0,
        "the default is the first row of [MS-OFFCRYPTO] 2.3.4.5's table"
    );

    // Not a byte comparison of two containers -- each call draws its own salt -- but a
    // comparison of what the two paths declare, which is what the parameter decides.
    let plain = plain_docx();
    let default = crate::classify(&crate::encrypt_ooxml_standard(&plain, PASSWORD).unwrap());
    let explicit = crate::classify(
        &crate::encrypt_ooxml_standard_with_key_bits(&plain, PASSWORD, DEFAULT_KEY_BITS).unwrap(),
    );
    assert_eq!(default.key_data, explicit.key_data);
    assert_eq!(default.family, explicit.family);
    assert_eq!(default.version, explicit.version);
}

/// **All three key sizes round-trip through the public read path, each with its own
/// wrong-password control.**
///
/// `decrypt_ooxml_with_policy` under the fail-closed default policy rather than a direct
/// call to `standard::decrypt`, so what is proved is the path a consumer takes: the
/// container, the header parse, the KDF at the declared key length, the verifier, the
/// package. The control is what makes it evidence: without it the test would pass on a
/// `verify_password` that accepted anything, and on a reader that ignored `KeySize` and
/// derived 16 bytes whatever the file said.
#[test]
fn every_key_size_round_trips_and_a_wrong_password_is_still_refused() {
    let plain = plain_docx();
    for (bits, _) in SPEC_ROWS {
        let container = crate::encrypt_ooxml_standard_with_key_bits(&plain, PASSWORD, bits)
            .unwrap_or_else(|e| panic!("AES-{bits} must be writable: {e}"));
        assert!(crate::is_cfb_office(&container));

        let crate::Decrypted {
            package: back,
            integrity: outcome,
        } = crate::decrypt_ooxml_with_policy(
            &container,
            PASSWORD,
            crate::IntegrityPolicy::default(),
        )
        .unwrap_or_else(|e| panic!("what this crate wrote at AES-{bits}, it must read: {e}"));
        assert_eq!(outcome, crate::IntegrityOutcome::NotApplicable);
        assert_eq!(back, plain, "AES-{bits} must survive the round trip");

        // The control, per key size.
        assert!(
            matches!(
                crate::decrypt_ooxml(&container, "not the password"),
                Err(Error::WrongPassword)
            ),
            "a wrong password must be refused at AES-{bits}, and by name"
        );
    }
}

/// `classify` reads the `AlgID` and the `KeySize` back out of each artifact, and they are
/// the pair [MS-OFFCRYPTO] §2.3.2 requires.
///
/// Two readings, because `classify` reports only one of the two fields: `key_bits` comes
/// from the classifier, and the `AlgID` is read out of the `\EncryptionInfo` stream at
/// the offset §2.3.4.5's layout puts it. The expectations are [`SPEC_ROWS`], typed from
/// the specification rather than from this module's own table — the point is that the
/// bytes on disk match the document, not that two of our constants match each other.
#[test]
fn classify_reads_back_the_key_size_and_the_alg_id_of_every_artifact() {
    let plain = plain_docx();
    for (bits, alg_id) in SPEC_ROWS {
        let container =
            crate::encrypt_ooxml_standard_with_key_bits(&plain, PASSWORD, bits).unwrap();

        let c = crate::classify(&container);
        assert_eq!(c.family, crate::Family::Standard, "AES-{bits}");
        assert_eq!(c.version, Some(STANDARD_VERSION));
        assert_eq!(c.data_integrity, crate::IntegrityDeclaration::NotApplicable);
        assert_eq!(
            c.key_data.expect("the EncryptionHeader is present"),
            crate::AlgorithmParams {
                cipher: Some(crate::CipherAlgorithm::Aes),
                // SHA-1 and the 16-byte salt do not move with the key size: 2.3.4.5
                // fixes AlgIDHash and 2.3.3 fixes SaltSize.
                hash: Some(crate::HashAlgorithm::Sha1),
                key_bits: Some(bits),
                block_size: None,
                salt_size: Some(16),
                spin_count: None,
            },
            "classify must report AES-{bits}"
        );

        // The stream itself. Offsets from 2.3.4.5: vMajor(2) vMinor(2) Flags-copy(4)
        // EncryptionHeaderSize(4), then the header -- Flags(4) SizeExtra(4) AlgID(4)
        // AlgIDHash(4) KeySize(4) -- so AlgID is at 20 and KeySize at 28.
        let stream = stream_of(&container, "/EncryptionInfo");
        let u32_at = |at: usize| u32::from_le_bytes(stream[at..at + 4].try_into().unwrap());
        assert_eq!(u32_at(20), alg_id, "AlgID for AES-{bits}");
        assert_eq!(u32_at(28), bits, "KeySize for AES-{bits}");
        assert_eq!(
            u32_at(12),
            FLAGS,
            "fCryptoAPI | fAES, whatever the key size"
        );
        assert_eq!(
            u32_at(8),
            140,
            "EncryptionHeaderSize does not move with the key"
        );
        assert_eq!(stream.len(), 224, "nor does the stream's length");
    }
}

/// The writer's output at every key size goes through `standard`'s own reader, and the
/// key it declares is the key it was encrypted under — `KeySize / 8` bytes, cut from the
/// same ladder.
///
/// The failure this catches is the quiet one: a writer that declared AES-256 and
/// encrypted under a 16-byte key would still round-trip through *this crate* if the
/// reader made the same mistake. Here the length is asserted against the arithmetic
/// §2.3.4.7 states, not against what either side chose.
#[test]
fn the_declared_key_size_is_the_key_length_the_verifier_was_built_under() {
    for (bits, _) in SPEC_ROWS {
        let key_size = AesKeySize::new(bits).unwrap();
        let m = generate(PASSWORD, key_size, &mut seeded()).unwrap();
        assert_eq!(m.key.with_secret(|k| k.len()), (bits / 8) as usize);

        let stream = write_encryption_info(
            m.key_size,
            &m.salt,
            &m.encrypted_verifier,
            &m.encrypted_verifier_hash,
        );
        let params = standard::parse_encryption_info(&stream[8..])
            .unwrap_or_else(|e| panic!("AES-{bits}: the writer's output must parse: {e}"));
        assert_eq!(params.key_size_bytes, (bits / 8) as usize);

        let key = derive_standard_key(PASSWORD, params.salt, params.key_size_bytes).unwrap();
        standard::verify_password(&key, &params)
            .unwrap_or_else(|e| panic!("AES-{bits}: the password must verify: {e}"));

        // The control: the same file read at the wrong key length must not verify --
        // which is why declaring the size correctly matters.
        if params.key_size_bytes != AES128_KEY_LEN {
            let short = derive_standard_key(PASSWORD, params.salt, AES128_KEY_LEN).unwrap();
            assert!(
                matches!(
                    standard::verify_password(&short, &params),
                    Err(Error::WrongPassword)
                ),
                "AES-{bits} must not verify under a 128-bit key"
            );
        }
    }
}

/// **The guard.** A key size [MS-OFFCRYPTO] §2.3.4.5 does not define is refused by name,
/// with both payload halves asserted and the nearest accepted size in `min` and `max`.
///
/// Delete the `Err` arm of `AesKeySize::new` — fall through to
/// `Self { bits: key_bits, alg_id: ALG_ID_AES_128 }` for an unrecognised size — and this
/// test fails on the first case, an `Ok` arriving where an error was demanded. The
/// failure that was actually observed is recorded in the report for this change.
///
/// The cases cover each way a value can be wrong: zero, below the smallest size, a legal
/// RC4 `KeySize` under §2.3.2's other row, values between two AES sizes, above them all,
/// and the `u32` ceiling. None of them is `UnsupportedByCipher`: on this path §2.3.4.5
/// enumerates AES's own three sizes, so the format refuses first and there is no value it
/// permits that AES cannot key. `AesKeySize::new` carries that argument.
#[test]
fn a_key_size_the_format_does_not_define_is_refused_before_the_password() {
    // (asked for, the nearest accepted size the refusal must name)
    for (got, nearest) in [
        (0u32, 128u32),
        (7, 128),
        (8, 128),
        (64, 128),
        (127, 128),
        (129, 128),
        // 160 is 32 from both 128 and 192: the tie resolves upwards, to the stronger key.
        (160, 192),
        (200, 192),
        (255, 256),
        (257, 256),
        (320, 256),
        (512, 256),
        (u32::MAX, 256),
    ] {
        let err = crate::encrypt_ooxml_standard_with_key_bits(&plain_docx(), PASSWORD, got)
            .expect_err("only 128, 192 and 256 are writable");
        assert!(
            matches!(
                err,
                Error::EncryptParams {
                    param: crate::EncryptParam::KeySize,
                    problem: crate::EncryptParamProblem::OutsideSpecRange,
                    got: g,
                    min,
                    max,
                } if g == got && min == nearest && max == nearest
            ),
            "{got} must be (KeySize, OutsideSpecRange) naming {nearest}, got: {err:?}"
        );

        // The rendered sentence, which a consumer with no match arm forwards: it must
        // name the field of the file being written, not the agile attribute.
        let text = err.to_string();
        assert!(
            text.contains("EncryptionHeader.KeySize") && !text.contains("keyBits"),
            "the message must name this format's field: {text}"
        );
    }

    // The negative control: the three the format does define are accepted, so the loop
    // above cannot be passing on a writer that refuses everything.
    for (bits, _) in SPEC_ROWS {
        AesKeySize::new(bits).unwrap_or_else(|e| panic!("AES-{bits} must be accepted: {e}"));
    }
}

/// The refusal costs no randomness and no key derivation — it happens before the first
/// draw, which is the `agile_encrypt::encrypt` ordering and the reason a caller can ask
/// the question cheaply.
///
/// Proved by consuming the RNG afterwards: if `encrypt` had drawn from it, the next 16
/// bytes would differ from a fresh generator's first 16.
#[test]
fn a_refused_key_size_consumes_no_rng_draw() {
    let mut used = seeded();
    let err = encrypt(&plain_docx(), PASSWORD, 200, &mut used)
        .expect_err("200 bits is not a key size this format defines");
    assert!(matches!(err, Error::EncryptParams { .. }));

    let (mut after, mut fresh_bytes) = ([0u8; 16], [0u8; 16]);
    fill(&mut used, &mut after).unwrap();
    fill(&mut seeded(), &mut fresh_bytes).unwrap();
    assert_eq!(
        after, fresh_bytes,
        "the generator must be untouched: the refusal precedes the first draw"
    );
}

/// The two other key sizes, written where the owner can double-click them, beside the
/// AES-128 artifact.
///
/// Word is the reference for this format and the only reader that can settle whether a
/// standard-encryption file at AES-192 or AES-256 opens; `cargo test` cannot drive it, so
/// this writes the files and prints their digests and the verdict is collected by hand.
/// **No assertion about an external reader is made here** — that would be a claim this
/// test cannot support.
#[test]
fn the_other_two_key_sizes_are_written_where_the_owner_can_open_them() {
    let dir = std::env::var_os("MSOFFICE_CRYPTO_ARTIFACT_DIR")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(std::env::temp_dir);
    for (bits, _) in SPEC_ROWS.iter().skip(1) {
        let path = dir.join(format!(
            "msoffice_crypto_encrypt_ooxml_standard_aes{bits}.docx"
        ));
        std::fs::write(
            &path,
            crate::encrypt_ooxml_standard_with_key_bits(&plain_docx(), PASSWORD, *bits).unwrap(),
        )
        .unwrap();
        let written = std::fs::read(&path).unwrap();
        let digest = hex(&crate::hash::HashAlgorithm::Sha256.digest(&written));
        println!(
            "standard AES-{bits} artifact written to {} ({} bytes, SHA-256 {digest})",
            path.display(),
            written.len()
        );
    }
}
