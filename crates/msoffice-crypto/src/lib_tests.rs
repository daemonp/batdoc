//! Tests for the crate root's public API.
//!
//! Split out of `lib.rs` rather than written inline, which is this crate's convention
//! everywhere else (`classify_tests.rs`, `encryption_info_tests.rs` and five more) and
//! which `lib.rs` and `sensitive.rs` were the only two modules not to follow.
//!
//! The move has a second effect worth stating, because it is why it happened now:
//! `.github/codeql/codeql-config.yml` can only exclude files that are *entirely* test
//! code, and an inline `#[cfg(test)] mod tests` inside a production file cannot be
//! filtered without blinding the scanner to the production half. Fifteen
//! `rust/hard-coded-cryptographic-value` alerts pointed at `"testpass"` in here, which is
//! the fixture password CLAUDE.md mandates. Test vectors in a known-answer test are
//! hard-coded keys by definition; that is what makes them known answers.

use super::*;

#[test]
fn test_is_cfb_office_magic() {
    let cfb = [0xD0u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00];
    assert!(is_cfb_office(&cfb));
}

#[test]
fn test_is_cfb_office_not_zip() {
    assert!(!is_cfb_office(b"PK\x03\x04something"));
}

#[test]
fn test_is_cfb_office_not_pdf() {
    assert!(!is_cfb_office(b"%PDF-1.4"));
}

#[test]
fn test_is_cfb_office_too_short() {
    assert!(!is_cfb_office(&[0xD0, 0xCF, 0x11]));
}

/// Everything past detection. Gated as a child module rather than by tagging each
/// test, so `cargo test --no-default-features` still runs the four `is_cfb_office`
/// cases above instead of finding an empty suite.
#[cfg(feature = "crypto-ops")]
mod crypto_ops {
    use crate::*;

    #[test]
    fn test_decrypt_non_cfb_returns_error() {
        let result = decrypt_ooxml(b"PK\x03\x04not a cfb", "password");
        assert!(matches!(result, Err(Error::NotACfbFile)));
    }

    /// End-to-end fixture tests. The fixtures are committed in tests/fixtures/
    /// and are NOT optional: a missing one fails the test rather than skipping,
    /// so the suite cannot go green while exercising nothing.
    #[test]
    fn test_agile_fixture_decrypts_to_zip() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/agile_encrypted.docx"
        );
        let data = std::fs::read(fixture_path)
            .expect("fixture must be present -- agile tests are not optional");
        assert!(is_cfb_office(&data));
        let plain = decrypt_ooxml(&data, "testpass").expect("Agile decrypt must succeed");
        assert!(
            plain.starts_with(b"PK\x03\x04"),
            "Decrypted output must be ZIP"
        );
    }

    #[test]
    fn test_agile_fixture_wrong_password() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/agile_encrypted.docx"
        );
        let data = std::fs::read(fixture_path)
            .expect("fixture must be present -- agile tests are not optional");
        let result = decrypt_ooxml(&data, "wrongpass");
        assert!(
            matches!(result, Err(Error::WrongPassword)),
            "Wrong password must return WrongPassword"
        );
    }

    // ---- non-SHA-512 agile fixtures (issue #11) ---------------------------------

    /// Every agile fixture that is not AES-256/SHA-512: the two GH #11 added and the
    /// three GH #13 added, all from `tools/gen_agile_fixtures.py`.
    const NON_SHA512_AGILE_FIXTURES: [&str; 5] = [
        "agile_aes256_sha384.docx",
        "agile_aes256_sha256.docx",
        "agile_aes128_sha1.docx",
        "agile_aes128_sha384.docx",
        "agile_aes192_sha384.docx",
    ];

    fn fixture(name: &str) -> Vec<u8> {
        let path = format!("{}/tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"));
        std::fs::read(&path)
            .unwrap_or_else(|e| panic!("fixture {name} must be present, not optional: {e}"))
    }

    /// An agile file may name any of four hashes on `<p:encryptedKey>`, and until
    /// issue #11 this crate ran SHA-512 for all of them: wrong `H_final`, wrong block
    /// keys, failed verifier, and the user told their password was wrong when it was
    /// right.
    ///
    /// **What these two fixtures are.** Written by msoffcrypto-tool 6.x — an
    /// independent MIT implementation with its own reading of [MS-OFFCRYPTO] — with
    /// its hardcoded parameter tuple replaced and nothing else changed, by
    /// `tools/gen_agile_fixtures.py`. Each was read back to a byte-identical
    /// `plain.docx` by msoffcrypto's own CLI before being committed, and is asserted
    /// here against that same known plaintext rather than against a round trip.
    ///
    /// **What they are not.** Evidence of agreement with Microsoft's writer. No real
    /// Office file with these tuples exists in any local corpus and none can be
    /// produced here; the container is the same shape as the SHA-512 fixture beside
    /// them, and that is as far as the claim goes. That fixture is **not** an Office
    /// artefact either: `agile_encrypted.docx` was written by msoffcrypto-tool over a
    /// python-docx `plain.docx`, like the other two — its `/EncryptionInfo` stream is
    /// byte-identical to msoffcrypto's `toEncryptionDescriptor()` template
    /// (`msoffcrypto/method/ecma376_agile.py:138-152`) rendered with the fixture's own
    /// attribute values, four-space indentation and `xmlns:c` included, where Office
    /// writes that stream unindented. **No Office-written agile fixture exists in this
    /// corpus, for any tuple** — the whole agile suite is one third-party writer read
    /// back by two readers.
    ///
    /// **Since GH #13, the four tuples that exist in the wild are all here** — the
    /// three LibreOffice writes besides Office 16's own, `(128, SHA1)`, `(128, SHA384)`
    /// and `(192, SHA384)`, produced by the same independent writer with its `keyBits`
    /// replaced. The AES-128/SHA-1 one is Word 2010's default. **Real Word 16 opens
    /// all five** (`tools/office_com_check.ps1`, recorded in CHANGELOG.md for
    /// 2026-09-05), which is evidence that each is a file Office recognises,
    /// separate from the evidence that this crate reads it — and it was not free:
    /// the SHA-1 file only opened once its `dataIntegrity` blobs were zero-padded,
    /// and the AES-192 file is the one whose 24-byte session key travels in a 32-byte
    /// blob. The SHA-1 fixture is also the only one whose blobs carry a `hashSize`
    /// pad at all, so it is the one that exercises `integrity::unwrap_blob`'s
    /// truncation on a real container.
    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn test_non_sha512_agile_fixtures_decrypt_to_the_known_plaintext() {
        let plain = fixture("plain.docx");
        for name in NON_SHA512_AGILE_FIXTURES {
            let data = fixture(name);
            assert!(is_cfb_office(&data));
            let crate::Decrypted {
                package: out,
                integrity: outcome,
            } = decrypt_ooxml_with_policy(&data, "testpass", IntegrityPolicy::Require)
                .unwrap_or_else(|e| panic!("{name} must decrypt: {e}"));
            assert_eq!(out, plain, "{name} must decrypt to the known plaintext");
            // Its dataIntegrity HMAC runs on the same non-SHA-512 algorithm, so
            // `Require` also proves `<keyData>`'s half is honoured end to end.
            assert_eq!(outcome, IntegrityOutcome::Verified, "{name}");
        }
    }

    /// The negative control the hash work needs: a wrong password on a non-SHA-512
    /// file must still be `WrongPassword`. Without it the test above cannot
    /// distinguish "the dispatch is wired" from "this tuple always errors" — which
    /// is exactly what the pre-#11 crate did, for every password.
    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn test_non_sha512_agile_fixtures_still_report_a_wrong_password() {
        for name in NON_SHA512_AGILE_FIXTURES {
            let result = decrypt_ooxml(&fixture(name), "wrongpass");
            assert!(
                matches!(result, Err(Error::WrongPassword)),
                "{name} with the wrong password got {:?}",
                result.map(|p| p.len())
            );
        }
    }

    /// The SHA-512 fixtures decrypt byte-identically to the same known plaintext.
    /// The hash work touched every derivation on the password path, so "still starts
    /// with PK" is not a strong enough regression assertion for it.
    #[test]
    fn test_sha512_agile_fixture_still_decrypts_byte_identically() {
        assert_eq!(
            decrypt_ooxml(&fixture("agile_encrypted.docx"), "testpass").unwrap(),
            fixture("plain.docx")
        );
    }

    /// The standard fixture decrypts to `plain.docx`, byte for byte.
    ///
    /// This asserted only `starts_with(b"PK\x03\x04")` until 2026-09-05, which a
    /// wrong-but-ZIP-shaped decrypt passes: a mis-derived key that happened to leave
    /// the first block intact, a segment boundary off by one, a truncation a few bytes
    /// early. The plaintext to compare against was already committed, and the
    /// expectation is external to this crate: `msoffcrypto-tool -p testpass` decrypts
    /// `standard_encrypted.docx` to a file whose SHA-256 is
    /// `285ce3ad5021f04436e3…`, identical to `tests/fixtures/plain.docx` (36 678
    /// bytes), measured before this assertion was written. The ZIP-shape check stays
    /// first so a failure is diagnosable rather than reducing to "36 678 bytes differ".
    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn test_standard_fixture_decrypts_to_zip() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/standard_encrypted.docx"
        );
        let data = std::fs::read(fixture_path)
            .expect("fixture must be present -- standard tests are not optional");
        assert!(is_cfb_office(&data));
        let plain = decrypt_ooxml(&data, "testpass").expect("Standard decrypt must succeed");
        assert!(
            plain.starts_with(b"PK\x03\x04"),
            "Decrypted output must be ZIP"
        );
        assert_eq!(
            plain,
            fixture("plain.docx"),
            "the standard fixture must decrypt to plain.docx byte for byte"
        );
    }

    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn test_standard_fixture_wrong_password() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/standard_encrypted.docx"
        );
        let data = std::fs::read(fixture_path)
            .expect("fixture must be present -- standard tests are not optional");
        let result = decrypt_ooxml(&data, "wrongpass");
        assert!(
            matches!(result, Err(Error::WrongPassword)),
            "Wrong password must return WrongPassword"
        );
    }

    // ---- dataIntegrity (S2 / F18) ------------------------------------------------

    fn agile_fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/agile_encrypted.docx"
        ))
        .expect("fixture must be present -- agile tests are not optional")
    }

    /// Flip one bit of ciphertext inside the CFB container's `EncryptedPackage`
    /// stream, well past the first segment so the decrypted ZIP header survives.
    ///
    /// Built at runtime rather than committed as a second binary fixture: a tampered
    /// `.docx` in the tree is indistinguishable from a corrupt one, and nothing in the
    /// file would record *which* byte was flipped or why.
    fn tamper_agile_fixture(offset: u64) -> Vec<u8> {
        use std::io::{Read, Seek, SeekFrom, Write};

        let mut cursor = std::io::Cursor::new(agile_fixture());
        {
            let mut container =
                cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB container");
            let mut stream = container
                .open_stream("/EncryptedPackage")
                .expect("fixture has an EncryptedPackage stream");

            stream.seek(SeekFrom::Start(offset)).unwrap();
            let mut byte = [0u8; 1];
            stream.read_exact(&mut byte).unwrap();
            byte[0] ^= 0x01;
            stream.seek(SeekFrom::Start(offset)).unwrap();
            stream.write_all(&byte).unwrap();
            stream.flush().unwrap();
        }
        cursor.into_inner()
    }

    /// Delete the whole `<dataIntegrity .../>` element from the fixture's
    /// `EncryptionInfo` XML.
    ///
    /// That XML is stored in the clear, so this touches no cryptographic parameter and
    /// needs no foreign writer: what comes back is a valid agile document that simply
    /// declares no tag. The base64 alphabet contains `/` but not `>`, so the first
    /// `/>` after the element name is its own terminator.
    ///
    /// This is not "an old file" — GH #12 found no writer that omits the element and
    /// no corpus file lacking it. It is the downgrade attack: the ~200 bytes an
    /// attacker deletes to turn off the integrity check, and nothing else changes.
    fn agile_fixture_without_data_integrity() -> Vec<u8> {
        use std::io::{Read, Seek, SeekFrom, Write};

        let mut cursor = std::io::Cursor::new(agile_fixture());
        {
            let mut container =
                cfb::CompoundFile::open(&mut cursor).expect("fixture is a CFB container");

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
            info.drain(start..end);

            let mut stream = container.open_stream("/EncryptionInfo").unwrap();
            stream.set_len(0).unwrap();
            stream.seek(SeekFrom::Start(0)).unwrap();
            stream.write_all(&info).unwrap();
            stream.flush().unwrap();
        }
        cursor.into_inner()
    }

    /// GH #12, the downgrade attack: an attacker who can modify the file deletes the
    /// `<dataIntegrity>` element and the tamper detection goes with it — no password,
    /// no error. The bytes here are exactly that attack, and the default policy must
    /// refuse them.
    ///
    /// This test asserted the opposite until #12 (`..._reports_not_declared`, where
    /// the loop below ran `VerifyIfPresent` as *the default* and expected `Ok`).
    #[test]
    fn test_agile_without_data_integrity_is_refused_by_default() {
        let data = agile_fixture_without_data_integrity();

        // The fix. `decrypt_ooxml` is the signature the consumer calls, so the fail-closed
        // behaviour must arrive without anyone passing a policy — asserted on the
        // specific variant, since `is_err()` would also pass if the rewrite had
        // simply broken the container.
        assert!(
            matches!(
                decrypt_ooxml(&data, "testpass"),
                Err(Error::IntegrityElementMissing)
            ),
            "the default policy must refuse an agile file whose tag was deleted"
        );
        assert!(matches!(
            decrypt_ooxml_with_policy(&data, "testpass", IntegrityPolicy::default()),
            Err(Error::IntegrityElementMissing)
        ));
        // `Require` is stricter still and refuses for the same reason, with the same
        // variant: the file is the problem, not the request.
        assert!(matches!(
            decrypt_ooxml_with_policy(&data, "testpass", IntegrityPolicy::Require),
            Err(Error::IntegrityElementMissing)
        ));

        // The negative control that makes the refusals mean something: the SAME bytes
        // decrypt under either explicit opt-out. Without this the test cannot tell
        // "the policy is wired" from "these bytes always fail", and the opt-out #12
        // promises could be unreachable while every assertion above still passed.
        for policy in [IntegrityPolicy::VerifyIfPresent, IntegrityPolicy::Skip] {
            let crate::Decrypted {
                package: plain,
                integrity: outcome,
            } = decrypt_ooxml_with_policy(&data, "testpass", policy).unwrap_or_else(|e| {
                panic!("{policy:?} must still decrypt a tag-less agile file: {e}")
            });
            // `Skip` reports `NotDeclared`, not `Skipped`: nothing was skipped.
            assert_eq!(outcome, IntegrityOutcome::NotDeclared, "{policy:?}");
            assert!(plain.starts_with(b"PK\x03\x04"), "{policy:?}");
        }

        // The second control: the same bytes with the tag still in place verify, so
        // the refusals come from the deleted element and not from the rewrite.
        let crate::Decrypted {
            integrity: outcome, ..
        } = decrypt_ooxml_with_policy(&agile_fixture(), "testpass", IntegrityPolicy::Require)
            .unwrap();
        assert_eq!(outcome, IntegrityOutcome::Verified);
    }

    /// The unmodified fixture verifies. This is the positive half of the guard: if the
    /// HMAC were computed over the wrong bytes, this fails rather than the tamper test.
    #[test]
    fn test_agile_fixture_integrity_verifies() {
        let crate::Decrypted {
            package: plain,
            integrity: outcome,
        } = decrypt_ooxml_with_policy(&agile_fixture(), "testpass", IntegrityPolicy::Require)
            .expect("the unmodified fixture must verify under Require");
        assert_eq!(outcome, IntegrityOutcome::Verified);
        assert!(plain.starts_with(b"PK\x03\x04"));
    }

    /// F18 itself: before this check existed, this input decrypted to garbage and
    /// returned `Ok`. The flipped byte sits ~20 KB in, so the ZIP magic still appears
    /// at the front of the plaintext — "it looks like a ZIP" is not a integrity check.
    #[test]
    fn test_agile_tampered_ciphertext_is_refused() {
        let tampered = tamper_agile_fixture(8 + 20_000);

        for policy in [
            IntegrityPolicy::default(),
            IntegrityPolicy::VerifyIfPresent,
            IntegrityPolicy::Require,
        ] {
            let result = decrypt_ooxml_with_policy(&tampered, "testpass", policy);
            assert!(
                matches!(result, Err(Error::IntegrityCheckFailed)),
                "{policy:?} must refuse a tampered package, got {:?}",
                result.map(|d| (d.package.len(), d.integrity))
            );
        }

        // The policy-free entry point the consumer calls refuses too. Note the variant: the
        // element is present and the HMAC is wrong, which is `IntegrityCheckFailed`,
        // not the `IntegrityElementMissing` of the deleted-element case.
        assert!(matches!(
            decrypt_ooxml(&tampered, "testpass"),
            Err(Error::IntegrityCheckFailed)
        ));
    }

    /// `Skip` still decrypts the same tampered file. This is what proves the policy is
    /// wired rather than the tamper being rejected by some unrelated check: the bytes
    /// are identical, only the policy differs.
    #[test]
    fn test_agile_tampered_ciphertext_decrypts_under_skip() {
        let tampered = tamper_agile_fixture(8 + 20_000);
        let crate::Decrypted {
            package: plain,
            integrity: outcome,
        } = decrypt_ooxml_with_policy(&tampered, "testpass", IntegrityPolicy::Skip)
            .expect("Skip must not check the HMAC");
        assert_eq!(outcome, IntegrityOutcome::Skipped);
        assert!(
            plain.starts_with(b"PK\x03\x04"),
            "the corruption is mid-package; the ZIP header still decrypts cleanly, \
             which is exactly why a structural sniff is not an integrity check"
        );

        // ... and it really is corrupt: the same offset in the clean fixture differs.
        let clean = decrypt_ooxml(&agile_fixture(), "testpass").unwrap();
        assert_ne!(clean, plain, "the flipped byte must change the plaintext");
    }

    /// A wrong password must still be reported as a wrong password, not as corruption.
    /// The password check runs first for exactly this reason.
    #[test]
    fn test_agile_wrong_password_is_not_reported_as_corruption() {
        let result =
            decrypt_ooxml_with_policy(&agile_fixture(), "wrongpass", IntegrityPolicy::Require);
        assert!(matches!(result, Err(Error::WrongPassword)));
    }

    /// ECMA-376 standard encryption has no integrity element by spec. Absence is
    /// reported, not treated as a failure.
    ///
    /// GH #12 item 3, and the single most likely way to break something while fixing
    /// the agile downgrade: the fail-closed default stops at formats that *define* an
    /// element, so `IntegrityPolicy::default()` and the bare `decrypt_ooxml` are in
    /// the loop below explicitly rather than left to be assumed equivalent to
    /// `VerifyIfPresent`.
    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn test_standard_reports_integrity_absent_not_failure() {
        let data = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/standard_encrypted.docx"
        ))
        .expect("fixture must be present -- standard tests are not optional");

        // Byte-identity against the committed plaintext, not a ZIP-shape check: see
        // `test_standard_fixture_decrypts_to_zip` for the provenance of that claim.
        let known = fixture("plain.docx");
        assert_eq!(
            decrypt_ooxml(&data, "testpass")
                .expect("the default must never refuse a standard file"),
            known,
            "the policy-free signature the consumer calls must still decrypt Office 2007 files"
        );

        for policy in [
            IntegrityPolicy::default(),
            IntegrityPolicy::VerifyIfPresent,
            IntegrityPolicy::Skip,
        ] {
            let crate::Decrypted {
                package: plain,
                integrity: outcome,
            } = decrypt_ooxml_with_policy(&data, "testpass", policy)
                .unwrap_or_else(|e| panic!("{policy:?} must decrypt a standard file: {e}"));
            assert_eq!(outcome, IntegrityOutcome::NotApplicable);
            assert_eq!(plain, known, "{policy:?}");
        }

        // `Require` is the caller demanding a guarantee the format cannot give, so it
        // is refused -- with a distinct error, never `IntegrityCheckFailed`.
        assert!(matches!(
            decrypt_ooxml_with_policy(&data, "testpass", IntegrityPolicy::Require),
            Err(Error::IntegrityUnavailable(_))
        ));
    }

    /// The no-regression half of GH #12: the policy-free signature still behaves
    /// exactly as it did for well-formed files, so an existing caller inherits the
    /// fail-closed default without a line changing on its side. A security fix that
    /// also broke every good file would not be one.
    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn test_default_policy_requires_a_tag_and_still_decrypts_good_files() {
        assert_eq!(
            IntegrityPolicy::default(),
            IntegrityPolicy::RequireWhereDefined
        );

        // Every committed agile fixture — all three carry the element — decrypts
        // unchanged through the bare signature, and reports `Verified` rather than
        // merely succeeding.
        for name in [
            "agile_encrypted.docx",
            "agile_aes256_sha384.docx",
            "agile_aes256_sha256.docx",
        ] {
            let data = fixture(name);
            let plain = decrypt_ooxml(&data, "testpass")
                .unwrap_or_else(|e| panic!("{name} must decrypt under the default: {e}"));
            let crate::Decrypted {
                package: with_policy,
                integrity: outcome,
            } = decrypt_ooxml_with_policy(&data, "testpass", IntegrityPolicy::Require).unwrap();
            assert_eq!(plain, with_policy, "{name}");
            assert_eq!(plain, fixture("plain.docx"), "{name}");
            assert_eq!(outcome, IntegrityOutcome::Verified, "{name}");
        }
    }

    // ---- the encrypt shape guard -----------------------------------------------

    /// The bug this guard exists for: encrypting a file that is already encrypted
    /// used to return `Ok` and produce a CFB wrapped in a CFB, indistinguishable
    /// from a single wrap without decrypting it. Found by a downstream consumer,
    /// which had to reimplement the CLI's guard to avoid it.
    ///
    /// The first `encrypt_ooxml` succeeding is the negative control: this test
    /// cannot pass by refusing everything.
    #[test]
    fn encrypt_ooxml_refuses_a_file_it_just_encrypted() {
        let plain = fixture("plain.docx");
        let sealed = encrypt_ooxml(&plain, "testpass").expect("a plain package encrypts");

        let again = encrypt_ooxml(&sealed, "testpass");
        assert!(
            matches!(
                again,
                Err(Error::AlreadyEncrypted {
                    family: Family::Agile,
                    document: Document::OoxmlPackage,
                })
            ),
            "a second wrap must be refused, naming what it found; got: {:?}",
            again.map(|c| c.len())
        );
    }

    /// The standard writer's half of the same bug.
    #[test]
    fn encrypt_ooxml_standard_refuses_a_file_it_just_encrypted() {
        let plain = fixture("plain.docx");
        let sealed = encrypt_ooxml_standard(&plain, "testpass").expect("a plain package encrypts");

        let again = encrypt_ooxml_standard(&sealed, "testpass");
        assert!(
            matches!(
                again,
                Err(Error::AlreadyEncrypted {
                    family: Family::Standard,
                    document: Document::OoxmlPackage,
                })
            ),
            "a second wrap must be refused, naming what it found; got: {:?}",
            again.map(|c| c.len())
        );
    }

    /// Cross-format double-wrapping is the same defect wearing a different hat, and
    /// nothing else here would catch it: each writer must refuse the other's output.
    #[test]
    fn the_two_writers_refuse_each_others_output() {
        let plain = fixture("plain.docx");
        let agile = encrypt_ooxml(&plain, "testpass").unwrap();
        let standard = encrypt_ooxml_standard(&plain, "testpass").unwrap();

        assert!(
            matches!(
                encrypt_ooxml_standard(&agile, "testpass"),
                Err(Error::AlreadyEncrypted {
                    family: Family::Agile,
                    ..
                })
            ),
            "the standard writer must refuse an agile container"
        );
        assert!(
            matches!(
                encrypt_ooxml(&standard, "testpass"),
                Err(Error::AlreadyEncrypted {
                    family: Family::Standard,
                    ..
                })
            ),
            "the agile writer must refuse a standard container"
        );
    }

    /// Eight bytes of CFB magic: a container this crate recognises and will not write
    /// into. Fixture-free on purpose — the fact is about the shape, not any document.
    #[test]
    fn a_bare_cfb_is_refused_as_not_a_plain_package() {
        let magic = [0xD0u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        assert!(matches!(
            check_encryptable(&magic),
            Err(Error::NotAPlainPackage)
        ));
        assert!(matches!(
            encrypt_ooxml(&magic, "testpass"),
            Err(Error::NotAPlainPackage)
        ));
        assert!(matches!(
            encrypt_ooxml_standard(&magic, "testpass"),
            Err(Error::NotAPlainPackage)
        ));
    }

    /// Bytes that are no container at all, including the empty input. `&[]` is
    /// `Container::Unknown`, so it is refused like any other non-package — the
    /// writer's handling of a degenerate *payload* is a separate fact, checked
    /// through the seeded cores in `agile_encrypt_tests` and `standard_encrypt_tests`.
    #[test]
    fn bytes_that_are_no_container_are_refused() {
        for input in [&b"sixteen bytes!!!"[..], &[][..]] {
            assert!(
                matches!(check_encryptable(input), Err(Error::UnknownContainer)),
                "{input:?} is not a container"
            );
            assert!(matches!(
                encrypt_ooxml(input, "testpass"),
                Err(Error::UnknownContainer)
            ));
            assert!(matches!(
                encrypt_ooxml_standard(input, "testpass"),
                Err(Error::UnknownContainer)
            ));
        }
    }

    /// A plain ZIP is the one thing the guard accepts, and four bytes of magic are
    /// the whole test — which is why a bare signature passes alongside a real
    /// package. Moved here from the CLI when the guard stopped being the CLI's.
    #[test]
    fn a_plain_zip_is_the_one_thing_the_guard_accepts() {
        check_encryptable(b"PK\x03\x04").expect("four bytes of zip magic are encryptable");
        check_encryptable(&fixture("plain.docx")).expect("a real package is encryptable");
    }

    /// The guard a caller runs before paying for a password is the guard the writer
    /// runs at the door. Not two functions that agree — one function, called twice —
    /// and this is what would fail if a future entry point grew its own copy.
    ///
    /// [`Error`] has no `PartialEq` by design, so the comparison is over a name.
    #[test]
    fn the_guard_a_caller_runs_is_the_guard_the_writer_runs() {
        fn kind(e: &Error) -> &'static str {
            match e {
                Error::AlreadyEncrypted { .. } => "already-encrypted",
                Error::NotAPlainPackage => "not-a-plain-package",
                Error::UnknownContainer => "unknown-container",
                _ => "other",
            }
        }
        fn verdict(r: Result<Vec<u8>, Error>) -> String {
            r.map_or_else(|e| kind(&e).to_string(), |_| "ok".to_string())
        }

        let plain = fixture("plain.docx");
        let sealed = encrypt_ooxml(&plain, "testpass").unwrap();
        let magic = [0xD0u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

        for input in [
            &plain[..],
            &sealed[..],
            &magic[..],
            &b"sixteen bytes!!!"[..],
            &[][..],
        ] {
            let asked = check_encryptable(input)
                .map_or_else(|e| kind(&e).to_string(), |()| "ok".to_string());
            assert_eq!(
                asked,
                verdict(encrypt_ooxml(input, "testpass")),
                "the agile writer disagreed with the guard a caller would have run"
            );
            assert_eq!(
                asked,
                verdict(encrypt_ooxml_standard(input, "testpass")),
                "the standard writer disagreed with the guard a caller would have run"
            );
        }
    }
}
