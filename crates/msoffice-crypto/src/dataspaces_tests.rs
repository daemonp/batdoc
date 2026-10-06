//! Does a container the `cfb` crate wrote pass for an encrypted OOXML document?
//!
//! Plan slice **S4** / GH #5, promoted out of `tests/cfb_dataspaces_container.rs` into the
//! crate in GH #6 step 1. **Why it moved:** the builder and the four blob generators used
//! to live in the test file, so the byte-identity proof below covered a copy that the
//! encrypt path would not have used. `build_container` is `pub(crate)`, which an
//! integration test cannot reach, so the tests came to the code rather than the code being
//! published to reach the tests — the same reason `classify_tests.rs` and
//! `malformed_input.rs` live in `src/`.
//!
//! # The experiment
//!
//! Take `tests/fixtures/agile_encrypted.docx`, lift its `EncryptionInfo` and
//! `EncryptedPackage` streams out **verbatim**, and rebuild the whole container from
//! nothing with `cfb` 0.14 — including the six-stream `\x06DataSpaces` subtree. The
//! cryptography is then byte-identical to a file that already works, so the only variable
//! left is who wrote the container.
//!
//! # What these prove, and what they do not
//!
//! Automated here: the four generated blobs are byte-identical to the ones in a real
//! encrypted document; the rebuilt container carries the same ten directory entries; this
//! crate decrypts it to the same plaintext as the original; `office-crypto` — an
//! independent MIT implementation — agrees; and a one-bit change to the payload inside the
//! rebuilt container is still caught, so the round-trip is not passing by accident.
//!
//! **Not** automated: real Word. Driving Word through COM needs Office installed and an
//! interactive desktop session, which no CI has, so the acceptance result is recorded
//! rather than re-run — see [`the_rebuilt_container_is_written_where_word_can_be_pointed_at_it`],
//! which writes the artifact it was measured on, and `CHANGELOG.md` § *S4 (GH #5)*.
//!
//! Also not proven, by anything here: Excel or PowerPoint against a `cfb`-built container,
//! any Word but Office 16 desktop, any tuple but agile AES-256/SHA-512, a re-*save*
//! through Word rather than an open, or standard (2007) encryption.

use super::*;
use std::io::Read;

// ---------------------------------------------------------------------------
// Fixture access
// ---------------------------------------------------------------------------

/// The agile fixture, whose password is `testpass`.
///
/// A hard failure, never a skip: a test that silently passes when its fixture is missing
/// tests nothing, which is a bug this crate has already shipped once.
fn agile_fixture() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/agile_encrypted.docx"
    ))
    .expect("fixture must be present -- container tests are not optional")
}

/// Read one stream out of a CFB container held in memory.
fn stream_of(data: &[u8], path: &str) -> Vec<u8> {
    let mut container =
        cfb::CompoundFile::open(Cursor::new(data)).expect("input is a CFB container");
    let mut bytes = Vec::new();
    container
        .open_stream(path)
        .unwrap_or_else(|e| panic!("container has no {path}: {e}"))
        .read_to_end(&mut bytes)
        .expect("in-memory read");
    bytes
}

/// The fixture's two encrypted streams, lifted out verbatim.
fn encrypted_streams() -> (Vec<u8>, Vec<u8>) {
    let fixture = agile_fixture();
    (
        stream_of(&fixture, ENCRYPTION_INFO),
        stream_of(&fixture, ENCRYPTED_PACKAGE),
    )
}

/// The fixture's streams rebuilt into a container `cfb` wrote from scratch.
fn rebuilt_container() -> Vec<u8> {
    let (info, package) = encrypted_streams();
    build_container(&info, &package).expect("an in-memory container is always writable")
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

/// The 452 bytes are right, and they were derived rather than copied.
///
/// This is what licenses the "generated from field definitions" claim in the module doc.
/// Every constant there — the LP-P4 padding rule, `Length` = 104 counting itself,
/// `TransformLength` = 88 stopping after `TransformID`, `Reserved` = 4, the two-`u16`
/// versions — is a place a plausible misreading produces a plausible-looking blob, and a
/// byte-for-byte diff against a real encrypted document is the only thing that catches
/// all of them at once. Two of the four sizes (112 and 200) are reached by a back-patched
/// length field, so a wrong one shifts nothing and would go unnoticed by a length check.
#[test]
fn the_generated_dataspaces_blobs_match_a_real_documents_byte_for_byte() {
    let fixture = agile_fixture();

    for (path, generated, expected_len) in [
        (VERSION, data_space_version_info(), 76),
        (DATA_SPACE_MAP, data_space_map(), 112),
        (
            STRONG_ENCRYPTION_DATA_SPACE,
            data_space_definition(),
            64_usize,
        ),
        (PRIMARY, transform_info(), 200),
    ] {
        assert_eq!(
            generated.len(),
            expected_len,
            "{path} is the wrong length -- an LP-P4 pad or a header field is off"
        );
        assert_eq!(
            generated,
            stream_of(&fixture, path),
            "{path} differs from the same stream in agile_encrypted.docx"
        );
    }

    let total = data_space_version_info().len()
        + data_space_map().len()
        + data_space_definition().len()
        + transform_info().len();
    assert_eq!(total, 452, "the whole DataSpaces payload is 452 bytes");
}

/// The rebuilt container holds the ten entries an encrypted document has, and `cfb`
/// wrote the container header the corpus has.
///
/// Six streams and four storages, none of which this crate's own reader would notice
/// were missing — `cfb_reader` opens two streams and ignores the rest, which is right for
/// reading and silently insufficient for writing. Only a structural check sees the other
/// four, so this is the test that would fail if the encrypt path quietly dropped a stream.
#[test]
fn the_rebuilt_container_holds_all_ten_entries_a_real_document_has() {
    let rebuilt = rebuilt_container();
    let container = cfb::CompoundFile::open(Cursor::new(&rebuilt))
        .expect("the rebuilt file is a CFB container");

    assert_eq!(
        container.version(),
        cfb::Version::V3,
        "Office writes CFB v3 (512-byte sectors); cfb's constructors default to V4"
    );

    for path in STORAGE_PATHS {
        assert!(container.is_storage(path), "{path} must be a storage");
    }
    // Every level, including the ones `create_storage_all` filled in on the way down.
    assert!(container.is_storage("/\u{6}DataSpaces"));
    assert!(container.is_storage("/\u{6}DataSpaces/TransformInfo"));

    let fixture = agile_fixture();
    for path in [
        VERSION,
        DATA_SPACE_MAP,
        STRONG_ENCRYPTION_DATA_SPACE,
        PRIMARY,
        ENCRYPTION_INFO,
        ENCRYPTED_PACKAGE,
    ] {
        assert!(container.is_stream(path), "{path} must be a stream");
        assert_eq!(
            stream_of(&rebuilt, path),
            stream_of(&fixture, path),
            "{path} must survive the rebuild byte for byte"
        );
    }

    let entries = container.walk().count();
    assert_eq!(
        entries, 11,
        "root + 4 storages + 6 streams; a stray entry means a path was mistyped"
    );
}

/// `cfb` owns the directory tree, so herumi's hardcoded colours are ours to ignore.
///
/// **This is the question GH #5 exists to answer.** herumi lists a colour, a left and a
/// right sibling and a child for all eleven entries (`make_dataspace.hpp:77-87`) because
/// it hand-writes the directory sector and has to. `cfb` has no such API: `Color`,
/// `DirEntry` and the sibling pointers are all private to the crate, and the only
/// directory type it re-exports — `cfb::Entry` — exposes name, kind, length, CLSID, state
/// bits and timestamps and nothing else.
///
/// The claim the compiler cannot check is the one this test is evidence for: the tree
/// `cfb` built is demonstrably *not* herumi's, and every reader tried — this crate,
/// `office-crypto`, `msoffcrypto-tool` and real Word 16 — accepts it anyway. Dumping the
/// raw directory sectors of both containers shows the fixture reproducing herumi's eleven
/// rows down to a **RED root**, which MS-CFB §2.6.4 forbids and `cfb` tolerates only on
/// read, against an all-black rebuilt tree with different sibling and child pointers.
#[test]
fn the_directory_tree_is_cfbs_to_build_and_the_colours_are_not_our_problem() {
    let rebuilt = rebuilt_container();
    let container = cfb::CompoundFile::open(Cursor::new(&rebuilt))
        .expect("the rebuilt file is a CFB container");

    // The two containers hold the same eleven entries, and `walk` cannot tell them
    // apart: it is an *in-order* traversal of a BST keyed on the name, so its output is
    // the names in MS-CFB sort order and is a fact about the contents rather than about
    // the tree. That is itself the finding restated -- the shape is invisible from
    // outside, which is exactly why nothing downstream can depend on it.
    //
    // (An earlier draft asserted the two orders *differed*, on the assumption that `walk`
    // was pre-order. It is not, and the assertion failed on containers that do differ.)
    let walk_of = |data: &[u8]| -> Vec<String> {
        cfb::CompoundFile::open(Cursor::new(data.to_vec()))
            .expect("a CFB container")
            .walk()
            .map(|e| e.name().to_string())
            .collect()
    };
    let ours = walk_of(&rebuilt);
    assert_eq!(ours.len(), 11);
    assert_eq!(ours[0], "Root Entry");
    assert_eq!(
        ours,
        walk_of(&agile_fixture()),
        "the same eleven entries in the same sort order, from two different trees"
    );

    // And the trees really are different, which the byte-level dump in `CHANGELOG.md`
    // § S4 spells out (a RED root and herumi's exact colour table on one side, all-black
    // on the other). The observable half of that difference is here: identical entries,
    // identical stream bytes -- asserted by the test above -- and files that are not the
    // same bytes.
    assert_ne!(
        rebuilt,
        agile_fixture(),
        "if the rebuild ever reproduced the source byte for byte, every claim in this \
         file about writing a *different* container would be vacuous"
    );

    // What `cfb` will tell us about an entry -- deliberately, no colour among it.
    let version = container
        .entry(VERSION)
        .expect("the rebuilt container has a Version stream");
    assert!(version.is_stream());
    assert_eq!(version.len(), 76);
    // MS-CFB 2.6.1: a stream's creation and modified times must be zero. `cfb` does that
    // for us; herumi stamps every entry with the current time, which is why this crate's
    // own fixtures fail `open_strict` below and the rebuilt file does not.
    //
    // A zero CFB timestamp surfaces as the FILETIME epoch, 1601-01-01, which is
    // 11 644 473 600 seconds before the Unix one -- not as `UNIX_EPOCH`.
    let cfb_epoch = std::time::UNIX_EPOCH - std::time::Duration::from_secs(11_644_473_600);
    assert_eq!(version.created(), cfb_epoch);
    assert_eq!(version.modified(), cfb_epoch);

    // The strongest available statement about the tree without reaching into `cfb`'s
    // internals: its own validator, which checks name ordering unconditionally and
    // red/black adjacency under `strict`, accepts what it built. `agile_encrypted.docx`
    // does not -- so the rebuilt container is *more* conformant than its own source.
    cfb::CompoundFile::open_strict(Cursor::new(&rebuilt))
        .expect("cfb's own strict validator accepts the tree it built");
    assert!(
        cfb::CompoundFile::open_strict(Cursor::new(agile_fixture())).is_err(),
        "the control: the fixture's hand-written directory fails strict validation, so \
         the assertion above is about the rebuilt tree and not about strict being lenient"
    );
}

/// Word 16 opens this file. Recorded, not re-run.
///
/// `cargo test` cannot drive Word — it needs Office installed and an interactive desktop
/// session — so the acceptance result lives in the report and in `CHANGELOG.md`, and this
/// test writes the artifact that produced it. Reproduce with:
///
/// ```powershell
/// $w = New-Object -ComObject Word.Application
/// $w.Visible = $false; $w.DisplayAlerts = 0
/// try {
///   $doc = $w.Documents.Open($path, $false, $true, $false, 'testpass')
///   $doc.Content.Text; $doc.Close(0)
/// } finally { $w.Quit() }
/// ```
///
/// The destination is [`std::env::temp_dir`], not `CARGO_TARGET_TMPDIR`: cargo defines
/// that variable for integration tests and benches only, and this became a unit test when
/// the builder moved into `src/`. The path is printed, so `cargo test -- --nocapture`
/// still says where to point Word.
///
/// Word's verdicts are distinguishable and worth reading by HRESULT rather than by its
/// localised message, which contains the full file path: `0x800A17C8` = no usable
/// `EncryptionInfo`, `0x800A141F` = the stream or its XML was rejected, `0x800A1520` =
/// everything parsed and the password verifier missed, no throw = opened. The container
/// question is settled the moment Word reaches `0x800A1520`, because getting that far
/// means it accepted the container.
///
/// **Measured 2026-09-04, Word 16 desktop on Windows 11**, on the exact file this test
/// writes:
///
/// | file | password | verdict |
/// | --- | --- | --- |
/// | the unmodified fixture (control) | `testpass` | OPENED |
/// | this artifact | `testpass` | OPENED, identical `$doc.Content.Text` |
/// | same builder, `Reserved` 0x40 → 0x00 | `testpass` | REFUSED `0x800A141F` |
/// | this artifact | `WRONGPASS` | REFUSED `0x800A1520` |
///
/// The two refusals are what make the two opens mean something. The first says Word
/// inspects the file rather than opening whatever it is handed. The second is the
/// stronger one: reaching "the password is incorrect" means Word walked the container,
/// found `EncryptionInfo`, accepted its header and XML and ran the KDF — every step that
/// depends on the container being well-formed — before failing on the one thing that was
/// meant to fail.
///
/// Word also opened a variant carrying no `\x06DataSpaces` subtree at all. Write the
/// subtree anyway: every writer in the corpus does, and it is 452 bytes of constants.
///
/// Always `$w.Quit()` in a `finally` — a leaked WINWORD.EXE blocks the next run — and
/// keep `DisplayAlerts = 0` with `Close(0)` so no modal dialog can appear.
#[test]
fn the_rebuilt_container_is_written_where_word_can_be_pointed_at_it() {
    let path = std::env::temp_dir().join("cfb_rebuilt_agile.docx");
    std::fs::write(&path, rebuilt_container()).expect("the temp dir is writable");
    println!("rebuilt container written to {}", path.display());
}

/// Round-trip: this crate reads what `cfb` wrote, and gets the original plaintext.
///
/// The plaintext is compared against the plaintext of the *unmodified fixture* rather
/// than against a ZIP magic check, so a container that decrypted to something plausible
/// but wrong would still fail.
#[test]
fn a_cfb_built_container_decrypts_through_this_crate() {
    let expected = crate::decrypt_ooxml(&agile_fixture(), "testpass")
        .expect("the unmodified fixture decrypts -- otherwise nothing below means anything");

    let rebuilt = crate::decrypt_ooxml(&rebuilt_container(), "testpass")
        .expect("a container cfb built must decrypt exactly like the one it was copied from");

    assert_eq!(
        rebuilt, expected,
        "the rebuilt container must yield the same package, byte for byte"
    );
    assert!(rebuilt.starts_with(b"PK\x03\x04"));
}

/// The negative control for the test above: the rebuild carries the payload, it does not
/// merely produce a file that happens to decrypt.
///
/// One bit of `EncryptedPackage` ciphertext is flipped 4 KiB into the stream, and the
/// container is rebuilt around the damaged stream. Without this, the round-trip test
/// could not distinguish "the payload made it through the new container" from "any
/// container with these stream names decrypts".
///
/// Where that byte lands, precisely, because it is easy to get wrong: `package` is the
/// **whole** stream including its 8-byte little-endian plaintext-size prefix, and
/// `agile.rs` strips that prefix before segmenting (`&encrypted_package[8..]`, then
/// `.chunks(4096)`). Stream index 4096 is therefore ciphertext offset 4088 — AES block
/// 255, the *last* full block **of segment 0**, not the first block of segment 1.
///
/// Which segment it lands in is immaterial to the assertion, and that is the point: the
/// package HMAC covers the whole stream and `check_integrity` runs *before*
/// `decrypt_package` (`agile.rs`, step 5 before step 6), so the flip is caught with no
/// plaintext produced at all. There is no decrypted ZIP header on this path for a parse
/// error to trip over. Tampering in the ciphertext body rather than the header or
/// container is what CLAUDE.md § *Testing Rules* asks for, and offset 4096 satisfies it
/// anywhere in the body; the arithmetic above is recorded so a later tamper test on the
/// encrypt path — where an integrity failure may be raised *after* per-segment
/// decryption — can pick its offset from the real layout instead of from this one.
#[test]
fn a_cfb_built_container_around_a_tampered_payload_is_still_caught() {
    let (info, mut package) = encrypted_streams();
    package[4096] ^= 0x01;
    let container = build_container(&info, &package).expect("in-memory container");

    // `.map(..)` before the assertion: on a *failing* run the `Ok` arm is the whole
    // decrypted document, and `{result:?}` would dump 36 KB of it into the test output
    // and bury the assertion that failed.
    let result = crate::decrypt_ooxml(&container, "testpass").map(|p| p.len());
    assert!(
        matches!(result, Err(Error::IntegrityCheckFailed)),
        "a flipped ciphertext bit must fail the package HMAC, got: {result:?}"
    );
}

/// The differential oracle: an independent implementation reads it too.
///
/// `office-crypto` is a separate MIT crate with its own CFB reader and its own parse of
/// [MS-OFFCRYPTO]. Agreeing with it is evidence about the file; agreeing only with
/// `decrypt_ooxml` would be evidence about this crate's own two-stream reader, which
/// never opens the `\x06DataSpaces` subtree at all and so cannot notice whether it is
/// well-formed.
#[test]
fn an_independent_implementation_reads_the_cfb_built_container_too() {
    let expected = office_crypto::decrypt_from_bytes(agile_fixture(), "testpass")
        .expect("office-crypto reads the unmodified fixture");
    let rebuilt = office_crypto::decrypt_from_bytes(rebuilt_container(), "testpass")
        .expect("office-crypto must read a container cfb built");

    assert_eq!(rebuilt, expected);
}

/// Classification is unchanged by the rebuild.
///
/// `classify` is the detection build's entire view of a file, and it reads the container
/// header and `/EncryptionInfo` — both of which the rebuild rewrites. This is the test
/// that would catch a rebuilt container that still decrypts but no longer *describes*
/// itself the same way.
///
/// It no longer runs in the detection build. It used to: the builder lived in the test
/// file, so it compiled with no cipher crate in the graph. Now that the builder is
/// `crypto-ops` production code there is nothing to exercise in a build that cannot
/// write a container, and gating the module rather than sprinkling `allow(dead_code)` is
/// the rule `limits.rs` already follows. The assertion is unchanged in value — `classify`
/// is the same code in both configurations.
#[test]
fn the_rebuilt_container_classifies_identically_to_its_source() {
    let before = crate::classify(&agile_fixture());
    let after = crate::classify(&rebuilt_container());
    assert_eq!(
        format!("{before:?}"),
        format!("{after:?}"),
        "rebuilding the container must not change what the file says it is"
    );
}

/// The container is byte-deterministic: the same streams give the same bytes, twice.
///
/// `cfb` stamps the root entry and each storage with the clock as it creates them, and
/// the seeded golden over the whole write path (`agile_encrypt_tests.rs`) is what noticed
/// — same seed, same 41 984 bytes, different SHA-256 across two processes, with the
/// directory timestamps the only difference. `build_container` zeroes every one of them
/// as its last act, and this pins that: every timestamped entry reads back as the CFB
/// zero, and two builds are identical. The stream assertion is `cfb`'s own doing and was
/// already true; the root and storages are ours.
#[test]
fn the_container_carries_no_timestamps_and_is_byte_deterministic() {
    let (info, package) = encrypted_streams();
    let a = build_container(&info, &package).unwrap();
    let b = build_container(&info, &package).unwrap();
    assert_eq!(
        a, b,
        "two builds of the same streams must be the same bytes"
    );

    let container = cfb::CompoundFile::open(Cursor::new(&a)).unwrap();
    for path in TIMESTAMPED_ENTRIES {
        let entry = container
            .entry(path)
            .unwrap_or_else(|e| panic!("{path}: {e}"));
        assert_eq!(entry.created(), cfb_zero_time(), "{path}: created");
        assert_eq!(entry.modified(), cfb_zero_time(), "{path}: modified");
    }
    let stream = container.entry(ENCRYPTED_PACKAGE).unwrap();
    assert_eq!(
        stream.created(),
        cfb_zero_time(),
        "streams are zero by the spec"
    );
}
