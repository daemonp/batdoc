//! The `\x06DataSpaces` subtree and the CFB container that carries it.
//!
//! Writing an encrypted OOXML document means writing six streams into a CFB container:
//! the two this crate's reader cares about — `/EncryptionInfo` and `/EncryptedPackage` —
//! and four fixed blobs under `\x06DataSpaces` that name the transform applied to the
//! payload. [MS-OFFCRYPTO] §2.1 defines all four; together they are 452 bytes and they
//! are identical in every encrypted document Office writes.
//!
//! # The four blobs are generated, not copied
//!
//! herumi's `include/resource.hpp` holds the same 452 bytes as four escaped string
//! literals. That file is BSD-3 and portable with attribution, so copying it would be
//! *allowed* — it is simply worse. Building the blobs from their field definitions keeps
//! this crate free of anyone's expression, and it makes every constant a reader can check
//! against the spec rather than a byte string they must take on faith.
//!
//! The claim that the derivation is right is not an assertion, it is a test:
//! `the_generated_dataspaces_blobs_match_a_real_documents_byte_for_byte` diffs all four
//! against the same streams lifted out of `tests/fixtures/agile_encrypted.docx`. Two of
//! the four lengths (112 and 200) are reached through a back-patched length field, so a
//! misread field shifts nothing and would survive a length check — only a byte-for-byte
//! diff catches every case at once.
//!
//! Section numbers and field names follow `ms-offcrypto-writer` (MIT/Apache)
//! `src/lib.rs:228-373`, which reads them from [MS-OFFCRYPTO] §2.1.2 and §2.1.4-2.1.9.
//!
//! # The directory tree belongs to `cfb`
//!
//! herumi hand-writes the directory sector and therefore hardcodes a colour, a left and
//! right sibling and a child for all eleven entries (`include/make_dataspace.hpp:77-87`).
//! GH #5 existed to find out whether that table is a property of the *format* — in which
//! case `cfb` would be unusable here, since it exposes no colour or sibling API at all —
//! or a property of *herumi's tree-builder*.
//!
//! It is the latter. `cfb` builds a demonstrably different tree (every entry Black
//! against the fixture's RED root, different sibling and child pointers, zero timestamps
//! against herumi's `now()`) and real Word 16, `office-crypto`, `msoffcrypto-tool` and
//! this crate all open the result. So this module writes six streams and lets `cfb` own
//! everything else — design **D2**, one container layer rather than three.
//!
//! # Gating
//!
//! `cfg(feature = "crypto-ops")`, since GH #6 step 6 made `encrypt_ooxml` its production
//! caller. It was `cfg(all(test, crypto-ops))` for the four steps before that, on the
//! rule `odf-crypto` `manifest.rs:594` states: `cfg` says "no caller yet" where
//! `allow(dead_code)` would only silence it. A detection build classifies files and
//! writes none, so it never wants this module.

use crate::error::Error;
use std::io::{Cursor, Write};

// ---------------------------------------------------------------------------
// Stream paths
// ---------------------------------------------------------------------------

/// `\x06DataSpaces/Version` — [`data_space_version_info`].
pub(crate) const VERSION: &str = "/\u{6}DataSpaces/Version";
/// `\x06DataSpaces/DataSpaceMap` — [`data_space_map`].
pub(crate) const DATA_SPACE_MAP: &str = "/\u{6}DataSpaces/DataSpaceMap";
/// The data space definition, named by [`data_space_map`]'s single entry.
pub(crate) const STRONG_ENCRYPTION_DATA_SPACE: &str =
    "/\u{6}DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace";
/// The transform description, named by [`data_space_definition`].
pub(crate) const PRIMARY: &str =
    "/\u{6}DataSpaces/TransformInfo/StrongEncryptionTransform/\u{6}Primary";
/// The version header plus the agile XML or the standard binary header.
pub(crate) const ENCRYPTION_INFO: &str = "/EncryptionInfo";
/// The 8-byte little-endian plaintext size followed by the ciphertext.
pub(crate) const ENCRYPTED_PACKAGE: &str = "/EncryptedPackage";

/// The two storages that must exist before the four `\x06DataSpaces` streams can be
/// created. `create_storage_all` makes every missing level of one path in a single call,
/// so `/\x06DataSpaces` and `/\x06DataSpaces/TransformInfo` are created implicitly.
pub(crate) const STORAGE_PATHS: [&str; 2] = [
    "/\u{6}DataSpaces/DataSpaceInfo",
    "/\u{6}DataSpaces/TransformInfo/StrongEncryptionTransform",
];

/// Every directory entry that carries a timestamp: the root and all four storages —
/// streams cannot, by [MS-CFB] §2.6.1. All are set to the CFB zero, see
/// [`build_container`].
const TIMESTAMPED_ENTRIES: [&str; 5] = [
    "/",
    "/\u{6}DataSpaces",
    "/\u{6}DataSpaces/DataSpaceInfo",
    "/\u{6}DataSpaces/TransformInfo",
    "/\u{6}DataSpaces/TransformInfo/StrongEncryptionTransform",
];

/// The CFB "uninitialized" timestamp — a FILETIME of zero, which is 1601-01-01 and sits
/// 11 644 473 600 seconds before the Unix epoch. `cfb` reads a zero field back as this
/// `SystemTime` and writes this `SystemTime` back as a zero field.
///
/// The return type is [`web_time::SystemTime`] because that is what `cfb` 0.14 accepts.
/// On native targets `web_time` re-exports `std::time`, so this is the same value the
/// upstream function returned. On `wasm32-unknown-unknown` the two `SystemTime` types
/// differ and a `std::time` value does not compile (batdoc fork; see `BATDOC-FORK.md`).
pub(crate) fn cfb_zero_time() -> web_time::SystemTime {
    // 1601-01-01 is before the Unix epoch. `std::time::SystemTime` can represent it;
    // `web_time`'s wasm `SystemTime` cannot (`checked_sub` returns `None`). Decrypt
    // never calls this — only the encrypt container builder does — so a wasm build
    // still links. Calling encrypt on wasm hits the `expect` until `cfb` can store a
    // pre-epoch time without going through that type.
    web_time::UNIX_EPOCH
        .checked_sub(std::time::Duration::from_secs(11_644_473_600))
        .expect("CFB epoch (1601-01-01) is representable on this platform")
}

// ---------------------------------------------------------------------------
// The four fixed blobs, generated from their field definitions
// ---------------------------------------------------------------------------

/// Append a `UNICODE-LP-P4` ([MS-OFFCRYPTO] §2.1.2): a `u32` byte count, the UTF-16LE
/// code units, then zero padding up to the next multiple of four.
///
/// The length is a **byte** count of the string data and excludes any terminator — there
/// is none. The trailing zeros that look like one are the padding: three of the seven
/// strings in these blobs need two bytes of it.
///
/// The `u32` conversion cannot fail here and is not a hostile-input path: every caller in
/// this module passes a `&'static str` literal, and no file-derived string ever reaches
/// it. Keeping the function private is what holds that true.
fn write_lp_p4(buffer: &mut Vec<u8>, text: &str) {
    let units: Vec<u16> = text.encode_utf16().collect();
    let byte_len = units.len() * 2;
    buffer.extend_from_slice(&u32::try_from(byte_len).expect("literal fits").to_le_bytes());
    for unit in units {
        buffer.extend_from_slice(&unit.to_le_bytes());
    }
    // Pad to a 4-byte boundary.
    buffer.resize(buffer.len() + (4 - (byte_len % 4)) % 4, 0);
}

/// Append a `Version` ([MS-OFFCRYPTO] §2.1.4): two `u16`s, major then minor.
///
/// Two `u16`s, not one `u32`. They serialize identically while the minor is zero, which
/// is what makes a `u32` reading of these fields look correct.
fn write_version(buffer: &mut Vec<u8>, major: u16, minor: u16) {
    buffer.extend_from_slice(&major.to_le_bytes());
    buffer.extend_from_slice(&minor.to_le_bytes());
}

/// `DataSpaceVersionInfo` ([MS-OFFCRYPTO] §2.1.5) — the `Version` stream. 76 bytes.
pub(crate) fn data_space_version_info() -> Vec<u8> {
    let mut buffer = Vec::new();
    write_lp_p4(&mut buffer, "Microsoft.Container.DataSpaces");
    write_version(&mut buffer, 1, 0); // ReaderVersion
    write_version(&mut buffer, 1, 0); // UpdaterVersion
    write_version(&mut buffer, 1, 0); // WriterVersion
    buffer
}

/// `DataSpaceMap` ([MS-OFFCRYPTO] §2.1.6) — the `DataSpaceMap` stream. 112 bytes.
///
/// One entry, mapping the stream named `EncryptedPackage` to the data space named
/// `StrongEncryptionDataSpace`. `Length` counts itself: 4 + 4 + 4 + 36 + 56 = 104, which
/// is the whole entry and not just what follows the field.
pub(crate) fn data_space_map() -> Vec<u8> {
    let mut entry = Vec::new();
    entry.extend_from_slice(&1u32.to_le_bytes()); // ReferenceComponentCount
    entry.extend_from_slice(&0u32.to_le_bytes()); // ReferenceComponentType: 0 = stream
    write_lp_p4(&mut entry, "EncryptedPackage"); // ReferenceComponent
    write_lp_p4(&mut entry, "StrongEncryptionDataSpace"); // DataSpaceName

    let mut buffer = Vec::new();
    buffer.extend_from_slice(&8u32.to_le_bytes()); // HeaderLength
    buffer.extend_from_slice(&1u32.to_le_bytes()); // EntryCount
    let entry_len = u32::try_from(entry.len() + 4).expect("literal fits");
    buffer.extend_from_slice(&entry_len.to_le_bytes()); // MapEntries[0].Length
    buffer.extend_from_slice(&entry);
    buffer
}

/// `DataSpaceDefinition` ([MS-OFFCRYPTO] §2.1.7) — the `StrongEncryptionDataSpace`
/// stream. 64 bytes.
///
/// The stream's *name* is what `DataSpaceMap` points at; its *content* names the
/// transform storage that holds `\x06Primary`.
pub(crate) fn data_space_definition() -> Vec<u8> {
    let mut buffer = Vec::new();
    buffer.extend_from_slice(&8u32.to_le_bytes()); // HeaderLength
    buffer.extend_from_slice(&1u32.to_le_bytes()); // TransformReferenceCount
    write_lp_p4(&mut buffer, "StrongEncryptionTransform"); // TransformReferences[0]
    buffer
}

/// `TransformInfoHeader` (§2.1.8) followed by `EncryptionTransformInfo` (§2.1.9) — the
/// `\x06Primary` stream. 200 bytes.
///
/// `TransformLength` covers only itself, `TransformType` and `TransformID`
/// (4 + 4 + 80 = 88) — not `TransformName` and not the three versions after it.
///
/// The GUID is the ECMA-376 encryption transform's identifier, stored as an ASCII-in-
/// UTF-16 brace string rather than as a 16-byte binary CLSID. Every value above is
/// [MS-OFFCRYPTO] §2.3.4.3's — `TransformType 0x00000001`, that `TransformID`, the
/// `Microsoft.Container.EncryptionTransform` name, and `1.0` for all three versions.
///
/// # `EncryptionTransformInfo` is written as Office writes it, not as §2.1.9 states it
///
/// §2.1.9 says `EncryptionBlockSize` "MUST be 0x00000010 as specified by the Advanced
/// Encryption Standard (AES)" and that `EncryptionName` "MUST be the name of an
/// encryption algorithm, such as 'AES 128'". Office writes **zero** for the block size
/// and an empty `UTF-8-LP-P4` for the name — in both fixtures, agile and standard alike:
/// the `\x06Primary` streams of `agile_encrypted.docx` and `standard_encrypted.docx` are
/// byte-identical 200-byte blobs ending `00000000 00000000 00000000 04000000` (measured
/// 2026-09-05). §2.3.4.3 requires the null name for agile only, so the standard file's
/// empty name is Office diverging from §2.1.9 as well.
///
/// It costs nothing to follow the file rather than the prose: §2.3.4.3 makes
/// `EncryptionInfo` authoritative over this structure outright ("if the algorithms
/// specified in the EncryptionTransformInfo structure differ from the algorithms
/// specified in the EncryptionInfo stream ... the EncryptionInfo stream MUST be
/// considered authoritative"), so no reader may act on these four words. Matching Office
/// keeps one blob for both formats and keeps the byte-identity test meaningful.
/// `ms-offcrypto-writer` reaches the same values and records the field as "at best
/// underspecified" (`src/lib.rs:363-366`).
pub(crate) fn transform_info() -> Vec<u8> {
    let mut header = Vec::new();
    header.extend_from_slice(&1u32.to_le_bytes()); // TransformType
    write_lp_p4(&mut header, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"); // TransformID

    let mut buffer = Vec::new();
    let transform_len = u32::try_from(header.len() + 4).expect("literal fits");
    buffer.extend_from_slice(&transform_len.to_le_bytes()); // TransformLength
    buffer.extend_from_slice(&header);

    write_lp_p4(&mut buffer, "Microsoft.Container.EncryptionTransform"); // TransformName
    write_version(&mut buffer, 1, 0); // ReaderVersion
    write_version(&mut buffer, 1, 0); // UpdaterVersion
    write_version(&mut buffer, 1, 0); // WriterVersion

    buffer.extend_from_slice(&0u32.to_le_bytes()); // EncryptionName: empty LP-P4
    buffer.extend_from_slice(&0u32.to_le_bytes()); // EncryptionBlockSize
    buffer.extend_from_slice(&0u32.to_le_bytes()); // CipherMode
    buffer.extend_from_slice(&4u32.to_le_bytes()); // Reserved
    buffer
}

// ---------------------------------------------------------------------------
// The container
// ---------------------------------------------------------------------------

/// Build a complete encrypted-OOXML container around two already-encrypted streams.
///
/// The cryptography is the caller's: this writes whatever `encryption_info` and
/// `encrypted_package` bytes it is handed, surrounded by the four constant blobs and the
/// storage skeleton. Splitting it that way is what let GH #5 answer the container
/// question with a fixture's real streams before any encrypt code existed.
///
/// Two choices here are deliberate and are the only ones a caller could get wrong:
///
/// * **`Version::V3`**, 512-byte sectors. `CompoundFile::create` and
///   `OpenOptions::create_with` both default to V4; every encrypted document in the corpus
///   is V3, herumi hardcodes it (`include/cfb.hpp:125-126`), and `ms-offcrypto-writer`
///   passes it explicitly. The same tree under V4 is three times the size.
/// * **`create_new_stream`**, not `create_stream`. Each stream is written exactly once
///   into an empty container, so an existing entry is a bug; `create_stream` would
///   silently truncate and return it.
///
/// Every `Stream` is dropped before `into_inner()`. `Stream` holds only a `Weak` to the
/// allocator, so consuming the container first lets `Arc::try_unwrap` succeed while the
/// stream is still alive — and that stream's later `Drop` then swallows an
/// `io::Error("CompoundFile was dropped")` and loses its buffered bytes with no
/// diagnostic. Scoping each write is the whole fix.
///
/// **Every timestamp in the directory is zero.** `cfb` stamps the root entry and each
/// storage with the clock as it creates them (`lib.rs:1417`, `directory.rs:275`), which
/// made two builds of the same streams differ in exactly those bytes and nowhere else —
/// the golden that pins the whole write path under a seeded RNG found it. Zero is what
/// [MS-CFB] §2.6.1 calls the uninitialized value; it is what Office itself writes for the
/// root's creation time (measured on `word16_agile.docx`, where `cfb`'s `now()` is the one
/// place it diverged from Office); and a file that does not record when it was made is
/// one fewer thing the ciphertext's owner did not choose to publish. Streams are already
/// zero by the spec's requirement, and `cfb` refuses to set them otherwise.
///
/// # Errors
///
/// [`Error::Io`] if the in-memory container cannot be written. Every path and
/// blob below is a compile-time constant, so there is no input-shaped failure here — but
/// the writes still go through `io::Write` and this crate does not `unwrap` on a code
/// path a caller can reach.
pub(crate) fn build_container(
    encryption_info: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, Error> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut container = cfb::CompoundFile::create_with_version(cfb::Version::V3, &mut cursor)?;

        for path in STORAGE_PATHS {
            container.create_storage_all(path)?;
        }

        for (path, bytes) in [
            (VERSION, data_space_version_info()),
            (DATA_SPACE_MAP, data_space_map()),
            (STRONG_ENCRYPTION_DATA_SPACE, data_space_definition()),
            (PRIMARY, transform_info()),
            (ENCRYPTION_INFO, encryption_info.to_vec()),
            (ENCRYPTED_PACKAGE, encrypted_package.to_vec()),
        ] {
            let mut stream = container.create_new_stream(path)?;
            stream.write_all(&bytes)?;
            stream.flush()?;
        }

        // After every write: creating a stream re-stamps its parent storages' modified
        // time, so this has to be the last thing that touches the directory.
        for path in TIMESTAMPED_ENTRIES {
            container.set_created_time(path, cfb_zero_time())?;
            container.set_modified_time(path, cfb_zero_time())?;
        }

        container.flush()?;
    }
    Ok(cursor.into_inner())
}

#[cfg(test)]
#[path = "dataspaces_tests.rs"]
mod tests;
