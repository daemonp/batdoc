//! Every number a file declares about itself must produce an error — never a panic,
//! never a hang.
//!
//! This crate's whole input is a document somebody else wrote. `spinCount`, `keyBits`,
//! `EncryptionHeader.KeySize` and the `EncryptionVerifier` length are all chosen by that
//! file, all reached **before** the password is checked, and each one used to be a loop
//! count or a slice index with nothing between it and the arithmetic. A panic escapes
//! the `Result` contract entirely — a caller writes `decrypt_ooxml(..).map_err(..)`, and
//! `map_err` never sees an unwind — so "returns the wrong error" and "aborts the caller's
//! process" are not the same failure.
//!
//! The containers here are built at runtime rather than committed as fixtures: a
//! deliberately malformed `.docx` in the tree is indistinguishable from a corrupt one,
//! and nothing in the file would record *which* field was poisoned or why. Each test
//! carries a control differing only in the field under test, so a green result cannot
//! come from the synthetic container being rejected for some unrelated reason.

use crate::{decrypt_ooxml, Error};
use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use std::io::Write;

/// `AlgID` for AES-128 — [MS-OFFCRYPTO] §2.3.2's `CALG_AES_128`, which is what a
/// conforming Office 2007 file carries. This constant used to be `0x00006801` here and in
/// `standard.rs`, which is RC4's identifier; the shipped fixture declares the
/// spec-forbidden `fAES` + `0x6801` pair, so the wrong value went unnoticed.
const ALG_ID_AES_128: u32 = 0x0000_660E;

/// `AlgID` for RC4 — the CryptoAPI cipher this crate does not implement.
const ALG_ID_RC4: u32 = 0x0000_6801;

/// `EncryptionHeader.Flags` as a conforming ECMA-376 standard writer emits it:
/// `fCryptoAPI | fAES` ([MS-OFFCRYPTO] §2.3.1, §2.3.4.5).
const FLAGS_AES: u32 = 0x0000_0024;

/// The same header without `fAES` — an RC4 CryptoAPI file, which is written with the
/// same `vMinor = 2` and is therefore not distinguishable by the version pair.
const FLAGS_RC4: u32 = 0x0000_0004;

/// Wrap the two streams `decrypt_ooxml` reads in a minimal CFB container.
fn build_cfb(encryption_info: &[u8], encrypted_package: &[u8]) -> Vec<u8> {
    let mut cursor = std::io::Cursor::new(Vec::new());
    {
        let mut container = cfb::CompoundFile::create(&mut cursor).unwrap();
        for (path, bytes) in [
            ("/EncryptionInfo", encryption_info),
            ("/EncryptedPackage", encrypted_package),
        ] {
            let mut stream = container.create_stream(path).unwrap();
            stream.write_all(bytes).unwrap();
            stream.flush().unwrap();
        }
        container.flush().unwrap();
    }
    cursor.into_inner()
}

/// An `EncryptedPackage` stream: the 8-byte little-endian plaintext size, then one
/// AES block of ciphertext. Never actually decrypted by these tests — every one of
/// them is refused before the package is touched.
fn encrypted_package() -> Vec<u8> {
    let mut v = 16u64.to_le_bytes().to_vec();
    v.extend_from_slice(&[0u8; 16]);
    v
}

/// An agile `EncryptionInfo` stream: the 8-byte version header (`vMajor=4 vMinor=4`)
/// followed by the XML, with `spinCount` and `keyBits` under the caller's control.
///
/// Every base64 blob is zeros of the length the parser expects, so the file is
/// well-formed in every respect except the field being tested.
fn agile_encryption_info(spin_count: u32, key_bits: u32) -> Vec<u8> {
    agile_encryption_info_with_hash(spin_count, key_bits, "SHA512", 64)
}

/// The same stream with `p:encryptedKey/@hashAlgorithm` and `@hashSize` under the
/// caller's control too.
///
/// The hash is a field a file declares about itself exactly like `spinCount` and
/// `keyBits`, and until issue #11 it was the one such field this crate ignored — so the
/// two are varied by the same builder. `encryptedVerifierHashValue`'s length follows the
/// declared `hashSize`, the way a writer's would: `roundUp(hashSize, blockSize)`.
///
/// `<keyData>` keeps SHA-512 throughout. That is deliberate: this crate reads each
/// element's own hash whether or not they match — §2.3.4.10 tells a writer to match them
/// and says nothing to a reader — and holding one of them fixed is what makes a failure
/// attributable to the element under test.
fn agile_encryption_info_with_hash(
    spin_count: u32,
    key_bits: u32,
    hash_name: &str,
    hash_size: usize,
) -> Vec<u8> {
    agile_encryption_info_full(
        spin_count,
        key_bits,
        KEY_DATA_KEY_BITS,
        hash_name,
        hash_size,
        "ChainingModeCBC",
        AGILE_RESERVED,
    )
}

/// `keyData/@keyBits` as every real file carries it, and the value `encryptedKeyValue`'s
/// 32 zero bytes agree with: AES-256 means a 32-byte session key.
const KEY_DATA_KEY_BITS: u32 = 256;

/// `EncryptionInfo.Reserved` for agile encryption, as [MS-OFFCRYPTO] §2.3.4.10 requires
/// it — `0x40`, not zero, which is the one thing about this field that is easy to get
/// backwards.
const AGILE_RESERVED: u32 = 0x0000_0040;

/// The same stream with `cipherChaining`, `keyData/@keyBits` and the header's `Reserved`
/// word under the caller's control too.
///
/// All three are things the *file* declares about itself in exactly the way `spinCount`
/// and `p:encryptedKey/@keyBits` are, and all three went unread for the same kind of
/// reason: `cipherChaining` because its `cipherAlgorithm` sibling was checked and it was
/// not, `keyData/@keyBits` because its `<p:encryptedKey>` namesake was, and `Reserved`
/// because bytes 4..8 were skipped on the way to the XML.
///
/// `key_bits` is `<p:encryptedKey>`'s and `key_data_key_bits` is `<keyData>`'s. They are
/// two declarations about two different keys — the one that wraps the session key, and
/// the session key itself — and a file is free to disagree with itself about them.
fn agile_encryption_info_full(
    spin_count: u32,
    key_bits: u32,
    key_data_key_bits: u32,
    hash_name: &str,
    hash_size: usize,
    chaining: &str,
    reserved: u32,
) -> Vec<u8> {
    let zeros = |n: usize| BASE64.encode(vec![0u8; n]);
    let salt = zeros(16);
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="{key_data_key_bits}" hashSize="64" cipherAlgorithm="AES"
           cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{salt}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin_count}" saltSize="16" blockSize="16"
                      keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES"
                      cipherChaining="{chaining}" hashAlgorithm="{hash_name}"
                      saltValue="{salt}" encryptedVerifierHashInput="{vi}"
                      encryptedVerifierHashValue="{vh}" encryptedKeyValue="{kv}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        vi = zeros(16),
        vh = zeros(hash_size.div_ceil(16) * 16),
        kv = zeros(32),
    );
    let mut out = vec![0x04, 0x00, 0x04, 0x00];
    out.extend_from_slice(&reserved.to_le_bytes());
    out.extend_from_slice(xml.as_bytes());
    out
}

/// A standard-encryption `EncryptionInfo` stream: the 8-byte version header
/// (`vMajor=4 vMinor=2`), a **34**-byte `EncryptionHeader` — 32 fixed fields plus an
/// immediately-terminated UTF-16LE `CSPName` — then an `EncryptionVerifier` of the
/// caller's chosen length.
///
/// The verifier length is a file-chosen quantity in exactly this way: the stream simply
/// ends after `EncryptionHeaderSize` says the header does, and the reader has to decide
/// what is left is enough.
fn standard_encryption_info(key_size_bits: u32, verifier_len: usize) -> Vec<u8> {
    standard_encryption_info_full(FLAGS_AES, ALG_ID_AES_128, key_size_bits, verifier_len)
}

/// The same stream with `EncryptionHeader.Flags` and `AlgID` under the caller's control.
///
/// Those two fields together name the cipher, and only one of them used to be read —
/// against RC4's identifier at that. Both are things the *file* declares in exactly the
/// way `KeySize` is.
fn standard_encryption_info_full(
    flags: u32,
    alg_id: u32,
    key_size_bits: u32,
    verifier_len: usize,
) -> Vec<u8> {
    let mut out = vec![0x04, 0x00, 0x02, 0x00];
    // EncryptionHeader.Flags — *not* agile's Reserved word, though it occupies the same
    // four bytes. `standard_encrypted.docx` carries zeros here, which is why the agile
    // Reserved check is gated to the (4, 4) arm.
    out.extend_from_slice(&0u32.to_le_bytes());
    // EncryptionHeaderSize = **34**: the 32 fixed fields plus the two bytes of CSPName
    // below, which is the whole header. It said 32 while the reader located the verifier
    // by scanning CSPName for its terminator — the two disagreed and only the scan was
    // consulted, so the lie cost nothing. Now the field is the offset the verifier is
    // sliced at ([MS-OFFCRYPTO] §2.3.4.5), and 32 would put `SaltSize` on the NUL.
    out.extend_from_slice(&34u32.to_le_bytes()); // EncryptionHeaderSize
    for field in [
        flags,         // Flags     <- names the cipher with AlgID
        0,             // SizeExtra
        alg_id,        // AlgID     <- names the cipher with Flags
        0x0000_8004,   // AlgIDHash (SHA-1)
        key_size_bits, // KeySize
        0x0000_0018,   // ProviderType (AES)
        0,             // Reserved1
        0,             // Reserved2
    ] {
        out.extend_from_slice(&field.to_le_bytes());
    }
    out.extend_from_slice(&[0x00, 0x00]); // CSPName: an immediate UTF-16LE NUL

    let mut verifier = vec![0u8; verifier_len];
    if verifier_len >= 4 {
        verifier[..4].copy_from_slice(&16u32.to_le_bytes()); // SaltSize = 16
    }
    if verifier_len >= 40 {
        // VerifierHashSize = 20 — [MS-OFFCRYPTO] §2.3.4.9 step 3 fixes it, and
        // `standard::parse_encryption_info` refuses anything else. These containers exist
        // to exercise one hostile field each, so every *other* field is the conforming
        // value; a zero here would stop the parse before the field under test.
        verifier[36..40].copy_from_slice(&20u32.to_le_bytes());
    }
    out.extend_from_slice(&verifier);
    out
}

// ---- agile: p:encryptedKey/@keyBits -------------------------------------------------

/// `keyBits / 8` truncates a fixed 64-byte SHA-512 digest inside `derive_block_key`,
/// so anything above 512 sliced out of range — reached from `verify_password`, i.e.
/// before any password was checked. `aes_cbc_decrypt`'s key-length guard does not
/// cover it: the slice panics one call frame earlier.
#[test]
fn agile_key_bits_out_of_range_is_an_error_not_a_panic() {
    // Above 512 first: those are the ones that used to slice off the end of the digest.
    for key_bits in [520u32, 4096, u32::MAX, 0, 1, 8, 264] {
        let data = build_cfb(&agile_encryption_info(0, key_bits), &encrypted_package());
        let err = decrypt_ooxml(&data, "irrelevant")
            .expect_err("keyBits outside the ECMA-376 set must be refused");
        assert!(
            matches!(err, Error::BadParameters(_)) && err.to_string().contains("keyBits"),
            "keyBits={key_bits} must be refused by name, got: {err}"
        );
    }
}

/// The control that keeps the previous test honest: the same container with a legal
/// `keyBits` gets all the way to the verifier comparison and is refused there, so the
/// rejections above are attributable to the field and not to the synthetic container.
#[test]
fn agile_key_bits_256_passes_the_bound() {
    let data = build_cfb(&agile_encryption_info(0, 256), &encrypted_package());
    assert!(
        matches!(
            decrypt_ooxml(&data, "irrelevant"),
            Err(Error::WrongPassword)
        ),
        "a legal keyBits must reach the verifier comparison"
    );
}

/// 128 and 192 reach the verifier comparison, like 256 — the cipher dispatches on the
/// key length the file declares (GH #13). Until then this test asserted the opposite: that
/// both were refused one frame after the parser with "AES-256 needs a 32-byte key". A
/// `WrongPassword` on zero blobs is the same evidence the 256 case gives above: the block
/// keys were derived at 16 and 24 bytes, AES-128 and AES-192 ran on them, and the
/// comparison — not a length check — is what said no.
#[test]
fn agile_key_bits_128_and_192_reach_the_verifier_comparison() {
    for key_bits in [128u32, 192] {
        let data = build_cfb(&agile_encryption_info(0, key_bits), &encrypted_package());
        assert!(
            matches!(
                decrypt_ooxml(&data, "irrelevant"),
                Err(Error::WrongPassword)
            ),
            "keyBits={key_bits} must reach the verifier comparison"
        );
    }
}

// ---- agile: keyData/@keyBits ---------------------------------------------------------

/// `<keyData>` and `<p:encryptedKey>` each declare a `keyBits`, about two different keys,
/// and only the second was ever read. A file naming 128 or 192 on `<keyData>` alongside
/// 256 on `<p:encryptedKey>` therefore unwrapped a full 32-byte session key and handed all
/// of it to AES-256, where the writer had used `keyData.keyBits / 8` of it (herumi
/// resizes to exactly that, `include/decode.hpp:131-133`).
///
/// With no `<dataIntegrity>` element — a shape the default policy accepted until GH #12
/// made it fail closed — nothing downstream could notice: the password verified, the key
/// was the right *length* for AES-256, and the caller received rubbish under `Ok`. The
/// refusal below is by name, at parse time, and does not depend on that: it fires long
/// before any integrity policy is consulted.
#[test]
fn agile_key_data_key_bits_disagreeing_with_the_session_key_blob_is_refused_by_name() {
    // The blob is 32 bytes. keyData 128 wants 16 and is refused; keyData 192 wants 24
    // padded to 32 — the AES-192 wire format Word writes and opens (GH #13) — so it is
    // *not* a disagreement, gets past this check, and is refused by the verifier instead.
    let data = build_cfb(
        &agile_encryption_info_full(0, 256, 128, "SHA512", 64, "ChainingModeCBC", AGILE_RESERVED),
        &encrypted_package(),
    );
    let err = decrypt_ooxml(&data, "irrelevant")
        .expect_err("a keyData/@keyBits the session key cannot satisfy must be refused");
    assert!(
        matches!(&err, Error::BadParameters(msg)
            if msg.contains("keyData/@keyBits")),
        "keyData keyBits=128 must be refused by name, got: {err:?}"
    );
    let data = build_cfb(
        &agile_encryption_info_full(0, 256, 192, "SHA512", 64, "ChainingModeCBC", AGILE_RESERVED),
        &encrypted_package(),
    );
    assert!(
        matches!(
            decrypt_ooxml(&data, "irrelevant"),
            Err(Error::WrongPassword)
        ),
        "keyData keyBits=192 over a 32-byte blob is the padded AES-192 shape and must reach the verifier"
    );

    // Out of the ECMA-376 set entirely, on this element too.
    for key_data_key_bits in [0u32, 255, 512, u32::MAX] {
        let data = build_cfb(
            &agile_encryption_info_full(
                0,
                256,
                key_data_key_bits,
                "SHA512",
                64,
                "ChainingModeCBC",
                AGILE_RESERVED,
            ),
            &encrypted_package(),
        );
        let err = decrypt_ooxml(&data, "irrelevant").expect_err("out-of-set keyBits is refused");
        assert!(
            matches!(&err, Error::BadParameters(msg)
                if msg.contains("keyData/@keyBits")),
            "keyData keyBits={key_data_key_bits} got: {err:?}"
        );
    }
}

/// The control: the same container with the two elements agreeing reaches the verifier
/// comparison, so the refusals above come from `<keyData>`'s attribute and not from the
/// synthetic file. Without this, the test above cannot tell "the attribute is checked"
/// from "this container never decrypts".
#[test]
fn agile_key_data_key_bits_of_256_reaches_the_verifier_comparison() {
    let data = build_cfb(
        &agile_encryption_info_full(0, 256, 256, "SHA512", 64, "ChainingModeCBC", AGILE_RESERVED),
        &encrypted_package(),
    );
    assert!(matches!(
        decrypt_ooxml(&data, "irrelevant"),
        Err(Error::WrongPassword)
    ));
}

// ---- agile: p:encryptedKey/@hashAlgorithm -------------------------------------------

/// Issue #11, through the public entry point rather than the parser: a hash name this
/// crate does not implement must reach the caller as its own variant.
///
/// The failure it replaces was not an error at all. The attribute was never read, the
/// KDF ran SHA-512 regardless, the verifier comparison failed, and `decrypt_ooxml`
/// returned `WrongPassword` — the one answer that is actively misleading, because the
/// password may be exactly right and there is nothing the user can do about it.
#[test]
fn agile_unimplemented_hash_name_is_its_own_error_not_a_wrong_password() {
    for (name, size) in [("MD5", 16usize), ("SHA3-512", 64), ("BLAKE2b", 64)] {
        let data = build_cfb(
            &agile_encryption_info_with_hash(0, 256, name, size),
            &encrypted_package(),
        );
        let err = decrypt_ooxml(&data, "irrelevant")
            .expect_err("an unimplemented hashAlgorithm must be refused");
        assert!(
            matches!(&err, Error::UnsupportedAlgorithm { what, name: got }
                if *what == "p:encryptedKey/@hashAlgorithm" && got == name),
            "hashAlgorithm={name} got: {err:?}"
        );
    }
}

/// The negative control for the test above and the positive half of the hash work: the
/// three hashes that pair with AES-256 are accepted, run their own KDF, and are refused
/// by the *verifier* — `WrongPassword`, from a container whose blobs are all zeros.
///
/// Without this the previous test cannot distinguish "unknown names are rejected" from
/// "every non-SHA-512 name is rejected", which is what the crate did before #11 and what
/// `office-crypto` still does deliberately (`src/lib.rs:29`).
#[test]
fn agile_implemented_hash_names_reach_the_verifier_comparison() {
    for (name, size) in [("SHA256", 32usize), ("SHA384", 48), ("SHA512", 64)] {
        let data = build_cfb(
            &agile_encryption_info_with_hash(0, 256, name, size),
            &encrypted_package(),
        );
        assert!(
            matches!(
                decrypt_ooxml(&data, "irrelevant"),
                Err(Error::WrongPassword)
            ),
            "hashAlgorithm={name} must reach the verifier comparison"
        );
    }
}

/// `keyBits / 8` is a truncation length on the digest of the hash the file names, so
/// `AGILE_KEY_BITS_ALLOWED` stopped covering the slice the moment the digest length
/// became file-controlled: SHA-1 yields 20 bytes and `keyBits="256"` wants 32. This is a
/// crafted file, not one any writer produces — LibreOffice's four accepted tuples and
/// herumi's two emitted ones all avoid it — and it is refused by name rather than padded
/// (see `agile::derive_block_key` for why).
#[test]
fn agile_key_bits_longer_than_the_named_digest_is_refused_by_name() {
    for key_bits in [192u32, 256] {
        let data = build_cfb(
            &agile_encryption_info_with_hash(0, key_bits, "SHA1", 20),
            &encrypted_package(),
        );
        let err =
            decrypt_ooxml(&data, "irrelevant").expect_err("SHA-1 cannot fill a 24- or 32-byte key");
        let text = err.to_string();
        assert!(
            matches!(err, Error::BadParameters(_))
                && text.contains("keyBits")
                && text.contains("SHA1"),
            "keyBits={key_bits} got: {text}"
        );
    }

    // The control: the same SHA-1 file with a key it *can* fill -- 16 bytes from a
    // 20-byte digest -- gets past the pairing check, reaches AES-128 and the verifier
    // comparison, and is refused there, by the comparison. This is GH #13's warning made
    // a test: the guard is `keyBits / 8 > digest_len`, never `hash == SHA-1`, and a guard
    // written the second way would refuse Word 2010's default tuple here.
    let data = build_cfb(
        &agile_encryption_info_with_hash(0, 128, "SHA1", 20),
        &encrypted_package(),
    );
    assert!(
        matches!(
            decrypt_ooxml(&data, "irrelevant"),
            Err(Error::WrongPassword)
        ),
        "AES-128/SHA-1 must reach the verifier comparison"
    );
}

// ---- agile: cipherChaining ------------------------------------------------------------

/// The sibling of `cipherAlgorithm`, and the more consequential of the two to have left
/// unread. `ChainingModeCFB` is legal ECMA-376 that this crate does not implement, and
/// decrypting such a file as CBC is a **silent** wrong answer rather than a loud one:
/// the `dataIntegrity` HMAC is computed over ciphertext, so the file would verify as
/// `IntegrityOutcome::Verified` and the caller would receive rubbish.
#[test]
fn agile_cipher_chaining_other_than_cbc_is_its_own_error() {
    for chaining in ["ChainingModeCFB", "ChainingModeGCM", "ChainingModeECB"] {
        let data = build_cfb(
            &agile_encryption_info_full(
                0,
                256,
                KEY_DATA_KEY_BITS,
                "SHA512",
                64,
                chaining,
                AGILE_RESERVED,
            ),
            &encrypted_package(),
        );
        let err = decrypt_ooxml(&data, "irrelevant")
            .expect_err("an unimplemented cipherChaining must be refused");
        assert!(
            matches!(&err, Error::UnsupportedAlgorithm { what, name }
                if *what == "p:encryptedKey/@cipherChaining" && name == chaining),
            "cipherChaining={chaining} got: {err:?}"
        );
    }
}

/// The control: the same container declaring the chaining mode this crate *does*
/// implement reaches the verifier comparison, so the refusals above come from the
/// attribute and not from the synthetic file.
#[test]
fn agile_chaining_mode_cbc_reaches_the_verifier_comparison() {
    let data = build_cfb(
        &agile_encryption_info_full(
            0,
            256,
            KEY_DATA_KEY_BITS,
            "SHA512",
            64,
            "ChainingModeCBC",
            AGILE_RESERVED,
        ),
        &encrypted_package(),
    );
    assert!(matches!(
        decrypt_ooxml(&data, "irrelevant"),
        Err(Error::WrongPassword)
    ));
}

// ---- agile: EncryptionInfo.Reserved ---------------------------------------------------

/// [MS-OFFCRYPTO] §2.3.4.10 fixes bytes 4..8 of an agile `EncryptionInfo` stream at
/// `0x00000040`. Those four bytes used to be skipped unread on the way to the XML — the
/// same as msoffcrypto-tool, which seeks past them (`format/ooxml.py:64`). herumi checks
/// them, in this crate's own shape, on the two lines before its XML parse
/// (`include/crypto_util.hpp:322-323`); LibreOffice checks them too
/// (`AgileEngine.cxx:522-530`, behaviour only). Both do it *before* handing the rest of
/// the stream to a parser, which is the point: one comparison standing in front of
/// everything else.
///
/// `0` is the first case for a reason — it is what a reader would write if it assumed
/// "reserved" meant "must be zero", and it is what standard encryption legitimately
/// carries in the same four bytes.
#[test]
fn agile_reserved_word_must_be_0x40_not_zero() {
    for reserved in [0u32, 1, 0x0000_0004, 0x0000_0041, 0x4000_0000, u32::MAX] {
        let data = build_cfb(
            &agile_encryption_info_full(
                0,
                256,
                KEY_DATA_KEY_BITS,
                "SHA512",
                64,
                "ChainingModeCBC",
                reserved,
            ),
            &encrypted_package(),
        );
        let err =
            decrypt_ooxml(&data, "irrelevant").expect_err("a wrong Reserved word must be refused");
        assert!(
            matches!(&err, Error::BadParameters(msg) if msg.contains("Reserved")),
            "reserved={reserved:#010x} must be refused by name, got: {err:?}"
        );
    }
}

/// The control: `0x40` in the same four bytes, everything else identical, gets all the
/// way to the verifier comparison. Without it the test above cannot distinguish "the
/// Reserved word is checked" from "this container never decrypts".
#[test]
fn agile_reserved_word_of_0x40_is_accepted() {
    let data = build_cfb(
        &agile_encryption_info_full(
            0,
            256,
            KEY_DATA_KEY_BITS,
            "SHA512",
            64,
            "ChainingModeCBC",
            AGILE_RESERVED,
        ),
        &encrypted_package(),
    );
    assert!(matches!(
        decrypt_ooxml(&data, "irrelevant"),
        Err(Error::WrongPassword)
    ));
}

/// ... and the check is gated to agile. Standard encryption's bytes 4..8 are
/// `EncryptionHeader.Flags`, an unrelated field: `standard_encrypted.docx` carries
/// `00 00 00 00` there, so a Reserved check hoisted above the version dispatch would
/// reject every Office 2007 file in existence.
#[test]
fn the_reserved_check_does_not_apply_to_standard_encryption() {
    let data = build_cfb(&standard_encryption_info(128, 72), &encrypted_package());
    assert!(
        matches!(
            decrypt_ooxml(&data, "irrelevant"),
            Err(Error::WrongPassword)
        ),
        "standard encryption writes Flags, not 0x40, in the same four bytes"
    );
}

// ---- agile: p:encryptedKey/@spinCount -----------------------------------------------

/// `spin_hash` runs before the password check, the integrity check, and every other
/// gate, so an unbounded `spinCount` is a hang the caller cannot defend against:
/// `catch_unwind` never fires and a Rust worker thread cannot be interrupted.
///
/// The elapsed-time assertion is the load-bearing half. Without the bound this input
/// still ends in an error — after roughly fifty minutes of one core.
#[test]
fn agile_spin_count_above_the_ceiling_is_refused_before_the_spin_runs() {
    let data = build_cfb(&agile_encryption_info(u32::MAX, 256), &encrypted_package());

    let start = std::time::Instant::now();
    let err = decrypt_ooxml(&data, "irrelevant").expect_err("an unbounded spinCount is refused");
    let elapsed = start.elapsed();

    assert!(
        matches!(err, Error::BadParameters(_)) && err.to_string().contains("spinCount"),
        "the refusal must name spinCount, got: {err}"
    );
    assert!(
        elapsed < std::time::Duration::from_secs(5),
        "spinCount must be bounded before the spin loop, not after it (took {elapsed:?})"
    );
}

// ---- agile: what the error message is allowed to carry -------------------------------

/// The same stream with one attribute value replaced verbatim, escaping nothing.
///
/// Every other builder here poisons a *number* the file declares. This one poisons the
/// attribute *text*, which is a different attack surface: the value is unescaped by the
/// XML parser before this crate ever sees it, so a value can make the parser itself
/// produce a message.
fn agile_with_attr_value(attr: &str, value: &str) -> Vec<u8> {
    let info = agile_encryption_info(100_000, 256);
    let (header, xml_bytes) = info.split_at(8);
    let xml = std::str::from_utf8(xml_bytes).expect("the builder writes UTF-8");
    let needle = format!("{attr}=\"");
    let start = xml.find(&needle).expect("the builder emits this attribute") + needle.len();
    let end = start
        + xml[start..]
            .find('"')
            .expect("the attribute value is quoted");

    let mut out = header.to_vec();
    out.extend_from_slice(&xml.as_bytes()[..start]);
    out.extend_from_slice(value.as_bytes());
    out.extend_from_slice(&xml.as_bytes()[end..]);
    out
}

/// **The file's own text must never come back in the error message.**
///
/// `encryptedKeyValue` is unescaped by quick-xml before this crate reads it, and an
/// unrecognized entity makes quick-xml's `Display` quote the text between `&` and the
/// next `;` verbatim -- bounded only by `ENCRYPTION_INFO_READ_CAP`, 1 MiB. `agile.rs`
/// forwarded that `Display` into `Error::XmlParse` until `xml_error` classified it.
///
/// Both halves are load-bearing and neither implies the other: a message could be short
/// and still quote a truncated marker, or marker-free and still be a megabyte of the
/// parser's opinion. `Debug` is checked beside `Display` because both are public API and
/// the leak this guards against reached a UI through `Display` in the consumer that
/// reported it.
///
/// Proven by reverting `xml_error` to `Error::XmlParse(e.to_string())`: the assertion
/// fails with `the file's own text reached the message: EncryptionInfo XML parse error:
/// at 1..4609: unrecognized entity `CANARYcanaryCANARY...`.
#[test]
fn a_hostile_entity_in_an_attribute_value_never_reaches_the_error_message() {
    const MARKER: &str = "CANARYcanaryCANARY";
    let body = MARKER.repeat(256);
    let data = build_cfb(
        &agile_with_attr_value("encryptedKeyValue", &format!("&{body};")),
        &encrypted_package(),
    );

    let err = decrypt_ooxml(&data, "irrelevant").expect_err("an unrecognized entity is refused");
    let display = err.to_string();
    let debug = format!("{err:?}");

    assert!(matches!(err, Error::XmlParse(_)), "got: {err}");
    assert!(
        !display.contains(MARKER),
        "the file's own text reached the message: {display}"
    );
    assert!(
        !debug.contains(MARKER),
        "the file's own text reached Debug: {debug}"
    );
    assert!(
        display.len() < 200,
        "the message is unbounded ({} bytes): {display}",
        display.len()
    );
}

/// The negative control for the test above.
///
/// Without it, that test cannot tell "the message is clean" from "this container is
/// rejected before the XML is read at all", which would make it pass against a parser
/// that never ran. The same builder, the same attribute, a *well-formed* value: the file
/// must get past XML parsing and fail somewhere later.
#[test]
fn the_same_attribute_with_a_well_formed_value_gets_past_the_xml() {
    let data = build_cfb(
        &agile_with_attr_value("encryptedKeyValue", &BASE64.encode([0u8; 32])),
        &encrypted_package(),
    );

    let err = decrypt_ooxml(&data, "irrelevant").expect_err("the package is not decryptable");

    assert!(
        !matches!(err, Error::XmlParse(_)),
        "a well-formed value must parse; the poisoned test proves nothing otherwise: {err}"
    );
}

/// The base64 decoder is the other foreign `Display` on this path.
///
/// `DecodeError::InvalidByte` carries the offending character. The offset is kept because
/// it is a position; the character is dropped because it is content.
#[test]
fn a_bad_base64_symbol_is_reported_by_offset_and_not_by_value() {
    // `*` is outside the base64 alphabet and is not an XML metacharacter, so it reaches
    // the decoder rather than the parser.
    let mut value = BASE64.encode([0u8; 32]);
    value.replace_range(4..5, "*");
    let data = build_cfb(
        &agile_with_attr_value("encryptedKeyValue", &value),
        &encrypted_package(),
    );

    let err = decrypt_ooxml(&data, "irrelevant").expect_err("invalid base64 is refused");
    let msg = err.to_string();

    assert!(matches!(err, Error::XmlParse(_)), "got: {err}");
    assert!(
        msg.contains("offset 4"),
        "the offset locates the fault: {msg}"
    );
    assert!(
        !msg.contains('*'),
        "the offending symbol is content and must not be quoted: {msg}"
    );
}

// ---- standard: EncryptionVerifier length --------------------------------------------

/// The guard admitted 52 bytes; the last field ends at 72. Every verifier length in
/// 52..=71 passed the check and panicked on `&v[40..72]`, before key derivation, so no
/// password work stood between a crafted file and the crash.
#[test]
fn standard_short_encryption_verifier_is_an_error_not_a_panic() {
    for verifier_len in [0usize, 4, 51, 52, 53, 60, 71] {
        let data = build_cfb(
            &standard_encryption_info(128, verifier_len),
            &encrypted_package(),
        );
        let err = decrypt_ooxml(&data, "irrelevant")
            .expect_err("a truncated EncryptionVerifier must be refused");
        assert!(
            matches!(err, Error::MissingStream(_)),
            "verifier_len={verifier_len} must be a MissingStream error, got: {err}"
        );
    }
}

/// The control for both standard-encryption tests: a full 72-byte verifier and a legal
/// `KeySize` get all the way to the password comparison.
#[test]
fn standard_full_encryption_verifier_reaches_the_password_check() {
    let data = build_cfb(&standard_encryption_info(128, 72), &encrypted_package());
    assert!(
        matches!(
            decrypt_ooxml(&data, "irrelevant"),
            Err(Error::WrongPassword)
        ),
        "72 bytes is the full AES EncryptionVerifier and must be accepted"
    );
}

// ---- standard: EncryptionHeader Flags + AlgID ----------------------------------------

/// The cipher is named by `fAES` and `AlgID` together, and the gate used to read `AlgID`
/// alone — against `0x00006801`, which is **RC4**'s identifier ([MS-OFFCRYPTO] §2.3.2;
/// herumi `include/standard_encryption.hpp:56-63` accepts `0x660e/0x660f/0x6610` and has
/// `0x6801` commented out as `AlgoRC4`).
///
/// Both halves of that were wrong answers reachable from a real file. A genuine RC4
/// CryptoAPI header — same `vMinor = 2`, `fAES` clear — cleared the gate, derived an AES
/// key, failed the AES verifier and came back `WrongPassword`: the one answer that is
/// actively misleading, because the password may be exactly right and the cipher is
/// simply not implemented.
#[test]
fn standard_rc4_cryptoapi_is_named_not_reported_as_a_wrong_password() {
    for (flags, alg_id) in [(FLAGS_RC4, ALG_ID_RC4), (FLAGS_RC4, 0)] {
        let data = build_cfb(
            &standard_encryption_info_full(flags, alg_id, 128, 72),
            &encrypted_package(),
        );
        let err = decrypt_ooxml(&data, "irrelevant").expect_err("RC4 CryptoAPI is not implemented");
        assert!(
            matches!(&err, Error::UnsupportedAlgorithm { what, name }
                if *what == "EncryptionHeader/@AlgID" && name.contains("RC4")),
            "flags={flags:#x} algId={alg_id:#x} got: {err:?}"
        );
    }
}

/// The other half: a conforming Office 2007 AES-128 file (`AlgID = 0x660E`, the value
/// LibreOffice and Office write) was refused as "Office XP/2003 RC4 encryption" — a file
/// that is neither RC4 nor pre-2007. It now reaches the verifier comparison like any
/// other AES file, which is also the control for the RC4 test above: the two containers
/// differ only in `Flags` and `AlgID`.
#[test]
fn standard_conforming_aes128_alg_id_reaches_the_password_check() {
    for (flags, alg_id) in [
        (FLAGS_AES, ALG_ID_AES_128),
        (FLAGS_RC4, ALG_ID_AES_128),
        // The shipped fixture's own combination: the spec-forbidden `fAES` + RC4 AlgID
        // pair, which `classify` already resolves in favour of `fAES`.
        (FLAGS_AES, ALG_ID_RC4),
    ] {
        let data = build_cfb(
            &standard_encryption_info_full(flags, alg_id, 128, 72),
            &encrypted_package(),
        );
        assert!(
            matches!(
                decrypt_ooxml(&data, "irrelevant"),
                Err(Error::WrongPassword)
            ),
            "flags={flags:#x} algId={alg_id:#x} must reach the verifier comparison"
        );
    }
}

// ---- standard: EncryptionHeader.KeySize ----------------------------------------------

/// `KeySize / 8` truncated the fixed 40-byte XOR-ladder buffer (panic above 320 bits)
/// and the result was then handed to AES-128, which panics inside `GenericArray` for
/// anything but 16 bytes — two distinct crashes from one unvalidated field, both before
/// the verifier comparison. 256 is the one that shows they are distinct: it clears the
/// 40-byte slice and dies in the cipher.
///
/// **These containers all declare `AlgID = 0x0000660E`**, so 256 is still refused here
/// even though [MS-OFFCRYPTO] §2.3.4.5 permits the value: AES-128's identifier paired
/// with a 256-bit `KeySize` describes no cipher the format defines. The row that says
/// the reader now *opens* AES-256 is
/// [`standard_aes192_and_aes256_reach_the_password_check`] below, and the two together
/// are what distinguish "the value is out of range" from "the two fields disagree".
#[test]
fn standard_key_size_other_than_128_is_an_error_not_a_panic() {
    for key_size_bits in [0u32, 8, 64, 256, 320, 512, u32::MAX] {
        let data = build_cfb(
            &standard_encryption_info(key_size_bits, 72),
            &encrypted_package(),
        );
        let err = decrypt_ooxml(&data, "irrelevant").expect_err("AlgID 0x660E means a 128-bit key");
        assert!(
            matches!(err, Error::BadParameters(_)) && err.to_string().contains("KeySize"),
            "KeySize={key_size_bits} must be refused by name, got: {err}"
        );
    }
}

/// `AlgID` for AES-192 and AES-256 — [MS-OFFCRYPTO] §2.3.2, verbatim: "This value MUST
/// be 0x0000660E (AES-128), 0x0000660F (AES-192), or 0x00006610 (AES-256)."
const ALG_ID_AES_192: u32 = 0x0000_660F;
const ALG_ID_AES_256: u32 = 0x0000_6610;

/// A conforming Office 2007 document may declare AES-192 or AES-256, and until
/// 2026-09-20 this crate refused both **by name** — `standard::require_aes_128` returned
/// `UnsupportedAlgorithm` before the header had finished parsing.
///
/// That refusal was then cited as evidence that the format offered no choice of key
/// size, which is the decision restated as its own justification. Both directions were
/// unfaithful, and the read direction is the one that matters: a file the format defines,
/// that its owner already holds, that this crate would not open.
///
/// Reaching `WrongPassword` is the whole assertion — the password here is deliberately
/// wrong, and the point is *where* the file is refused. `UnsupportedAlgorithm` would mean
/// the cipher was never accepted; `BadParameters` would mean a length check refused it.
/// The end-to-end proof that the right password opens such a file, with a real verifier
/// and a real package, is `standard::tests::all_three_aes_key_lengths_decrypt_end_to_end`.
#[test]
fn standard_aes192_and_aes256_reach_the_password_check() {
    for (alg_id, key_size_bits) in [
        (ALG_ID_AES_128, 128u32),
        (ALG_ID_AES_192, 192),
        (ALG_ID_AES_256, 256),
    ] {
        let data = build_cfb(
            &standard_encryption_info_full(FLAGS_AES, alg_id, key_size_bits, 72),
            &encrypted_package(),
        );
        assert!(
            matches!(
                decrypt_ooxml(&data, "irrelevant"),
                Err(Error::WrongPassword)
            ),
            "AlgID {alg_id:#010x} with KeySize {key_size_bits} must reach the verifier \
             comparison, not be refused by name"
        );
    }
}

// ---- agile: keyData/@hashAlgorithm, independently of p:encryptedKey's -----------------

/// An agile `EncryptionInfo` whose two elements name **different** hashes, carrying real
/// blobs from a real key schedule and no `<dataIntegrity>` element.
///
/// The shape `agile_encryption_info_full` cannot build: it holds `<keyData>` at SHA-512
/// so that a failure is attributable to `<p:encryptedKey>`. This is the other direction —
/// `<p:encryptedKey>` stays SHA-512, and `<keyData>` is the variable — with the password
/// encryptor generated by `agile_encrypt::generate` under a seeded RNG so the file is
/// genuinely openable rather than well-formed-but-inert.
fn agile_encryption_info_mixed(
    key_data_hash: &str,
    key_data_hash_size: usize,
    spin_count: u32,
    m: &crate::agile_encrypt::AgileKeyMaterial,
) -> Vec<u8> {
    let b64 = |b: &[u8]| BASE64.encode(b);
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="{key_data_hash_size}" cipherAlgorithm="AES"
           cipherChaining="ChainingModeCBC" hashAlgorithm="{key_data_hash}" saltValue="{kds}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin_count}" saltSize="16" blockSize="16"
                      keyBits="256" hashSize="64" cipherAlgorithm="AES"
                      cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"
                      saltValue="{ps}" encryptedVerifierHashInput="{vi}"
                      encryptedVerifierHashValue="{vh}" encryptedKeyValue="{kv}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        kds = b64(&m.key_data_salt),
        ps = b64(&m.password_salt),
        vi = b64(&m.encrypted_verifier_hash_input),
        vh = b64(&m.encrypted_verifier_hash_value),
        kv = b64(&m.encrypted_key_value),
    );
    let mut out = vec![0x04, 0x00, 0x04, 0x00];
    out.extend_from_slice(&AGILE_RESERVED.to_le_bytes());
    out.extend_from_slice(xml.as_bytes());
    out
}

/// Encrypt a package the way a writer that chose `key_data_hash` for `<keyData>` would:
/// 4096-byte segments, each under `H(keyData.saltValue || LE32(i))[..16]` for that hash.
fn package_under(
    plaintext: &[u8],
    key_data_hash: crate::hash::HashAlgorithm,
    m: &crate::agile_encrypt::AgileKeyMaterial,
) -> Vec<u8> {
    // Since GH #6 step 6 this is the production writer, not a second copy of it.
    crate::agile_encrypt::encrypt_package(
        plaintext,
        &m.session_key,
        &m.key_data_salt,
        key_data_hash,
    )
    .expect("the tuple is in range")
}

/// A file whose `<keyData>` names SHA-1 while `<p:encryptedKey>` names SHA-512 decrypts —
/// under SHA-1 segment IVs, which is what its writer used.
///
/// The file is **non-conforming**: [MS-OFFCRYPTO] §2.3.4.10 says a
/// `PasswordKeyEncryptor`'s "hashing algorithm specified MUST be the same as the hashing
/// algorithm specified for the Encryption.keyData element". That is an obligation on
/// writers, and this crate is a reader; the two attributes still drive disjoint halves of
/// the algorithm — `<p:encryptedKey>`'s the password KDF, `<keyData>`'s every package
/// segment IV and the `dataIntegrity` HMAC — so the only way to be sure each half runs on
/// its own declaration is to build the file where they differ and watch both come out
/// right. No guard rejects the combination: refusing it would buy nothing (a
/// disagreeing file is one no writer produced, not one that threatens us) and would cost
/// this test, which is what pins that the halves are not crossed.
///
/// Until 2026-09-05 no test exercised `<keyData>` at anything but SHA-512 above the unit
/// level — `agile_encryption_info_with_hash` holds it there on purpose, and every fixture
/// in the corpus is SHA-512 on both elements.
///
/// `VerifyIfPresent` because the file carries no `<dataIntegrity>`; the default policy
/// refuses that outright, and the last assertion pins that it does, so the file's
/// taglessness is a fact and not an accident of the builder.
///
/// The control is the same package under a `<keyData>` that lies about its hash: it
/// decrypts to the wrong bytes, with nothing refusing it, which is precisely the silent
/// failure the per-element hash exists to prevent. This is not first coverage of the IV
/// derivation — `agile::tests::package_segment_ivs_follow_the_key_data_hash_algorithm`
/// pins it at the unit level and fails under the same mutation, alongside five others —
/// but it is the only statement of the property through the public entry point on a
/// complete file.
#[test]
fn agile_key_data_may_name_a_different_hash_from_the_password_encryptor() {
    use rand::SeedableRng as _;
    const PASSWORD: &str = "testpass";
    const SPIN: u32 = 1_000; // the KDF is not what this exercises
    let mut rng = chacha20::ChaCha12Rng::from_seed([0x5Au8; 32]);
    let m = crate::agile_encrypt::generate(
        PASSWORD,
        crate::EncryptParams {
            spin_count: SPIN,
            ..Default::default()
        },
        &mut rng,
    )
    .unwrap();

    // Three segments, so indices 0, 1 and 2 all take part -- a one-segment package
    // would only ever exercise LE32(0).
    let mut plaintext = b"PK\x03\x04".to_vec();
    plaintext.extend((0..9_000u32).map(|i| (i % 251) as u8));

    let honest = build_cfb(
        &agile_encryption_info_mixed("SHA1", 20, SPIN, &m),
        &package_under(&plaintext, crate::hash::HashAlgorithm::Sha1, &m),
    );
    let crate::Decrypted {
        package: plain,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(
        &honest,
        PASSWORD,
        crate::IntegrityPolicy::VerifyIfPresent,
    )
    .expect("a legal mixed-hash file must decrypt");
    assert_eq!(outcome, crate::IntegrityOutcome::NotDeclared);
    assert_eq!(
        plain, plaintext,
        "SHA-1 keyData: every segment IV must follow it"
    );

    // The control: the same SHA-1-encrypted package under a <keyData> claiming SHA-512.
    // It opens -- nothing in a tagless file can refuse it -- and it is wrong.
    let lying = build_cfb(
        &agile_encryption_info_mixed("SHA512", 64, SPIN, &m),
        &package_under(&plaintext, crate::hash::HashAlgorithm::Sha1, &m),
    );
    let crate::Decrypted { package: wrong, .. } =
        crate::decrypt_ooxml_with_policy(&lying, PASSWORD, crate::IntegrityPolicy::VerifyIfPresent)
            .expect("nothing in a tagless file refuses a wrong keyData hash");
    assert_ne!(
        wrong, plaintext,
        "the IV must actually depend on keyData's hash"
    );

    // And the file really is tagless: the default policy says so by name.
    assert!(matches!(
        decrypt_ooxml(&honest, PASSWORD),
        Err(Error::IntegrityElementMissing)
    ));
}
