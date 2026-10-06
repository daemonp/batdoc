//! ECMA-376 Standard Encryption (Office 2007): the read side, and the primitives the
//! write side in `standard_encrypt` shares with it.
//!
//! EncryptionInfo stream (after the 8-byte header = 4-byte version + a 4-byte copy of
//! `EncryptionHeader.Flags`):
//!   `EncryptionHeaderSize` (4) + binary EncryptionHeader (that many bytes) +
//!   EncryptionVerifier (72)
//!
//! Algorithm:
//!
//! ```text
//! 1. Parse binary EncryptionHeader → AlgID (AES-128, AES-192 or AES-256),
//!    AlgIDHash (SHA-1), KeySize (which must agree with the AlgID)
//! 2. Find EncryptionVerifier at the declared EncryptionHeaderSize
//! 3. Spin hash: H_0 = SHA1(salt + password_utf16le)
//!               H_i = SHA1(LE32(i) + H_{i-1})  for i in 0..50_000
//! 4. H_final = SHA1(H_n + LE32(0))   — the block-0 step
//!    X1 = SHA1(H_final XOR 0x36-pad64), X2 = SHA1(H_final XOR 0x5C-pad64)
//!    Final key = (X1 || X2)[..key_size]   — concatenated, not hashed again
//! 5. Verify password via AES-ECB on EncryptedVerifier
//! 6. Decrypt EncryptedPackage in 4096-byte ECB segments
//! ```
//!
//! **One key length is not one KDF.** Steps 3 and 4 do not branch on the key size at
//! all: [MS-OFFCRYPTO] §2.3.4.7 fixes `H` at SHA-1 and the iteration count at 50 000,
//! and the only key-size-dependent step in the whole derivation is step 6 of the
//! enumerated key-derivation method — "let keyDerived be equal to the first
//! `cbRequiredKeyLength` bytes of X3" — with step 1 capping `cbRequiredKeyLength` at 40,
//! inside which AES-256's 32 bytes sit. So AES-192 and AES-256 are the same forty bytes
//! of ladder output cut in a different place, and nothing else. Until 2026-09-20 they
//! were nevertheless refused **by name**.
//!
//! Since GH #7 the steps a writer inverts are seams rather than inline code:
//! [`parse_encryption_info`] (1-2), [`derive_standard_key`] (3-4) and
//! [`verify_password`] (5), with [`decrypt`] composing them in that order exactly as it
//! did before. `standard_encrypt` runs the same `derive_standard_key` and the encrypt
//! half of the same ECB helper, so the two directions cannot drift the way two
//! derivations written side by side can.
use crate::error::Error;
use crate::limits;
use crate::sensitive::{
    utf16le_password, DerivedKey, PadBlock, PasswordDigest, SpinState, VerifierPlaintext,
};
use aes::{Aes128, Aes192, Aes256};
use ecb::cipher::{block_padding::NoPadding, BlockDecryptMut, BlockEncryptMut, KeyInit};
use secure_gate::{ConstantTimeEq, RevealSecret, RevealSecretMut};
use sha1::{Digest, Sha1};

/// `AlgID` values — [MS-OFFCRYPTO] §2.3.2, the `wincrypt.h` `CALG_*` constants. The same
/// four `classify` reads (`classify.rs`), spelled once per module because the detection
/// build does not compile this one.
///
/// `ALG_ID_AES_128` used to be `0x00006801` here, which is **RC4**'s identifier. Three
/// references agree on the AES values — herumi accepts `0x660e/0x660f/0x6610` and has
/// `0x6801` commented out as `AlgoRC4` (`include/standard_encryption.hpp:56-63`),
/// msoffcrypto-tool maps `0x0000660E` to AES-128 (`msoffcrypto/format/ooxml.py:48-54`),
/// LibreOffice defines `ENCRYPT_ALGO_AES128 = 0x0000660E` and
/// `ENCRYPT_ALGO_RC4 = 0x00006801` (`include/filter/msfilter/mscodec.hxx:417-420`,
/// constants only) — and so does this crate's own classifier. The wrong constant worked
/// only because `standard_encrypted.docx` declares the spec-forbidden `fAES` + `0x6801`
/// pair, and it cost two answers: a conforming `0x660E` file was refused, and a genuine
/// RC4 CryptoAPI file reached the AES path and came back `WrongPassword`.
///
/// All three AES identifiers are `pub(crate)` because `standard_encrypt` now **emits**
/// one of them: the reader's table and the writer's are the same three numbers, declared
/// once here, so a writer cannot name a cipher this reader would refuse.
pub(crate) const ALG_ID_AES_128: u32 = 0x0000_660E;
pub(crate) const ALG_ID_AES_192: u32 = 0x0000_660F;
pub(crate) const ALG_ID_AES_256: u32 = 0x0000_6610;
const ALG_ID_RC4: u32 = 0x0000_6801;

/// `AlgIDHash` for SHA-1 — `CALG_SHA1`, [MS-OFFCRYPTO] §2.3.2. The only hash this format
/// defines, and the one [`derive_standard_key`] runs; see [`require_sha1`].
const ALG_ID_HASH_SHA1: u32 = 0x0000_8004;

/// `EncryptionHeaderFlags.fAES` — [MS-OFFCRYPTO] §2.3.1.
///
/// This bit, not the version pair, is what separates ECMA-376 standard (AES) encryption
/// from RC4 CryptoAPI: both are written with `vMinor = 2`. `classify` reads it the same
/// way and for the same reason (`classify.rs`, `FLAG_AES`).
pub(crate) const FLAG_AES: u32 = 0x0000_0020;

/// [MS-OFFCRYPTO] §2.3.4.7 fixes the iteration count; the file does not carry it.
pub(crate) const SPIN_COUNT: u32 = 50_000;

/// SHA-1 digest length. Every hash on this path is SHA-1 by [MS-OFFCRYPTO] §2.3.4.7, so
/// the number is a property of the format rather than of any field in the file.
pub(crate) const SHA1_LEN: usize = 20;

/// [MS-OFFCRYPTO] §2.3.3, AES variant:
///   SaltSize (4) + Salt (16) + EncryptedVerifier (16) + VerifierHashSize (4) +
///   EncryptedVerifierHash (32) = **72** bytes.
///
/// One constant for the length guard and for the end of the last field it protects.
/// They used to be two literals — a guard at 52 and a read to 72 — so every verifier of
/// 52..=71 bytes passed the check and panicked on the slice. Where it *starts* is chosen
/// by the file, via `EncryptionHeaderSize`; how long it is is not.
pub(crate) const ENCRYPTION_VERIFIER_LEN: usize = 72;

/// The `EncryptionHeader`'s fixed fields ahead of `CSPName` — [MS-OFFCRYPTO] §2.3.2:
/// `Flags`, `SizeExtra`, `AlgID`, `AlgIDHash`, `KeySize`, `ProviderType`, `Reserved1`,
/// `Reserved2`, eight `u32`s. The floor `EncryptionHeaderSize` is checked against, and
/// the offset `CSPName` begins at. Declared here rather than in `standard_encrypt`
/// because the reader is the module that must not trust the number.
pub(crate) const HEADER_FIXED_LEN: usize = 32;

/// AES-128 key length in bytes — what `standard_encrypt` writes, and the smallest of
/// the three [`AES_KEY_LENS`] this module's cipher calls accept.
pub(crate) const AES128_KEY_LEN: usize = 16;

/// The three AES key lengths in bytes, in the order [`limits::STANDARD_KEY_BITS_AES`]
/// gives them: AES-128, AES-192, AES-256.
///
/// Fixed by AES, not by [MS-OFFCRYPTO] — but §2.3.2's `KeySize` table enumerates the
/// same three in bits, so on this path the format and the cipher agree and the header
/// is checked against the format's list (see [`aes_key_bits`]). This one exists for
/// [`check_aes_key`], whose subject is the cipher's requirement rather than the file's
/// declaration.
const AES_KEY_LENS: [usize; 3] = [16, 24, 32];

const _: () = {
    assert!(AES_KEY_LENS[0] == AES128_KEY_LEN);
    assert!(AES_KEY_LENS[0] * 8 == limits::STANDARD_KEY_BITS_AES[0] as usize);
    assert!(AES_KEY_LENS[1] * 8 == limits::STANDARD_KEY_BITS_AES[1] as usize);
    assert!(AES_KEY_LENS[2] * 8 == limits::STANDARD_KEY_BITS_AES[2] as usize);
    // §2.3.4.7 step 1: "cbRequiredKeyLength MUST be less than or equal to 40", and the
    // ladder yields exactly 2 × SHA1_LEN = 40. AES-256's key is the largest this format
    // can name and it fits with 8 bytes to spare, which is why one KDF serves all three.
    assert!(AES_KEY_LENS[2] <= 2 * SHA1_LEN);
};

/// What a standard `EncryptionInfo` stream declares, borrowed from it — every field
/// [`decrypt`] acts on and nothing it does not.
///
/// The lengths are fixed by [`parse_encryption_info`] before this exists: the salt and
/// the verifier are 16 bytes each and the verifier hash is 32, the split of the 72-byte
/// [`ENCRYPTION_VERIFIER_LEN`]. `key_size_bytes` is the one number that comes from a
/// field rather than from the layout, and on the way here it is pinned to 16, 24 or 32 —
/// the `KeySize` the file declares, checked against the cipher its `AlgID` names.
pub(crate) struct StandardParams<'a> {
    /// `EncryptionVerifier.Salt`.
    pub(crate) salt: &'a [u8],
    /// `EncryptionVerifier.EncryptedVerifier` — 16 random bytes under the derived key.
    pub(crate) encrypted_verifier: &'a [u8],
    /// `EncryptionVerifier.EncryptedVerifierHash` — SHA-1 of those bytes, padded to 32,
    /// under the same key.
    pub(crate) encrypted_verifier_hash: &'a [u8],
    /// `EncryptionHeader.KeySize / 8`.
    pub(crate) key_size_bytes: usize,
}

/// Decrypt a standard-encryption `EncryptedPackage`.
///
/// `info` is the EncryptionInfo stream with the 8-byte header stripped
/// (4-byte version + 4-byte `Flags` copy already consumed by the caller).
///
/// On error, no plaintext is produced. This format defines no integrity element.
///
/// # Errors
///
/// [`Error::MissingStream`] if the header is truncated;
/// [`Error::BadParameters`] if `EncryptionHeaderSize` or a sibling field is
/// out of range; [`Error::UnsupportedAlgorithm`] if `AlgID` names a cipher other than
/// AES or `AlgIDHash` is not SHA-1; [`Error::WrongPassword`] if the verifier does not
/// match; [`Error::CipherError`] if an AES-ECB step rejects a block.
pub(crate) fn decrypt(
    info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, Error> {
    let params = parse_encryption_info(info)?;
    let derived_key = derive_standard_key(password, params.salt, params.key_size_bytes)?;
    verify_password(&derived_key, &params)?;
    decrypt_package(&derived_key, encrypted_package)
}

/// Steps 1 and 2: the binary header and the verifier, with every field this crate acts
/// on checked before anything is derived from it.
///
/// `info` is the stream after its 8-byte prefix. Every refusal here happens before the
/// 50 000-round KDF runs, so a malformed file costs nothing to reject.
pub(crate) fn parse_encryption_info(info: &[u8]) -> Result<StandardParams<'_>, Error> {
    // MS-OFFCRYPTO §2.3.4.5: after the 8-byte version+flags header (already stripped by
    // the caller), the stream is:
    //   EncryptionHeaderSize (4 bytes LE u32)
    //   EncryptionHeader     (EncryptionHeaderSize bytes)
    //     Flags       (4) + SizeExtra (4) + AlgID (4) + AlgIDHash (4) +
    //     KeySize     (4) + ProviderType (4) + Reserved1 (4) + Reserved2 (4) = 32 bytes
    //     CSPName     (variable, null-terminated UTF-16LE)
    //   EncryptionVerifier   (72 bytes — see ENCRYPTION_VERIFIER_LEN for the fields)
    if info.len() < 36 {
        return Err(Error::MissingStream("EncryptionInfo too short"));
    }

    // `EncryptionHeaderSize` is the length of the `EncryptionHeader` that follows it, and
    // therefore the offset the `EncryptionVerifier` begins at. It is read and bounded
    // here rather than skipped: this reader used to locate the verifier by scanning
    // `CSPName` for its UTF-16 terminator, which agrees with the size field on every
    // conforming file — the shipped fixture declares 140 and its terminator sits at 140 —
    // and disagrees the moment anything follows `CSPName` inside the declared header.
    // Then the salt and both verifier blobs are read out of the file's own padding and a
    // correct password comes back `WrongPassword`. `rc4_cryptoapi::parse` has always used
    // the size field for this identical structure; the two now agree about where the
    // fields live.
    //
    // Both bounds are the ones that structure already carries: the header cannot be
    // shorter than its own 32 fixed bytes, and `RC4_ENCRYPTION_HEADER_SIZE_MAX` is the
    // ceiling on 32 fixed bytes plus a CSP name. The field is an unconstrained `u32`
    // reached before any password work, so it is checked before it is used as an offset.
    let header_size = read_u32(info, 0) as usize;
    if header_size < HEADER_FIXED_LEN {
        return Err(Error::BadParameters(format!(
            "EncryptionHeaderSize is {header_size}; the EncryptionHeader's fixed fields \
             alone take {HEADER_FIXED_LEN}"
        )));
    }
    if header_size > limits::RC4_ENCRYPTION_HEADER_SIZE_MAX {
        return Err(Error::BadParameters(format!(
            "EncryptionHeaderSize is {header_size}, over the {} this crate allows for 32 \
             fixed bytes and a CSP name",
            limits::RC4_ENCRYPTION_HEADER_SIZE_MAX
        )));
    }

    // The header itself. `info[0..4]` is the size field, so the header runs from 4 and
    // the verifier from `4 + header_size`.
    let Some(h) = info.get(4..4 + header_size) else {
        return Err(Error::BadParameters(format!(
            "EncryptionHeaderSize is {header_size} but only {} bytes follow it",
            info.len() - 4
        )));
    };

    let flags = read_u32(h, 0);
    let _size_extra = read_u32(h, 4);
    let alg_id = read_u32(h, 8);
    let alg_id_hash = read_u32(h, 12);
    let key_size_bits = read_u32(h, 16);
    let _provider_type = read_u32(h, 20);
    // Reserved1 at 24, Reserved2 at 28 — ignored

    let named_key_bits = aes_key_bits(flags, alg_id)?;
    require_sha1(alg_id_hash)?;

    // `KeySize` is an unconstrained u32 out of the file and it sizes two buffers below —
    // `KeySize / 8` truncates the fixed 40-byte XOR-ladder output (a panic above 320
    // bits) and the result is then handed to AES (a panic for anything but 16, 24 or
    // 32). Neither needs a password to reach: both run before the verifier comparison.
    //
    // Two checks, because the file makes two separate statements. First, the value must
    // be one [MS-OFFCRYPTO] §2.3.4.5 permits: "This value MUST be 0x00000080 (AES-128),
    // 0x000000C0 (AES-192), or 0x00000100 (AES-256)."
    if !limits::STANDARD_KEY_BITS_AES.contains(&key_size_bits) {
        return Err(Error::BadParameters(format!(
            "EncryptionHeader.KeySize is {key_size_bits}; [MS-OFFCRYPTO] 2.3.4.5 permits \
             {:?} for AES",
            limits::STANDARD_KEY_BITS_AES
        )));
    }
    // Second, where the `AlgID` names a key length of its own, `KeySize` must agree with
    // it: §2.3.2 gives AES-128, AES-192 and AES-256 three distinct AlgIDs *and* three
    // distinct `KeySize` values, so a header pairing `0x0000660E` with 256 describes no
    // cipher at all. Taking either field alone would silently pick a winner — and the
    // wrong pick is a 32-byte key run against AES-128's schedule, i.e. `WrongPassword`
    // for a password that was right.
    if let Some(named) = named_key_bits {
        if key_size_bits != named {
            return Err(Error::BadParameters(format!(
                "EncryptionHeader.KeySize is {key_size_bits} but AlgID {alg_id:#010x} names \
                 AES-{named}; [MS-OFFCRYPTO] 2.3.2 pairs each AlgID with one KeySize"
            )));
        }
    }

    let key_size_bytes = (key_size_bits / 8) as usize;

    // `CSPName` fills the rest of the header (`h[32..]`) and nothing here reads it: it
    // names a Windows cryptographic service provider, and §2.3.4.5 makes it a SHOULD
    // whose value cannot change how the file decrypts. Its only former job was marking
    // where the verifier began, which `EncryptionHeaderSize` now does.
    //
    // EncryptionVerifier begins where the header ends.
    let Some(v) = info.get(4 + header_size..) else {
        return Err(Error::MissingStream(
            "EncryptionVerifier missing or truncated",
        ));
    };
    if v.len() < ENCRYPTION_VERIFIER_LEN {
        return Err(Error::MissingStream(
            "EncryptionVerifier missing or truncated",
        ));
    }

    let salt_size = read_u32(v, 0) as usize;
    if salt_size != 16 {
        return Err(Error::MissingStream(
            "unexpected EncryptionVerifier salt size",
        ));
    }

    // `VerifierHashSize` — [MS-OFFCRYPTO] §2.3.4.9 step 3: "The number of bytes used by
    // the decrypted Verifier hash is given by the VerifierHashSize field, which MUST be
    // 20." `verify_password` compares exactly `SHA1_LEN` bytes and never consults this
    // field, so before this check a file could declare any width at all and be compared
    // over 20 regardless — the file's own declaration and the reader's behaviour
    // disagreeing with nothing in the path able to notice. `rc4_cryptoapi::parse` checks
    // its copy of the same field for the same reason; the two now agree.
    let verifier_hash_size = read_u32(v, 36);
    if verifier_hash_size as usize != SHA1_LEN {
        return Err(Error::BadParameters(format!(
            "EncryptionVerifier.VerifierHashSize is {verifier_hash_size}; \
             [MS-OFFCRYPTO] 2.3.4.9 requires {SHA1_LEN} (SHA-1)"
        )));
    }

    Ok(StandardParams {
        salt: &v[4..20],
        encrypted_verifier: &v[20..36],
        // verifier_hash_size at v[36..40], checked above; EncryptedVerifierHash is the
        // 32-byte blob that fills the rest — §2.3.4.9 step 3 again, "the number of bytes
        // used by the encrypted Verifier hash MUST be 32".
        encrypted_verifier_hash: &v[40..ENCRYPTION_VERIFIER_LEN],
        key_size_bytes,
    })
}

/// Step 5: the password check, [MS-OFFCRYPTO] §2.3.4.9.
///
/// The key is read once; both verifier plaintexts are wrapped, and only the comparison
/// result leaves the closures. The inverse of `standard_encrypt::generate`'s verifier
/// construction, and the function the writer's tests drive rather than a check written
/// beside them.
pub(crate) fn verify_password(key: &DerivedKey, params: &StandardParams<'_>) -> Result<(), Error> {
    let (dec_verifier, dec_hash) = key.with_secret(
        |k| -> Result<(VerifierPlaintext, VerifierPlaintext), Error> {
            Ok((
                VerifierPlaintext::new(aes_ecb_decrypt(k, params.encrypted_verifier)?),
                VerifierPlaintext::new(aes_ecb_decrypt(k, params.encrypted_verifier_hash)?),
            ))
        },
    )?;

    // Only the first `SHA1_LEN` bytes of the decrypted hash blob are the digest; the rest
    // is the writer's padding (zeros from this crate and from LibreOffice; Word compares
    // the whole blob, so a writer must pad with zeros, and the reader here follows every
    // other reader in ignoring the tail). `get` rather than a slice: `parse_encryption_info`
    // fixes the blob at 32 bytes, but this is a seam and a short blob must be a refusal.
    //
    // Constant-time, matching agile's two comparisons -- see the ct-eq note in
    // Cargo.toml. One comparison idiom across the crate, so the wrong one cannot be
    // arrived at by pattern-matching on the other.
    let matches = dec_verifier.with_secret(|v| {
        let computed = Sha1::digest(v);
        dec_hash.with_secret(|h| {
            h.get(..SHA1_LEN)
                .is_some_and(|digest| computed.as_slice().ct_eq(digest))
        })
    });
    if !matches {
        return Err(Error::WrongPassword);
    }
    Ok(())
}

/// The cipher this file names: AES at one of three key lengths, or a refusal **by name**.
///
/// Returns the key length in bits that `Flags` and `AlgID` between them declare, or
/// `None` where they name AES without naming a length — in which case `KeySize` is the
/// file's only statement of it and the caller takes that instead.
///
/// **Both fields decide it and both are read**, in the precedence [MS-OFFCRYPTO] §2.3.1
/// gives: "If the fAES encryption bit is set, a block cipher that supports ECB mode MUST
/// be used." RC4 is a stream cipher, so with the bit set the `AlgID` does not get to name
/// RC4 — and §2.3.2's combination table, which the two fields "MUST be set to one of",
/// carries no row pairing `fAES` with `0x00006801` at all. That is not academic —
/// `standard_encrypted.docx` declares exactly that forbidden pair and is AES-128 in fact,
/// and `classify` already reads it this way (`classify::classify_standard`). Reading
/// `AlgID` alone, as this function's predecessor did, made the decryptor and the
/// classifier disagree about the same bytes.
///
/// **The three AES rows, from §2.3.2's `AlgID` table verbatim:** `0x0000660E` is
/// "128-bit AES", `0x0000660F` "192-bit AES", `0x00006610` "256-bit AES"; §2.3.4.5 says
/// of the same field "This value MUST be 0x0000660E (AES-128), 0x0000660F (AES-192), or
/// 0x00006610 (AES-256)." Until 2026-09-20 this function was `require_aes_128` and
/// refused the latter two outright, so a conforming Office 2007 document using either was
/// unopenable — an owner locked out of their own file by this crate's choice and not by
/// the format's. Nothing else on the path needed to change: §2.3.4.7's KDF does not
/// branch on the key length (see this module's header), and `AlgIDHash` and the iteration
/// count are fixed by the format at SHA-1 and 50 000 whatever `AlgID` says.
///
/// The `None` arm is the forbidden `fAES` + `0x00006801` pair. §2.3.1 makes it AES, and
/// §2.3.2's combination table has no row for it and therefore states no key length, so
/// there is nothing here to agree or disagree with `KeySize` about. `AlgID = 0` is
/// different and is **not** this case: the table's `fCryptoAPI=1, fAES=1, fExternal=0,
/// AlgID=0x00000000` row reads "128-bit AES" in so many words, so the length is declared.
///
/// The refusal is [`Error::UnsupportedAlgorithm`], never
/// `UnsupportedEncryptionVersion`: an RC4 CryptoAPI file is a well-formed `vMinor = 2`
/// document this crate has not implemented, and the version pair is not what is wrong
/// with it. It used to come back `WrongPassword` — the AES verifier comparison failing on
/// an RC4 file — which is the one answer that is actively misleading.
fn aes_key_bits(flags: u32, alg_id: u32) -> Result<Option<u32>, Error> {
    const WHAT: &str = "EncryptionHeader/@AlgID";

    // The three AlgIDs that name a cipher and a key length at once. Checked before the
    // fAES branch, in both directions: the file is unambiguous about the cipher, and
    // `standard_encrypted.docx` shows the flag is not always trustworthy on its own.
    match alg_id {
        ALG_ID_AES_128 => return Ok(Some(limits::STANDARD_KEY_BITS_AES[0])),
        ALG_ID_AES_192 => return Ok(Some(limits::STANDARD_KEY_BITS_AES[1])),
        ALG_ID_AES_256 => return Ok(Some(limits::STANDARD_KEY_BITS_AES[2])),
        _ => {}
    }

    if flags & FLAG_AES != 0 {
        // fAES set with an AlgID that is not one of the three. `0x00000000` means
        // "determined by Flags" (§2.3.2) and the combination table's fAES row for it
        // says 128-bit AES; anything else (in practice RC4's `0x00006801`) is a
        // combination the table does not carry, so the length is left to `KeySize`.
        return Ok(if alg_id == 0 {
            Some(limits::STANDARD_KEY_BITS_AES[0])
        } else {
            None
        });
    }

    // fAES clear. `AlgID = 0` means "determined by Flags" ([MS-OFFCRYPTO] §2.3.2), and
    // with the AES bit clear that is RC4.
    match alg_id {
        ALG_ID_RC4 | 0 => Err(Error::UnsupportedAlgorithm {
            what: WHAT,
            name: "RC4 CryptoAPI (0x00006801)".to_string(),
        }),
        other => Err(Error::UnsupportedAlgorithm {
            what: WHAT,
            name: format!("{other:#010x}"),
        }),
    }
}

/// The hash this file names, refusing everything but SHA-1 **by name**.
///
/// [MS-OFFCRYPTO] §2.3.4.5 fixes the field for this format — "This value MUST be
/// 0x00008004 (SHA-1)" — and §2.3.2's `AlgIDHash` table reads `0x00000000` with
/// `fExternal` clear the same way. Both are accepted; §2.3.4.7 then says the hashing
/// algorithm "MUST be SHA-1" outright, so there is no combination in which the field
/// legitimately names anything else.
///
/// Refused rather than ignored because [`derive_standard_key`] is SHA-1 and nothing
/// else. A file declaring SHA-256 was previously run through the SHA-1 KDF anyway and
/// came back [`Error::WrongPassword`] — the answer that sends the caller
/// looking for a typo in a password that was right. `rc4_cryptoapi::parse` has always
/// refused its copy of this field by name; this is the standard path saying the same
/// thing about the same bytes.
fn require_sha1(alg_id_hash: u32) -> Result<(), Error> {
    match alg_id_hash {
        ALG_ID_HASH_SHA1 | 0 => Ok(()),
        other => Err(Error::UnsupportedAlgorithm {
            what: "EncryptionHeader/@AlgIDHash",
            name: format!("{other:#010x}"),
        }),
    }
}

/// Steps 3 and 4: the 50 000-round SHA-1 spin and the XOR ladder,
/// [MS-OFFCRYPTO] §2.3.4.7.
///
/// Shared with `standard_encrypt::generate`, which is what makes a file this crate
/// writes a file this crate opens: there is one derivation, not a writer's and a
/// reader's. herumi's `getEncryptionKey` (`include/standard_encryption.hpp:120-133`)
/// computes only X1, which is all a 16-byte key needs; both halves are computed here
/// because the spec defines the key as a prefix of `X1 || X2` and a reader that
/// accepted `KeySize` up to 320 bits would need the second digest.
///
/// **What is wrapped.** The UTF-16LE password buffer the first round hashes — a
/// [`crate::sensitive::Utf16Password`], wrapped since rc.4 and a bare growing `Vec`
/// before that — `H_final`, the standard path's [`PasswordDigest`], the value the
/// secure-gate skill's table places here, and the ladder output, which is the
/// [`DerivedKey`] itself. Since rc.5 the rest is wrapped too: the spin state is one
/// reused [`SpinState`] slot rather than a bare `[u8; 20]`, and the two ladder inputs are
/// [`PadBlock`]s. Both were left plain earlier on the reading that intermediate spin
/// states are only worth inverting to the password — which misses that hashing *forward*
/// from any of them reaches `H_final` for free, and the last is one SHA-1 from it. What
/// is still out of reach is inside `sha1`: the hasher's block buffer and its `finalize`
/// temporaries, which have no zeroize (the skill's *Residue* section).
pub(crate) fn derive_standard_key(
    password: &str,
    salt: &[u8],
    key_size_bytes: usize,
) -> Result<DerivedKey, Error> {
    // `key_size_bytes` is `EncryptionHeader.KeySize / 8`, a file field. `decrypt`
    // pins it to 128 bits before calling; re-checked here so a future caller cannot
    // slice past the 40 bytes two SHA-1 digests actually supply — and checked first,
    // so a request this cannot meet costs no hashing.
    if key_size_bytes > 2 * SHA1_LEN {
        return Err(Error::BadParameters(format!(
            "EncryptionHeader.KeySize asks for a {}-byte key; the SHA-1 XOR ladder \
             yields {}",
            key_size_bytes,
            2 * SHA1_LEN
        )));
    }

    // H_0 = SHA1(salt + password_utf16le)
    //
    // The UTF-16LE re-encoding is wrapped and scoped to this statement — the same
    // change, for the same reason, as `agile::spin_hash`'s first round: `collect()`
    // into a `Vec<u8>` under-reserves and abandons unwiped blocks holding a prefix of
    // the password. See `sensitive::utf16le_password`.
    //
    // The digest is finalized straight into the wrapped slot rather than returned and
    // copied in, so there is no named array of it outside the wrapper.
    let mut h = SpinState::new_with(|h| {
        utf16le_password(password).with_secret(|pw| {
            Sha1::new()
                .chain_update(salt)
                .chain_update(pw)
                .finalize_into(h.into());
        });
    });

    // H_i = SHA1(LE32(i) + H_{i-1}), in place. `chain_update` has copied `H_{i-1}` into
    // the hasher before `finalize_into` borrows the slot to overwrite it, so one slot
    // serves every round — and it is wiped on drop, which the bare `[u8; 20]` this was
    // until rc.5 never was. Every `H_i` is key-equivalent (see `SpinState`), the last
    // one most of all. `&h[..]`, not the `*h` clippy suggests for an array: `*h` passes
    // the array by value, a fresh unwiped stack copy of the state every round.
    h.with_secret_mut(|h| {
        for i in 0u32..SPIN_COUNT {
            let hasher = Sha1::new()
                .chain_update(i.to_le_bytes())
                .chain_update(&h[..]);
            hasher.finalize_into(h.into());
        }
    });

    // H_final = SHA1(H_n + LE32(0)) — the block-0 step (MS-OFFCRYPTO §2.3.4.7).
    let h_final = h.with_secret(|h| {
        PasswordDigest::new(
            Sha1::new()
                .chain_update(h)
                .chain_update(0u32.to_le_bytes())
                .finalize()
                .to_vec(),
        )
    });
    drop(h);

    // X1 = SHA1(H_final XOR 0x36-pad), X2 = SHA1(H_final XOR 0x5C-pad): 64-byte pads
    // with the digest XOR'd into the first cbHash bytes only. The key is the first
    // `key_size_bytes` of X1 || X2, cut inside the wrapper so nothing is copied out.
    // Each pad block is one hash from the key, so each is built in a wrapped slot.
    Ok(h_final.with_secret(|hf| {
        let pad = |fill: u8| {
            PadBlock::new_with(|block| {
                block.fill(fill);
                for (b, byte) in block.iter_mut().zip(hf.iter()) {
                    *b ^= byte;
                }
            })
        };
        let buf1 = pad(0x36);
        let buf2 = pad(0x5C);
        let mut derived = Vec::with_capacity(2 * SHA1_LEN);
        buf1.with_secret(|b| derived.extend_from_slice(&Sha1::digest(b)));
        buf2.with_secret(|b| derived.extend_from_slice(&Sha1::digest(b)));
        derived.truncate(key_size_bytes);
        DerivedKey::new(derived)
    }))
}

fn decrypt_package(key: &DerivedKey, encrypted_package: &[u8]) -> Result<Vec<u8>, Error> {
    if encrypted_package.len() < 8 {
        return Err(Error::MissingStream("EncryptedPackage too short"));
    }
    let declared_size = u64::from_le_bytes(encrypted_package[..8].try_into().unwrap());
    let data = &encrypted_package[8..];

    // The identical guard `agile::decrypt_package` carries, for the identical field, and
    // it matters more here: [MS-OFFCRYPTO] §2.3.4.5 defines no integrity element for
    // standard encryption, so there is no HMAC that incidentally covers this prefix and
    // nothing else in the path can notice. The number's only use is `output.truncate(..)`
    // below, and `Vec::truncate` past the current length is a silent no-op — so any
    // declared size at or above the ciphertext length was *accepted* and the caller
    // quietly received the writer's final-segment padding appended to their ZIP.
    //
    // Compared as `u64`, before the `as usize` cast: on a 32-bit target that cast
    // truncates, and a declared size of `2^32 + 16` would otherwise pass as 16 — silently
    // returning a 16-byte prefix of the document as though it were the whole thing.
    if declared_size > data.len() as u64 {
        return Err(Error::BadParameters(format!(
            "EncryptedPackage declares {declared_size} plaintext bytes but carries only {} \
             bytes of ciphertext",
            data.len()
        )));
    }
    let plaintext_size = declared_size as usize;

    let mut output = Vec::with_capacity(data.len());
    for chunk in data.chunks(4096) {
        // Pad final chunk to AES block boundary
        let mut padded = chunk.to_vec();
        let rem = padded.len() % 16;
        if rem != 0 {
            padded.resize(padded.len() + (16 - rem), 0);
        }
        let dec = key.with_secret(|k| aes_ecb_decrypt(k, &padded))?;
        output.extend_from_slice(&dec);
    }

    output.truncate(plaintext_size);
    Ok(output)
}

/// The key length rule both ECB directions share.
///
/// Checked rather than asserted, mirroring `agile::check_cbc_lengths` for the same
/// reason: `GenericArray::from_slice` panics on a length mismatch, and on the decrypt
/// side the length is reachable from `EncryptionHeader.KeySize`, an unconstrained u32 in
/// the file. The check belongs here *as well as* at the parse boundary — the parse check
/// states what the *format* permits and what the file's own `AlgID` agrees with, this one
/// states what the cipher requires, and a future caller that skips the first still hits
/// the second.
fn check_aes_key(key: &[u8]) -> Result<(), Error> {
    if !AES_KEY_LENS.contains(&key.len()) {
        return Err(Error::BadParameters(format!(
            "AES-ECB takes a 16-, 24- or 32-byte key; this file's parameters produced {}",
            key.len()
        )));
    }
    Ok(())
}

/// AES-ECB decrypt with NoPadding, AES-128/192/256 by key length — the read half of
/// [`aes_ecb_encrypt`], with the same key rule. Data must be a multiple of 16 bytes;
/// anything else is [`Error::CipherError`], never a panic.
///
/// **AES-128, AES-192 and AES-256, chosen by key length** — since 2026-09-20. Before that
/// this was `aes_ecb_decrypt` and the only key length in the module, which is why
/// `parse_encryption_info` had to refuse AES-192 and AES-256 by name one frame earlier.
/// The dispatch is `agile::aes_cbc_decrypt`'s, deliberately: the two modules now decide
/// the cipher the same way, from the same fact, so neither can be fixed without the
/// other's shape being visible.
pub(crate) fn aes_ecb_decrypt(key: &[u8], ciphertext: &[u8]) -> Result<Vec<u8>, Error> {
    use aes::cipher::generic_array::GenericArray;
    check_aes_key(key)?;
    let mut out = vec![0u8; ciphertext.len()];
    // The key length *is* the cipher, and `check_aes_key` has already refused every
    // length AES does not define — which is what makes the third arm total rather than
    // a guess.
    macro_rules! run {
        ($cipher:ty) => {
            ecb::Decryptor::<$cipher>::new(GenericArray::from_slice(key))
                .decrypt_padded_b2b_mut::<NoPadding>(ciphertext, &mut out)
        };
    }
    match key.len() {
        16 => run!(Aes128),
        24 => run!(Aes192),
        _ => run!(Aes256),
    }
    .map_err(|_| Error::CipherError)?;
    Ok(out)
}

/// AES-ECB encrypt with NoPadding, AES-128/192/256 by key length — the write half of
/// [`aes_ecb_decrypt`], with the same key rule. Data must be a multiple of 16 bytes;
/// anything else is [`Error::CipherError`], never a panic.
///
/// ECB has no IV and no chaining, which is why the standard format needs no segment
/// iterator and no per-segment IV derivation: every block is independent, and the
/// writer's job is only to pad the tail to a block. That independence is also why the
/// format is the weaker one — see `standard_encrypt`'s header.
///
/// `standard_encrypt` hands this a 16-byte key and nothing else today; the other two arms
/// exist because the reader's tests drive this function to build the AES-192 and AES-256
/// containers no writer of ours yet emits, and because letting the two directions differ
/// in what they accept is how a crate ends up writing a file it cannot read.
pub(crate) fn aes_ecb_encrypt(key: &[u8], plaintext: &[u8]) -> Result<Vec<u8>, Error> {
    use aes::cipher::generic_array::GenericArray;
    check_aes_key(key)?;
    let mut out = vec![0u8; plaintext.len()];
    macro_rules! run {
        ($cipher:ty) => {
            ecb::Encryptor::<$cipher>::new(GenericArray::from_slice(key))
                .encrypt_padded_b2b_mut::<NoPadding>(plaintext, &mut out)
        };
    }
    match key.len() {
        16 => run!(Aes128),
        24 => run!(Aes192),
        _ => run!(Aes256),
    }
    .map_err(|_| Error::CipherError)?;
    Ok(out)
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A conforming `EncryptionInfo` body — the stream with its 8-byte
    /// version + `Flags` copy prefix already removed, exactly what
    /// [`parse_encryption_info`] takes — with every field [MS-OFFCRYPTO] §2.3.4.5
    /// mandates, and two of them left to the caller so a test can put one wrong value in
    /// an otherwise valid header.
    ///
    /// Built here rather than taken from `standard_encrypt::write_encryption_info`: the
    /// reader's guards are about bytes any writer might produce, and a test that can only
    /// express what this crate's own writer emits cannot reach them.
    fn header_body(alg_id_hash: u32, verifier_hash_size: u32) -> Vec<u8> {
        header_body_full(alg_id_hash, verifier_hash_size, 0, None, 0xAA)
    }

    const CSP: &str = "Microsoft Enhanced RSA and AES Cryptographic Provider";

    /// [`header_body`] with the two fields that decide *where* the verifier begins under
    /// the caller's control as well.
    ///
    /// `trailing_header_bytes` are written after `CSPName`'s terminator and inside the
    /// declared header, which is what tells a size-driven reader apart from a
    /// terminator-scanning one: the scan stops at the terminator and the size does not.
    /// `declared_size` overrides `EncryptionHeaderSize` outright, for the bounds tests.
    /// `salt_byte` fills the salt so a test can say *which* 16 bytes it got.
    fn header_body_full(
        alg_id_hash: u32,
        verifier_hash_size: u32,
        trailing_header_bytes: usize,
        declared_size: Option<u32>,
        salt_byte: u8,
    ) -> Vec<u8> {
        let csp_len = (CSP.len() + 1) * 2;
        let real_size = (32 + csp_len + trailing_header_bytes) as u32;
        let mut out = Vec::new();
        out.extend_from_slice(&declared_size.unwrap_or(real_size).to_le_bytes());
        for field in [
            FLAG_AES | 0x0000_0004, // Flags: fAES | fCryptoAPI
            0,                      // SizeExtra
            ALG_ID_AES_128,
            alg_id_hash,
            128, // KeySize
            0x0000_0018,
            0, // Reserved1
            0, // Reserved2
        ] {
            out.extend_from_slice(&field.to_le_bytes());
        }
        for unit in CSP.encode_utf16().chain(std::iter::once(0u16)) {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        // Padding inside the declared header. `0xFF` rather than zero so that a reader
        // that landed here would produce an obviously wrong SaltSize rather than a
        // plausible one.
        out.extend_from_slice(&vec![0xFFu8; trailing_header_bytes]);
        out.extend_from_slice(&16u32.to_le_bytes()); // SaltSize
        out.extend_from_slice(&[salt_byte; 16]); // Salt
        out.extend_from_slice(&[0u8; 16]); // EncryptedVerifier
        out.extend_from_slice(&verifier_hash_size.to_le_bytes());
        out.extend_from_slice(&[0u8; 32]); // EncryptedVerifierHash
        out
    }

    /// The `EncryptionVerifier` begins at `EncryptionHeaderSize`, not wherever the first
    /// UTF-16 NUL happens to be — [MS-OFFCRYPTO] §2.3.4.5 lays the stream out as
    /// `EncryptionVersionInfo`, `EncryptionHeader.Flags`, `EncryptionHeaderSize`,
    /// `EncryptionHeader` *of that size*, `EncryptionVerifier`, and the size field is the
    /// only thing that says where the last one starts.
    ///
    /// The scan this replaces agrees with the size field on every conforming file — the
    /// shipped fixture declares 140 and its terminator sits at 140 — and disagrees the
    /// moment anything follows `CSPName` inside the header. Then the salt, the verifier
    /// and its hash are all read 16 bytes early, from the file's own padding, and the
    /// answer is `WrongPassword` for a correct password. `rc4_cryptoapi::parse` has
    /// always used the size field for the identical structure.
    #[test]
    fn the_verifier_is_located_by_encryption_header_size_not_by_scanning_cspname() {
        for trailing in [0usize, 2, 4, 16, 64] {
            let body = header_body_full(0x0000_8004, 20, trailing, None, 0xAA);
            let params = parse_encryption_info(&body)
                .unwrap_or_else(|e| panic!("trailing={trailing} is a conforming header: {e}"));
            assert_eq!(
                params.salt, &[0xAAu8; 16],
                "trailing={trailing}: the salt must come from the verifier, not the \
                 header's padding"
            );
            assert_eq!(params.encrypted_verifier, &[0u8; 16]);
            assert_eq!(params.key_size_bytes, AES128_KEY_LEN);
        }
    }

    /// `EncryptionHeaderSize` is a `u32` the file declares and it is the offset the
    /// verifier is sliced at, so every way it can lie is a refusal: below the 32 bytes
    /// the header's own fixed fields take, past the bytes actually present, and past the
    /// ceiling `limits::RC4_ENCRYPTION_HEADER_SIZE_MAX` puts on a header of 32 fixed
    /// bytes plus a CSP name.
    #[test]
    fn encryption_header_size_is_bounded_before_it_slices() {
        let good = (32 + (CSP.len() + 1) * 2) as u32;
        for declared in [0u32, 4, 31, good + 1, 4096, u32::MAX] {
            let body = header_body_full(0x0000_8004, 20, 0, Some(declared), 0xAA);
            let err = parse_encryption_info(&body)
                .err()
                .unwrap_or_else(|| panic!("EncryptionHeaderSize {declared} must be refused"));
            assert!(
                matches!(err, Error::BadParameters(_) | Error::MissingStream(_)),
                "declared={declared} got {err:?}"
            );
        }
        // The control: the honest size parses.
        assert!(
            parse_encryption_info(&header_body_full(0x0000_8004, 20, 0, Some(good), 0xAA)).is_ok()
        );
    }

    /// `EncryptionHeader.AlgIDHash` — [MS-OFFCRYPTO] §2.3.4.5: "This value MUST be
    /// 0x00008004 (SHA-1)"; §2.3.2 additionally reads `0x00000000` with `fExternal`
    /// clear as SHA-1.
    ///
    /// The KDF in `derive_standard_key` is SHA-1 and nothing else, so a file naming
    /// SHA-256 here was decrypted under SHA-1 regardless and came back `WrongPassword` —
    /// the one answer that sends a user looking for a typo in a password that was right.
    /// `rc4_cryptoapi::parse` has always refused the same field by name; this is the
    /// standard path catching up with its sibling.
    #[test]
    fn alg_id_hash_naming_anything_but_sha1_is_refused_by_name() {
        // CALG_SHA_256, CALG_SHA_384, CALG_SHA_512, CALG_MD5 — real wincrypt.h values a
        // file could plausibly carry.
        for alg_id_hash in [0x0000_800Cu32, 0x0000_800D, 0x0000_800E, 0x0000_8003] {
            let err = parse_encryption_info(&header_body(alg_id_hash, 20))
                .err()
                .unwrap_or_else(|| panic!("AlgIDHash {alg_id_hash:#010x} must be refused"));
            match &err {
                Error::UnsupportedAlgorithm { what, name } => {
                    assert_eq!(*what, "EncryptionHeader/@AlgIDHash");
                    assert!(
                        name.contains(&format!("{alg_id_hash:#010x}")),
                        "the refusal must name the value, got: {name}"
                    );
                }
                other => panic!("expected UnsupportedAlgorithm, got {other:?}"),
            }
        }
        // The controls: the two values §2.3.2 reads as SHA-1 both parse.
        for alg_id_hash in [0x0000_8004u32, 0] {
            assert!(
                parse_encryption_info(&header_body(alg_id_hash, 20)).is_ok(),
                "AlgIDHash {alg_id_hash:#010x} names SHA-1 and must be accepted"
            );
        }
    }

    /// `EncryptionVerifier.VerifierHashSize` — [MS-OFFCRYPTO] §2.3.4.9 step 3: "The
    /// number of bytes used by the decrypted Verifier hash is given by the
    /// VerifierHashSize field, which MUST be 20."
    ///
    /// `verify_password` compares exactly `SHA1_LEN` bytes and never reads the field, so
    /// a file declaring 32 was silently compared over 20 — the writer's declaration and
    /// the reader's behaviour disagreeing with nothing to notice it. Checked here, at
    /// parse, where `rc4_cryptoapi::parse` checks its own copy.
    #[test]
    fn verifier_hash_size_that_is_not_sha1s_twenty_is_refused() {
        for verifier_hash_size in [0u32, 16, 19, 21, 32, u32::MAX] {
            let err = parse_encryption_info(&header_body(0x0000_8004, verifier_hash_size))
                .err()
                .unwrap_or_else(|| panic!("VerifierHashSize {verifier_hash_size} must be refused"));
            assert!(
                matches!(err, Error::BadParameters(_)),
                "expected BadParameters for {verifier_hash_size}, got {err:?}"
            );
            assert!(
                err.to_string().contains("VerifierHashSize"),
                "the refusal must name the field, got: {err}"
            );
        }
        // The control.
        assert!(parse_encryption_info(&header_body(0x0000_8004, 20)).is_ok());
    }

    #[test]
    fn test_standard_key_derivation_length() {
        // AES-128 key = 16 bytes
        let key = derive_standard_key("Password", b"1234567890123456", 16).unwrap();
        assert_eq!(key.with_secret(|k| k.len()), 16);
    }

    #[test]
    fn test_standard_key_deterministic() {
        let k1 = derive_standard_key("test", b"1234567890123456", 16).unwrap();
        let k2 = derive_standard_key("test", b"1234567890123456", 16).unwrap();
        // `Dynamic` deliberately has no `PartialEq`; read both out explicitly.
        assert!(k1.with_secret(|a| k2.with_secret(|b| a == b)));
    }

    /// Known-answer test derived from msoffcrypto-tool's doctest:
    ///   password='Password1234_', salt=e882664990c55bd1eebd2b4394e3f830ef[:16],
    ///   expected=40b13a71f90b966e375408f2d181a1aa
    #[test]
    fn test_standard_key_known_answer() {
        let salt = [
            0xe8u8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8,
            0x30, 0xef,
        ];
        let expected = [
            0x40u8, 0xb1, 0x3a, 0x71, 0xf9, 0x0b, 0x96, 0x6e, 0x37, 0x54, 0x08, 0xf2, 0xd1, 0x81,
            0xa1, 0xaa,
        ];
        let key = derive_standard_key("Password1234_", &salt, 16).unwrap();
        key.with_secret(|k| {
            assert_eq!(
                k, &expected,
                "Standard key derivation does not match reference"
            );
        });
    }

    /// `EncryptionHeader.KeySize / 8` is a truncation length on the 40 bytes two SHA-1
    /// digests supply. `decrypt` bounds it to 128 bits first; this is the second layer,
    /// so the slice cannot go out of range for any future caller either.
    #[test]
    fn key_size_beyond_the_xor_ladder_is_an_error_not_a_panic() {
        for key_size_bytes in [41usize, 64, 0xFFFF] {
            assert!(
                matches!(
                    derive_standard_key("pw", b"1234567890123456", key_size_bytes),
                    Err(Error::BadParameters(_))
                ),
                "key_size_bytes={key_size_bytes} must be an error"
            );
        }
        // The control: 40 bytes is exactly what the ladder yields, so it is accepted.
        assert!(derive_standard_key("pw", b"1234567890123456", 40).is_ok());
    }

    /// `GenericArray::from_slice` panics on any length AES does not define, and the
    /// length reaches here from `EncryptionHeader.KeySize`. This is the mirror of the
    /// guard `agile::aes_cbc_decrypt` already had — the sibling function did not get it,
    /// so a `KeySize` of 0 or 64 bits cleared the slice above and crashed here.
    /// Both directions share the rule, so both are asserted.
    ///
    /// 32 used to be in the refusal list, because this module was AES-128 only. It is in
    /// the control list now: AES-256 is [MS-OFFCRYPTO] §2.3.2's `0x00006610`, and the
    /// length the cipher refuses is the length AES refuses, not the length this crate
    /// had implemented.
    #[test]
    fn aes_key_length_is_checked_not_asserted() {
        for key_len in [0usize, 8, 15, 17, 23, 25, 31, 33, 40] {
            assert!(
                matches!(
                    aes_ecb_decrypt(&vec![0u8; key_len], &[0u8; 16]),
                    Err(Error::BadParameters(_))
                ),
                "a {key_len}-byte key must be an error, not a panic"
            );
            assert!(
                matches!(
                    aes_ecb_encrypt(&vec![0u8; key_len], &[0u8; 16]),
                    Err(Error::BadParameters(_))
                ),
                "a {key_len}-byte key must be an error on the encrypt side too"
            );
        }
        for key_len in AES_KEY_LENS {
            assert!(aes_ecb_decrypt(&vec![0u8; key_len], &[0u8; 16]).is_ok());
            assert!(aes_ecb_encrypt(&vec![0u8; key_len], &[0u8; 16]).is_ok());
        }
    }

    /// `NoPadding` means the caller pads; a length that is not a block multiple is a
    /// refusal from the cipher layer, not a panic and not silent truncation.
    #[test]
    fn a_partial_block_is_a_cipher_error_in_both_directions() {
        for len in [1usize, 15, 17, 4095] {
            assert!(
                matches!(
                    aes_ecb_encrypt(&[0u8; 16], &vec![0u8; len]),
                    Err(Error::CipherError)
                ),
                "{len} bytes must be refused by the encrypt side"
            );
            assert!(
                matches!(
                    aes_ecb_decrypt(&[0u8; 16], &vec![0u8; len]),
                    Err(Error::CipherError)
                ),
                "{len} bytes must be refused by the decrypt side"
            );
        }
    }

    /// The sibling of `agile::tests::a_declared_plaintext_size_beyond_the_ciphertext_is_
    /// refused_not_silently_ignored`, for the identical field in the identical stream.
    ///
    /// It matters more on this path than on that one: [MS-OFFCRYPTO] §2.3.4.5 defines no
    /// integrity element for standard encryption, so nothing else in the pipeline can
    /// notice. `Vec::truncate` past the current length is a silent no-op, so any declared
    /// size at or above the ciphertext length was accepted and the caller received the
    /// writer's final-segment padding appended to their ZIP — no error, no diagnostic.
    ///
    /// The stream is written by the production `standard_encrypt::encrypt_package` — it
    /// was a test-side copy until GH #7 gave the crate a real one.
    #[test]
    fn a_declared_plaintext_size_beyond_the_ciphertext_is_refused_not_silently_ignored() {
        let plaintext: Vec<u8> = (0..5000u32).map(|i| (i % 251) as u8).collect();
        let key = DerivedKey::new(vec![0x5Au8; 16]);
        let stream = crate::standard_encrypt::encrypt_package(&plaintext, &key).unwrap();
        let ciphertext_len = (stream.len() - 8) as u64;

        for declared in [ciphertext_len + 1, ciphertext_len + 16, 1 << 40, u64::MAX] {
            let mut forged = stream.clone();
            forged[..8].copy_from_slice(&declared.to_le_bytes());
            // The length, not the bytes: an accepted forgery hands back thousands of
            // them and `Vec<u8>`'s `Debug` would bury the assertion.
            let got = decrypt_package(&key, &forged).map(|p| p.len());
            assert!(
                matches!(&got, Err(Error::BadParameters(msg))
                    if msg.contains("EncryptedPackage declares")),
                "declared={declared} got: {got:?}"
            );
        }

        // Two controls, the same pair the agile test carries. The honest size round-trips,
        // and the largest size the ciphertext could legitimately carry — every padding
        // byte counted as plaintext — is accepted rather than rejected by an off-by-one,
        // since the bound is "not more than the ciphertext", not "exactly the plaintext".
        assert_eq!(decrypt_package(&key, &stream).unwrap(), plaintext);

        let mut at_the_ceiling = stream.clone();
        at_the_ceiling[..8].copy_from_slice(&ciphertext_len.to_le_bytes());
        let out = decrypt_package(&key, &at_the_ceiling).unwrap();
        assert_eq!(out.len(), ciphertext_len as usize);
        assert!(out.starts_with(&plaintext), "the padding is what follows");
    }

    const FAES: u32 = FLAG_AES | 0x0000_0004; // fAES | fCryptoAPI, as writers emit
    const CRYPTO_API: u32 = 0x0000_0004; // fCryptoAPI alone — RC4 CryptoAPI

    /// `AlgID` alone used to decide this, against RC4's own identifier. Both halves of
    /// that were wrong answers, and both are asserted here: a conforming `0x660E` file is
    /// accepted, and an RC4 CryptoAPI header is refused **by name** rather than
    /// misdiagnosed.
    ///
    /// Since 2026-09-20 the function also *returns* the key length its two fields name,
    /// so every accepted row asserts that too — the value `parse_encryption_info` then
    /// requires `KeySize` to agree with.
    #[test]
    fn the_cipher_is_decided_by_faes_first_then_alg_id() {
        // Accepted, with the key length each row declares. The three AES AlgIDs name
        // their own ([MS-OFFCRYPTO] §2.3.2); `fAES` with `AlgID = 0` is the combination
        // table's "128-bit AES" row; `fAES` with RC4's AlgID is the spec-forbidden pair
        // the shipped fixture carries, which names AES but no length, so `None`.
        for (flags, alg_id, named) in [
            (FAES, ALG_ID_AES_128, Some(128)),
            (CRYPTO_API, ALG_ID_AES_128, Some(128)),
            (FAES, ALG_ID_AES_192, Some(192)),
            (CRYPTO_API, ALG_ID_AES_192, Some(192)),
            (FAES, ALG_ID_AES_256, Some(256)),
            (CRYPTO_API, ALG_ID_AES_256, Some(256)),
            (FAES, 0, Some(128)),
            (FAES, ALG_ID_RC4, None),
        ] {
            assert_eq!(
                aes_key_bits(flags, alg_id)
                    .unwrap_or_else(|e| panic!("flags={flags:#x} algId={alg_id:#x}: {e}")),
                named,
                "flags={flags:#x} algId={alg_id:#x}"
            );
        }

        // Refused by name. RC4 with `fAES` clear is the one that used to reach the AES
        // verifier and come back `WrongPassword`.
        for (flags, alg_id, expected) in [
            (CRYPTO_API, ALG_ID_RC4, "RC4"),
            (CRYPTO_API, 0u32, "RC4"),
            (CRYPTO_API, 0x0000_6802, "0x00006802"),
        ] {
            let err = aes_key_bits(flags, alg_id).expect_err("must be refused");
            assert!(
                matches!(&err, Error::UnsupportedAlgorithm { what, name }
                    if *what == "EncryptionHeader/@AlgID" && name.contains(expected)),
                "flags={flags:#x} algId={alg_id:#x} got: {err:?}"
            );
        }
    }

    #[test]
    fn test_aes_ecb_roundtrip_at_all_three_key_lengths() {
        for key_len in AES_KEY_LENS {
            let key_bytes = vec![0u8; key_len];
            let plaintext = b"test block data!"; // exactly 16 bytes
            let enc = aes_ecb_encrypt(&key_bytes, plaintext).unwrap();
            assert_ne!(
                &enc[..],
                &plaintext[..],
                "ECB under a zero {key_len}-byte key is not the identity"
            );
            let dec = aes_ecb_decrypt(&key_bytes, &enc).unwrap();
            assert_eq!(&dec, plaintext, "key_len={key_len}");
        }

        // The three are genuinely three ciphers and not one arm reached three ways: the
        // same 16 zero bytes under a zero key of each length must give three different
        // blocks. Without this, a dispatch that fell through to `Aes256` for everything
        // would pass every other assertion in this file.
        let blocks: Vec<Vec<u8>> = AES_KEY_LENS
            .iter()
            .map(|&n| aes_ecb_encrypt(&vec![0u8; n], &[0u8; 16]).unwrap())
            .collect();
        assert_ne!(blocks[0], blocks[1], "AES-128 and AES-192 must differ");
        assert_ne!(blocks[1], blocks[2], "AES-192 and AES-256 must differ");
        assert_ne!(blocks[0], blocks[2], "AES-128 and AES-256 must differ");
    }

    /// `verify_password` is a seam now, so a blob shorter than a digest must be a
    /// refusal there and not a slice panic — `parse_encryption_info` fixes the length at
    /// 32, but the seam does not get to rely on every caller having gone through it.
    #[test]
    fn a_verifier_hash_blob_shorter_than_a_digest_is_a_wrong_password_not_a_panic() {
        let key = DerivedKey::new(vec![0x5Au8; 16]);
        let salt = [0u8; 16];
        let encrypted_verifier = aes_ecb_encrypt(&[0x5Au8; 16], &[7u8; 16]).unwrap();
        let short_hash = aes_ecb_encrypt(&[0x5Au8; 16], &[0u8; 16]).unwrap();
        let params = StandardParams {
            salt: &salt,
            encrypted_verifier: &encrypted_verifier,
            encrypted_verifier_hash: &short_hash,
            key_size_bytes: 16,
        };
        assert!(matches!(
            verify_password(&key, &params),
            Err(Error::WrongPassword)
        ));
    }

    // ---- AES-192 and AES-256 on the read path ---------------------------------------
    //
    // [MS-OFFCRYPTO] §2.3.2 defines three AES `AlgID`s for this format — `0x0000660E`
    // (AES-128), `0x0000660F` (AES-192), `0x00006610` (AES-256) — and §2.3.4.5 repeats
    // all three against the three `KeySize` values. Until 2026-09-20 this reader refused
    // the latter two by name, so a conforming Office 2007 document using either was
    // unopenable here.
    //
    // **There is no fixture and there cannot be one from Office**, which writes AES-128
    // for this format and offers no setting that changes it (`development-record.md:225`
    // records the registry policy being ignored on the agile path for the same reason).
    // So the containers below are built at runtime, in the style of `malformed_input.rs`,
    // and built from the *reader's own* primitives inverted: the verifier blobs come from
    // `derive_standard_key` + `aes_ecb_encrypt`, which is exactly what §2.3.4.8 says a
    // writer does. That makes these tests evidence about the parse, the key length and
    // the cipher dispatch — not about interoperability, which only a real reader
    // (`tools/acceptance_gate.py`) can supply and which the write half of this work owes.

    const PW: &str = "testpass";

    /// A standard `EncryptionInfo` body — the stream after its 8-byte prefix, which is
    /// what [`parse_encryption_info`] takes — at a caller-chosen `AlgID` and `KeySize`,
    /// with the two verifier blobs supplied.
    ///
    /// `AlgID` and `KeySize` are independent parameters on purpose: a file is free to
    /// disagree with itself about them, and that disagreement is its own test below.
    fn aes_body(
        alg_id: u32,
        key_size_bits: u32,
        salt: &[u8; 16],
        enc_verifier: &[u8; 16],
        enc_hash: &[u8; 32],
    ) -> Vec<u8> {
        let csp_len = (CSP.len() + 1) * 2;
        let mut out = Vec::new();
        out.extend_from_slice(&((HEADER_FIXED_LEN + csp_len) as u32).to_le_bytes());
        for field in [
            FAES,
            0, // SizeExtra
            alg_id,
            ALG_ID_HASH_SHA1,
            key_size_bits,
            0x0000_0018, // ProviderType (AES)
            0,           // Reserved1
            0,           // Reserved2
        ] {
            out.extend_from_slice(&field.to_le_bytes());
        }
        for unit in CSP.encode_utf16().chain(std::iter::once(0u16)) {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out.extend_from_slice(&16u32.to_le_bytes()); // SaltSize
        out.extend_from_slice(salt);
        out.extend_from_slice(enc_verifier);
        out.extend_from_slice(&(SHA1_LEN as u32).to_le_bytes()); // VerifierHashSize
        out.extend_from_slice(enc_hash);
        out
    }

    /// The `EncryptionVerifier` a conforming writer produces for `password` at this key
    /// length — [MS-OFFCRYPTO] §2.3.4.8, steps 3 to 6 — plus the key itself.
    fn real_verifier(
        password: &str,
        key_size_bytes: usize,
        salt: &[u8; 16],
    ) -> (DerivedKey, [u8; 16], [u8; 32]) {
        let key = derive_standard_key(password, salt, key_size_bytes).unwrap();
        let verifier = [0x5Au8; 16];
        let enc_verifier: [u8; 16] = key
            .with_secret(|k| aes_ecb_encrypt(k, &verifier))
            .unwrap()
            .try_into()
            .unwrap();
        // §2.3.3 step 6 then step 7: the SHA-1 digest of the verifier, written into a
        // 32-byte slot whose tail is zero (§2.3.4.9 step 3 fixes the encrypted blob at
        // 32, and zeros are the pad every reader here ignores and Word requires).
        let mut hash_blob = [0u8; 32];
        hash_blob[..SHA1_LEN].copy_from_slice(&Sha1::digest(verifier));
        let enc_hash: [u8; 32] = key
            .with_secret(|k| aes_ecb_encrypt(k, &hash_blob))
            .unwrap()
            .try_into()
            .unwrap();
        (key, enc_verifier, enc_hash)
    }

    /// An `EncryptedPackage` stream under `key`: the 8-byte little-endian plaintext size
    /// then the ECB ciphertext, with the final segment zero-padded to a block
    /// ([MS-OFFCRYPTO] §2.3.4.4).
    fn encrypted_package(key: &DerivedKey, plaintext: &[u8]) -> Vec<u8> {
        let mut padded = plaintext.to_vec();
        let rem = padded.len() % 16;
        if rem != 0 {
            padded.resize(padded.len() + (16 - rem), 0);
        }
        let mut out = (plaintext.len() as u64).to_le_bytes().to_vec();
        out.extend_from_slice(&key.with_secret(|k| aes_ecb_encrypt(k, &padded)).unwrap());
        out
    }

    /// The whole read path at all three AES key lengths: parse, derive, verify, decrypt.
    ///
    /// The two that matter are AES-192 and AES-256, which this reader refused by name
    /// until 2026-09-20; AES-128 is the control, and it is the row that proves the
    /// container builder above produces something the reader accepts for reasons other
    /// than the key length.
    ///
    /// **Delete the `aes_key_bits` AES-192/AES-256 arms and this fails at
    /// `UnsupportedAlgorithm`; delete the `KeySize`-agrees-with-`AlgID` check and it
    /// still passes** — which is why the disagreement test below exists separately.
    #[test]
    fn all_three_aes_key_lengths_decrypt_end_to_end() {
        let plaintext: Vec<u8> = (0..3000u32).map(|i| (i % 251) as u8).collect();
        for (alg_id, key_size_bits) in [
            (ALG_ID_AES_128, 128u32),
            (ALG_ID_AES_192, 192),
            (ALG_ID_AES_256, 256),
        ] {
            let salt = [0x3Cu8; 16];
            let key_size_bytes = (key_size_bits / 8) as usize;
            let (key, enc_verifier, enc_hash) = real_verifier(PW, key_size_bytes, &salt);
            let body = aes_body(alg_id, key_size_bits, &salt, &enc_verifier, &enc_hash);

            let params = parse_encryption_info(&body)
                .unwrap_or_else(|e| panic!("AES-{key_size_bits} is [MS-OFFCRYPTO] 2.3.2's {alg_id:#010x} and must parse: {e}"));
            assert_eq!(params.key_size_bytes, key_size_bytes);
            assert_eq!(params.salt, &salt);

            let derived = derive_standard_key(PW, params.salt, params.key_size_bytes).unwrap();
            assert_eq!(derived.with_secret(|k| k.len()), key_size_bytes);
            verify_password(&derived, &params).unwrap_or_else(|e| {
                panic!("AES-{key_size_bits}: the right password must verify: {e}")
            });

            let package = encrypted_package(&key, &plaintext);
            assert_eq!(
                decrypt_package(&derived, &package).unwrap(),
                plaintext,
                "AES-{key_size_bits} must round-trip the package"
            );

            // The negative control, per size: without it the verifier could be accepting
            // anything and every assertion above would still hold.
            let wrong =
                derive_standard_key("not the password", params.salt, key_size_bytes).unwrap();
            assert!(
                matches!(verify_password(&wrong, &params), Err(Error::WrongPassword)),
                "AES-{key_size_bits}: a wrong password must be WrongPassword"
            );
        }
    }

    /// One KDF, three cut lengths — [MS-OFFCRYPTO] §2.3.4.7's key-derivation method,
    /// whose only key-size-dependent step is step 6, "the first `cbRequiredKeyLength`
    /// bytes of X3".
    ///
    /// So the three keys are not three derivations: each is a **prefix** of the next. If
    /// a future change ever made the derivation branch on the key size, this fails while
    /// every round-trip above still passes, because a round trip only needs the writer
    /// and the reader to agree with each other.
    #[test]
    fn the_three_key_lengths_are_prefixes_of_one_forty_byte_ladder() {
        let salt = [0x3Cu8; 16];
        let full = derive_standard_key(PW, &salt, 2 * SHA1_LEN).unwrap();
        for key_size_bytes in AES_KEY_LENS {
            let cut = derive_standard_key(PW, &salt, key_size_bytes).unwrap();
            assert!(
                cut.with_secret(|c| full.with_secret(|f| c == &f[..key_size_bytes])),
                "the {key_size_bytes}-byte key must be a prefix of the 40-byte ladder"
            );
        }
    }

    /// `AlgID` and `KeySize` each name the key length and a file can make them disagree.
    ///
    /// [MS-OFFCRYPTO] §2.3.2 pairs each AlgID with one `KeySize` (`0x0000660E` with
    /// `0x00000080`, and so on), so a header pairing AES-128's identifier with 256 bits
    /// describes no cipher the format defines. Taking either field alone would silently
    /// pick a winner, and the wrong pick derives a key of one length and runs it through
    /// another cipher's schedule — which surfaces as `WrongPassword` for a password that
    /// was right, the answer this crate has twice been bitten by.
    ///
    /// The control is the matching pair, which is the test above.
    #[test]
    fn key_size_disagreeing_with_the_alg_id_is_refused_by_name() {
        let salt = [0x3Cu8; 16];
        for (alg_id, key_size_bits) in [
            (ALG_ID_AES_128, 192u32),
            (ALG_ID_AES_128, 256),
            (ALG_ID_AES_192, 128),
            (ALG_ID_AES_192, 256),
            (ALG_ID_AES_256, 128),
            (ALG_ID_AES_256, 192),
        ] {
            let body = aes_body(alg_id, key_size_bits, &salt, &[0u8; 16], &[0u8; 32]);
            let err = parse_encryption_info(&body).err().unwrap_or_else(|| {
                panic!("AlgID {alg_id:#010x} with KeySize {key_size_bits} must be refused")
            });
            assert!(
                matches!(&err, Error::BadParameters(msg)
                    if msg.contains("KeySize") && msg.contains("AlgID")),
                "AlgID={alg_id:#010x} KeySize={key_size_bits} must name both fields, got: {err}"
            );
        }
    }

    /// The other half of the `KeySize` rule: a value outside the three §2.3.4.5
    /// enumerates is refused whatever the `AlgID` says.
    ///
    /// The field is an unconstrained `u32` reached before any password work, and
    /// `KeySize / 8` is a truncation length on a 40-byte buffer as well as an AES key
    /// length — 320 and above used to slice out of range, and anything that cleared that
    /// slice but was not 16 panicked inside `GenericArray`.
    ///
    /// **`ALG_ID_RC4` is in the sweep and it is the row that earns the check.** Deleting
    /// the enumeration check and running this test with only the three AES AlgIDs passes
    /// — every one of those names a key length, so the *agreement* check refuses the
    /// value first and the enumeration check never runs. `fAES` with RC4's AlgID is the
    /// shipped fixture's own spec-forbidden combination, and it is the one shape where
    /// `KeySize` is the file's only statement of the key length: there this check is the
    /// sole guard, and without it the file travels as far as `derive_standard_key`'s
    /// 40-byte ceiling or the cipher's key-length check and is refused by a message that
    /// names neither the field nor the format.
    #[test]
    fn key_size_outside_the_specs_three_is_refused_by_name() {
        let salt = [0x3Cu8; 16];
        for key_size_bits in [0u32, 8, 64, 127, 129, 160, 255, 320, 512, u32::MAX] {
            for alg_id in [
                ALG_ID_AES_128,
                ALG_ID_AES_192,
                ALG_ID_AES_256,
                ALG_ID_RC4, // with fAES set: names AES, names no length
            ] {
                let body = aes_body(alg_id, key_size_bits, &salt, &[0u8; 16], &[0u8; 32]);
                let err = parse_encryption_info(&body)
                    .err()
                    .unwrap_or_else(|| panic!("KeySize {key_size_bits} must be refused"));
                assert!(
                    matches!(&err, Error::BadParameters(msg) if msg.contains("KeySize")),
                    "KeySize={key_size_bits} algId={alg_id:#010x} got: {err}"
                );
            }
        }
        // The controls: each of the three, paired with its own AlgID — and each of the
        // three under the fixture's `fAES` + RC4-AlgID pair, where `KeySize` alone
        // decides and all three are therefore legitimate.
        for (alg_id, key_size_bits) in [
            (ALG_ID_AES_128, 128u32),
            (ALG_ID_AES_192, 192),
            (ALG_ID_AES_256, 256),
            (ALG_ID_RC4, 128),
            (ALG_ID_RC4, 192),
            (ALG_ID_RC4, 256),
        ] {
            assert!(
                parse_encryption_info(&aes_body(
                    alg_id,
                    key_size_bits,
                    &salt,
                    &[0u8; 16],
                    &[0u8; 32]
                ))
                .is_ok(),
                "AlgID {alg_id:#010x} with KeySize {key_size_bits} is conforming"
            );
        }
    }
}
