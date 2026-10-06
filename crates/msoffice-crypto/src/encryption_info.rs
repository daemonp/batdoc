//! Serialise the agile `\EncryptionInfo` stream — the inverse of `agile`'s parser.
//!
//! GH #6 step 3. The stream is an 8-byte binary header followed by a UTF-8 XML document:
//!
//! ```text
//! u16 vMajor = 4 | u16 vMinor = 4 | u32 Reserved = 0x00000040 | <encryption …>
//! ```
//!
//! # The shape is Word's, measured rather than composed
//!
//! Word 16, Excel 16 and PowerPoint 16 write a **byte-identical 1 289-byte stream** for
//! the same tuple — verified across `tests/fixtures/{word16_agile.docx,
//! excel16_agile.xlsx, powerpoint16_agile.pptx}`, whose XML differs only in the base64
//! values. That is the target this module reproduces exactly, down to the `\r\n` after the
//! declaration and the absence of a space before every `/>`.
//!
//! herumi writes the same document from a `snprintf` template (`crypto_util.hpp:356-390`,
//! BSD-3) and differs in exactly two places: a bare `\n` after the declaration, and the
//! `xmlns:c` certificate namespace only when its `isOffice2013` flag is set. Word emits
//! `\r\n` and always declares `xmlns:c`, so **this follows Word**. Both readings of the
//! format are fine; matching the one that ships in the product is the cheaper bet when
//! GH #8's acceptance bar is real Word opening the result.
//!
//! Nothing here is copied from herumi. The document is written from the fixture bytes
//! this repository already carries, which is a fact about Microsoft's output rather than
//! anyone's expression — the same posture as `dataspaces`.
//!
//! # Nothing secret and nothing attacker-controlled is interpolated
//!
//! Every value written below is either an integer this module chose or base64 of bytes
//! the caller generated. The password never appears in this document — that is the point
//! of the verifier blobs — and base64's alphabet cannot produce `<`, `>`, `&` or a quote,
//! so there is no escaping to get wrong and no injection to defend against. Keeping the
//! only string inputs base64 is what holds that true; do not add a free-text attribute
//! here without revisiting it.
//!
//! # The writer refuses to emit what the reader would reject
//!
//! Every length checked in [`write`] is one `agile::parse_encryption_info` cross-checks on
//! the way back in. All seven are **derived from the caller's [`EncryptParams`]** rather
//! than stated a second time here, which is the same idea one layer up: the writer and the
//! generator cannot disagree about `saltSize`, `keyBits` or the hash because they were
//! handed the same value. `hashSize` follows from `params.hash` and `blockSize` is AES's
//! 16 — neither is a parameter, because neither is free. A writer that can emit a file its
//! own reader refuses is a bug generator, and this crate is both halves.

use crate::error::Error;
use crate::EncryptParams;
use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};

/// `EncryptionInfo.vMajor` / `vMinor` for agile encryption — [MS-OFFCRYPTO] §2.3.4.10.
const AGILE_VERSION: (u16, u16) = (4, 4);

/// The AES block size, which is also `blockSize` on both elements.
const AES_BLOCK_LEN: usize = 16;

/// The cipher this crate writes, on both elements: AES in CBC.
///
/// The one part of the tuple that is **not** a parameter. [MS-OFFCRYPTO] §2.3.4.10 MUSTs
/// the two `cipherAlgorithm` attributes equal, and AES-CBC is what `agile_encrypt`'s key
/// schedule performs and what `agile`'s reader implements; a third value would be a new
/// cipher, not a new number. Everything else on both elements comes from
/// [`EncryptParams`]: `spinCount`, `hashAlgorithm`, the two `keyBits` and the two
/// `saltSize`. `hashSize` is the named hash's digest length and `blockSize` is AES's 16,
/// both derived in [`write()`] rather than repeated.
const CIPHER_ALGORITHM: &str = "AES";
const CIPHER_CHAINING: &str = "ChainingModeCBC";

/// What Office 16 writes, and [`EncryptParams::default`]'s `spin_count`.
///
/// 100 000 SHA-512 rounds, measured identically in all three real-Office fixtures. Well
/// under `crate::limits::SPIN_COUNT_MAX`, which [`EncryptParams::validate`] checks against
/// on the way into [`write()`]: the ceiling exists to bound a *hostile* file, and emitting
/// one this crate would refuse to read is the failure this module is shaped to prevent.
pub(crate) const OFFICE_SPIN_COUNT: u32 = 100_000;

/// `roundUp(n, blockSize)` — the pad every encrypted blob in this document carries.
///
/// A **spec requirement** rather than something Word merely tolerates: §2.3.4.13 pads the
/// three `p:encryptedKey` blobs to the block size with 0x00, and §2.3.4.14 step 3 does the
/// same for the two `dataIntegrity` blobs. `blockSize` is [`AES_BLOCK_LEN`], fixed by the
/// cipher and not by a parameter.
///
/// Saturating rather than wrapping so that no arithmetic here can overflow; every input is
/// a value [`EncryptParams::validate`] has already bounded, so the saturation is
/// unreachable and exists because an unreachable panic is still a panic.
fn round_up_to_block(n: usize) -> usize {
    n.div_ceil(AES_BLOCK_LEN).saturating_mul(AES_BLOCK_LEN)
}

/// A validated `u32` parameter as a length.
///
/// [`EncryptParams::validate`] has already bounded every value reaching this — `saltSize`
/// at 65 536 and `keyBits` at 256 — so the fallback cannot be reached on any target this
/// crate builds for. It saturates rather than unwrapping for the reason above, and a
/// saturated value can only make the length comparison in [`write()`] fail.
fn to_len(n: u32) -> usize {
    usize::try_from(n).unwrap_or(usize::MAX)
}

/// Everything [`write()`] needs that is not fixed by the tuple.
///
/// All of it is **public by construction**: every field is written into the document in
/// the clear, so none is wrapped and none may be. The session key, the block keys and the
/// spin hash are what produced these blobs and none of them appears here — that separation
/// is the whole reason this struct exists rather than the writer taking the key material.
pub(crate) struct EncryptionInfoParams<'a> {
    /// The tuple the blobs below were generated under.
    ///
    /// **One value, not four loose fields**, and that is the point of it. Every number
    /// this module writes that is not a blob — both `saltSize`, both `keyBits`,
    /// `hashAlgorithm`, `spinCount` — and every length it asserts comes from here, so the
    /// writer cannot be told about a hash or a key size the generator did not use. Loose
    /// fields would make that disagreement expressible; a single [`EncryptParams`] makes
    /// it unrepresentable, and the caller passes the same value to both halves.
    pub(crate) params: EncryptParams,
    /// `keyData/@saltValue` — the package salt.
    pub(crate) key_data_salt: &'a [u8],
    /// `dataIntegrity/@encryptedHmacKey`.
    pub(crate) encrypted_hmac_key: &'a [u8],
    /// `dataIntegrity/@encryptedHmacValue`.
    pub(crate) encrypted_hmac_value: &'a [u8],
    /// `p:encryptedKey/@saltValue` — the password salt.
    pub(crate) password_salt: &'a [u8],
    /// `p:encryptedKey/@encryptedVerifierHashInput`.
    pub(crate) encrypted_verifier_hash_input: &'a [u8],
    /// `p:encryptedKey/@encryptedVerifierHashValue`.
    pub(crate) encrypted_verifier_hash_value: &'a [u8],
    /// `p:encryptedKey/@encryptedKeyValue` — the wrapped session key.
    pub(crate) encrypted_key_value: &'a [u8],
}

/// Serialise the whole `\EncryptionInfo` stream, header included.
///
/// # Errors
///
/// [`Error::EncryptParams`] if [`EncryptParams::validate`] refuses the tuple — that is the
/// caller's own choice being judged, before a byte is written and by the same function a
/// caller can run itself.
///
/// [`Error::BadParameters`] if any blob is not the length that tuple fixes. Deliberately
/// **not** [`Error::EncryptParams`]: by the time that check runs `validate` has passed, so
/// a mismatch means the generator and this writer disagree about a parameter they were
/// both handed — the crate contradicting itself, not the caller choosing badly. Every one
/// of the seven is a check the parser makes on the way back in, so failing here is the
/// writer declining to produce a file this crate could not read.
pub(crate) fn write(p: &EncryptionInfoParams<'_>) -> Result<Vec<u8>, Error> {
    // The caller's tuple first, and in exactly one place. Every length below is derived
    // from it, so judging a blob against a `keyBits` this crate would refuse to write is
    // measuring with a ruler that is itself out of range. This is also where `spinCount`
    // meets `SPIN_COUNT_MAX`: this function used to check that ceiling itself, and one
    // fact checked in two places is one of them going stale.
    p.params.validate()?;
    let params = p.params;

    // `hashSize` is **derived, never a parameter** — §2.3.4.10 MUSTs it equal to the
    // digest length of the named `hashAlgorithm`, so a file may not disagree and neither
    // may a caller.
    let hash_size = params.hash.digest_len();

    // 1 and 2. Each `saltValue` against **its own element's** `saltSize` (§2.3.4.10 states
    // the MUST once per element, which is why `EncryptParams` carries two of them).
    let key_data_salt_len = to_len(params.key_data_salt_size);
    let password_salt_len = to_len(params.password_salt_size);

    // 3 and 4. Both `dataIntegrity` blobs are `roundUp(hashSize, blockSize)` under
    // `keyData`'s hash.
    //
    // **The spec contradicts itself here, and this follows the worked example.**
    // §2.3.4.14 step 2 sizes the HMAC key at `keyData.saltSize` bytes; §3.11's own worked
    // example carries a 32-byte `encryptedHmacKey` for `saltSize = 16`, `blockSize = 16`,
    // `hashSize = 20` — which is `roundUp(hashSize, blockSize)` and not the salt size.
    // Every real file measured in this repository matches the example, as does `agile`'s
    // reader, so the two halves of this crate agree with §3.11 rather than with §2.3.4.14
    // step 2. Under the default tuple the two rules coincide at 64, so only a non-default
    // hash can tell them apart.
    let hash_blob = round_up_to_block(hash_size);

    // 5. §2.3.4.13: the verifier input is `saltSize` random bytes from the **password**
    // element, encrypted — so the blob is that length padded to the block size.
    let verifier_input_blob = round_up_to_block(password_salt_len);

    // 7. §2.3.4.13 step 1: the wrapped key is the package key, whose size is
    // `Encryptor.KeyData.keyBits` — **`keyData`'s, not `p:encryptedKey`'s** — padded to
    // the block size. Not `keyBits / 8`: AES-192 puts a 24-byte key in a 32-byte blob,
    // which `agile`'s reader requires and which real Word 16 writes in
    // `tests/fixtures/agile_aes192_sha384.docx`.
    let key_value_blob = round_up_to_block(to_len(params.key_data_key_bits / 8));

    for (what, got, want) in [
        (
            "keyData/@saltValue",
            p.key_data_salt.len(),
            key_data_salt_len,
        ),
        (
            "p:encryptedKey/@saltValue",
            p.password_salt.len(),
            password_salt_len,
        ),
        (
            "dataIntegrity/@encryptedHmacKey",
            p.encrypted_hmac_key.len(),
            hash_blob,
        ),
        (
            "dataIntegrity/@encryptedHmacValue",
            p.encrypted_hmac_value.len(),
            hash_blob,
        ),
        (
            "p:encryptedKey/@encryptedVerifierHashInput",
            p.encrypted_verifier_hash_input.len(),
            verifier_input_blob,
        ),
        (
            // 6. §2.3.4.13: the digest of the verifier input, padded the same way.
            "p:encryptedKey/@encryptedVerifierHashValue",
            p.encrypted_verifier_hash_value.len(),
            hash_blob,
        ),
        (
            "p:encryptedKey/@encryptedKeyValue",
            p.encrypted_key_value.len(),
            key_value_blob,
        ),
    ] {
        if got != want {
            // Every value interpolated here is a caller-chosen integer or a fixed enum
            // name — no attacker text reaches this message, so there is nothing to bound
            // or truncate. It names the parameters rather than a tuple this crate no
            // longer fixes: "AES-256/SHA-512" was true when the tuple was a constant and
            // would now be a typed lie about whatever the caller actually asked for.
            return Err(Error::BadParameters(format!(
                "{what} is {got} bytes; hashAlgorithm={hash} (hashSize {hash_size}), \
                 keyData keyBits={kd_bits} saltSize={kd_salt}, p:encryptedKey \
                 keyBits={pw_bits} saltSize={pw_salt} and blockSize={block} fix it at \
                 {want}",
                hash = params.hash.name(),
                kd_bits = params.key_data_key_bits,
                kd_salt = params.key_data_salt_size,
                pw_bits = params.password_key_bits,
                pw_salt = params.password_salt_size,
                block = AES_BLOCK_LEN,
            )));
        }
    }

    let b64 = |bytes: &[u8]| BASE64.encode(bytes);

    // One line after the declaration, no indentation, no space before `/>`. Word's exact
    // layout -- see the module header for why it is Word's and not herumi's.
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <encryption \
         xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" \
         xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" \
         xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">\
         <keyData saltSize=\"{kd_salt_size}\" blockSize=\"{block}\" keyBits=\"{kd_key_bits}\" \
         hashSize=\"{hash_size}\" cipherAlgorithm=\"{cipher}\" cipherChaining=\"{chaining}\" \
         hashAlgorithm=\"{hash}\" saltValue=\"{key_data_salt}\"/>\
         <dataIntegrity encryptedHmacKey=\"{hmac_key}\" encryptedHmacValue=\"{hmac_value}\"/>\
         <keyEncryptors>\
         <keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">\
         <p:encryptedKey spinCount=\"{spin}\" saltSize=\"{pw_salt_size}\" \
         blockSize=\"{block}\" \
         keyBits=\"{pw_key_bits}\" hashSize=\"{hash_size}\" cipherAlgorithm=\"{cipher}\" \
         cipherChaining=\"{chaining}\" hashAlgorithm=\"{hash}\" saltValue=\"{password_salt}\" \
         encryptedVerifierHashInput=\"{verifier_input}\" \
         encryptedVerifierHashValue=\"{verifier_value}\" \
         encryptedKeyValue=\"{key_value}\"/>\
         </keyEncryptor></keyEncryptors></encryption>",
        block = AES_BLOCK_LEN,
        cipher = CIPHER_ALGORITHM,
        chaining = CIPHER_CHAINING,
        hash = params.hash.name(),
        kd_salt_size = params.key_data_salt_size,
        kd_key_bits = params.key_data_key_bits,
        pw_salt_size = params.password_salt_size,
        pw_key_bits = params.password_key_bits,
        key_data_salt = b64(p.key_data_salt),
        hmac_key = b64(p.encrypted_hmac_key),
        hmac_value = b64(p.encrypted_hmac_value),
        spin = params.spin_count,
        password_salt = b64(p.password_salt),
        verifier_input = b64(p.encrypted_verifier_hash_input),
        verifier_value = b64(p.encrypted_verifier_hash_value),
        key_value = b64(p.encrypted_key_value),
    );

    let mut stream = Vec::with_capacity(8 + xml.len());
    stream.extend_from_slice(&AGILE_VERSION.0.to_le_bytes());
    stream.extend_from_slice(&AGILE_VERSION.1.to_le_bytes());
    stream.extend_from_slice(&crate::AGILE_ENCRYPTION_RESERVED.to_le_bytes());
    stream.extend_from_slice(xml.as_bytes());
    Ok(stream)
}

#[cfg(test)]
#[path = "encryption_info_tests.rs"]
mod tests;
