//! Verify the ECMA-376 agile `dataIntegrity` HMAC over an `EncryptedPackage` stream.
//!
//! [MS-OFFCRYPTO] §2.3.4.14 defines two extra block keys and two encrypted blobs
//! carried on `<dataIntegrity>`:
//!
//! ```text
//! iv1      = H(keyData.saltValue || 5F B2 AD 01 0C B9 E1 F6)[..keyData.blockSize]
//! iv2      = H(keyData.saltValue || A0 67 7F 02 B2 2C 84 33)[..keyData.blockSize]
//! hmac_key = AES-CBC-Dec(encryptedHmacKey,   session_key, iv1)[..keyData.hashSize]
//! expected = AES-CBC-Dec(encryptedHmacValue, session_key, iv2)[..keyData.hashSize]
//! actual   = HMAC-H(hmac_key, whole EncryptedPackage stream)
//! ```
//!
//! Four details are easy to get wrong and each is silent when wrong:
//!
//! 1. **The salt is `keyData.saltValue`, not `p:encryptedKey.saltValue`.** The two are
//!    different 16-byte values in every real file. Crossing them yields a garbage IV.
//! 2. **The IV is truncated to `blockSize` (16), not `keyBits / 8` (32).** The sibling
//!    derivation for the *verifier* keys truncates to `keyBits / 8`, and both take a
//!    hash plus a block key, so the two look interchangeable and are not.
//!    (herumi `crypto_util.hpp:432-441` `generateIv` vs `:447-452` `generateKey`.)
//! 3. **The AES key is the session key** — the same key that decrypts the package,
//!    recovered from `encryptedKeyValue`. The dataIntegrity block keys feed *only* the
//!    IVs. That inverts the verifier-blob relationship, where the block key derives the
//!    key and the salt is the IV.
//! 4. **The HMAC covers ciphertext, including the 8-byte little-endian size prefix**,
//!    and including the writer's block padding on the final segment. It is
//!    encrypt-then-MAC over the stream exactly as stored. The spec says so in as many
//!    words (§2.3.4.14 step 5: "the entire EncryptedPackage stream (1), including the
//!    StreamSize field, MUST be used as the message").
//!
//! # The one place this deliberately does not follow the spec's prose
//!
//! §2.3.4.14 step 2 says the HMAC key is "a random array of bytes, known as Salt, of the
//! same length as the value of the **KeyData.saltSize** attribute". Every real file
//! contradicts it: the blob is `roundUp(hashSize, blockSize)` bytes, so the key is
//! `hashSize` long, not `saltSize`. Measured on the fixtures here (2026-09-05) —
//! `agile_encrypted.docx`, `word16_agile.docx`, `excel16_agile.xlsx` and
//! `powerpoint16_agile.pptx` all declare `saltSize="16" hashSize="64"` and carry a
//! **64**-byte `encryptedHmacKey`; `agile_aes128_sha1.docx` declares
//! `saltSize="16" hashSize="20"` and carries **32** bytes, which is 20 rounded up to a
//! block. Sixteen would be one block in every one of those files.
//!
//! herumi (`include/encode.hpp:53-72`), msoffcrypto-tool
//! (`ecma376_agile.py:483-497`) and LibreOffice (`AgileEngine.cxx:390-451`, behaviour
//! only) all key the HMAC from the whole `hashSize`-long blob, so the four independent
//! readers and Microsoft's own writer agree against the sentence. This crate follows
//! the files: [`verify`] truncates to `hashSize` and [`generate`] draws `hashSize`
//! bytes. Reading `saltSize` instead would fail the HMAC on every document Office has
//! ever written.
//!
//! Sources: herumi/msoffice (BSD-3) `include/decode.hpp:69-85` (`VerifyIntegrity`) and
//! `include/crypto_util.hpp:33-41, 432-441`; msoffcrypto-tool (MIT)
//! `msoffcrypto/method/ecma376_agile.py:483-497`. Behaviour cross-checked against
//! LibreOffice `oox/source/crypto/AgileEngine.cxx:390-451` (MPL-2.0, read-only — no
//! expression from it is reproduced here).

use crate::error::Error;
use crate::hash::{derive_iv, HashAlgorithm};
use crate::sensitive::{IntegrityKey, IntegrityTag, SessionKey};
use secure_gate::{ConstantTimeEq, RevealSecret};

/// [MS-OFFCRYPTO] §2.3.4.14 step 3 — the IV for `encryptedHmacKey`.
/// herumi `include/crypto_util.hpp:34`.
const BLOCK_DATA_INTEGRITY_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
/// [MS-OFFCRYPTO] §2.3.4.14 step 6 — the IV for `encryptedHmacValue`.
/// herumi `include/crypto_util.hpp:36`.
const BLOCK_DATA_INTEGRITY_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

/// AES block size in bytes. `keyData/@blockSize` must equal this for a CBC file.
const AES_BLOCK_LEN: usize = 16;

/// What the caller wants done about the package HMAC.
///
/// The default is [`IntegrityPolicy::RequireWhereDefined`]: fail closed on any format
/// that defines an integrity element, and report absence only where the *format* has
/// none. This is the single statement of intent behind the two places that enforce it —
/// `agile::check_integrity` for the agile half, the version dispatch in `lib.rs` for the
/// standard half — and both cite it rather than restating it.
///
/// The default was [`IntegrityPolicy::VerifyIfPresent`] until GH #12, which is a silent
/// downgrade: an attacker who can edit the file deletes ~200 bytes of XML and the check
/// evaporates, with no password and no error. See that variant's own doc for what
/// choosing it now means, and `CHANGELOG.md` for the evidence behind the flip.
///
/// `#[non_exhaustive]`, so a policy can be added without breaking a consumer's `match`.
/// Choosing one is unaffected: every variant is a unit variant and names freely.
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{
///     decrypt_ooxml_with_policy, IntegrityOutcome, IntegrityPolicy,
/// };
///
/// let decrypted = decrypt_ooxml_with_policy(
///     include_bytes!("../tests/fixtures/agile_encrypted.docx"),
///     "testpass",
///     IntegrityPolicy::Require,
/// )?;
/// assert_eq!(decrypted.integrity, IntegrityOutcome::Verified);
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`IntegrityOutcome`] reports what actually happened. [`crate::IntegrityDeclaration`]
/// is the pre-flight from [`crate::classify()`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum IntegrityPolicy {
    /// Refuse the file unless an integrity tag is present **and** verifies — in *every*
    /// format, including the ones that define no such element.
    ///
    /// Stricter than the default by exactly one case: ECMA-376 standard encryption
    /// (Office 2007), which has no integrity element by spec, is refused with
    /// [`crate::Error::IntegrityUnavailable`] rather than decrypted. Choose
    /// it when unauthenticated plaintext is unacceptable whatever the file claims to be;
    /// choose the default when a 2007-era file should still open.
    Require,
    /// **The default.** Require and verify a tag in every format that defines one, and
    /// report [`IntegrityOutcome::NotApplicable`] for a format that defines none.
    ///
    /// Concretely: an agile file must carry `<dataIntegrity>` — its absence is
    /// [`crate::Error::IntegrityElementMissing`] — and the tag must verify, or
    /// the file is refused with [`crate::Error::IntegrityCheckFailed`];
    /// ECMA-376 standard encryption decrypts and reports
    /// [`IntegrityOutcome::NotApplicable`].
    ///
    /// Agile is fail-closed here because absence is not a legitimate shape for it. The
    /// element is `minOccurs="0"` in the [MS-OFFCRYPTO] §2.3.4.10 schema, but Appendix A
    /// footnote `<22>` scopes that optionality to writers of *non*-ECMA-376 documents and
    /// states that every ECMA-376 document Office encrypts with agile encryption carries
    /// one — and this crate's input domain is exactly ECMA-376. Every writer read emits
    /// it unconditionally (herumi `include/crypto_util.hpp:369-370`, msoffcrypto-tool
    /// `method/ecma376_agile.py:143`, ms-offcrypto-writer `src/lib.rs:631-639`,
    /// LibreOffice `AgileEngine.cxx:799-802` — behaviour only), and Word 16 and
    /// PowerPoint 16 both refuse a file the element has been deleted from.
    #[default]
    RequireWhereDefined,
    /// Verify when an integrity tag is present; accept its absence.
    ///
    /// **This is a downgrade, and it is what GH #12 removed from the default.** An
    /// attacker who can modify the file deletes the `<dataIntegrity>` element and this
    /// policy decrypts happily — no password, no error, no tamper detection, and the only
    /// sign is an [`IntegrityOutcome::NotDeclared`] the caller has to think to look at.
    ///
    /// It stays reachable because no corpus is the whole world: if a legitimate writer
    /// that omits the element ever turns up, this is the escape hatch that keeps its
    /// files openable. It must be **chosen**, never inherited — that is the whole point
    /// of it no longer being `#[default]`. Prefer reading
    /// [`crate::IntegrityDeclaration`] up front over reaching for this blind.
    VerifyIfPresent,
    /// Do not verify, even when a tag is present.
    ///
    /// This returns unauthenticated plaintext. It exists so a caller salvaging a
    /// damaged document can opt into it explicitly rather than getting it by default —
    /// which is what the crate did before this check existed.
    Skip,
}

/// What actually happened to the integrity tag, reported alongside the plaintext.
///
/// `#[non_exhaustive]`, which for an enum a caller makes a *security* decision on is a
/// hazard as much as a courtesy: an exhaustive `match` would have caught a new variant at
/// compile time, and a wildcard arm silently files it wherever the caller guessed. So the
/// question a caller actually has — "was this plaintext authenticated?" — is answered by
/// [`IntegrityOutcome::is_authenticated`], a total predicate maintained here, and a new
/// variant is added to it in the same change that adds the variant.
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{
///     decrypt_ooxml_with_policy, IntegrityOutcome, IntegrityPolicy,
/// };
///
/// let decrypted = decrypt_ooxml_with_policy(
///     include_bytes!("../tests/fixtures/agile_encrypted.docx"),
///     "testpass",
///     IntegrityPolicy::RequireWhereDefined,
/// )?;
/// assert_eq!(decrypted.integrity, IntegrityOutcome::Verified);
/// assert!(decrypted.integrity.is_authenticated());
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`IntegrityPolicy`] is the request; this is the result. [`crate::IntegrityDeclaration`]
/// is what [`crate::classify()`] reported before any decrypt.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum IntegrityOutcome {
    /// A `<dataIntegrity>` tag was present and the HMAC matched.
    Verified,
    /// An agile file that declares no `<dataIntegrity>` element.
    ///
    /// **Unreachable under the default policy** since GH #12: absence is an error there,
    /// not an outcome. It survives as an outcome because the two explicit opt-outs still
    /// have to report something honest, and neither sibling would be true —
    /// [`IntegrityOutcome::Skipped`] claims a tag existed and was passed over, and
    /// [`IntegrityOutcome::NotApplicable`] claims the format defines none. Folding it
    /// into the error would leave `VerifyIfPresent` and `Skip` returning a lie about the
    /// file, which is worse than a variant that is rare by design.
    ///
    /// [`IntegrityPolicy::Skip`] on a tag-less agile file reports this rather than
    /// `Skipped`, deliberately: nothing was skipped, because nothing was there.
    NotDeclared,
    /// The format defines no integrity element at all — ECMA-376 standard
    /// encryption (Office 2007). Absence here is the spec, not a defect.
    NotApplicable,
    /// A tag was present but [`IntegrityPolicy::Skip`] said not to look at it.
    /// The returned plaintext is **unauthenticated**.
    Skipped,
}

impl IntegrityOutcome {
    /// `true` only when the package HMAC was present **and** verified.
    ///
    /// The one question this enum exists to answer, as a predicate rather than a `match`,
    /// because the enum is `#[non_exhaustive]`: a consumer's wildcard arm cannot know
    /// which side a future variant belongs on, and this can. Every variant is named here
    /// by hand — no `_` — so adding one without deciding is a compile error.
    ///
    /// `NotApplicable` is `false` deliberately. ECMA-376 standard encryption defines no
    /// integrity element, so its plaintext is unauthenticated by construction; that the
    /// format could not have done better does not make the bytes trustworthy.
    ///
    /// # Examples
    ///
    /// ```
    /// use msoffice_crypto::IntegrityOutcome;
    ///
    /// assert!(IntegrityOutcome::Verified.is_authenticated());
    /// assert!(!IntegrityOutcome::NotApplicable.is_authenticated());
    /// assert!(!IntegrityOutcome::NotDeclared.is_authenticated());
    /// assert!(!IntegrityOutcome::Skipped.is_authenticated());
    /// ```
    #[must_use]
    pub fn is_authenticated(self) -> bool {
        match self {
            Self::Verified => true,
            Self::NotDeclared | Self::NotApplicable | Self::Skipped => false,
        }
    }
}

#[cfg(test)]
mod outcome_tests {
    use super::IntegrityOutcome as O;

    /// All four arms by name. `Verified` alone is `true`; the three ways of *not* having
    /// checked — declared nothing, format defines nothing, told not to look — are each
    /// `false` on their own, so a mutation that promotes any one of them fails here rather
    /// than passing on a wildcard.
    #[test]
    fn only_verified_is_authenticated() {
        assert!(O::Verified.is_authenticated());
        assert!(!O::NotDeclared.is_authenticated());
        assert!(
            !O::NotApplicable.is_authenticated(),
            "the format could not do better; the bytes are still unauthenticated"
        );
        assert!(!O::Skipped.is_authenticated());
    }
}

/// Everything `<keyData>` and `<dataIntegrity>` contribute to the check.
///
/// All of it is public by construction — salts, sizes and ciphertext blobs are written
/// to `EncryptionInfo` in the clear — so none of it is wrapped.
pub(crate) struct IntegrityMaterial<'a> {
    /// `keyData/@saltValue`. **Not** `p:encryptedKey/@saltValue`.
    pub key_data_salt: &'a [u8],
    /// `keyData/@blockSize`.
    pub block_size: u32,
    /// `keyData/@hashSize`.
    pub hash_size: u32,
    /// `keyData/@hashAlgorithm`.
    pub hash: HashAlgorithm,
    /// `dataIntegrity/@encryptedHmacKey`, base64-decoded.
    pub encrypted_hmac_key: &'a [u8],
    /// `dataIntegrity/@encryptedHmacValue`, base64-decoded.
    pub encrypted_hmac_value: &'a [u8],
}

/// Verify the package HMAC. Returns `Ok(())` only when the tag matches.
///
/// `encrypted_package` must be the **whole** `EncryptedPackage` stream as stored,
/// starting at its 8-byte little-endian plaintext-size prefix.
///
/// The caller must have verified the password first: a wrong password yields a wrong
/// session key, and every blob below then decrypts to noise, so the mismatch would be
/// reported as corruption rather than as a bad password.
pub(crate) fn verify(
    session_key: &SessionKey,
    material: &IntegrityMaterial<'_>,
    encrypted_package: &[u8],
) -> Result<(), Error> {
    let hash = material.hash;
    let hash_size = material.hash_size as usize;
    let block_size = material.block_size as usize;

    // `hashSize` is a truncation length taken straight from the file. Pair it against
    // the algorithm before trusting it, exactly as herumi does when parsing
    // (`include/crypto_util.hpp:156-166`).
    if hash_size != hash.digest_len() {
        return Err(Error::BadParameters(format!(
            "keyData hashSize {} does not match hashAlgorithm {} (expected {})",
            hash_size,
            hash.name(),
            hash.digest_len()
        )));
    }
    if block_size != AES_BLOCK_LEN {
        return Err(Error::BadParameters(format!(
            "keyData blockSize {block_size} is not the AES block size ({AES_BLOCK_LEN})"
        )));
    }

    let iv1 = derive_iv(
        hash,
        material.key_data_salt,
        &BLOCK_DATA_INTEGRITY_KEY,
        block_size,
    )?;
    let iv2 = derive_iv(
        hash,
        material.key_data_salt,
        &BLOCK_DATA_INTEGRITY_VALUE,
        block_size,
    )?;

    let hmac_key = IntegrityKey::new(unwrap_blob(
        session_key,
        material.encrypted_hmac_key,
        &iv1,
        hash_size,
        "encryptedHmacKey",
    )?);
    let expected = IntegrityTag::new(unwrap_blob(
        session_key,
        material.encrypted_hmac_value,
        &iv2,
        hash_size,
        "encryptedHmacValue",
    )?);

    let actual = IntegrityTag::new(hmac_key.with_secret(|k| hash.hmac(k, encrypted_package)));

    // Constant-time. Both operands stay inside their closures; only the bool escapes.
    // This one is a genuine MAC-forgery oracle if it short-circuits — unlike
    // `verify_password`, the right-hand side is a value the attacker supplies in the
    // file — which is why secure-gate's `ct-eq` feature is switched on for this crate.
    let matches =
        actual.with_secret(|a| expected.with_secret(|e| a.as_slice().ct_eq(e.as_slice())));

    if !matches {
        return Err(Error::IntegrityCheckFailed);
    }
    Ok(())
}

/// AES-CBC-decrypt one dataIntegrity blob and cut it back to the digest length.
///
/// The writer pads the `hashSize`-byte plaintext up to a `blockSize` multiple before
/// encrypting, so the blob is `roundUp(hashSize, blockSize)` bytes. Dropping that tail
/// is mandatory: keep it and the value is compared against something longer than any
/// digest, and — under any pad but zero — the HMAC is keyed with the wrong bytes too. It
/// is a no-op for SHA-256/384/512 (32/48/64 are all multiples of 16) and load-bearing for
/// SHA-1 (20 → 32). This reader is pad-agnostic, as herumi (`decode.hpp:79-80`) and
/// LibreOffice (`AgileEngine.cxx:413, :440`, behaviour only) are; Word 16 is not, see
/// [`generate_with_key`].
///
/// A blob **shorter** than `hash_size` is an error. LibreOffice zero-extends it
/// (`AgileEngine.cxx:413`, behaviour only), which turns a truncated file into a
/// wrong-but-successful decrypt.
fn unwrap_blob(
    session_key: &SessionKey,
    blob: &[u8],
    iv: &[u8],
    hash_size: usize,
    what: &'static str,
) -> Result<Vec<u8>, Error> {
    if blob.len() < hash_size || blob.len() % AES_BLOCK_LEN != 0 {
        return Err(Error::BadParameters(format!(
            "dataIntegrity {what} is {} bytes; expected a non-zero multiple of {} at least {} long",
            blob.len(),
            AES_BLOCK_LEN,
            hash_size
        )));
    }
    let mut plain = session_key.with_secret(|k| crate::agile::aes_cbc_decrypt(blob, k, iv))?;
    plain.truncate(hash_size);
    Ok(plain)
}

/// The two blobs a writer puts on `<dataIntegrity>` — public ciphertext, unwrapped.
pub(crate) struct IntegrityBlobs {
    /// `dataIntegrity/@encryptedHmacKey`, before base64.
    pub(crate) encrypted_hmac_key: Vec<u8>,
    /// `dataIntegrity/@encryptedHmacValue`, before base64.
    pub(crate) encrypted_hmac_value: Vec<u8>,
}

/// Produce the `<dataIntegrity>` pair for a package — the inverse of [`verify`], and the
/// encrypt half of GH #3. [MS-OFFCRYPTO] §2.3.4.14; herumi `include/encode.hpp:53-72`
/// (`GenerateIntegrityParameter`, BSD-3, attribution in `NOTICE`).
///
/// Draws the HMAC key from the injected `rng` — `hashSize` bytes, wrapped as
/// [`IntegrityKey`] the moment they exist — and hands off to [`generate_with_key`].
/// Randomness enters this way and no other (plan D3), so a seeded RNG makes the output a
/// committable golden and the production caller passes `rand::rngs::SysRng` through the
/// same function.
pub(crate) fn generate<R: rand::TryRng + rand::TryCryptoRng>(
    session_key: &SessionKey,
    hash: HashAlgorithm,
    key_data_salt: &[u8],
    block_size: usize,
    encrypted_package: &[u8],
    rng: &mut R,
) -> Result<IntegrityBlobs, Error> {
    let hmac_key = IntegrityKey::from_rng(hash.digest_len(), rng)
        .map_err(crate::agile_encrypt::random_source)?;
    generate_with_key(
        session_key,
        hash,
        key_data_salt,
        block_size,
        encrypted_package,
        &hmac_key,
    )
}

/// [`generate`] with the HMAC key supplied — deterministic, which is what makes it
/// checkable against Office's own output: unwrap the key from a real fixture's
/// `encryptedHmacKey`, hand it back in here, and the two blobs that come out must be the
/// fixture's, byte for byte.
///
/// ```text
/// iv1      = H(keyData.saltValue || 5F B2 AD 01 0C B9 E1 F6)[..blockSize]
/// iv2      = H(keyData.saltValue || A0 67 7F 02 B2 2C 84 33)[..blockSize]
/// key blob = AES-CBC-Enc(pad_zero(hmac_key),                 session_key, iv1)
/// tag blob = AES-CBC-Enc(pad_zero(HMAC-H(hmac_key, package)), session_key, iv2)
/// ```
///
/// `package` is the **whole** `EncryptedPackage` stream, 8-byte size prefix included —
/// the same bytes [`verify`] HMACs, because the two must agree on exactly one thing and
/// this is it. Both plaintexts are padded up to a `blockSize` multiple **with zeros** —
/// a no-op for SHA-256/384/512 and load-bearing for SHA-1 (20 → 32).
///
/// Zero is the spec's byte, not a convention this crate settled on: §2.3.4.14 step 3 says
/// to "pad the array with 0x00 to the next integral multiple of blockSize bytes" (and
/// §2.3.4.13 says the same of the six encrypted blobs). That is a requirement, so it
/// outranks the measurement below — which is kept because it is the independent
/// corroboration, and because it is what tells you *why* a writer cannot quietly pick a
/// different filler and still be opened.
///
/// LibreOffice pads these blobs with 0x36
/// (`AgileEngine.cxx:654-656, :688-689`, behaviour only), and this crate did the same
/// until the SHA-1 fixture was put in front of real Word 16: Word passed the verifier and
/// then refused the file with `0x800A1066` ("Command failed"). Thirteen same-length
/// variants of that file, each changing one pad or one HMAC input (2026-09-05,
/// CHANGELOG), pinned the rule — Word keys the HMAC with the **whole** decrypted key
/// blob, compares the **whole** decrypted value blob against its own HMAC zero-extended
/// to the blob length, and does the same to `encryptedVerifierHashValue`. A 0x36 tail in
/// any of the three is a refusal; zero tails in all three open. So LibreOffice's writer
/// choice is latent only because LibreOffice never writes SHA-1, and a writer that wants
/// Word to open a SHA-1 file has exactly one pad byte available. Readers that truncate to
/// `hashSize` — this one, herumi, LibreOffice — cannot tell the difference, and HMAC
/// itself zero-extends a short key, which is why msoffcrypto-tool's untruncated key
/// (`ecma376_agile.py:490-497`) also verifies a zero-padded file.
pub(crate) fn generate_with_key(
    session_key: &SessionKey,
    hash: HashAlgorithm,
    key_data_salt: &[u8],
    block_size: usize,
    encrypted_package: &[u8],
    hmac_key: &IntegrityKey,
) -> Result<IntegrityBlobs, Error> {
    // The same pin `verify` applies: this crate writes AES-CBC and nothing else.
    if block_size != AES_BLOCK_LEN {
        return Err(Error::BadParameters(format!(
            "keyData blockSize {block_size} is not the AES block size ({AES_BLOCK_LEN})"
        )));
    }
    let iv1 = derive_iv(hash, key_data_salt, &BLOCK_DATA_INTEGRITY_KEY, block_size)?;
    let iv2 = derive_iv(hash, key_data_salt, &BLOCK_DATA_INTEGRITY_VALUE, block_size)?;

    // The padded length of a blob: §2.3.4.14 step 3's "next integral multiple of
    // blockSize bytes". `checked_mul` rather than `*` because a panic is a vulnerability
    // in this crate even where the operands are this well behaved -- `block_size` is
    // pinned to 16 four lines above and the lengths are digest lengths, so the overflow
    // is unreachable and the error arm is there to keep it that way if either ever stops
    // being true.
    let padded_len = |len: usize| -> Result<usize, Error> {
        len.div_ceil(block_size)
            .checked_mul(block_size)
            .ok_or_else(|| {
                Error::BadParameters(format!(
                    "padding {len} bytes up to a multiple of {block_size} overflows"
                ))
            })
    };

    // Pad INTO a wrapped buffer allocated at the final length -- never draw short and
    // resize. The obvious `v.to_vec()` followed by `resize` is wrong twice over, and both
    // faults are the ones `standard_encrypt.rs:204-215` already documents:
    //
    // 1. `to_vec` allocates exactly `hashSize` bytes, and `resize` to a larger length
    //    cannot fit, so it reallocates and frees the block holding the HMAC key -- or the
    //    tag -- **unwiped**. Nothing in the wrapper can reach an allocation it no longer
    //    owns. It is a no-op at SHA-512, whose 64 bytes are already a block multiple; it
    //    goes live the moment anything writes SHA-1 (20 -> 32), which is a reachable
    //    parameter, not a hypothetical.
    // 2. The padded copy was a bare `Vec<u8>` holding the whole dataIntegrity HMAC key --
    //    the value §2.3.4.14 exists to protect, the one that forges a tag for any package
    //    this session key encrypts -- wrapped nowhere and zeroized never.
    //
    // `Dynamic::new_with` closes both: one allocation at the final size, and a slot
    // secure-gate guarantees is pre-zeroed, so the 0x00 tail the spec requires is the
    // guarantee rather than something a `resize` writes on the way to a new block. The
    // form is the one `standard_encrypt.rs:217` and `rc4_cryptoapi.rs:310` already take.
    // The slice bound is safe by construction: `padded_len(n) >= n`.
    let encrypted_hmac_key = hmac_key.with_secret(|k| {
        let padded = IntegrityKey::new_with(padded_len(k.len())?, |slot| {
            slot[..k.len()].copy_from_slice(k);
        });
        padded.with_secret(|p| {
            session_key.with_secret(|sk| crate::agile::aes_cbc_encrypt(p, sk, &iv1))
        })
    })?;

    // `IntegrityTag::new` takes ownership of the `Vec` the HMAC returns -- exact-sized, so
    // nothing is abandoned on the way in -- and the padded copy is a second wrapped
    // buffer rather than a grown first one, for the reason above.
    let tag = IntegrityTag::new(hmac_key.with_secret(|k| hash.hmac(k, encrypted_package)));
    let encrypted_hmac_value = tag.with_secret(|t| {
        let padded = IntegrityTag::new_with(padded_len(t.len())?, |slot| {
            slot[..t.len()].copy_from_slice(t);
        });
        padded.with_secret(|p| {
            session_key.with_secret(|sk| crate::agile::aes_cbc_encrypt(p, sk, &iv2))
        })
    })?;

    Ok(IntegrityBlobs {
        encrypted_hmac_key,
        encrypted_hmac_value,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::generic_array::GenericArray;
    use aes::Aes256;
    use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};

    /// Encrypt with AES-256-CBC/NoPadding — the writer half of `unwrap_blob`.
    fn cbc_encrypt(plain: &[u8], key: &[u8], iv: &[u8]) -> Vec<u8> {
        let mut out = vec![0u8; plain.len()];
        cbc::Encryptor::<Aes256>::new(GenericArray::from_slice(key), GenericArray::from_slice(iv))
            .encrypt_padded_b2b_mut::<NoPadding>(plain, &mut out)
            .unwrap();
        out
    }

    /// Build a complete, self-consistent dataIntegrity pair the way a writer would:
    /// pad both values up to a block multiple with `pad_byte` before encrypting. Zero is
    /// what Word requires and what `generate_with_key` writes; 0x36 is what LibreOffice
    /// writes (`AgileEngine.cxx:654-656, :688-689`, behaviour only), and the reader tests
    /// use it on purpose — a non-zero tail is the only thing that makes the key's
    /// truncation observable, because HMAC zero-extends a short key itself.
    fn make_material(
        hash: HashAlgorithm,
        session_key: &[u8; 32],
        salt: &[u8],
        package: &[u8],
        hmac_key: &[u8],
        pad_byte: u8,
    ) -> (Vec<u8>, Vec<u8>) {
        let block = AES_BLOCK_LEN;
        let pad = |v: &[u8]| {
            let mut p = v.to_vec();
            let target = v.len().div_ceil(block) * block;
            p.resize(target, pad_byte);
            p
        };
        let iv1 = derive_iv(hash, salt, &BLOCK_DATA_INTEGRITY_KEY, block).unwrap();
        let iv2 = derive_iv(hash, salt, &BLOCK_DATA_INTEGRITY_VALUE, block).unwrap();
        let tag = hash.hmac(hmac_key, package);
        (
            cbc_encrypt(&pad(hmac_key), session_key, &iv1),
            cbc_encrypt(&pad(&tag), session_key, &iv2),
        )
    }

    // ---- the write half (GH #6 step 5) ----------------------------------------------
    //
    // `make_material` above stays exactly what it was: a second, independent
    // implementation of the write side, written for the *verify* tests before `generate`
    // existed. It is not rewritten to call `generate`, because then every verify test
    // would be checking the writer against itself. It is the oracle instead.

    /// One real fixture's streams, its parameters, and its session key, recovered through
    /// the same code `decrypt` runs.
    fn office_fixture(name: &str) -> (Vec<u8>, crate::agile::AgileParams, SessionKey) {
        let path = format!("{}/tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"));
        let data = std::fs::read(&path).unwrap_or_else(|e| panic!("{path}: {e}"));
        let streams = crate::cfb_reader::read_cfb_streams(&data).expect("an encrypted container");
        let params = crate::agile::parse_encryption_info(&streams.encryption_info[8..])
            .expect("Office writes a parseable EncryptionInfo");
        let session_key =
            crate::agile::recover_session_key(&params, "testpass").expect("the password");
        (streams.encrypted_package, params, session_key)
    }

    /// **Given Office's own HMAC key, the writer reproduces Office's own blobs, byte for
    /// byte** — on all three real-Office fixtures.
    ///
    /// The strongest statement available about a writer whose only random input is the
    /// key: unwrap the key from the fixture's `encryptedHmacKey` (exactly what `verify`
    /// does on the way in), hand it back to `generate_with_key`, and require both blobs
    /// to be the fixture's. It pins the IV block-key order, the IV truncation, that the
    /// HMAC runs over the whole stream *including* the 8-byte prefix, and that the key is
    /// encrypted under the session key rather than a derived one — every place a
    /// plausible reading of §2.3.4.14 produces a plausible blob that is not Office's.
    ///
    /// Not a round-trip. `generate` → `verify` agreeing proves the two halves agree,
    /// which they would even if both were wrong about the format.
    #[test]
    #[cfg_attr(
        not(fixture_corpus),
        ignore = "needs the fixture corpus, which the published crate does not ship"
    )]
    fn offices_own_blobs_are_reproduced_byte_for_byte_from_offices_own_hmac_key() {
        for name in [
            "word16_agile.docx",
            "excel16_agile.xlsx",
            "powerpoint16_agile.pptx",
        ] {
            let (package, params, session_key) = office_fixture(name);
            let di = params
                .data_integrity()
                .expect("Office always declares the element");
            let m = params
                .integrity_material(di)
                .expect("Office's keyData is complete");

            let iv1 = derive_iv(m.hash, m.key_data_salt, &BLOCK_DATA_INTEGRITY_KEY, 16).unwrap();
            let offices_key = IntegrityKey::new(
                unwrap_blob(
                    &session_key,
                    m.encrypted_hmac_key,
                    &iv1,
                    m.hash_size as usize,
                    "key",
                )
                .expect("Office's key blob unwraps"),
            );

            let blobs = generate_with_key(
                &session_key,
                m.hash,
                m.key_data_salt,
                m.block_size as usize,
                &package,
                &offices_key,
            )
            .expect("Office's parameters are writable");

            assert_eq!(
                blobs.encrypted_hmac_key, m.encrypted_hmac_key,
                "{name}: encryptedHmacKey differs from what Office wrote"
            );
            assert_eq!(
                blobs.encrypted_hmac_value, m.encrypted_hmac_value,
                "{name}: encryptedHmacValue differs from what Office wrote"
            );
        }
    }

    /// `generate` draws its key from the injected RNG and nowhere else: the same seed
    /// reproduces the pair, a different seed does not, and what it produces verifies —
    /// and stops verifying the moment a package byte moves.
    #[test]
    fn generated_blobs_verify_and_follow_the_injected_rng() {
        use rand::SeedableRng as _;
        let sk = SessionKey::new(vec![0x42u8; 32]);
        let salt = [0x33u8; 16];
        let mut package = vec![0u8; 8];
        package.extend_from_slice(&[0x9Cu8; 4096 + 300]);
        let gen = |seed: u8| {
            generate(
                &sk,
                HashAlgorithm::Sha512,
                &salt,
                16,
                &package,
                &mut chacha20::ChaCha12Rng::from_seed([seed; 32]),
            )
            .unwrap()
        };

        let a = gen(1);
        let b = gen(1);
        let c = gen(2);
        assert_eq!(
            a.encrypted_hmac_key, b.encrypted_hmac_key,
            "seeded: reproducible"
        );
        assert_eq!(a.encrypted_hmac_value, b.encrypted_hmac_value);
        assert_ne!(
            a.encrypted_hmac_key, c.encrypted_hmac_key,
            "seeded: seed-dependent"
        );

        fn material<'a>(salt: &'a [u8], blobs: &'a IntegrityBlobs) -> IntegrityMaterial<'a> {
            IntegrityMaterial {
                key_data_salt: salt,
                block_size: 16,
                hash_size: 64,
                hash: HashAlgorithm::Sha512,
                encrypted_hmac_key: &blobs.encrypted_hmac_key,
                encrypted_hmac_value: &blobs.encrypted_hmac_value,
            }
        }
        verify(&sk, &material(&salt, &a), &package).expect("what generate wrote must verify");

        let mut tampered = package.clone();
        tampered[8 + 4096 + 7] ^= 0x80;
        assert!(matches!(
            verify(&sk, &material(&salt, &a), &tampered),
            Err(Error::IntegrityCheckFailed)
        ));
    }

    /// `generate_with_key` agrees with `make_material`, the independent test-side writer,
    /// on every hash — two implementations of §2.3.4.14 written at different times for
    /// different reasons producing the same bytes.
    #[test]
    fn generate_agrees_with_the_independent_test_side_writer_on_every_hash() {
        let session_key = [0x77u8; 32];
        let sk = SessionKey::new(session_key.to_vec());
        let salt = [0x44u8; 16];
        let mut package = vec![0u8; 8];
        package.extend_from_slice(&[0x5Au8; 1000]);
        for hash in [
            HashAlgorithm::Sha1,
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ] {
            let key = vec![0xABu8; hash.digest_len()];
            let (want_key, want_value) =
                make_material(hash, &session_key, &salt, &package, &key, 0);
            let got =
                generate_with_key(&sk, hash, &salt, 16, &package, &IntegrityKey::new(key)).unwrap();
            assert_eq!(got.encrypted_hmac_key, want_key, "{hash:?}");
            assert_eq!(got.encrypted_hmac_value, want_value, "{hash:?}");
        }
    }

    /// Under SHA-1 both plaintexts are 20 bytes and both blobs are 32, and the twelve
    /// bytes the reader cuts away are **zero** — the one pad byte real Word accepts.
    ///
    /// This pinned 0x36 until 2026-09-05, on the reasoning that the pad is invisible to
    /// every reader and matching LibreOffice's byte was the conservative choice. Word 16
    /// then refused a 0x36-padded SHA-1 file and opened the same file zero-padded (see
    /// `generate_with_key`). The Office fixtures cannot pin this — 64 is already a block
    /// multiple — so this is the only test that does, and it is the one that fails if the
    /// pad byte drifts back.
    #[test]
    fn sha1_blobs_are_padded_to_a_block_with_zeros() {
        let sk = SessionKey::new(vec![0x11u8; 32]);
        let salt = [0x22u8; 16];
        let package = b"\x08\x00\x00\x00\x00\x00\x00\x00eight bytes here".to_vec();
        let key = IntegrityKey::new(vec![0xCDu8; 20]);
        let blobs = generate_with_key(&sk, HashAlgorithm::Sha1, &salt, 16, &package, &key).unwrap();
        assert_eq!(blobs.encrypted_hmac_key.len(), 32, "roundUp(20, 16)");
        assert_eq!(blobs.encrypted_hmac_value.len(), 32);

        let iv1 = derive_iv(HashAlgorithm::Sha1, &salt, &BLOCK_DATA_INTEGRITY_KEY, 16).unwrap();
        let raw = sk
            .with_secret(|k| crate::agile::aes_cbc_decrypt(&blobs.encrypted_hmac_key, k, &iv1))
            .unwrap();
        assert_eq!(&raw[..20], &[0xCDu8; 20], "the key survives");
        assert_eq!(&raw[20..], &[0u8; 12], "the pad is zero, as Word requires");

        let iv2 = derive_iv(HashAlgorithm::Sha1, &salt, &BLOCK_DATA_INTEGRITY_VALUE, 16).unwrap();
        let raw = sk
            .with_secret(|k| crate::agile::aes_cbc_decrypt(&blobs.encrypted_hmac_value, k, &iv2))
            .unwrap();
        assert_eq!(
            &raw[..20],
            &HashAlgorithm::Sha1.hmac(&[0xCDu8; 20], &package)[..]
        );
        assert_eq!(&raw[20..], &[0u8; 12], "the value's pad is zero too");
    }

    /// SHA-1 is the only agile shape where `hashSize` (20) is not already a multiple of
    /// the AES block size, so it is the only one that exercises the truncation in
    /// `unwrap_blob`. No fixture in the corpus is SHA-1, which is exactly why this test
    /// is synthetic: without it the truncation is unexercised code that passes either way.
    #[test]
    fn sha1_blobs_need_the_hash_size_truncation() {
        let session_key = [7u8; 32];
        let salt = [0x11u8; 16];
        let package = b"\x10\x00\x00\x00\x00\x00\x00\x00ciphertext bytes".to_vec();
        let hmac_key = [0xABu8; 20];

        // 0x36, not zero: a zero tail is indistinguishable to HMAC, which zero-extends a
        // short key itself, so only a non-zero pad can show the key truncation working.
        let (enc_key, enc_value) = make_material(
            HashAlgorithm::Sha1,
            &session_key,
            &salt,
            &package,
            &hmac_key,
            0x36,
        );
        // roundUp(20, 16) = 32 — the blobs carry twelve pad bytes each.
        assert_eq!(enc_key.len(), 32);
        assert_eq!(enc_value.len(), 32);

        let sk = SessionKey::new(session_key.to_vec());
        let material = IntegrityMaterial {
            key_data_salt: &salt,
            block_size: 16,
            hash_size: 20,
            hash: HashAlgorithm::Sha1,
            encrypted_hmac_key: &enc_key,
            encrypted_hmac_value: &enc_value,
        };
        verify(&sk, &material, &package).expect("SHA-1 dataIntegrity must verify");

        // The truncation is what makes it work: keeping the 0x36 tail gives a 32-byte
        // key and a 32-byte "expected" that no 20-byte digest can equal.
        let untruncated_key = sk
            .with_secret(|k| {
                crate::agile::aes_cbc_decrypt(
                    &enc_key,
                    k,
                    &derive_iv(HashAlgorithm::Sha1, &salt, &BLOCK_DATA_INTEGRITY_KEY, 16).unwrap(),
                )
            })
            .unwrap();
        assert_eq!(untruncated_key.len(), 32);
        assert_ne!(
            HashAlgorithm::Sha1.hmac(&untruncated_key, &package),
            HashAlgorithm::Sha1.hmac(&hmac_key, &package),
            "a 32-byte key must not produce the same HMAC as the 20-byte one"
        );
    }

    #[test]
    fn sha512_blobs_verify_and_a_flipped_package_byte_is_refused() {
        let session_key = [3u8; 32];
        let salt = [0x22u8; 16];
        let mut package = vec![0u8; 8];
        package.extend_from_slice(&[0x5Au8; 64]);
        let hmac_key = [0xCDu8; 64];

        let (enc_key, enc_value) = make_material(
            HashAlgorithm::Sha512,
            &session_key,
            &salt,
            &package,
            &hmac_key,
            0,
        );
        let sk = SessionKey::new(session_key.to_vec());
        let material = IntegrityMaterial {
            key_data_salt: &salt,
            block_size: 16,
            hash_size: 64,
            hash: HashAlgorithm::Sha512,
            encrypted_hmac_key: &enc_key,
            encrypted_hmac_value: &enc_value,
        };
        verify(&sk, &material, &package).expect("SHA-512 dataIntegrity must verify");

        let mut tampered = package.clone();
        tampered[20] ^= 0x01;
        assert!(matches!(
            verify(&sk, &material, &tampered),
            Err(Error::IntegrityCheckFailed)
        ));

        // The 8-byte size prefix is inside the HMAC too.
        let mut prefix_tampered = package.clone();
        prefix_tampered[0] ^= 0x01;
        assert!(matches!(
            verify(&sk, &material, &prefix_tampered),
            Err(Error::IntegrityCheckFailed)
        ));
    }

    #[test]
    fn hash_size_must_match_the_named_algorithm() {
        let sk = SessionKey::new(vec![0u8; 32]);
        let material = IntegrityMaterial {
            key_data_salt: &[0u8; 16],
            block_size: 16,
            hash_size: 32, // SHA-512 is 64
            hash: HashAlgorithm::Sha512,
            encrypted_hmac_key: &[0u8; 64],
            encrypted_hmac_value: &[0u8; 64],
        };
        assert!(matches!(
            verify(&sk, &material, b""),
            Err(Error::BadParameters(_))
        ));
    }

    #[test]
    fn malformed_blob_lengths_are_errors_not_panics() {
        let sk = SessionKey::new(vec![0u8; 32]);
        for (k, v) in [
            (vec![0u8; 0], vec![0u8; 64]),  // empty
            (vec![0u8; 48], vec![0u8; 64]), // shorter than hashSize 64
            (vec![0u8; 70], vec![0u8; 64]), // not a block multiple
            (vec![0u8; 64], vec![0u8; 16]), // value blob too short
        ] {
            let material = IntegrityMaterial {
                key_data_salt: &[0u8; 16],
                block_size: 16,
                hash_size: 64,
                hash: HashAlgorithm::Sha512,
                encrypted_hmac_key: &k,
                encrypted_hmac_value: &v,
            };
            assert!(matches!(
                verify(&sk, &material, b""),
                Err(Error::BadParameters(_))
            ));
        }
    }

    #[test]
    fn block_size_must_be_the_aes_block_size() {
        let sk = SessionKey::new(vec![0u8; 32]);
        let material = IntegrityMaterial {
            key_data_salt: &[0u8; 16],
            block_size: 32,
            hash_size: 64,
            hash: HashAlgorithm::Sha512,
            encrypted_hmac_key: &[0u8; 64],
            encrypted_hmac_value: &[0u8; 64],
        };
        assert!(matches!(
            verify(&sk, &material, b""),
            Err(Error::BadParameters(_))
        ));
    }

    #[test]
    fn hash_algorithm_parses_the_names_office_writes() {
        assert_eq!(HashAlgorithm::parse("SHA512"), Some(HashAlgorithm::Sha512));
        assert_eq!(HashAlgorithm::parse("SHA-1"), Some(HashAlgorithm::Sha1));
        assert_eq!(HashAlgorithm::parse("SHA384"), Some(HashAlgorithm::Sha384));
        assert_eq!(HashAlgorithm::parse("MD5"), None);
    }

    /// The one-line fact GH #12 turns on: a caller who passes no policy gets the
    /// fail-closed one. Everything else in that fix is downstream of this attribute.
    #[test]
    fn default_policy_requires_a_tag_wherever_the_format_defines_one() {
        assert_eq!(
            IntegrityPolicy::default(),
            IntegrityPolicy::RequireWhereDefined
        );
        assert_ne!(IntegrityPolicy::default(), IntegrityPolicy::VerifyIfPresent);
    }
}
