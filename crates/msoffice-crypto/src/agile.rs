//! Decrypt ECMA-376 agile encryption (Office 2010+).
//!
//! EncryptionInfo stream (after 8-byte header):
//!   XML containing keyData (outer salt) + p:encryptedKey (password params).
//!
//! Algorithm:
//!
//! ```text
//! 1. Parse XML → outer_salt, inner_salt, spinCount, keyBits, two hashes, encrypted fields
//! 2. Spin hash: H_0 = Hp(inner_salt + password_utf16le)
//!               H_i = Hp(LE32(i) + H_{i-1})  for i in 0..spinCount
//! 3. Derive block keys → decrypt encryptedVerifierHashInput/Value → verify password
//! 4. Decrypt encryptedKeyValue → obtain encryption_key
//! 5. Verify the dataIntegrity HMAC over the EncryptedPackage ciphertext
//! 6. Decrypt EncryptedPackage in 4096-byte segments with per-segment IVs
//! ```
//!
//! **`Hp` is `p:encryptedKey/@hashAlgorithm`, not `keyData/@hashAlgorithm`.** Steps 2, 3
//! and 4 run on the password encryptor's hash; steps 5 and 6 run on `<keyData>`'s. The
//! two elements carry their own copies, and although [MS-OFFCRYPTO] §2.3.4.10 obliges a
//! *writer* to give them the same value, a reader that assumes it silently runs the wrong
//! hash over whichever half it guessed on any file that does not. Three implementations
//! that keep them apart agree on the split:
//! herumi holds two `CipherParam`s (`include/crypto_util.hpp:327-343`) and passes
//! `encryptedKey.hashName` to the spin hash (`include/decode.hpp:91-93`) while passing
//! `keyData` to both `VerifyIntegrity` and `DecContent` (`include/decode.hpp:122, 133`);
//! msoffcrypto-tool splits them at its caller (`msoffcrypto/format/ooxml.py:69-96`);
//! office-crypto names the two fields `key_data_hash_algorithm` and
//! `password_hash_algorithm` (`src/crypto.rs:40-51`). LibreOffice cannot answer the
//! question — its handler merges every shared attribute into one flat struct
//! (`AgileEngine.cxx:99-131`, behaviour only) — so reading it for this is how you get it
//! backwards.
//!
//! Step 5 runs **before** step 6 on purpose. The HMAC covers ciphertext, so it can be
//! checked without decrypting anything, and doing it first means no unauthenticated
//! plaintext is ever produced — let alone returned. LibreOffice takes the other order
//! (`DocumentDecryption.cxx:210-218`: decrypt the whole package into the caller's
//! stream, then compare) and relies on the caller discarding it.
use crate::classify::local_name;
use crate::error::Error;
use crate::hash::{fit_iv, HashAlgorithm};
use crate::integrity::{self, IntegrityMaterial, IntegrityOutcome, IntegrityPolicy};
use crate::limits;
use crate::segments::Segments;
use crate::sensitive::{
    utf16le_password, DerivedKey, PasswordDigest, SessionKey, VerifierPlaintext,
};
use aes::{Aes128, Aes192, Aes256};
use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use cbc::cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};
use quick_xml::{events::Event, Reader};
use secure_gate::{ConstantTimeEq, RevealSecret, RevealSecretMut};

// Block key constants — verified against office-crypto reference implementation.
// The two dataIntegrity constants live in `integrity.rs` beside their only user.
pub(crate) const BLOCK_VERIFIER_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
pub(crate) const BLOCK_VERIFIER_HASH: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
pub(crate) const BLOCK_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

/// The `keyEncryptor/@uri` values [MS-OFFCRYPTO] §2.3.4.10 enumerates
/// (`ST_PasswordKeyEncryptorUri` in each of the two schemas it gives).
const URI_PASSWORD_KEY_ENCRYPTOR: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
const URI_CERTIFICATE_KEY_ENCRYPTOR: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";

/// Which `KeyEncryptor` the `<encryptedKey>` currently being read belongs to.
///
/// [MS-OFFCRYPTO] §2.3.4.10 defines **two** elements named `encryptedKey`, in two
/// namespaces: `CT_PasswordKeyEncryptor` (`p:`) and `CT_CertificateKeyEncryptor` (`c:`),
/// and a file may carry "exactly one PasswordKeyEncryptor" alongside "zero or more
/// CertificateKeyEncryptor elements". The two share the attribute name
/// `encryptedKeyValue` and wrap the *same* intermediate key — one under a
/// password-derived key, the other under an RSA public key — so a parser that matches on
/// the local name alone takes whichever comes last and reports a correct password as
/// wrong. See `a_certificate_key_encryptor_does_not_displace_the_password_one`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum KeyEncryptorKind {
    /// Inside a `<keyEncryptor>` whose `uri` names the password schema.
    Password,
    /// Inside a `<keyEncryptor>` whose `uri` names the certificate schema.
    Certificate,
    /// Not inside a `<keyEncryptor>`, or inside one carrying no `uri` — the attribute is
    /// optional in the schema (`<xs:attribute name="uri" type="xs:token"/>`, no
    /// `use="required"`), so the element's own shape has to settle it.
    Unstated,
}

/// The three AES key lengths, in bytes: `keyBits` 128, 192 and 256. Every cipher call in
/// this module dispatches on which of the three it was handed.
const AES_KEY_LENS: [usize; 3] = [16, 24, 32];
/// AES block / CBC IV length in bytes.
const AES_BLOCK_LEN: usize = 16;

pub(crate) struct AgileParams {
    /// `keyData/@saltValue` — the *package* salt. Seeds the per-segment content IVs and
    /// both dataIntegrity IVs. Distinct from `inner_salt`; crossing the two is silent.
    ///
    /// Its length is checked at parse time against **this element's own** `saltSize`.
    /// Nothing downstream would catch a wrong one: `derive_iv` hashes the salt at any
    /// length and yields a different-but-valid IV, so an over- or under-long `<keyData>`
    /// salt decrypts every segment to noise without a single error being raised.
    outer_salt: Vec<u8>,
    /// `p:encryptedKey/@saltValue` — the *password* salt. Both the spin-hash salt and
    /// the literal CBC IV for the three password-encryptor blobs.
    inner_salt: Vec<u8>,
    spin_count: u32,
    /// `p:encryptedKey/@keyBits` — the length of the key that *wraps* the session key,
    /// i.e. the truncation length applied to each block key derived from `H_final`.
    /// Its `<keyData>` sibling is a different number about a different key; see
    /// `key_data_key_bits`.
    key_bits: u32,
    encrypted_key_value: Vec<u8>,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    /// `<dataIntegrity>` — `Option` because the *parser* must be able to say "the file
    /// declared none", not because absence is a legitimate agile shape. It is not: the
    /// element is required in practice and `check_integrity` refuses a file without one
    /// under the default policy. See [`IntegrityPolicy::RequireWhereDefined`]'s doc for
    /// the [MS-OFFCRYPTO] Appendix A footnote `<22>` and the writer survey behind that.
    ///
    /// The `Option` earns its keep on exactly two paths: the explicitly chosen
    /// `VerifyIfPresent` and `Skip` policies, which report `NotDeclared` rather than
    /// erroring. Everything else treats `None` as a defect.
    ///
    /// Its three companion parameters come from `<keyData>`, never from
    /// `<p:encryptedKey>`. The two elements carry their own `blockSize` / `hashSize` /
    /// `hashAlgorithm`; the first two are per-element by [MS-OFFCRYPTO] §2.3.4.10's own
    /// schema and the third is one a writer MUST match, which a reader cannot rely on.
    /// Office writes all three identically, which is what hides the bug in round-trip
    /// tests against your own writer (herumi `include/encode.hpp:146-147`).
    data_integrity: Option<DataIntegrity>,
    /// `keyData/@hashAlgorithm`, resolved at parse time.
    ///
    /// Required, and required *whether or not* a `<dataIntegrity>` tag is present:
    /// every per-segment package IV is derived with it too. Leaving it optional is what
    /// let the two consumers of `<keyData>` disagree — the integrity check honouring
    /// the file while `decrypt_package` hardcoded SHA-512.
    key_data_hash: HashAlgorithm,
    /// `keyData/@blockSize` — the truncation length of every IV seeded by `<keyData>`.
    /// Pinned to the AES block size at parse time; this crate implements CBC only.
    key_data_block_size: u32,
    /// `keyData/@keyBits` — the length of the **package** key, which is a different
    /// statement from `key_bits` and is the one attribute of `<keyData>` this crate used
    /// to discard.
    ///
    /// The session key is the plaintext of `encryptedKeyValue`, so its length is that
    /// blob's length; `<keyData>` says independently how long the package cipher's key
    /// is. herumi reconciles the two by resizing — `normalizeKey(secretKey,
    /// keyData.keyBits / 8)` immediately before `DecContent`
    /// (`include/decode.hpp:131-133`), truncating or 0x36-padding to fit. This crate
    /// refuses instead, at parse time, where the same fact is available on the
    /// *ciphertext*: normalizing is a guess about which of the file's two declarations
    /// the writer meant, and a wrong guess is unobservable — the package decrypts under
    /// the wrong key, and with no `<dataIntegrity>` element nothing says so. That is the
    /// same silent-wrong-answer shape `require_aes_cbc` and `key_data_hash` already
    /// refuse. The one resize it does make is the cut herumi's call also makes for
    /// AES-192, whose 24-byte key travels padded to a 32-byte blob — the parse-time
    /// check admits exactly that shape, and `recover_session_key` makes the cut.
    key_data_key_bits: u32,
    /// `keyData/@hashSize` — only the dataIntegrity blobs are cut to it, so unlike its
    /// two siblings it is not needed to decrypt a file that declares no tag.
    ///
    /// `Option` therefore means "the file need not declare this", not "we will guess":
    /// when it *is* declared it is checked against `key_data_hash` at parse time like
    /// every other number here, and when it is absent `check_integrity` turns the
    /// absence into a `BadParameters` before anything can use it as a length.
    key_data_hash_size: Option<u32>,
    /// `p:encryptedKey/@hashAlgorithm`, resolved at parse time.
    ///
    /// The sibling of `key_data_hash` and deliberately a separate field: this one drives
    /// the spin hash, all three block keys and the verifier digest, and `key_data_hash`
    /// drives nothing on the password path. Collapsing the two into one is the bug
    /// LibreOffice's flat parser has (`AgileEngine.cxx:99-131`, behaviour only) and the
    /// one this crate already fixed once for the package IVs.
    password_hash: HashAlgorithm,
    /// `p:encryptedKey/@saltSize`, checked at parse time against `inner_salt`'s length.
    ///
    /// Not redundant with that length: `roundUp(saltSize, blockSize)` is how long the
    /// decrypted `encryptedVerifierHashInput` must be, and hashing anything else is a
    /// silently-wrong verifier (herumi `include/encode.hpp:168-171` pads the random
    /// `saltSize` bytes up to a block multiple and hashes the *padded* buffer).
    password_salt_size: u32,
    /// `p:encryptedKey/@blockSize` — the rounding modulus for both verifier blobs, and
    /// the block size of the cipher that decrypts them. Pinned to the AES block size at
    /// parse time, like its `<keyData>` counterpart.
    password_block_size: u32,
}

/// The two base64 blobs on `<dataIntegrity>`.
pub(crate) struct DataIntegrity {
    encrypted_hmac_key: Vec<u8>,
    encrypted_hmac_value: Vec<u8>,
}

impl AgileParams {
    /// The `<dataIntegrity>` element, if the file declared one.
    pub(crate) fn data_integrity(&self) -> Option<&DataIntegrity> {
        self.data_integrity.as_ref()
    }

    /// The CBC IV for the three `<p:encryptedKey>` blobs — [MS-OFFCRYPTO] §2.3.4.12
    /// case 2: "If a blockKey is not provided, let IV be equal to the following value:
    /// KeySalt". §2.3.4.13 uses that IV for `encryptedVerifierHashInput`,
    /// `encryptedVerifierHashValue` and `encryptedKeyValue` alike, which is why all
    /// three read it from here rather than reaching for `inner_salt` themselves.
    ///
    /// The salt is put through [`fit_iv`] rather than used raw, because §2.3.4.12's
    /// third step applies to this case as much as to the hashed one: an IV shorter than
    /// `blockSize` is 0x36-padded and a longer one truncated. `saltSize` is
    /// 1..=65 536 ([`limits::AGILE_SALT_SIZE`]) and `blockSize` is pinned to the AES
    /// block, so a conforming file may declare a salt of any length and used to be
    /// refused by `check_cbc_lengths` — "AES-CBC needs a 16-byte IV" — one frame after
    /// the parser had accepted it. Office writes 16, so this is the identity for every
    /// real file; the `blockSize` is `<p:encryptedKey>`'s, the element whose blobs it
    /// decrypts.
    pub(crate) fn password_blob_iv(&self) -> Vec<u8> {
        fit_iv(&self.inner_salt, self.password_block_size as usize)
    }

    /// Everything `integrity::verify` — and `integrity::generate`'s byte-identity test —
    /// needs from `<keyData>` and `<dataIntegrity>`.
    ///
    /// Takes the element rather than looking it up, so there is no "present but absent"
    /// arm to write: the caller has already decided what a missing element means.
    ///
    /// `hashSize` is the one `<keyData>` attribute only the tag needs — it cuts the two
    /// decrypted blobs back to a digest length — so a file without `<dataIntegrity>`
    /// still decrypts even if `<keyData>` omits it. `blockSize` and `hashAlgorithm` are
    /// not in that position: `decrypt_package` needs both, so they are required at
    /// parse time and are already resolved here.
    pub(crate) fn integrity_material<'a>(
        &'a self,
        di: &'a DataIntegrity,
    ) -> Result<IntegrityMaterial<'a>, Error> {
        let hash_size = self.key_data_hash_size.ok_or_else(|| {
            Error::BadParameters(
                "<dataIntegrity> is present but keyData/@hashSize is missing".into(),
            )
        })?;
        Ok(IntegrityMaterial {
            key_data_salt: &self.outer_salt,
            block_size: self.key_data_block_size,
            hash_size,
            hash: self.key_data_hash,
            encrypted_hmac_key: &di.encrypted_hmac_key,
            encrypted_hmac_value: &di.encrypted_hmac_value,
        })
    }
}

/// Steps 1 through 4: the spin hash, the password check, and the session key out of
/// `encryptedKeyValue`.
///
/// Extracted from `decrypt` so a test can recover a real fixture's session key through
/// the same code the decrypt path runs — the encrypt side's byte-identity tests need
/// Office's own session key to reproduce Office's own blobs — rather than re-deriving it
/// beside the thing under test. A wrong password is refused here, before any key
/// exists: `verify_password` runs first, so what comes back is never a key the caller
/// merely hopes is right.
pub(crate) fn recover_session_key(
    params: &AgileParams,
    password: &str,
) -> Result<SessionKey, Error> {
    let h_final = spin_hash(
        params.password_hash,
        password,
        &params.inner_salt,
        params.spin_count,
    );

    // Verify password before touching the session key -- and before the integrity
    // check, which would otherwise report a wrong password as corruption.
    verify_password(params, &h_final)?;

    let key_deriv = derive_block_key(
        params.password_hash,
        &h_final,
        &BLOCK_KEY_VALUE,
        params.key_bits,
    )?;
    let mut encryption_key = key_deriv.with_secret(|k| {
        aes_cbc_decrypt(&params.encrypted_key_value, k, &params.password_blob_iv())
    })?;
    // The blob is the session key padded up to a block multiple (see the length check in
    // `parse_encryption_info`); the key is the first `keyData/@keyBits / 8` bytes of it.
    // A no-op for AES-128 and AES-256, and the eight bytes that make AES-192 open —
    // herumi's `normalizeKey` (`include/decode.hpp:131-133`), LibreOffice's
    // `mKey.resize(nKeySize, 0)` (`AgileEngine.cxx:368`, behaviour only).
    encryption_key.truncate((params.key_data_key_bits / 8) as usize);
    Ok(SessionKey::new(encryption_key))
}

/// Decrypt an agile `EncryptedPackage` given the XML half of `EncryptionInfo`.
///
/// `xml_data` is the EncryptionInfo stream content with the 8-byte version header
/// stripped; `encrypted_package` is the **whole** EncryptedPackage stream, starting at
/// its 8-byte little-endian plaintext-size prefix (the HMAC covers that prefix).
///
/// On error, no plaintext is produced. The HMAC is checked before decryption.
///
/// # Errors
///
/// [`Error::XmlParse`] if the XML is malformed, a required attribute is
/// missing, or the file's only key encryptors are certificate ones;
/// [`Error::BadParameters`] if a declared length, spin count or sibling field
/// is out of range, or the file declares more than one `PasswordKeyEncryptor`; [`Error::UnsupportedAlgorithm`] if a named
/// cipher or hash is not implemented; [`Error::WrongPassword`] if the verifier
/// does not match; [`Error::IntegrityElementMissing`],
/// [`Error::IntegrityCheckFailed`] or [`Error::BadParameters`] from
/// the HMAC policy; [`Error::CipherError`] if an AES-CBC step rejects a block;
/// [`Error::MissingStream`] if the package prefix is truncated.
pub(crate) fn decrypt(
    xml_data: &[u8],
    encrypted_package: &[u8],
    password: &str,
    policy: IntegrityPolicy,
) -> Result<(Vec<u8>, IntegrityOutcome), Error> {
    let params = parse_encryption_info(xml_data)?;

    // `p:encryptedKey/@hashAlgorithm`, never `<keyData>`'s: this is step 2 of the module
    // header, and every consumer of `h_final` below (both verifier block keys, the
    // verifier digest, the session-key block key) runs on `password_hash` too. Passing
    // the `<keyData>` hash here was the last surviving instance of the conflation the
    // rest of this module was rewritten to prevent — invisible in every fixture and in
    // every round trip against a writer that sets both elements from one pair (herumi
    // `include/encode.hpp:146-147`), and a `WrongPassword` for a correct password on any
    // file that does not.
    let encryption_key = recover_session_key(&params, password)?;

    let outcome = check_integrity(&params, &encryption_key, encrypted_package, policy)?;

    let plaintext = decrypt_package(&encryption_key, &params, encrypted_package)?;
    Ok((plaintext, outcome))
}

/// Apply `policy` to whatever `<dataIntegrity>` the file did or did not declare.
///
/// Returns `Err` on a genuine mismatch, on a missing element under a policy that
/// requires one (the default included), and on a declared-but-unusable parameter set.
/// Never returns `Ok` for a tag that failed.
fn check_integrity(
    params: &AgileParams,
    session_key: &SessionKey,
    encrypted_package: &[u8],
    policy: IntegrityPolicy,
) -> Result<IntegrityOutcome, Error> {
    let Some(di) = params.data_integrity() else {
        // The agile half of `IntegrityPolicy`'s contract; the standard half is the
        // version dispatch in `lib.rs`. Absence is a defect here, not a shape: agile
        // encryption always writes the element (see `RequireWhereDefined`'s doc for the
        // spec footnote and the writer survey), so a file without it was truncated,
        // hand-written, or stripped by someone who wanted the HMAC not to be checked.
        //
        // Matched exhaustively on purpose. The wildcard this replaces (`_ => Ok(..)`)
        // is how a policy added later would be absorbed into the permissive answer,
        // quietly undoing the fail-closed default GH #12 landed. A new variant must be
        // a compile error here, not a downgrade.
        return match policy {
            IntegrityPolicy::Require | IntegrityPolicy::RequireWhereDefined => {
                Err(Error::IntegrityElementMissing)
            }
            // `NotDeclared` for `Skip` too, not `Skipped`: nothing was skipped, because
            // nothing was there. Stated rather than inherited from arm ordering.
            IntegrityPolicy::VerifyIfPresent | IntegrityPolicy::Skip => {
                Ok(IntegrityOutcome::NotDeclared)
            }
        };
    };

    if policy == IntegrityPolicy::Skip {
        return Ok(IntegrityOutcome::Skipped);
    }

    let material = params.integrity_material(di)?;
    integrity::verify(session_key, &material, encrypted_package)?;
    Ok(IntegrityOutcome::Verified)
}

/// `Hp(decrypted encryptedVerifierHashInput)` must equal the decrypted
/// `encryptedVerifierHashValue`, cut to `Hp`'s digest length.
///
/// Three lengths are involved and none of them may be a literal:
///
/// * **What is hashed** is the *whole* decrypted `encryptedVerifierHashInput`, not a
///   `saltSize`-long prefix of it. The writer draws `saltSize` random bytes, pads them up
///   to a `blockSize` multiple, and hashes the padded buffer (herumi
///   `include/encode.hpp:168-171`; msoffcrypto-tool
///   `msoffcrypto/method/ecma376_agile.py:327-338`). Slicing to `saltSize` would agree
///   with the writer only while `saltSize` happens to be a whole number of blocks.
/// * **How long that blob must be** is therefore `roundUp(saltSize, blockSize)`, checked
///   rather than assumed. It is 16 for every real file, which is where the literal `16`
///   this function used to carry came from — but the blob is attacker-controlled base64,
///   and without the check a hostile file could hand us a megabyte to hash.
/// * **How many bytes are compared** is `Hp`'s digest length, and the decrypted
///   `encryptedVerifierHashValue` is `roundUp(hashSize, blockSize)` bytes — longer than
///   the digest for SHA-1 (20 → 32) and exactly equal for the other three. herumi
///   truncates the value before comparing (`include/decode.hpp:99`
///   `.substr(0, hashedVerifier.size())`) and LibreOffice compares exactly `hashSize`
///   bytes (`AgileEngine.cxx:338, 356`, behaviour only). **msoffcrypto-tool does not**
///   (`ecma376_agile.py:466` compares against the whole blob), which is why its own
///   `verify_password` returns false for a correct password on any SHA-1 agile file. Do
///   not port that one.
///
/// Exposed to the crate so `agile_encrypt`'s tests can put a generated encryptor through
/// the real verification path — write, parse, verify — rather than re-deriving the check
/// beside the generator, which would prove only that the test agrees with itself.
pub(crate) fn verify_password(params: &AgileParams, h_final: &PasswordDigest) -> Result<(), Error> {
    let hash = params.password_hash;
    let block_size = params.password_block_size as usize;
    let digest_len = hash.digest_len();
    // `saltSize`, `hashSize` and `blockSize` are all validated at parse time — the first
    // against `inner_salt`'s own length, the second against the named algorithm, the
    // third pinned to the AES block size — so the two lengths below are derived from
    // checked numbers rather than from raw file input. That pin is also what makes the
    // `div_ceil` total: `blockSize` is the one operand that could divide by zero, and no
    // `AgileParams` with a zero `password_block_size` can come out of a file.
    let input_len = (params.password_salt_size as usize).div_ceil(block_size) * block_size;
    let value_len = digest_len.div_ceil(block_size) * block_size;

    // Checked on the *ciphertext*, before either blob is decrypted: AES-CBC preserves
    // length, so this is the same test one frame earlier and it bounds the work a
    // hostile file can buy. An under-length blob used to panic on the slice instead.
    if params.encrypted_verifier_hash_input.len() != input_len
        || params.encrypted_verifier_hash_value.len() != value_len
    {
        return Err(Error::BadParameters(format!(
            "encryptedVerifierHashInput/Value are {}/{} bytes; {} with saltSize {} and \
             blockSize {} requires {}/{}",
            params.encrypted_verifier_hash_input.len(),
            params.encrypted_verifier_hash_value.len(),
            hash.name(),
            params.password_salt_size,
            block_size,
            input_len,
            value_len
        )));
    }

    let verifier_input_key =
        derive_block_key(hash, h_final, &BLOCK_VERIFIER_INPUT, params.key_bits)?;
    let verifier_input = verifier_input_key.with_secret(|k| {
        aes_cbc_decrypt(
            &params.encrypted_verifier_hash_input,
            k,
            &params.password_blob_iv(),
        )
    })?;
    let verifier_input = VerifierPlaintext::new(verifier_input);

    let verifier_hash_key = derive_block_key(hash, h_final, &BLOCK_VERIFIER_HASH, params.key_bits)?;
    let verifier_hash = verifier_hash_key.with_secret(|k| {
        aes_cbc_decrypt(
            &params.encrypted_verifier_hash_value,
            k,
            &params.password_blob_iv(),
        )
    })?;
    let verifier_hash = VerifierPlaintext::new(verifier_hash);

    // Both operands stay inside their closures; only the boolean escapes. Constant-time
    // for the same reason the dataIntegrity comparison is -- see the ct-eq note in
    // Cargo.toml. The channel here is much weaker (both sides derive from the password
    // being guessed), but two comparisons that look identical should not have been
    // reasoned about differently.
    // **Over `saltSize` bytes, not the whole decrypted blob**, and the difference is only
    // visible when `saltSize` is not a multiple of `blockSize`.
    //
    // [MS-OFFCRYPTO] §2.3.4.13, `encryptedVerifierHashValue` step 1: "Obtain the hash
    // value of the random array of bytes generated in step 1 of the steps for
    // encryptedVerifierHashInput" — and that step 1 is "Generate a random array of bytes
    // with the number of bytes used specified by the saltSize attribute". The `0x00`
    // padding to a block multiple arrives in step 3, at *encryption* time, after the hash
    // has been taken. So the hashed array is `saltSize` bytes and the blob it travels in
    // is `roundUp(saltSize, blockSize)`.
    //
    // This digested the whole blob until 2026-09-21, and `agile_encrypt` hashed the
    // padded buffer to match. The two agreed with each other and disagreed with the
    // format, so every round trip in this crate passed and every file it wrote at a
    // non-block-multiple `saltSize` was refused by real Word 16 with `0x800A1520` — the
    // wrong-password code, on a correct password. Found by running the artifact set past
    // Word rather than by any test here, which is the argument for that gate.
    //
    // The read half was the shipped defect: a *conforming* file from another writer with
    // `saltSize = 8` would have been refused as a wrong password. That is the same class
    // as the `spinCount` ceiling — an owner locked out of their own document by our
    // choice, not the format's.
    //
    // The slice is in range because `input_len` above is `roundUp(salt_len, block_size)`,
    // which is `>= salt_len`, and the blob's length was checked equal to it.
    let salt_len = params.password_salt_size as usize;
    let matches = verifier_input.with_secret(|vi| {
        // Wrapped like the value it is compared against: when the password is right,
        // the two are the same bytes. `digest` allocates exactly `digest_len`, so the
        // wrapper takes it by move with nothing left behind.
        let computed = VerifierPlaintext::new(hash.digest(&vi[..salt_len]));
        computed
            .with_secret(|c| verifier_hash.with_secret(|vh| c.as_slice().ct_eq(&vh[..digest_len])))
    });
    if !matches {
        return Err(Error::WrongPassword);
    }
    Ok(())
}

/// Decrypt the EncryptedPackage segments using the session encryption key.
///
/// Each segment gets its own IV, `H(keyData.saltValue || LE32(segment_index))`
/// truncated to `keyData/@blockSize`. Both the hash and the truncation length come from
/// `<keyData>`, which is why this takes the whole parameter set rather than just the
/// salt: hardcoding SHA-512 and 16 here made a file whose `keyData/@hashAlgorithm` is
/// not SHA-512 decrypt under the wrong IVs *after* `check_integrity` had already
/// reported `Verified` under the right ones. herumi derives it from `keyData.hashName`
/// (`include/crypto_util.hpp:432-441`, via `DecContent`) and LibreOffice from
/// `mInfo.hashAlgorithm` (`AgileEngine.cxx:497`, behaviour only) — both read it from
/// the file, and so does the integrity path in `integrity.rs`.
fn decrypt_package(
    encryption_key: &SessionKey,
    params: &AgileParams,
    encrypted_package: &[u8],
) -> Result<Vec<u8>, Error> {
    if encrypted_package.len() < 8 {
        return Err(Error::MissingStream("EncryptedPackage too short"));
    }

    // The session key must be the length `<keyData>` says the package key is. herumi
    // makes the same statement one line before `DecContent` and reconciles a mismatch by
    // resizing (`include/decode.hpp:131-133`, `normalizeKey(secretKey, keyData.keyBits /
    // 8)`); this crate refuses, because normalizing is a guess about which of the file's
    // two `keyBits` declarations the writer meant and a wrong guess is unobservable —
    // there may be no `<dataIntegrity>` element to fail.
    //
    // The parser already checked the same fact on the *ciphertext*, where AES-CBC's
    // length preservation makes it available before any password work. This is the repeat
    // at the point of use, for the same reason `derive_block_key` repeats the parser's
    // `keyBits` / digest pairing: one comparison against a value already in a register,
    // and a future caller reaching this function another way still hits it.
    let key_len = (params.key_data_key_bits / 8) as usize;
    let actual = encryption_key.with_secret(|k| k.len());
    if actual != key_len {
        return Err(Error::BadParameters(format!(
            "keyData/@keyBits is {} ({key_len} bytes) but the session key recovered \
             from encryptedKeyValue is {actual} bytes",
            params.key_data_key_bits
        )));
    }

    let declared_size = u64::from_le_bytes(encrypted_package[..8].try_into().unwrap());
    let data = &encrypted_package[8..];

    // The ceiling, repeated at the point of allocation. `cfb_reader` already refused a
    // larger stream, so this is unreachable through `decrypt_ooxml` — it is here because
    // `decrypt_package` takes a `&[u8]` and a later caller (the streaming API, a test, a
    // second entry point) reaches it without passing that reader. GH #10 pattern 3: one
    // named ceiling aliased per allocation path, never a second number.
    if data.len() > limits::PAYLOAD_CEILING {
        return Err(Error::BadParameters(format!(
            "EncryptedPackage carries {} bytes of ciphertext; this crate decrypts at most \
             {} (see limits::PAYLOAD_CEILING)",
            data.len(),
            limits::PAYLOAD_CEILING
        )));
    }

    // The declared plaintext size is the file's own claim about its payload and it is
    // checked, not trusted. AES-CBC preserves length and each 4096-byte segment is padded
    // up to a 16-byte multiple, so a well-formed file always satisfies
    // `declared <= ciphertext length`, with under 16 bytes of slack per segment. Without
    // this the only use of the number is `output.truncate(..)` below, and `Vec::truncate`
    // past the current length is a silent no-op — so `plaintext_size = u64::MAX` was
    // *accepted* and the caller quietly received the writer's final-segment padding
    // appended to their ZIP.
    //
    // Compared as `u64`, before the `as usize` cast: on a 32-bit target that cast
    // truncates, and a declared size of `2^32 + 16` would otherwise pass as 16.
    if declared_size > data.len() as u64 {
        return Err(Error::BadParameters(format!(
            "EncryptedPackage declares {declared_size} plaintext bytes but carries only {} \
             bytes of ciphertext",
            data.len()
        )));
    }
    let plaintext_size = declared_size as usize;

    // One segmentation and one IV derivation, shared with the encrypt path (plan D4) —
    // see `crate::segments` for why this is a type and not a `chunks(4096)` here.
    let segments = Segments::new(
        data,
        params.key_data_hash,
        &params.outer_salt,
        params.key_data_block_size as usize,
    )?;

    let mut output = Vec::with_capacity(data.len());
    for segment in segments {
        let segment = segment?;
        let dec =
            encryption_key.with_secret(|k| aes_cbc_decrypt(&segment.padded(), k, &segment.iv))?;
        output.extend_from_slice(&dec);
    }

    output.truncate(plaintext_size);
    Ok(output)
}

/// Iterated spin hash under `p:encryptedKey/@hashAlgorithm`.
///
/// The counter is **prepended** to the previous digest, for every algorithm — that is a
/// property of the format, not of SHA-512 (herumi `include/decode.hpp:91-93` via
/// `hashPassword`; msoffcrypto-tool `ecma376_agile.py:184-188`; LibreOffice's
/// `IterCount::PREPEND` at `AgileEngine.cxx:309-314`, behaviour only).
pub(crate) fn spin_hash(
    hash: HashAlgorithm,
    password: &str,
    salt: &[u8],
    spin_count: u32,
) -> PasswordDigest {
    // H_0 = H(salt + password_utf16le)
    //
    // The UTF-16LE re-encoding is wrapped and scoped to this statement: it is the
    // password verbatim, it is used exactly once, and the block that held it is wiped
    // before the spin loop starts rather than at the end of the function. Built with
    // `sensitive::utf16le_password` and not `collect()` — see that function for the
    // reallocation the collect abandoned, and `docs/design/heap-residue.md` for why an
    // abandoned block is unreachable to every wrapper in the crate.
    let len = hash.digest_len();
    let mut current = PasswordDigest::new_with(len, |slot| {
        utf16le_password(password).with_secret(|pw| hash.digest_two_into(salt, pw, slot));
    });

    // H_i = H(LE32(i) + H_{i-1})  — counter PREPENDED
    //
    // Two wrapped buffers, alternated: each round reads one and writes the other, then
    // the two swap (the wrappers swap, not their bytes). Until rc.5 each round
    // allocated a fresh `Vec` and dropped the last one unwiped -- up to `spinCount`
    // abandoned digests per decrypt, on the reading that intermediate states are only
    // worth inverting. That reading missed the forward direction: any `H_i` reaches
    // `H_final` by `spinCount - i` public hashes, and the last abandoned one was a
    // single hash from it. Now both buffers are wiped on drop, and the loop allocates
    // nothing.
    let mut next = PasswordDigest::new_with(len, |_| {});
    for i in 0u32..spin_count {
        current.with_secret(|prev| {
            next.with_secret_mut(|out| hash.digest_two_into(&i.to_le_bytes(), prev, out));
        });
        std::mem::swap(&mut current, &mut next);
    }
    current
}

/// Derive a block key: `Hp(H_final || block_key)` truncated to `key_bits / 8` bytes.
///
/// **Truncation only, never padding — a deliberate divergence from the spec's prose.**
/// [MS-OFFCRYPTO] §2.3.4.11 ends: "If the size of the resulting Hfinal is smaller than
/// that of PasswordKeyEncryptor.keyBits, the key MUST be padded by appending bytes with
/// a value of 0x36. If the hash value is larger in size than
/// PasswordKeyEncryptor.keyBits, the key is obtained by truncating the hash value." The
/// truncating half is implemented; the padding half is refused, and the reason is that
/// the pad would silently manufacture a key no writer ever used.
///
/// `key_bits / 8` can exceed the digest length once the hash is the file's choice —
/// SHA-1 gives 20 bytes and `keyBits="256"` asks for 32 — and the four references
/// genuinely disagree about that case, which is itself evidence that no file exercises
/// it. herumi 0x36-pads via its
/// shared `normalizeKey` (`include/crypto_util.hpp:38-41`, applied by `generateKey` at
/// `:447-452`); msoffcrypto-tool truncates and then dies inside `algorithms.AES` with
/// `Invalid key size (160)` (`ecma376_agile.py:201`); LibreOffice truncates with no guard
/// and admits the gap in its own comment (`AgileEngine.cxx:277-279, 499`); office-crypto
/// truncates but never reaches the case, refusing every non-SHA-512 agile file outright
/// (`src/lib.rs:29`). In Rust the naive form was worse than any of those — `digest[..32]`
/// on a 20-byte digest is a panic, i.e. a vulnerability under CLAUDE.md's first design
/// value — so the combination is refused instead, at parse time, matching this crate's
/// own precedent for the `derive_iv` pad branch. No shipping writer emits it: LibreOffice
/// accepts four tuples and herumi's encoder emits two, and none of the six reaches the
/// pad.
///
/// The check is repeated here rather than left to the parser. This function is reachable
/// from three call sites with a `key_bits` the parser vouched for, and the cost of not
/// trusting that is one comparison against a value already in a register. Repeated, not
/// re-expressed: both this and the parse-time check ask
/// [`HashAlgorithm::can_carry_key_bits`], so the duplication is of the *call* and not of
/// the bound, and the two cannot come to disagree about which pairs are refusable.
pub(crate) fn derive_block_key(
    hash: HashAlgorithm,
    h_final: &PasswordDigest,
    block_key: &[u8; 8],
    key_bits: u32,
) -> Result<DerivedKey, Error> {
    let key_len = (key_bits / 8) as usize;
    if !hash.can_carry_key_bits(key_bits) {
        return Err(unusable_key_bits(key_bits, hash));
    }
    // `key_len <= hash.digest_len()` now holds, which is exactly `digest_two_into`'s
    // contract — the check above is the one that discharges it.
    //
    // The digest is cut straight into the wrapper's own slot and never exists as a
    // `Vec`. Written the obvious way — `let digest = hash.digest_two(hf, block_key);
    // DerivedKey::new(digest[..key_len].to_vec())` — it existed twice over: a heap
    // `Vec` holding the **whole** digest of `H_final`, dropped unwiped, plus a second
    // allocation for the prefix the wrapper kept. Both went to the allocator carrying
    // block-key material. See `docs/design/heap-residue.md`.
    Ok(h_final.with_secret(|hf| {
        DerivedKey::new_with(key_len, |slot| hash.digest_two_into(hf, block_key, slot))
    }))
}

/// The one message for "this file's `keyBits` asks for more bytes than its own
/// `hashAlgorithm` can produce", shared by the parse-time check and `derive_block_key`
/// so the two cannot drift apart.
fn unusable_key_bits(key_bits: u32, hash: HashAlgorithm) -> Error {
    Error::BadParameters(format!(
        "p:encryptedKey/@keyBits is {} ({} bytes) but hashAlgorithm {} yields only {}; \
         no writer produces this combination and this crate refuses to pad the digest",
        key_bits,
        key_bits / 8,
        hash.name(),
        hash.digest_len()
    ))
}

/// Build [`Error::UnsupportedAlgorithm`] from a name the *file* chose.
///
/// The name is bounded here and nowhere else: it is XML attribute text limited only by
/// `limits::ENCRYPTION_INFO_READ_CAP`, so a 1 MiB `hashAlgorithm="AAAA…"` would otherwise
/// land whole in an error string. 32 characters is generous — the longest legitimate
/// value is `SHA-512` — and the cut is on a char boundary, not a byte one.
fn unsupported_algorithm(what: &'static str, name: &str) -> Error {
    const MAX_CHARS: usize = 32;
    let mut shown: String = name.chars().take(MAX_CHARS).collect();
    if name.chars().nth(MAX_CHARS).is_some() {
        shown.push('…');
    }
    Error::UnsupportedAlgorithm { what, name: shown }
}

/// AES-CBC decrypt with NoPadding, AES-128/192/256 by key length. `ciphertext` must be a
/// multiple of 16 bytes.
///
/// The key and IV lengths are checked rather than asserted. `GenericArray::from_slice`
/// panics on a length mismatch, and the key length is file-derived: it is `keyBits / 8`,
/// the length [`derive_block_key`] cuts the named digest to. The IV no longer arrives raw
/// — the password-encryptor blobs come through [`AgileParams::password_blob_iv`] and the
/// package and `dataIntegrity` IVs through [`crate::hash::derive_iv`], and both fit their
/// bytes to a `blockSize` the parser has pinned to 16. The check stays because this is a
/// seam: a future caller reaching it another way must get an error, not a crash in the
/// caller's process.
///
/// **AES-128, AES-192 and AES-256, chosen by key length** — since GH #13. Before it this
/// was `aes256_cbc_decrypt`, and three of the four agile tuples in the wild, including
/// Word 2010's own AES-128/SHA-1 default, were refused one frame after the parser had
/// accepted them.
pub(crate) fn aes_cbc_decrypt(ciphertext: &[u8], key: &[u8], iv: &[u8]) -> Result<Vec<u8>, Error> {
    use aes::cipher::generic_array::GenericArray;
    check_cbc_lengths(key, iv)?;
    let iv = GenericArray::from_slice(iv);
    let mut out = vec![0u8; ciphertext.len()];
    // The key length *is* the cipher: `keyBits` 128, 192 and 256 name AES-128, AES-192 and
    // AES-256, and `check_cbc_lengths` has already refused anything else, which is what
    // makes the third arm total rather than a guess.
    macro_rules! run {
        ($cipher:ty) => {
            cbc::Decryptor::<$cipher>::new(GenericArray::from_slice(key), iv)
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

/// AES-CBC encrypt with NoPadding, AES-128/192/256 by key length. `plaintext` must be a
/// multiple of 16 bytes.
///
/// The twin of [`aes_cbc_decrypt`], deliberately next to it: the two share every
/// length rule, and the failure mode of letting them drift is a file this crate writes and
/// cannot read. [`check_cbc_lengths`] is the shared half, so a change to one cannot miss the
/// other.
///
/// The caller pads. `NoPadding` here is not an oversight — the format's own padding rules
/// differ by blob (`roundUp(saltSize, blockSize)` for the verifier input,
/// `roundUp(hashSize, blockSize)` for its hash, zero-fill for the final package segment),
/// and only the caller knows which applies. A block-cipher padding mode would append
/// bytes the reader is not expecting and lengthen every blob past what `<keyData>` says.
///
pub(crate) fn aes_cbc_encrypt(plaintext: &[u8], key: &[u8], iv: &[u8]) -> Result<Vec<u8>, Error> {
    use aes::cipher::generic_array::GenericArray;
    use cbc::cipher::BlockEncryptMut;
    check_cbc_lengths(key, iv)?;
    let iv = GenericArray::from_slice(iv);
    let mut out = vec![0u8; plaintext.len()];
    macro_rules! run {
        ($cipher:ty) => {
            cbc::Encryptor::<$cipher>::new(GenericArray::from_slice(key), iv)
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

/// The key and IV length rules both CBC directions share.
///
/// Checked rather than asserted: `GenericArray::from_slice` panics on a length mismatch,
/// and on the *decrypt* side the key length is file-derived — `keyBits / 8` is the length
/// [`derive_block_key`] cuts the named digest to. Every in-crate IV now arrives already
/// fitted to a `blockSize` the parser pinned to 16, so the IV arm guards the seam rather
/// than a live path. A key of any length AES does not define must produce an error rather
/// than a crash in the caller's process.
fn check_cbc_lengths(key: &[u8], iv: &[u8]) -> Result<(), Error> {
    if !AES_KEY_LENS.contains(&key.len()) {
        return Err(Error::BadParameters(format!(
            "AES-CBC takes a 16-, 24- or 32-byte key; this file's parameters produced {}",
            key.len()
        )));
    }
    if iv.len() != AES_BLOCK_LEN {
        return Err(Error::BadParameters(format!(
            "AES-CBC needs a {}-byte IV; this file's parameters produced {}",
            AES_BLOCK_LEN,
            iv.len()
        )));
    }
    Ok(())
}

/// Parse the Agile EncryptionInfo XML and extract all needed parameters.
/// `pub(crate)` so `encryption_info`'s tests can hand it what the writer produced. The
/// writer is this function's inverse, and the cheapest statement of that is the parser
/// accepting a synthetic document that no fixture supplies.
pub(crate) fn parse_encryption_info(xml_data: &[u8]) -> Result<AgileParams, Error> {
    let mut reader = Reader::from_reader(xml_data);
    reader.config_mut().trim_text(true);

    let mut outer_salt: Option<Vec<u8>> = None;
    let mut inner_salt: Option<Vec<u8>> = None;
    let mut spin_count: Option<u32> = None;
    let mut key_bits: Option<u32> = None;
    let mut encrypted_key_value: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_value: Option<Vec<u8>> = None;
    let mut key_data_block_size: Option<u32> = None;
    let mut key_data_hash_size: Option<u32> = None;
    let mut key_data_key_bits: Option<u32> = None;
    let mut key_data_salt_size: Option<u32> = None;
    let mut key_data_hash_algorithm: Option<String> = None;
    let mut key_data_cipher_algorithm: Option<String> = None;
    let mut key_data_cipher_chaining: Option<String> = None;
    let mut password_hash_algorithm: Option<String> = None;
    let mut password_cipher_algorithm: Option<String> = None;
    let mut password_cipher_chaining: Option<String> = None;
    let mut password_hash_size: Option<u32> = None;
    let mut password_salt_size: Option<u32> = None;
    let mut password_block_size: Option<u32> = None;
    let mut saw_data_integrity = false;
    let mut encrypted_hmac_key: Option<Vec<u8>> = None;
    let mut encrypted_hmac_value: Option<Vec<u8>> = None;
    // Which `<keyEncryptor>` we are inside, and what the sequence has held so far.
    // `<keyEncryptor>` elements do not nest, so one flag is the whole state.
    let mut current_encryptor = KeyEncryptorKind::Unstated;
    let mut saw_password_encryptor = false;
    let mut saw_certificate_encryptor = false;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                // Strip namespace prefix for matching
                let name_owned = e.name().as_ref().to_vec();
                let local = local_name(&name_owned);

                match local {
                    // The wrapper whose `uri` says which of §2.3.4.10's two
                    // `encryptedKey` elements the child is. Read before the child, which
                    // is the only ordering XML allows.
                    b"keyEncryptor" => {
                        current_encryptor = KeyEncryptorKind::Unstated;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"uri" {
                                let uri = decode_str_attr(&attr)?;
                                current_encryptor = if uri == URI_PASSWORD_KEY_ENCRYPTOR {
                                    KeyEncryptorKind::Password
                                } else if uri == URI_CERTIFICATE_KEY_ENCRYPTOR {
                                    KeyEncryptorKind::Certificate
                                } else {
                                    // An unrecognised uri is not a refusal: the schema
                                    // types it `xs:token` and says extensibility is the
                                    // point ("To ensure extensibility, arbitrary
                                    // elements can be defined to encrypt the
                                    // intermediate key"). The child element decides.
                                    KeyEncryptorKind::Unstated
                                };
                            }
                        }
                    }
                    // <keyData> governs the package and the dataIntegrity tag. Its
                    // blockSize / hashSize / hashAlgorithm are deliberately kept apart
                    // from <p:encryptedKey>'s identically-named attributes: the spec
                    // gives each element its own set and only Office's habit of writing
                    // them identically makes a flat "last one wins" parser look correct
                    // (LibreOffice AgileEngine.cxx:92-172, behaviour only).
                    b"keyData" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"saltValue" => {
                                    outer_salt = Some(decode_b64_attr(&attr)?);
                                }
                                b"saltSize" => {
                                    key_data_salt_size = Some(parse_u32_attr(&attr)?);
                                }
                                b"blockSize" => {
                                    key_data_block_size = Some(parse_u32_attr(&attr)?);
                                }
                                b"hashSize" => {
                                    key_data_hash_size = Some(parse_u32_attr(&attr)?);
                                }
                                b"keyBits" => {
                                    key_data_key_bits = Some(parse_u32_attr(&attr)?);
                                }
                                b"hashAlgorithm" => {
                                    key_data_hash_algorithm = Some(decode_str_attr(&attr)?);
                                }
                                b"cipherAlgorithm" => {
                                    key_data_cipher_algorithm = Some(decode_str_attr(&attr)?);
                                }
                                b"cipherChaining" => {
                                    key_data_cipher_chaining = Some(decode_str_attr(&attr)?);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"dataIntegrity" => {
                        saw_data_integrity = true;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"encryptedHmacKey" => {
                                    encrypted_hmac_key = Some(decode_b64_attr(&attr)?);
                                }
                                b"encryptedHmacValue" => {
                                    encrypted_hmac_value = Some(decode_b64_attr(&attr)?);
                                }
                                _ => {}
                            }
                        }
                    }
                    // <p:encryptedKey> governs the password KDF: the spin hash, the three
                    // block keys and the verifier digest. Its hashAlgorithm / hashSize /
                    // blockSize / saltSize are read here and kept in their own fields,
                    // never merged with <keyData>'s. §2.3.4.10 makes matching
                    // `hashAlgorithm` a writer's obligation, not a reader's assumption,
                    // and Office's habit of writing every shared attribute identically
                    // is exactly what makes a merged parser look correct.
                    //
                    // `<c:encryptedKey>` has this same local name and is skipped: see
                    // [`KeyEncryptorKind`]. The certificate element is identified by its
                    // wrapper's `uri` where there is one, and otherwise by
                    // `X509Certificate`, which `CT_CertificateKeyEncryptor` requires and
                    // `CT_PasswordKeyEncryptor` does not define — so the two are told
                    // apart by a schema fact rather than by a namespace prefix a file is
                    // free to spell any way it likes.
                    b"encryptedKey" => {
                        let is_certificate = match current_encryptor {
                            KeyEncryptorKind::Certificate => true,
                            KeyEncryptorKind::Password => false,
                            KeyEncryptorKind::Unstated => e
                                .attributes()
                                .flatten()
                                .any(|a| a.key.as_ref() == b"X509Certificate"),
                        };
                        if is_certificate {
                            saw_certificate_encryptor = true;
                            buf.clear();
                            continue;
                        }
                        // "Exactly one PasswordKeyEncryptor MUST be present"
                        // (§2.3.4.10). Two of them is a file with two answers to which
                        // key wraps the session key, and taking the last is a guess whose
                        // only symptom is a correct password reported wrong.
                        if saw_password_encryptor {
                            return Err(Error::BadParameters(
                                "the EncryptionInfo carries more than one PasswordKeyEncryptor; \
                                 [MS-OFFCRYPTO] 2.3.4.10 admits exactly one"
                                    .into(),
                            ));
                        }
                        saw_password_encryptor = true;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"spinCount" => {
                                    spin_count = Some(parse_u32_attr(&attr)?);
                                }
                                b"hashAlgorithm" => {
                                    password_hash_algorithm = Some(decode_str_attr(&attr)?);
                                }
                                b"cipherAlgorithm" => {
                                    password_cipher_algorithm = Some(decode_str_attr(&attr)?);
                                }
                                b"cipherChaining" => {
                                    password_cipher_chaining = Some(decode_str_attr(&attr)?);
                                }
                                b"hashSize" => {
                                    password_hash_size = Some(parse_u32_attr(&attr)?);
                                }
                                b"saltSize" => {
                                    password_salt_size = Some(parse_u32_attr(&attr)?);
                                }
                                b"blockSize" => {
                                    password_block_size = Some(parse_u32_attr(&attr)?);
                                }
                                b"saltValue" => {
                                    inner_salt = Some(decode_b64_attr(&attr)?);
                                }
                                b"keyBits" => {
                                    key_bits = Some(parse_u32_attr(&attr)?);
                                }
                                b"encryptedKeyValue" => {
                                    encrypted_key_value = Some(decode_b64_attr(&attr)?);
                                }
                                b"encryptedVerifierHashInput" => {
                                    encrypted_verifier_hash_input = Some(decode_b64_attr(&attr)?);
                                }
                                b"encryptedVerifierHashValue" => {
                                    encrypted_verifier_hash_value = Some(decode_b64_attr(&attr)?);
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            // Leaving a `<keyEncryptor>`: the next `<encryptedKey>` is no longer covered
            // by this one's `uri`. Without this, a certificate encryptor's `uri` would
            // still be in force over a password element that followed it outside any
            // wrapper.
            Ok(Event::End(ref e)) => {
                let name_owned = e.name().as_ref().to_vec();
                if local_name(&name_owned) == b"keyEncryptor" {
                    current_encryptor = KeyEncryptorKind::Unstated;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(&e)),
            _ => {}
        }
        buf.clear();
    }

    // A file whose only key encryptors are certificate ones is encrypted to a
    // certificate holder, not to a password — a different fact from "this attribute is
    // missing", and the one the caller needs to hear. Without it the first `ok_or_else`
    // below reports `missing spinCount`, which invites the reader to look for a
    // truncated file.
    if saw_certificate_encryptor && !saw_password_encryptor {
        return Err(Error::XmlParse(
            "the EncryptionInfo carries only CertificateKeyEncryptor elements and no \
             PasswordKeyEncryptor; this crate opens password-protected documents only"
                .into(),
        ));
    }

    // A half-written <dataIntegrity> is malformed, not absent, and the two get
    // different errors. Collapsing them into `None` would report a broken element as
    // `IntegrityElementMissing` — "someone stripped the tag" — under the default
    // policy, and, under the two explicit opt-outs (`VerifyIfPresent`, `Skip`), would
    // hand the caller `NotDeclared` and decrypt a file whose tag is corrupt.
    // `BadParameters` says the true thing in both cases.
    let data_integrity = match (saw_data_integrity, encrypted_hmac_key, encrypted_hmac_value) {
        (false, _, _) => None,
        (true, Some(encrypted_hmac_key), Some(encrypted_hmac_value)) => Some(DataIntegrity {
            encrypted_hmac_key,
            encrypted_hmac_value,
        }),
        (true, _, _) => {
            return Err(Error::BadParameters(
                "<dataIntegrity> is missing encryptedHmacKey or encryptedHmacValue".into(),
            ))
        }
    };

    // Everything from here to the struct literal is validation, and one thing it does
    // *not* do is worth stating: no attribute is defaulted. Each is an `Option` that
    // reaches its check through `ok_or_else`, so "the file did not declare this" is a
    // distinct outcome from any value the file could have declared. LibreOffice solves
    // the same problem by zeroing five numeric fields before parsing
    // (`AgileEngine.cxx:532-536`, behaviour only) so that a missing attribute becomes 0
    // and fails a range check — which works for four of the five and leaks on the fifth:
    // `spinCount = 0` passes its own bound (`:554`), so a *missing* spinCount is
    // silently accepted as zero rounds of stretching. `Option` cannot express that
    // confusion, so there is no counterpart to LibreOffice's zero-init block here and
    // its absence is the design, not a gap.
    //
    // Bound the two `<p:encryptedKey>` numbers before anything acts on them. Both are
    // unconstrained u32s in the file's own XML and both are consumed *before* the
    // password is checked, so neither needs a valid password to reach:
    //   * spinCount is a loop count -- see limits::SPIN_COUNT_MAX for the measurement;
    //   * keyBits/8 is a truncation length on the digest of the hash the file names, so
    //     anything longer than that digest cannot be cut from it inside `derive_block_key`.
    // `aes_cbc_decrypt`'s key-length check is one call frame too late for the second.
    //
    // **The failure mode this guard prevents changed, and the guard did not.** It used
    // to be a panic: `derive_block_key` sliced `digest[..key_len]`, and an over-long
    // `key_len` was an out-of-range index. Since the digest is written into a wrapped
    // slot by `hash::digest_two_into`, the same input would instead produce a key whose
    // tail is the slot's zeros -- quieter and worse, because a zero-padded key is a key
    // that works, on a file no writer produced. The refusal is what stops both; only the
    // thing it stops is different now. `digest_two_into` carries a `debug_assert` saying
    // so at the other end.
    let spin_count = spin_count.ok_or_else(|| Error::XmlParse("missing spinCount".into()))?;
    if spin_count > limits::SPIN_COUNT_MAX {
        return Err(Error::BadParameters(format!(
            "p:encryptedKey/@spinCount is {spin_count}; [MS-OFFCRYPTO] 2.3.4.10 bounds it \
             at {} (Office writes 100000)",
            limits::SPIN_COUNT_MAX
        )));
    }
    let key_bits = check_key_bits(key_bits, "p:encryptedKey/@keyBits")?;

    // `<keyData>` governs the package IVs as well as the dataIntegrity tag, so all of
    // these are required for *any* agile file, not only one that carries a tag.
    let key_data_hash = resolve_hash(
        key_data_hash_algorithm,
        "keyData/@hashAlgorithm",
        "missing keyData.hashAlgorithm",
    )?;
    require_aes_cbc(
        key_data_cipher_algorithm,
        key_data_cipher_chaining,
        "keyData/@cipherAlgorithm",
        "keyData/@cipherChaining",
    )?;
    // The third member of herumi's cipher tuple, asked of this element as well as of its
    // sibling (`include/crypto_util.hpp:147-153`, reached for `<keyData>` at `:329`). It
    // is a separate function from `require_aes_cbc` rather than a fourth parameter
    // because the two refusals differ in kind — an out-of-set `keyBits` is a malformed
    // file, while `cipherAlgorithm="RC4"` is a well-formed one this crate has not
    // implemented — but the two calls stay adjacent on both elements so neither can be
    // added without the other.
    let key_data_key_bits = check_key_bits(key_data_key_bits, "keyData/@keyBits")?;
    let key_data_block_size = check_block_size(key_data_block_size, "keyData/@blockSize")?;
    let outer_salt = inner_or_outer_salt(outer_salt, "missing keyData.saltValue")?;
    check_salt_size(
        key_data_salt_size,
        &outer_salt,
        "keyData/@saltSize",
        "missing keyData.saltSize",
    )?;
    // `keyData/@hashSize` is the one attribute here a tag-less file may legitimately
    // omit — only the two dataIntegrity blobs are cut to it — so it stays `Option` and
    // `check_integrity` refuses the absence at the point of use. Declared, it is checked
    // against the named algorithm exactly like its `<p:encryptedKey>` twin: optional is
    // "need not be present", never "may contradict the file's own hashAlgorithm".
    if let Some(declared) = key_data_hash_size {
        check_hash_size(declared, key_data_hash, "keyData/@hashSize")?;
    }

    // The same five questions, asked again of `<p:encryptedKey>` and answered from that
    // element's own attributes. Asking them once against a merged parameter set is the
    // bug LibreOffice ships (`AgileEngine.cxx:99-131`, behaviour only); herumi asks them
    // per `CipherParam` (`include/crypto_util.hpp:126-167`, applied to `<keyData>` at
    // `:329` and to `<p:encryptedKey>` at `:339`), which is the shape ported here.
    let password_hash = resolve_hash(
        password_hash_algorithm,
        "p:encryptedKey/@hashAlgorithm",
        "missing encryptedKey.hashAlgorithm",
    )?;
    require_aes_cbc(
        password_cipher_algorithm,
        password_cipher_chaining,
        "p:encryptedKey/@cipherAlgorithm",
        "p:encryptedKey/@cipherChaining",
    )?;
    let password_hash_size = password_hash_size
        .ok_or_else(|| Error::XmlParse("missing encryptedKey.hashSize".into()))?;
    check_hash_size(
        password_hash_size,
        password_hash,
        "p:encryptedKey/@hashSize",
    )?;
    let password_block_size = check_block_size(password_block_size, "p:encryptedKey/@blockSize")?;
    let inner_salt = inner_or_outer_salt(inner_salt, "missing encryptedKey.saltValue")?;
    let password_salt_size = check_salt_size(
        password_salt_size,
        &inner_salt,
        "p:encryptedKey/@saltSize",
        "missing encryptedKey.saltSize",
    )?;

    // `keyBits / 8` is the block-key truncation length and `password_hash`'s digest is
    // what it truncates, so the pair has to be checked together — the `keyBits` bound
    // above cannot cover it on its own now that the digest length is the file's choice.
    // See `derive_block_key` for why this is a refusal rather than herumi's 0x36 pad,
    // and `HashAlgorithm::can_carry_key_bits` — the single predicate this and that
    // second, defence-in-depth check both ask — for why the comparison is not written
    // out here.
    if !password_hash.can_carry_key_bits(key_bits) {
        return Err(unusable_key_bits(key_bits, password_hash));
    }

    // `keyData/@keyBits` against the blob that carries the session key. The session key
    // is `keyBits / 8` bytes, padded up to a `blockSize` multiple before it is wrapped —
    // so the blob is 16 bytes for AES-128, 32 for AES-256, and **32 for AES-192**, whose
    // 24-byte key is the one size AES-CBC cannot carry unpadded. That is the wire format,
    // not one writer's habit: real Word 16 opens the AES-192 fixture whose blob is 32
    // bytes (measured 2026-09-05, CHANGELOG), herumi resizes the decrypted value to
    // `keyBits / 8` (`include/decode.hpp:131-133`) and so does LibreOffice
    // (`AgileEngine.cxx:368`, behaviour only). `recover_session_key` makes that cut; this
    // is the length the ciphertext must have for the cut to be meaningful, asked before
    // the password is checked and before either blob is decrypted.
    //
    // Without it `<keyData>` was the one element whose key size this crate never read:
    // a file declaring `keyBits="128"` there and `"256"` on `<p:encryptedKey>` handed all
    // 32 recovered bytes to AES-256 where the writer had used the first 16, and — with no
    // `<dataIntegrity>` element to fail — returned rubbish under `Ok`. An honest AES-128
    // file (both elements 128) opens: the cipher dispatches on the length this declares
    // (GH #13). This check is about the two elements *disagreeing*, not about the size.
    let encrypted_key_value =
        encrypted_key_value.ok_or_else(|| Error::XmlParse("missing encryptedKeyValue".into()))?;
    let session_key_len = (key_data_key_bits / 8) as usize;
    let block = key_data_block_size as usize;
    let wrapped_len = session_key_len.div_ceil(block) * block;
    if encrypted_key_value.len() != wrapped_len {
        return Err(Error::BadParameters(format!(
            "keyData/@keyBits is {key_data_key_bits} ({session_key_len} bytes, {wrapped_len} \
             once padded to keyData/@blockSize {key_data_block_size}) but encryptedKeyValue \
             decodes to {} bytes; the two must agree",
            encrypted_key_value.len()
        )));
    }

    Ok(AgileParams {
        outer_salt,
        inner_salt,
        spin_count,
        key_bits,
        encrypted_key_value,
        encrypted_verifier_hash_input: encrypted_verifier_hash_input
            .ok_or_else(|| Error::XmlParse("missing encryptedVerifierHashInput".into()))?,
        encrypted_verifier_hash_value: encrypted_verifier_hash_value
            .ok_or_else(|| Error::XmlParse("missing encryptedVerifierHashValue".into()))?,
        data_integrity,
        key_data_hash,
        key_data_block_size,
        key_data_key_bits,
        key_data_hash_size,
        password_hash,
        password_salt_size,
        password_block_size,
    })
}

/// Resolve one element's `hashAlgorithm` attribute, distinguishing "absent" from "named
/// something we do not implement".
///
/// The second is [`Error::UnsupportedAlgorithm`] rather than `BadParameters`
/// precisely so it can never be confused with a wrong password: the file is well-formed
/// and internally consistent, and the user's password may be exactly right.
fn resolve_hash(
    declared: Option<String>,
    what: &'static str,
    missing: &'static str,
) -> Result<HashAlgorithm, Error> {
    let name = declared.ok_or_else(|| Error::XmlParse(missing.into()))?;
    HashAlgorithm::parse(&name).ok_or_else(|| unsupported_algorithm(what, &name))
}

/// One element's `cipherAlgorithm` must be `AES` and its `cipherChaining`
/// `ChainingModeCBC`. Both are legal to declare otherwise and unimplemented here, which
/// is an unsupported algorithm, not a malformed file.
///
/// The two travel together because they are one decision — herumi validates them as a
/// single tuple alongside `keyBits` (`include/crypto_util.hpp:147-153`) and LibreOffice
/// pairs them in each of its four accepted combinations (`AgileEngine.cxx:574-612`,
/// behaviour only). Splitting them would let a reviewer add one and forget the other,
/// which is precisely how `cipherChaining` came to be the attribute this crate parsed
/// last: `cipherAlgorithm="AES" cipherChaining="ChainingModeCFB"` is well-formed
/// ECMA-376, and decrypting it as CBC is a **silent** wrong answer rather than a loud
/// one. The dataIntegrity HMAC covers ciphertext, so such a file verifies as
/// `IntegrityOutcome::Verified` and then hands back rubbish — the same failure shape the
/// crate already fixed once for `keyData/@hashAlgorithm`.
///
/// Both attributes are required rather than defaulted: [MS-OFFCRYPTO] §2.3.4.10 lists
/// them as mandatory, every writer emits them, and treating an absent `cipherChaining`
/// as "probably CBC" would be the one guess this parser makes about an
/// attacker-supplied document — a guess that, being wrong, is unobservable.
fn require_aes_cbc(
    cipher: Option<String>,
    chaining: Option<String>,
    cipher_what: &'static str,
    chaining_what: &'static str,
) -> Result<(), Error> {
    let cipher = cipher.ok_or_else(|| Error::XmlParse(format!("missing {cipher_what}")))?;
    if cipher != "AES" {
        return Err(unsupported_algorithm(cipher_what, &cipher));
    }
    let chaining = chaining.ok_or_else(|| Error::XmlParse(format!("missing {chaining_what}")))?;
    if chaining != "ChainingModeCBC" {
        return Err(unsupported_algorithm(chaining_what, &chaining));
    }
    Ok(())
}

/// One element's `keyBits`, bounded to the key sizes ECMA-376 defines for AES.
///
/// Asked of **both** elements, which is the point: `<keyData>`'s copy sizes the package
/// key and `<p:encryptedKey>`'s the key that wraps it, they are independent declarations,
/// and only the second was ever read. herumi validates the attribute per `CipherParam`
/// (`include/crypto_util.hpp:147-153`) and so reaches `<keyData>` at `:329` as well as
/// `<p:encryptedKey>` at `:339`; LibreOffice's flat parser cannot ask the question twice
/// (`AgileEngine.cxx:99-131`, behaviour only).
///
/// The bound is necessary and not sufficient on either element: `<p:encryptedKey>`'s
/// value is additionally paired with the named digest's length in the caller, and
/// `<keyData>`'s with the length of `encryptedKeyValue`.
fn check_key_bits(declared: Option<u32>, what: &'static str) -> Result<u32, Error> {
    let declared = declared.ok_or_else(|| Error::XmlParse(format!("missing {what}")))?;
    if !limits::AGILE_KEY_BITS_ALLOWED.contains(&declared) {
        return Err(Error::BadParameters(format!(
            "{what} is {declared}; ECMA-376 defines {:?} for AES",
            limits::AGILE_KEY_BITS_ALLOWED
        )));
    }
    Ok(declared)
}

/// One element's `blockSize`, pinned to the AES block size.
///
/// [MS-OFFCRYPTO] §2.3.4.10 leaves it a free number and LibreOffice bounds it to
/// `2..=4096` (`AgileEngine.cxx:551`, behaviour only), herumi to the same range plus
/// "must be even" (`include/crypto_util.hpp:141-143`). Neither range is the right one
/// here: this crate implements AES-CBC and nothing else, every IV seeded by `<keyData>`
/// is truncated to this length, and both verifier blob lengths are rounded up to a
/// multiple of it — so a `blockSize="4096"` file is not a file we decrypt differently,
/// it is a file we cannot decrypt at all. Pinning says that one frame earlier and with
/// a message naming the attribute, which `derive_iv`'s digest-length error would not.
fn check_block_size(declared: Option<u32>, what: &'static str) -> Result<u32, Error> {
    let declared = declared.ok_or_else(|| Error::XmlParse(format!("missing {what}")))?;
    if declared as usize != AES_BLOCK_LEN {
        return Err(Error::BadParameters(format!(
            "{what} is {declared}; this crate implements AES-CBC only, whose block size \
             is {AES_BLOCK_LEN}"
        )));
    }
    Ok(declared)
}

/// One element's `hashSize` against the algorithm that same element names.
///
/// The attribute is redundant with the algorithm and is therefore checked rather than
/// used: it is a truncation length in the writer's hands, so a file whose two
/// declarations disagree is one whose blobs we would cut in the wrong place. herumi
/// pairs the two at parse time for every hash it accepts
/// (`include/crypto_util.hpp:156-166`) and LibreOffice pairs them inside each of its
/// four accepted tuples (`AgileEngine.cxx:574-612`, behaviour only).
///
/// `integrity::verify` applies the identical predicate to `keyData/@hashSize` at the
/// point it becomes a length. That is deliberate duplication, not drift: this one is
/// reached by every agile file, that one only by a file carrying `<dataIntegrity>`, and
/// the value crosses a module boundary in between.
fn check_hash_size(declared: u32, hash: HashAlgorithm, what: &'static str) -> Result<(), Error> {
    if declared as usize != hash.digest_len() {
        return Err(Error::BadParameters(format!(
            "{what} is {} but hashAlgorithm {} produces {}",
            declared,
            hash.name(),
            hash.digest_len()
        )));
    }
    Ok(())
}

/// One element's `saltSize`, bounded and cross-checked against **that element's own**
/// decoded `saltValue`.
///
/// [MS-OFFCRYPTO] §2.3.4.10 gives both rules: `saltSize` "MUST be at least 1 and no
/// greater than 65,536", and "the number of bytes required to decode the saltValue
/// attribute MUST be equal to the value of the saltSize attribute". Both sentences are
/// quoted from LibreOffice's own comments (`AgileEngine.cxx:557, 564-565` — spec
/// citation and behaviour only); herumi enforces the same range independently at
/// `include/crypto_util.hpp:138-140`, which makes the figure available from a BSD-3
/// source as well.
///
/// **Per element is the whole point.** LibreOffice's parser merges every shared
/// attribute into one flat struct and discriminates only `saltValue`
/// (`AgileEngine.cxx:133-140`), so its cross-check compares `<keyData>`'s salt bytes
/// against `<p:encryptedKey>`'s `saltSize` — correct only because writers emit 16 on
/// both. Passing the pair in together is what makes that mistake unavailable here.
fn check_salt_size(
    declared: Option<u32>,
    salt: &[u8],
    what: &'static str,
    missing: &'static str,
) -> Result<u32, Error> {
    let declared = declared.ok_or_else(|| Error::XmlParse(missing.into()))?;
    if !limits::AGILE_SALT_SIZE.contains(&declared) || declared as usize != salt.len() {
        return Err(Error::BadParameters(format!(
            "{what} is {declared} but saltValue decodes to {} bytes (the spec requires \
             them equal, and saltSize in {}..={})",
            salt.len(),
            limits::AGILE_SALT_SIZE.start(),
            limits::AGILE_SALT_SIZE.end()
        )));
    }
    Ok(declared)
}

/// A required `saltValue`. Named for the two fields it produces so neither call site
/// reads as the other's.
fn inner_or_outer_salt(declared: Option<Vec<u8>>, missing: &'static str) -> Result<Vec<u8>, Error> {
    declared.ok_or_else(|| Error::XmlParse(missing.into()))
}

// The three helpers below read attribute values the same way `classify` does, and the
// `normalized_value(Implicit1_0)` argument is explained once at `classify::attr_str`
// rather than three more times here. Briefly: it is the same call the deprecated
// `unescape_value()` made, not a replacement for it.
fn decode_b64_attr(attr: &quick_xml::events::attributes::Attribute) -> Result<Vec<u8>, Error> {
    let val = attr
        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
        .map_err(|e| xml_error(&e))?;
    BASE64.decode(val.as_ref()).map_err(|e| base64_error(&e))
}

/// Classify a `quick-xml` failure into a fixed description of its *kind*, forwarding
/// nothing the parser wrote.
///
/// **This exists because `e.to_string()` here was a hole.** `quick_xml::Error::Escape`
/// carries the text between `&` and the next `;` of whatever was being unescaped
/// (`escape.rs:807` in 0.41.0), and the values this module unescapes are `saltValue`,
/// `encryptedKeyValue` and the two verifier blobs. Forwarding the `Display` therefore put
/// attacker-chosen text into the error string, bounded only by
/// `limits::ENCRYPTION_INFO_READ_CAP` -- 1 MiB. Measured against 0.41.0: the attribute
/// value `&AAAA...;` with a 4096-character body produced a 4130-byte message reading
/// ``at 1..4097: unrecognized entity `AAAA...` ``.
///
/// That is not key material -- the attacker supplies the text -- so it is not the leak it
/// first looks like. It is still wrong twice over: it is *unbounded* attacker-chosen
/// content, which is exactly what [`unsupported_algorithm`] truncates to 32 characters and
/// documents doing; and it is a standing conduit, because a quick-xml that echoed an
/// attribute *value* would put `encryptedKeyValue` into an error message without a line
/// changing here. Reported by the downstream consumer, who had shipped a fix for the same
/// shape that day: `rusqlite::Error::SqlInputError`'s `Display` prints the failing SQL, and
/// a `#[from]` variant rendered with `{0}` had carried a live SQLCipher key into a UI
/// string. The rule that follows from it -- *an error-hygiene rule governs the strings this
/// crate writes, never the strings its dependencies write* -- is why this is a function and
/// not a tidier `{e}`.
///
/// Matched exhaustively, with no `_` arm, on purpose: `quick_xml::Error` is not
/// `#[non_exhaustive]`, so a variant added upstream becomes a compile error here rather
/// than a silent forward. The same discipline as `IntegrityOutcome::is_authenticated`.
fn xml_error(e: &quick_xml::Error) -> Error {
    use quick_xml::Error as Qx;
    Error::XmlParse(
        match e {
            Qx::Io(_) => "the XML could not be read",
            Qx::Syntax(_) => "syntactically invalid XML",
            Qx::IllFormed(_) => "the XML is not well-formed",
            Qx::InvalidAttr(_) => "an attribute is malformed",
            Qx::Encoding(_) => "the XML is not valid UTF-8",
            Qx::Escape(_) => "an attribute value contains an unterminated or unrecognized entity",
            Qx::Namespace(_) => "the XML has a namespace error",
        }
        .to_string(),
    )
}

/// Classify a base64 failure, keeping the offset and dropping the byte.
///
/// `DecodeError::InvalidByte` and `InvalidLastSymbol` carry the offending byte, which is
/// one character of the blob being decoded. That blob is `encryptedKeyValue` or a verifier
/// -- ciphertext, already in the file, so one character of it is not a disclosure. It is
/// dropped anyway for the reason in [`xml_error`]: one character of diagnostics is not
/// worth keeping a foreign `Display` in the path. The offset is a position, which is the
/// class of value this crate's messages carry everywhere.
///
/// Exhaustive for the same reason as [`xml_error`]; `base64::DecodeError` is likewise not
/// `#[non_exhaustive]`.
fn base64_error(e: &base64::DecodeError) -> Error {
    use base64::DecodeError as B64;
    Error::XmlParse(match e {
        B64::InvalidByte(at, _) => format!("base64: invalid symbol at offset {at}"),
        B64::InvalidLength(at) => format!("base64: invalid length at offset {at}"),
        B64::InvalidLastSymbol(at, _) => format!("base64: invalid final symbol at offset {at}"),
        B64::InvalidPadding => "base64: incorrect padding".to_string(),
    })
}

/// An attribute whose value is a plain token (`hashAlgorithm="SHA512"`).
fn decode_str_attr(attr: &quick_xml::events::attributes::Attribute) -> Result<String, Error> {
    attr.normalized_value(quick_xml::XmlVersion::Implicit1_0)
        .map(|v| v.into_owned())
        .map_err(|e| xml_error(&e))
}

fn parse_u32_attr(attr: &quick_xml::events::attributes::Attribute) -> Result<u32, Error> {
    let val = attr
        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
        .map_err(|e| xml_error(&e))?;
    // `ParseIntError`'s own `Display` is content-free and would be safe to forward. It is
    // classified anyway, so that the invariant below is "no foreign `Display` reaches
    // `XmlParse`" rather than "no foreign `Display` except the ones audited as harmless" --
    // the first is checkable by reading this module, the second needs re-auditing on every
    // dependency bump.
    val.parse::<u32>()
        .map_err(|_| Error::XmlParse("not a decimal integer, or out of range for u32".into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const ALL_HASHES: [HashAlgorithm; 4] = [
        HashAlgorithm::Sha1,
        HashAlgorithm::Sha256,
        HashAlgorithm::Sha384,
        HashAlgorithm::Sha512,
    ];

    // ---- the <p:encryptedKey> hash governs the whole password path -------------------

    #[test]
    fn spin_hash_is_deterministic_and_follows_the_named_hash() {
        for hash in ALL_HASHES {
            let h1 = spin_hash(hash, "Password", b"saltsalt", 1000);
            let h2 = spin_hash(hash, "Password", b"saltsalt", 1000);
            // `Dynamic` has no `PartialEq` by design -- comparing wrapped secrets
            // with `==` is exactly the timing-unsafe habit the wrapper prevents.
            // Test material is not secret, so reading both out is fine here.
            assert!(h1.with_secret(|a| h2.with_secret(|b| a == b)), "{hash:?}");
            assert_eq!(h1.with_secret(|a| a.len()), hash.digest_len(), "{hash:?}");
        }

        // ... and the four are not interchangeable, which is the whole of issue #11:
        // a SHA-256 file spun with SHA-512 produces a different H_final, fails the
        // verifier, and used to be reported as a wrong password.
        let sha256 = spin_hash(HashAlgorithm::Sha256, "Password", b"saltsalt", 10);
        let sha512 = spin_hash(HashAlgorithm::Sha512, "Password", b"saltsalt", 10);
        let same = sha256.with_secret(|a| sha512.with_secret(|b| a[..20] == b[..20]));
        assert!(
            !same,
            "the spin hash must actually depend on the named algorithm"
        );
    }

    #[test]
    fn spin_hash_does_not_panic_on_an_empty_password() {
        for hash in ALL_HASHES {
            let h = spin_hash(hash, "", b"somesalt", 100);
            assert_eq!(h.with_secret(|v| v.len()), hash.digest_len());
        }
    }

    #[test]
    fn derive_block_key_truncates_to_key_bits() {
        let h = PasswordDigest::new(vec![0u8; 64]);
        for (hash, key_bits, want) in [
            (HashAlgorithm::Sha512, 256u32, 32usize),
            (HashAlgorithm::Sha512, 128, 16),
            (HashAlgorithm::Sha384, 256, 32),
            (HashAlgorithm::Sha256, 256, 32),
            (HashAlgorithm::Sha1, 128, 16),
        ] {
            let k = derive_block_key(hash, &h, &BLOCK_KEY_VALUE, key_bits).unwrap();
            assert_eq!(k.with_secret(|v| v.len()), want, "{hash:?}/{key_bits}");
        }
    }

    /// The slice in `derive_block_key` is only in range while `keyBits / 8` fits inside
    /// the named hash's digest. SHA-1 gives 20 bytes, so `keyBits="192"` and `"256"` —
    /// both in `AGILE_KEY_BITS_ALLOWED`, both reached before any password is checked —
    /// would ask for more bytes than it has. That was a panic while the key was cut with
    /// `digest[..key_len]`; since `hash::digest_two_into` writes into a pre-zeroed
    /// wrapped slot it would instead be a zero-padded key — silent, and a key that
    /// works. Either way the refusal below is what prevents it, which is why this test
    /// asserts the error and not the crash.
    #[test]
    fn key_bits_longer_than_the_named_digest_is_refused_not_padded() {
        let h = PasswordDigest::new(vec![0u8; 20]);
        for key_bits in [192u32, 256] {
            let err = derive_block_key(HashAlgorithm::Sha1, &h, &BLOCK_KEY_VALUE, key_bits)
                .expect_err("SHA-1 cannot produce that many key bytes");
            let text = err.to_string();
            assert!(matches!(err, Error::BadParameters(_)));
            assert!(
                text.contains("keyBits") && text.contains("SHA1"),
                "the refusal must name both halves of the pair, got: {text}"
            );
        }
        // The control: the same hash with a key it *can* fill is accepted, so the
        // refusals above come from the pairing and not from SHA-1 itself.
        assert!(derive_block_key(HashAlgorithm::Sha1, &h, &BLOCK_KEY_VALUE, 128).is_ok());
    }

    // ---- parse-time bounds on file-declared numbers ---------------------------------

    /// Just enough XML for `parse_encryption_info`. The parser matches on local element
    /// names wherever they appear, so the `<keyEncryptors>` nesting a real file has is
    /// not needed to exercise the attribute checks.
    ///
    /// The blob lengths are fixed at the SHA-512 shape because nothing in `parse` looks
    /// at them — `verify_password` is what checks a blob against the declared `saltSize`
    /// and `hashAlgorithm`, and it has its own tests below.
    fn minimal_xml(key_data_attrs: &str, encrypted_key_attrs: &str) -> Vec<u8> {
        let zeros = |n: usize| BASE64.encode(vec![0u8; n]);
        format!(
            r#"<encryption>
                 <keyData saltValue="{s16}" {key_data_attrs}/>
                 <p:encryptedKey saltValue="{s16}" encryptedKeyValue="{s32}"
                                 encryptedVerifierHashInput="{s16}"
                                 encryptedVerifierHashValue="{s64}" {encrypted_key_attrs}/>
               </encryption>"#,
            s16 = zeros(16),
            s32 = zeros(32),
            s64 = zeros(64),
        )
        .into_bytes()
    }

    /// The parsed parameters are discarded: `AgileParams` deliberately has no `Debug`
    /// (it holds the file's ciphertext blobs), and these tests only ask accept-or-why-not.
    fn parse_with(key_data_attrs: &str, encrypted_key_attrs: &str) -> Result<(), String> {
        parse_encryption_info(&minimal_xml(key_data_attrs, encrypted_key_attrs))
            .map(|_| ())
            .map_err(|e| e.to_string())
    }

    /// The same call, keeping the error so a test can assert the *variant* rather than a
    /// substring of its message.
    fn parse_err(key_data_attrs: &str, encrypted_key_attrs: &str) -> Error {
        parse_encryption_info(&minimal_xml(key_data_attrs, encrypted_key_attrs))
            .err()
            .expect("expected a parse failure")
    }

    /// A self-consistent `<keyData>` attribute set: AES-256-CBC / SHA-512 / 16-byte
    /// salt, matching what Office writes and what `minimal_xml`'s blob lengths assume.
    const KEY_DATA_DEFAULTS: &[(&str, &str)] = &[
        ("saltSize", "16"),
        ("blockSize", "16"),
        ("keyBits", "256"),
        ("hashSize", "64"),
        ("hashAlgorithm", "SHA512"),
        ("cipherAlgorithm", "AES"),
        ("cipherChaining", "ChainingModeCBC"),
    ];

    /// The `<p:encryptedKey>` twin. `spinCount` is 0 so a test that reaches the KDF
    /// costs nothing; `saltValue` and the three blobs come from `minimal_xml`.
    const ENC_KEY_DEFAULTS: &[(&str, &str)] = &[
        ("saltSize", "16"),
        ("blockSize", "16"),
        ("hashSize", "64"),
        ("hashAlgorithm", "SHA512"),
        ("cipherAlgorithm", "AES"),
        ("cipherChaining", "ChainingModeCBC"),
        ("spinCount", "0"),
        ("keyBits", "256"),
    ];

    /// Render one element's attributes from its defaults with named replacements —
    /// `Some(v)` substitutes a value, `None` omits the attribute entirely.
    ///
    /// Every parse test below is "the good file, differing in exactly one attribute",
    /// which is the control discipline `malformed_input` states and which spelled-out
    /// attribute strings quietly erode: adding a newly-required attribute to the parser
    /// meant editing eight literals, and any one of them left stale would turn a
    /// single-variable test into a two-variable one. Replacement is also the only safe
    /// way to override — appending a duplicate attribute does not work, because
    /// quick-xml's `Attributes` iterator reports a duplicate as an `AttrError` and the
    /// parser's `.flatten()` then drops it, so the *first* value survives.
    fn attrs(defaults: &[(&str, &str)], overrides: &[(&str, Option<&str>)]) -> String {
        // A misspelt override would otherwise be silently ignored, and the test would
        // then assert against the *unmodified* good file — passing for the wrong reason,
        // which is the one failure a control cannot catch.
        for (name, _) in overrides {
            assert!(
                defaults.iter().any(|(key, _)| key == name),
                "{name:?} is not an attribute of this element"
            );
        }
        let mut out = String::new();
        for (key, default) in defaults {
            let value = match overrides.iter().find(|(name, _)| name == key) {
                Some((_, None)) => continue,
                Some((_, Some(replacement))) => *replacement,
                None => *default,
            };
            out.push_str(&format!(" {key}=\"{value}\""));
        }
        out
    }

    fn key_data(overrides: &[(&str, Option<&str>)]) -> String {
        attrs(KEY_DATA_DEFAULTS, overrides)
    }

    fn enc_key(overrides: &[(&str, Option<&str>)]) -> String {
        attrs(ENC_KEY_DEFAULTS, overrides)
    }

    /// The unmodified pair, for the control half of every test here.
    fn good_key_data() -> String {
        key_data(&[])
    }
    fn good_enc_key() -> String {
        enc_key(&[])
    }

    /// The ceiling is inclusive, and one past it is refused by name. Testing this at
    /// the parser rather than through `decrypt_ooxml` is deliberate: accepting the
    /// ceiling means *running* it, and ten million SHA-512 rounds — about 7 seconds of
    /// one core — is not a unit test.
    ///
    /// One past the ceiling is `10_000_001`, which is what [MS-OFFCRYPTO] 2.3.4.10
    /// forbids rather than what this crate declines, so this is now a conformance test.
    #[test]
    fn spin_count_ceiling_is_inclusive_and_one_past_it_is_refused() {
        let at = limits::SPIN_COUNT_MAX.to_string();
        assert!(parse_with(&good_key_data(), &enc_key(&[("spinCount", Some(&at))])).is_ok());

        let over = (limits::SPIN_COUNT_MAX as u64 + 1).to_string();
        let err =
            parse_with(&good_key_data(), &enc_key(&[("spinCount", Some(&over))])).unwrap_err();
        assert!(err.contains("spinCount"), "got: {err}");
    }

    #[test]
    fn key_bits_outside_the_ecma376_set_is_refused_at_parse_time() {
        for key_bits in [128u32, 192, 256] {
            let bits = key_bits.to_string();
            assert!(
                parse_with(&good_key_data(), &enc_key(&[("keyBits", Some(&bits))])).is_ok(),
                "keyBits={key_bits} is legal ECMA-376 and must parse"
            );
        }
        for key_bits in [0u32, 255, 257, 512, 513, u32::MAX] {
            let bits = key_bits.to_string();
            let err =
                parse_with(&good_key_data(), &enc_key(&[("keyBits", Some(&bits))])).unwrap_err();
            assert!(err.contains("keyBits"), "keyBits={key_bits} got: {err}");
        }
    }

    /// `<keyData>` gets the same per-attribute treatment as `<p:encryptedKey>`, and
    /// every one of these is required for *every* agile file rather than only for one
    /// carrying a `<dataIntegrity>` tag: `hashAlgorithm` and `blockSize` derive each
    /// per-segment package IV, `saltValue`/`saltSize` seed it, and the cipher pair says
    /// which cipher those IVs are for. Leaving any of them optional is what let the two
    /// halves of the `<keyData>` path disagree about which hash the file named.
    #[test]
    fn key_data_self_consistency_is_checked_per_attribute() {
        assert!(parse_with(&good_key_data(), &good_enc_key()).is_ok());

        for (overrides, expected) in [
            (vec![("hashAlgorithm", None)], "hashAlgorithm"),
            (vec![("blockSize", None)], "blockSize"),
            (vec![("blockSize", Some("32"))], "blockSize"),
            // hashSize contradicting the hash this same element names. Only a file with
            // a <dataIntegrity> tag ever *uses* it, but a file that contradicts itself
            // is refused whether or not the contradiction would have been reached.
            (vec![("hashSize", Some("32"))], "hashSize"),
            // saltSize disagreeing with the 16 bytes keyData's own saltValue decodes to
            (vec![("saltSize", Some("32"))], "saltSize"),
            (vec![("saltSize", Some("0"))], "saltSize"),
            (vec![("saltSize", None)], "saltSize"),
        ] {
            let err = parse_with(&key_data(&overrides), &good_enc_key())
                .expect_err(&format!("keyData {overrides:?} must be refused"));
            assert!(err.contains(expected), "keyData {overrides:?} got: {err}");
        }

        // The control, and the one that keeps the `hashSize` case honest: omitting
        // `hashSize` entirely still parses, because a tag-less file need not declare it.
        // "Optional" and "may contradict the file" are different permissions.
        assert!(parse_with(&key_data(&[("hashSize", None)]), &good_enc_key()).is_ok());
    }

    /// Issue #11's headline: a hash name this crate does not implement must be its own
    /// error, on **either** element. `BadParameters` was already defensible for
    /// `<keyData>`, but the `<p:encryptedKey>` half had no site at all — the attribute
    /// was never read, SHA-512 ran regardless, and the verifier comparison then reported
    /// `WrongPassword` for a password that was right.
    #[test]
    fn an_unimplemented_hash_name_is_its_own_error_never_wrong_password() {
        let err = parse_err(
            &key_data(&[("hashAlgorithm", Some("MD5"))]),
            &good_enc_key(),
        );
        assert!(
            matches!(
                &err,
                Error::UnsupportedAlgorithm { what, name }
                    if *what == "keyData/@hashAlgorithm" && name == "MD5"
            ),
            "got: {err:?}"
        );

        for name in ["MD5", "SHA3-256", "RIPEMD160", ""] {
            let err = parse_err(&good_key_data(), &enc_key(&[("hashAlgorithm", Some(name))]));
            assert!(
                matches!(
                    &err,
                    Error::UnsupportedAlgorithm { what, name: got }
                        if *what == "p:encryptedKey/@hashAlgorithm" && got == name
                ),
                "hashAlgorithm={name:?} got: {err:?}"
            );
        }

        // The control: the same container naming a hash we *do* implement parses, so the
        // refusals above are attributable to the name and not to the synthetic XML.
        assert!(parse_with(&good_key_data(), &good_enc_key()).is_ok());
    }

    /// The name lands in an error string and comes out of an attacker's XML, bounded
    /// only by the 1 MiB `EncryptionInfo` read cap.
    #[test]
    fn an_unimplemented_hash_name_is_truncated_before_it_reaches_the_error() {
        let long = "A".repeat(4096);
        let err = parse_err(
            &good_key_data(),
            &enc_key(&[("hashAlgorithm", Some(&long))]),
        );
        let Error::UnsupportedAlgorithm { name, .. } = &err else {
            panic!("got: {err:?}");
        };
        assert_eq!(name.chars().count(), 33, "32 characters plus the ellipsis");
        assert!(err.to_string().len() < 200, "got: {err}");
    }

    /// `cipherAlgorithm` is required and must be AES on both elements. An RC4 or
    /// unknown-cipher file is well-formed and unimplemented here — the same distinction
    /// the hash name gets, for the same reason.
    #[test]
    fn cipher_algorithm_must_be_aes_on_both_elements() {
        for cipher in ["RC4", "AES-GCM", "DES"] {
            let err = parse_err(
                &key_data(&[("cipherAlgorithm", Some(cipher))]),
                &good_enc_key(),
            );
            assert!(
                matches!(&err, Error::UnsupportedAlgorithm { what, .. }
                    if *what == "keyData/@cipherAlgorithm"),
                "keyData cipher={cipher} got: {err:?}"
            );

            let err = parse_err(
                &good_key_data(),
                &enc_key(&[("cipherAlgorithm", Some(cipher))]),
            );
            assert!(
                matches!(&err, Error::UnsupportedAlgorithm { what, .. }
                    if *what == "p:encryptedKey/@cipherAlgorithm"),
                "encryptedKey cipher={cipher} got: {err:?}"
            );
        }

        // Absent is a parse error, not a silent "probably AES".
        for (kd, ek) in [
            (key_data(&[("cipherAlgorithm", None)]), good_enc_key()),
            (good_key_data(), enc_key(&[("cipherAlgorithm", None)])),
        ] {
            assert!(matches!(parse_err(&kd, &ek), Error::XmlParse(_)));
        }
    }

    /// `cipherChaining` gets the same treatment as its `cipherAlgorithm` sibling, and it
    /// is the more consequential of the two to have left unparsed. `ChainingModeCFB` is
    /// legal ECMA-376 that this crate does not implement, and reading such a file as CBC
    /// is a **silent** wrong answer: the `dataIntegrity` HMAC covers ciphertext, so the
    /// file verifies as `Verified` and the caller receives rubbish. Every other refusal
    /// in this module is loud.
    #[test]
    fn cipher_chaining_must_be_cbc_on_both_elements() {
        for chaining in ["ChainingModeCFB", "ChainingModeECB", ""] {
            let err = parse_err(
                &key_data(&[("cipherChaining", Some(chaining))]),
                &good_enc_key(),
            );
            assert!(
                matches!(&err, Error::UnsupportedAlgorithm { what, name }
                    if *what == "keyData/@cipherChaining" && name == chaining),
                "keyData chaining={chaining:?} got: {err:?}"
            );

            let err = parse_err(
                &good_key_data(),
                &enc_key(&[("cipherChaining", Some(chaining))]),
            );
            assert!(
                matches!(&err, Error::UnsupportedAlgorithm { what, name }
                    if *what == "p:encryptedKey/@cipherChaining" && name == chaining),
                "encryptedKey chaining={chaining:?} got: {err:?}"
            );
        }

        // Absent is a parse error, not a silent "probably CBC" — the guess would be
        // unobservable when wrong, which is the whole argument for making it.
        for (kd, ek) in [
            (key_data(&[("cipherChaining", None)]), good_enc_key()),
            (good_key_data(), enc_key(&[("cipherChaining", None)])),
        ] {
            assert!(matches!(parse_err(&kd, &ek), Error::XmlParse(_)));
        }

        // The control: `ChainingModeCBC` on both, everything else identical, parses.
        assert!(parse_with(&good_key_data(), &good_enc_key()).is_ok());
    }

    /// `hashSize`, `blockSize` and `saltSize` on `<p:encryptedKey>` are the file's own
    /// second opinion about numbers we can derive or measure, so each is checked rather
    /// than used. `hashSize` is the sharpest: it is a truncation length in the writer's
    /// hands, which is why `integrity::verify` already applies this predicate to
    /// `<keyData>`'s copy (herumi pairs them at parse time,
    /// `include/crypto_util.hpp:156-166`).
    #[test]
    fn password_encryptor_self_consistency_is_checked_per_attribute() {
        for (overrides, expected) in [
            // hashSize contradicting hashAlgorithm
            (vec![("hashSize", Some("32"))], "hashSize"),
            (vec![("hashSize", None)], "hashSize"),
            // blockSize that is not the AES block size
            (vec![("blockSize", Some("32"))], "blockSize"),
            (vec![("blockSize", None)], "blockSize"),
            // saltSize disagreeing with the 16 bytes saltValue actually decodes to
            (vec![("saltSize", Some("32"))], "saltSize"),
            // saltSize outside [MS-OFFCRYPTO] 2.3.4.10's range
            (vec![("saltSize", Some("0"))], "saltSize"),
            (vec![("saltSize", Some("65537"))], "saltSize"),
            (vec![("saltSize", None)], "saltSize"),
            // keyBits asking for more bytes than the named digest holds
            (
                vec![("hashSize", Some("20")), ("hashAlgorithm", Some("SHA1"))],
                "keyBits",
            ),
        ] {
            let text = parse_with(&good_key_data(), &enc_key(&overrides))
                .expect_err(&format!("p:encryptedKey {overrides:?} must be refused"));
            assert!(text.contains(expected), "{overrides:?} got: {text}");
        }

        // The controls: the same numbers, each at a self-consistent value.
        for overrides in [
            vec![],
            vec![("hashSize", Some("48")), ("hashAlgorithm", Some("SHA384"))],
            vec![
                ("hashSize", Some("20")),
                ("hashAlgorithm", Some("SHA1")),
                ("keyBits", Some("128")),
            ],
        ] {
            assert!(
                parse_with(&good_key_data(), &enc_key(&overrides)).is_ok(),
                "{overrides:?} must parse"
            );
        }
    }

    /// A file naming different hashes in each element must parse to two different
    /// values — with each element's own `hashSize` checked against its own
    /// `hashAlgorithm`, never against the sibling's. Such a file is non-conforming
    /// (§2.3.4.10 tells a writer to match them) and is not refused here: reading it
    /// correctly is what keeps a conforming file from being read under the wrong hash
    /// through the same code path. A parser that merged them —
    /// LibreOffice's shape, `AgileEngine.cxx:99-131`, behaviour only — cannot represent
    /// this file at all.
    #[test]
    fn the_two_hash_algorithms_are_parsed_independently() {
        let params = parse_encryption_info(&minimal_xml(
            &key_data(&[("hashAlgorithm", Some("SHA1")), ("hashSize", Some("20"))]),
            &good_enc_key(),
        ))
        .expect("a file naming two different hashes must parse, not be refused");
        assert_eq!(params.key_data_hash, HashAlgorithm::Sha1);
        assert_eq!(params.password_hash, HashAlgorithm::Sha512);
        assert_eq!(params.key_data_hash_size, Some(20));

        // ... and the cross-check is per element: `<keyData>` naming SHA-1 with
        // `<p:encryptedKey>`'s 64 is a contradiction, even though 64 is right next door.
        let err = parse_err(
            &key_data(&[("hashAlgorithm", Some("SHA1"))]),
            &good_enc_key(),
        );
        assert!(
            err.to_string().contains("keyData/@hashSize"),
            "got: {err:?}"
        );
    }

    // ---- the <keyData> hash governs the package IVs too ------------------------------

    fn test_params(key_data_hash: HashAlgorithm, outer_salt: Vec<u8>) -> AgileParams {
        password_params(key_data_hash, HashAlgorithm::Sha512, outer_salt)
    }

    /// `AgileParams` with the two hashes chosen independently, which is the whole point
    /// of there being two fields. The verifier blob lengths follow `password_hash`,
    /// exactly as a writer's would.
    fn password_params(
        key_data_hash: HashAlgorithm,
        password_hash: HashAlgorithm,
        outer_salt: Vec<u8>,
    ) -> AgileParams {
        let value_len = password_hash.digest_len().div_ceil(AES_BLOCK_LEN) * AES_BLOCK_LEN;
        AgileParams {
            outer_salt,
            inner_salt: vec![0u8; AES_BLOCK_LEN],
            spin_count: 0,
            key_bits: 256,
            encrypted_key_value: vec![0u8; 32],
            encrypted_verifier_hash_input: vec![0u8; AES_BLOCK_LEN],
            encrypted_verifier_hash_value: vec![0u8; value_len],
            data_integrity: None,
            key_data_hash,
            key_data_block_size: AES_BLOCK_LEN as u32,
            key_data_key_bits: 256,
            key_data_hash_size: Some(key_data_hash.digest_len() as u32),
            password_hash,
            password_salt_size: AES_BLOCK_LEN as u32,
            password_block_size: AES_BLOCK_LEN as u32,
        }
    }

    /// AES-256-CBC/NoPadding — the writer half of `aes_cbc_decrypt`.
    fn cbc_encrypt(plain: &[u8], key: &[u8], iv: &[u8]) -> Vec<u8> {
        use aes::cipher::generic_array::GenericArray;
        use cbc::cipher::{BlockEncryptMut, KeyIvInit};
        let mut out = vec![0u8; plain.len()];
        cbc::Encryptor::<Aes256>::new(GenericArray::from_slice(key), GenericArray::from_slice(iv))
            .encrypt_padded_b2b_mut::<NoPadding>(plain, &mut out)
            .unwrap();
        out
    }

    /// The writer half of `decrypt_package`: 4096-byte segments, each under its own
    /// `H(keyData.saltValue || LE32(index))` IV, behind the 8-byte size prefix.
    ///
    /// Runs on [`Segments`], the same type `decrypt_package` reads through, and this used
    /// to be a second `chunks(4096).enumerate()` with its own inline padding and its own
    /// `derive_iv` call. That is the drift plan D4 exists to prevent, and a test writer is
    /// the worst place to have it: writer and reader disagreeing would show up as a failing
    /// round-trip, but writer and reader agreeing *wrongly* would not show up at all.
    ///
    /// Sharing the segmentation does not make the round-trip tests below circular, because
    /// the shared half is pinned from outside: `segments_tests.rs` asserts the cut
    /// positions against literal lengths and re-derives every IV from the formula rather
    /// than from the iterator, and the real-Office fixtures decrypt through the same code
    /// under segmentation this crate did not choose.
    ///
    /// It is also a preview of GH #6 step 4 — modulo the session key being a bare array
    /// here, which is what a test may do and the encrypt path may not.
    fn encrypt_package(
        session_key: &[u8; 32],
        outer_salt: &[u8],
        hash: HashAlgorithm,
        plaintext: &[u8],
    ) -> Vec<u8> {
        // Since GH #6 step 6 this is the production writer, not a second copy of it.
        crate::agile_encrypt::encrypt_package(
            plaintext,
            &SessionKey::new(session_key.to_vec()),
            outer_salt,
            hash,
        )
        .expect("the test tuple is in range")
    }

    /// `keyData/@hashAlgorithm` is file-controlled and read separately from
    /// `p:encryptedKey`'s. `integrity.rs` honours it; `decrypt_package` used to
    /// hardcode SHA-512, so a mixed-algorithm file could be reported `Verified` — even
    /// under `IntegrityPolicy::Require` — and then decrypted with the wrong IV for
    /// every segment.
    #[test]
    fn package_segment_ivs_follow_the_key_data_hash_algorithm() {
        // Two full segments and a short tail, so segment indices 0, 1 and 2 all take
        // part: a single-segment file would only ever exercise LE32(0).
        let plaintext: Vec<u8> = (0..9000u32).map(|i| (i % 251) as u8).collect();
        let session_key = [0x5Au8; 32];
        let salt = vec![0x11u8; 16];
        let sk = SessionKey::new(session_key.to_vec());

        for hash in ALL_HASHES {
            let stream = encrypt_package(&session_key, &salt, hash, &plaintext);
            let out = decrypt_package(&sk, &test_params(hash, salt.clone()), &stream).unwrap();
            assert_eq!(out, plaintext, "{hash:?} package must round-trip");
        }

        // ... and the four are not interchangeable. Reading a SHA-1 file with SHA-512
        // IVs — what this function did unconditionally — corrupts the first block of
        // every segment while every other check in the crate still passes.
        let sha1_stream = encrypt_package(&session_key, &salt, HashAlgorithm::Sha1, &plaintext);
        let wrong = decrypt_package(
            &sk,
            &test_params(HashAlgorithm::Sha512, salt.clone()),
            &sha1_stream,
        )
        .unwrap();
        assert_ne!(
            wrong, plaintext,
            "the segment IV must actually depend on the named hash"
        );
    }

    /// The `EncryptedPackage` stream's 8-byte prefix is the file's own claim about how
    /// long its plaintext is, and its only use is `output.truncate(..)`. `Vec::truncate`
    /// past the current length is a **silent no-op**, so an absurd claim used to be
    /// accepted and the caller quietly handed the writer's final-segment padding
    /// appended to their ZIP — the shape of failure this crate's first design value
    /// exists to rule out, because nothing in the result says anything went wrong.
    ///
    /// AES-CBC preserves length and each segment is padded up to a 16-byte multiple, so
    /// a well-formed file always satisfies `declared <= ciphertext length`.
    #[test]
    fn a_declared_plaintext_size_beyond_the_ciphertext_is_refused_not_silently_ignored() {
        let plaintext: Vec<u8> = (0..5000u32).map(|i| (i % 251) as u8).collect();
        let session_key = [0x5Au8; 32];
        let salt = vec![0x11u8; 16];
        let sk = SessionKey::new(session_key.to_vec());
        let params = test_params(HashAlgorithm::Sha512, salt.clone());
        let stream = encrypt_package(&session_key, &salt, HashAlgorithm::Sha512, &plaintext);
        let ciphertext_len = (stream.len() - 8) as u64;

        for declared in [ciphertext_len + 1, ciphertext_len + 16, 1 << 40, u64::MAX] {
            let mut forged = stream.clone();
            forged[..8].copy_from_slice(&declared.to_le_bytes());
            // Report the length, not the bytes: an accepted forgery hands back
            // thousands of them, and `Vec<u8>`'s `Debug` would bury the assertion.
            let got = decrypt_package(&sk, &params, &forged).map(|p| p.len());
            assert!(
                matches!(&got, Err(Error::BadParameters(msg))
                    if msg.contains("EncryptedPackage declares")),
                "declared={declared} got: {got:?}"
            );
        }

        // Two controls. The honest size round-trips, and the largest size the ciphertext
        // could legitimately carry — every padding byte counted as plaintext — is
        // accepted rather than being rejected by an off-by-one, since the bound is
        // "not more than the ciphertext", not "exactly the plaintext".
        assert_eq!(decrypt_package(&sk, &params, &stream).unwrap(), plaintext);

        let mut at_the_ceiling = stream.clone();
        at_the_ceiling[..8].copy_from_slice(&ciphertext_len.to_le_bytes());
        let out = decrypt_package(&sk, &params, &at_the_ceiling).unwrap();
        assert_eq!(out.len(), ciphertext_len as usize);
        assert!(out.starts_with(&plaintext), "the padding is what follows");
    }

    /// `<keyData>` and `<p:encryptedKey>` each declare a `keyBits`, about two different
    /// keys, and `<keyData>`'s was the one attribute of that element this crate never
    /// read. It is not decoration: the session key is the plaintext of
    /// `encryptedKeyValue`, so `<keyData>` is the file stating how many of those bytes
    /// the package cipher actually uses (herumi resizes to it —
    /// `include/decode.hpp:131-133`).
    #[test]
    fn key_data_key_bits_is_parsed_bounded_and_matched_to_the_session_key_blob() {
        // `minimal_xml` carries a 32-byte `encryptedKeyValue`, so 256 is the value that
        // agrees with it: the control.
        assert!(parse_with(&good_key_data(), &good_enc_key()).is_ok());

        // Absent is a parse error, not a silent "probably the same as its sibling".
        assert!(matches!(
            parse_err(&key_data(&[("keyBits", None)]), &good_enc_key()),
            Error::XmlParse(_)
        ));

        // Outside the ECMA-376 set, refused by name on this element too.
        for key_bits in ["0", "255", "257", "512", "4294967295"] {
            let err = parse_err(&key_data(&[("keyBits", Some(key_bits))]), &good_enc_key());
            assert!(
                matches!(&err, Error::BadParameters(msg)
                    if msg.contains("keyData/@keyBits")),
                "keyData keyBits={key_bits} got: {err:?}"
            );
        }

        // In the set but disagreeing with the blob that carries the session key. This is
        // the silent case: `<p:encryptedKey keyBits="256">` still unwraps 32 bytes, so
        // nothing downstream is short, and all 32 would have gone to AES-256 where the
        // writer used the first 16.
        let err = parse_err(&key_data(&[("keyBits", Some("128"))]), &good_enc_key());
        assert!(
            matches!(&err, Error::BadParameters(msg)
                if msg.contains("keyData/@keyBits") && msg.contains("encryptedKeyValue")),
            "keyData keyBits=128 got: {err:?}"
        );

        // 192 is *not* a disagreement with a 32-byte blob: the 24-byte key is padded to
        // the block before it is wrapped, so 32 is the shape Word writes and opens (GH
        // #13), accepted here and cut to 24 in `recover_session_key`. The disagreement
        // for 192 is a 24-byte blob — a length no AES-CBC writer can produce — and the
        // rule is equality with the padded length, not "at least the key".
        assert!(parse_with(&key_data(&[("keyBits", Some("192"))]), &good_enc_key()).is_ok());
        let s32 = format!("encryptedKeyValue=\"{}\"", BASE64.encode([0u8; 32]));
        let s24 = format!("encryptedKeyValue=\"{}\"", BASE64.encode([0u8; 24]));
        let xml = String::from_utf8(minimal_xml(
            &key_data(&[("keyBits", Some("192"))]),
            &good_enc_key(),
        ))
        .unwrap();
        assert_eq!(xml.matches(&s32).count(), 1);
        let err = parse_encryption_info(xml.replace(&s32, &s24).as_bytes())
            .err()
            .expect("a 24-byte encryptedKeyValue must be refused");
        assert!(
            matches!(&err, Error::BadParameters(msg)
                if msg.contains("keyData/@keyBits") && msg.contains("24 bytes")),
            "keyData keyBits=192 over 24 bytes got: {err:?}"
        );
    }

    /// The point-of-use half of the same check: `decrypt_package` is where the session
    /// key becomes an AES key, and herumi's `normalizeKey(secretKey, keyData.keyBits / 8)`
    /// sits on the line before its own `DecContent` (`include/decode.hpp:131-133`).
    /// Reachable independently of the parser, so it is asserted independently of it.
    #[test]
    fn a_session_key_that_contradicts_key_data_key_bits_is_refused_at_the_cipher() {
        let plaintext: Vec<u8> = (0..2000u32).map(|i| (i % 251) as u8).collect();
        let session_key = [0x5Au8; 32];
        let salt = vec![0x11u8; 16];
        let sk = SessionKey::new(session_key.to_vec());
        let stream = encrypt_package(&session_key, &salt, HashAlgorithm::Sha512, &plaintext);

        // 32 bytes of session key against a `<keyData>` that claims 128 or 192 bits: the
        // crafted file of the finding, where `<p:encryptedKey keyBits="256">` unwraps a
        // full-length key and nothing downstream is short enough to notice.
        for key_bits in [128u32, 192] {
            let mut params = test_params(HashAlgorithm::Sha512, salt.clone());
            params.key_data_key_bits = key_bits;
            let got = decrypt_package(&sk, &params, &stream).map(|p| p.len());
            assert!(
                matches!(&got, Err(Error::BadParameters(msg))
                    if msg.contains("keyData/@keyBits") && msg.contains("session key")),
                "keyData keyBits={key_bits} got: {got:?}"
            );
        }

        // The control: the same key and the same stream under a `<keyData>` that agrees
        // round-trip, so the refusals above come from the mismatch and nothing else.
        let params = test_params(HashAlgorithm::Sha512, salt);
        assert_eq!(decrypt_package(&sk, &params, &stream).unwrap(), plaintext);
    }

    // ---- a whole file, both elements chosen independently ----------------------------

    /// Write a complete, self-consistent agile file — the `EncryptionInfo` XML as
    /// `decrypt` receives it (header already stripped) and the `EncryptedPackage`
    /// stream — with the two elements' hash algorithms chosen **separately**.
    ///
    /// That independence is the whole reason this builder exists: every fixture in the
    /// tree sets both elements from one pair, as Office and both writers this crate reads
    /// against do (herumi `include/encode.hpp:146-147`), so no fixture can tell the two
    /// apart and a crossed-element bug survives every round trip.
    ///
    /// `keyBits` is 256 on both elements, which is the only value the AES-256-only cipher
    /// path can carry end to end.
    fn write_agile_file(
        key_data_hash: HashAlgorithm,
        password_hash: HashAlgorithm,
        password: &str,
        plaintext: &[u8],
    ) -> (Vec<u8>, Vec<u8>) {
        write_agile_file_with_salt(
            key_data_hash,
            password_hash,
            password,
            plaintext,
            &[0x22u8; AES_BLOCK_LEN],
        )
    }

    /// [`write_agile_file`] with `p:encryptedKey/@saltValue` under the caller's control.
    ///
    /// The salt is the literal CBC IV for all three password blobs ([MS-OFFCRYPTO]
    /// §2.3.4.12 case 2, §2.3.4.13), so its length is the one parameter that decides
    /// whether the IV needs §2.3.4.12's third step at all. Every writer in the corpus
    /// emits 16 bytes and therefore cannot reach it; this one can.
    fn write_agile_file_with_salt(
        key_data_hash: HashAlgorithm,
        password_hash: HashAlgorithm,
        password: &str,
        plaintext: &[u8],
        inner_salt: &[u8],
    ) -> (Vec<u8>, Vec<u8>) {
        let key_bits = 256u32;
        let outer_salt = vec![0x11u8; AES_BLOCK_LEN];
        let session_key = [0x5Au8; 32];
        // The IV a conforming writer uses: the salt fitted to blockSize, not the salt.
        let iv = fit_iv(inner_salt, AES_BLOCK_LEN);

        let h_final = spin_hash(password_hash, password, inner_salt, 0);
        let (vi, vh) = encrypt_verifier(password_hash, password, inner_salt, key_bits);
        let key3 = derive_block_key(password_hash, &h_final, &BLOCK_KEY_VALUE, key_bits).unwrap();
        let encrypted_key_value = key3.with_secret(|k| cbc_encrypt(&session_key, k, &iv));

        let xml = format!(
            r#"<encryption>
                 <keyData saltSize="16" blockSize="16" keyBits="{key_bits}"
                          hashSize="{kd_hash_size}" cipherAlgorithm="AES"
                          cipherChaining="ChainingModeCBC" hashAlgorithm="{kd_hash}"
                          saltValue="{outer}"/>
                 <p:encryptedKey spinCount="0" saltSize="{salt_size}" blockSize="16"
                                 keyBits="{key_bits}" hashSize="{pw_hash_size}"
                                 cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                 hashAlgorithm="{pw_hash}" saltValue="{inner}"
                                 encryptedVerifierHashInput="{vi}"
                                 encryptedVerifierHashValue="{vh}"
                                 encryptedKeyValue="{kv}"/>
               </encryption>"#,
            kd_hash = key_data_hash.name(),
            kd_hash_size = key_data_hash.digest_len(),
            pw_hash = password_hash.name(),
            pw_hash_size = password_hash.digest_len(),
            salt_size = inner_salt.len(),
            outer = BASE64.encode(&outer_salt),
            inner = BASE64.encode(inner_salt),
            vi = BASE64.encode(&vi),
            vh = BASE64.encode(&vh),
            kv = BASE64.encode(&encrypted_key_value),
        )
        .into_bytes();

        let package = encrypt_package(&session_key, &outer_salt, key_data_hash, plaintext);
        (xml, package)
    }

    /// A `saltSize` that is not the AES block size still opens — [MS-OFFCRYPTO]
    /// §2.3.4.12 step 3, applied to the "no blockKey" case of the same section: the salt
    /// is the IV, an IV shorter than `blockSize` is padded with `0x36`, a longer one
    /// truncated.
    ///
    /// `saltSize` is 1..=65 536 by the schema (§2.3.4.10) and the parser accepts that
    /// whole range against the decoded `saltValue`, so the files below are conforming in
    /// every respect. They used to be refused one frame later, inside `check_cbc_lengths`
    /// — "AES-CBC needs a 16-byte IV; this file's parameters produced 8" — a well-formed
    /// document reported as malformed.
    #[test]
    fn a_password_salt_that_is_not_one_block_is_fitted_to_the_iv_not_refused() {
        let plaintext: Vec<u8> = (0..3000u32).map(|i| (i % 251) as u8).collect();
        // 1, 8 and 15 pad; 24 and 64 truncate; 16 is the control every real file uses
        // and must keep working unchanged.
        for salt_len in [1usize, 8, 15, 16, 24, 64] {
            let salt = vec![0x9Bu8; salt_len];
            let (xml, package) = write_agile_file_with_salt(
                HashAlgorithm::Sha512,
                HashAlgorithm::Sha512,
                "testpass",
                &plaintext,
                &salt,
            );
            let (got, _) = decrypt(&xml, &package, "testpass", IntegrityPolicy::VerifyIfPresent)
                .unwrap_or_else(|e| panic!("saltSize={salt_len} is conforming, got {e}"));
            assert_eq!(got, plaintext, "saltSize={salt_len}");

            // The control: the password still has to be right at every salt length, so
            // the fitted IV is not making the verifier pass on its own.
            assert!(
                matches!(
                    decrypt(&xml, &package, "wrong", IntegrityPolicy::VerifyIfPresent),
                    Err(Error::WrongPassword)
                ),
                "saltSize={salt_len}"
            );
        }
    }

    /// The two hashes drive disjoint halves of the algorithm: `<p:encryptedKey>`'s runs
    /// the spin hash, both verifier block keys and the session-key block key;
    /// `<keyData>`'s runs the package IVs and the dataIntegrity tag. Naming different
    /// ones is non-conforming (§2.3.4.10), and it is the only shape that can tell the two
    /// halves apart, which is why this test writes it. Feeding the spin hash the
    /// `<keyData>` value produced a wrong
    /// `H_final`, a failed verifier, and `WrongPassword` for a password that was
    /// right — issue #11's own failure mode, surviving inside the change that closed it,
    /// invisible to every fixture because every writer sets both from one pair.
    #[test]
    fn the_spin_hash_follows_the_password_element_not_key_data() {
        let plaintext: Vec<u8> = (0..5000u32).map(|i| (i % 251) as u8).collect();

        for (key_data_hash, password_hash) in [
            (HashAlgorithm::Sha512, HashAlgorithm::Sha256),
            (HashAlgorithm::Sha256, HashAlgorithm::Sha512),
            (HashAlgorithm::Sha384, HashAlgorithm::Sha512),
            (HashAlgorithm::Sha512, HashAlgorithm::Sha384),
            // The control: the two elements agreeing, which is every real file and the
            // configuration under which the bug above is unobservable.
            (HashAlgorithm::Sha512, HashAlgorithm::Sha512),
        ] {
            let (xml, package) =
                write_agile_file(key_data_hash, password_hash, "testpass", &plaintext);
            let (out, outcome) =
                decrypt(&xml, &package, "testpass", IntegrityPolicy::VerifyIfPresent)
                    .unwrap_or_else(|e| {
                        panic!("keyData={key_data_hash:?} password={password_hash:?}: {e}")
                    });
            assert_eq!(out, plaintext, "{key_data_hash:?}/{password_hash:?}");
            assert_eq!(outcome, IntegrityOutcome::NotDeclared);
        }

        // The negative control: a genuinely wrong password on the mixed-hash file is
        // still `WrongPassword`, so the successes above are not "accepts anything".
        let (xml, package) = write_agile_file(
            HashAlgorithm::Sha512,
            HashAlgorithm::Sha256,
            "testpass",
            &plaintext,
        );
        assert!(matches!(
            decrypt(
                &xml,
                &package,
                "nottestpass",
                IntegrityPolicy::VerifyIfPresent
            ),
            Err(Error::WrongPassword)
        ));
    }

    // ---- the password verifier ------------------------------------------------------

    /// The writer half of `verify_password`, built to herumi's rules
    /// (`include/encode.hpp:160-190`): draw `saltSize` bytes, pad up to a block multiple,
    /// hash the **padded** buffer, pad that digest up to a block multiple too, and
    /// CBC-encrypt both under block keys derived from the same `H_final` the reader will
    /// compute.
    ///
    /// The pad byte is 0x36 rather than msoffcrypto-tool's 0x00 so that a reader failing
    /// to cut the hash value back to the digest length would compare against a visibly
    /// different tail. LibreOffice pads with 0x36 (`AgileEngine.cxx:656, 689`, behaviour
    /// only); a reader that truncates, as this one does, cannot tell. **Word 16 is not
    /// such a reader** — it compares the whole blob against a zero-extended hash and
    /// refuses a 0x36 tail with `0x800A1520` (measured 2026-09-05, CHANGELOG), which is
    /// why the production writer pads with zeros and this test-side one, feeding only
    /// this crate's reader, deliberately does not.
    fn encrypt_verifier(
        hash: HashAlgorithm,
        password: &str,
        salt: &[u8],
        key_bits: u32,
    ) -> (Vec<u8>, Vec<u8>) {
        let pad = |v: &[u8]| {
            let mut p = v.to_vec();
            p.resize(v.len().div_ceil(AES_BLOCK_LEN) * AES_BLOCK_LEN, 0x36);
            p
        };
        let h_final = spin_hash(hash, password, salt, 0);
        let input = pad(&vec![0xA5u8; salt.len()]);
        // H(the `saltSize` bytes), not H(the padded blob) — §2.3.4.13's
        // `encryptedVerifierHashValue` step 1 hashes the array step 1 *generated*, which
        // is `saltSize` bytes, and the pad arrives a step later when it is encrypted.
        //
        // This helper hashed `&input` until 2026-09-21, which is the same error the
        // reader and the writer both carried. Three places agreeing is why no test could
        // see it: this one built files that matched the reader's mistake, so a salt of 1,
        // 8 or 15 round-tripped here while real Word refused the equivalent file.
        //
        // The `0x36` pad above is deliberately left as it is and is now irrelevant to
        // this value: the hash no longer covers the padding, so the byte cannot change
        // the digest. It stays because a non-conforming pad is a useful thing for these
        // synthetic files to carry — the reader must not care what is in the tail.
        let value = pad(&hash.digest(&input[..salt.len()]));
        // The IV is the salt fitted to `blockSize` ([MS-OFFCRYPTO] §2.3.4.12), which is
        // the salt itself for every length a real writer emits.
        let iv = fit_iv(salt, AES_BLOCK_LEN);

        let k1 = derive_block_key(hash, &h_final, &BLOCK_VERIFIER_INPUT, key_bits).unwrap();
        let k2 = derive_block_key(hash, &h_final, &BLOCK_VERIFIER_HASH, key_bits).unwrap();
        (
            k1.with_secret(|k| cbc_encrypt(&input, k, &iv)),
            k2.with_secret(|k| cbc_encrypt(&value, k, &iv)),
        )
    }

    /// Issue #11 at the unit level: a file naming SHA-256 or SHA-384 must accept a
    /// **correct** password. Before the hash was threaded, the spin hash, both verifier
    /// block keys and the verifier digest all ran SHA-512 regardless, so every one of
    /// these was a `WrongPassword` — or, once the length guard landed, a `BadParameters`
    /// naming SHA-512 for a file that named something else.
    ///
    /// SHA-1 is absent from this table because it pairs only with `keyBits="128"` (20
    /// digest bytes cannot fill a 32-byte key), and this test holds `keyBits` at 256. The
    /// AES-128/SHA-1 tuple is covered end to end instead, by
    /// `test_non_sha512_agile_fixtures_decrypt_to_the_known_plaintext` on a fixture an
    /// independent writer produced (GH #13).
    #[test]
    fn the_password_verifier_follows_the_password_hash_algorithm() {
        let salt = vec![0x11u8; AES_BLOCK_LEN];
        let reachable = [
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ];
        for hash in reachable {
            let (vi, vh) = encrypt_verifier(hash, "testpass", &salt, 256);
            let mut params = password_params(HashAlgorithm::Sha512, hash, vec![0u8; 16]);
            params.inner_salt = salt.clone();
            params.encrypted_verifier_hash_input = vi;
            params.encrypted_verifier_hash_value = vh;

            let right = spin_hash(hash, "testpass", &salt, 0);
            verify_password(&params, &right)
                .unwrap_or_else(|e| panic!("{hash:?} must accept the right password: {e}"));

            // The negative control: same blobs, same hash, wrong password. Without it
            // this test cannot tell "dispatch wired correctly" from "always accepts".
            let wrong = spin_hash(hash, "nottestpass", &salt, 0);
            assert!(
                matches!(verify_password(&params, &wrong), Err(Error::WrongPassword)),
                "{hash:?} must still report a genuinely wrong password as WrongPassword"
            );

            // ... and the three are not interchangeable. What refuses a crossed pair is
            // the **length** guard, not the digest comparison, and this assertion is
            // named for what it actually pins: `roundUp(digest_len, 16)` is 32 / 48 / 64
            // for the three, all distinct, so a crossed `encryptedVerifierHashValue` is
            // the wrong length and is refused on the ciphertext before either blob is
            // decrypted. The variant is asserted rather than `is_err()` for exactly that
            // reason — an `is_err()` here passed while claiming to exercise a comparison
            // it never reached.
            //
            // A crossed pair that *clears* the length guard is unreachable in this crate
            // today: only SHA-1 (20 → 32) and SHA-256 (32 → 32) share a value length, and
            // SHA-1 pairs only with `keyBits="128"`, which the AES-256-only cipher
            // refuses one frame later. The digest dispatch itself is pinned by the
            // positive assertion above and by `spin_hash_is_deterministic_and_follows_the
            // _named_hash`.
            for other in reachable {
                if other == hash {
                    continue;
                }
                let mut crossed = password_params(HashAlgorithm::Sha512, other, vec![0u8; 16]);
                crossed.inner_salt = salt.clone();
                crossed
                    .encrypted_verifier_hash_input
                    .clone_from(&params.encrypted_verifier_hash_input);
                crossed
                    .encrypted_verifier_hash_value
                    .clone_from(&params.encrypted_verifier_hash_value);
                let got = verify_password(&crossed, &right);
                assert!(
                    matches!(&got, Err(Error::BadParameters(msg))
                        if msg.contains("encryptedVerifierHashInput/Value")),
                    "a {hash:?} verifier read as {other:?} must be refused by the length \
                     guard, got: {got:?}"
                );
            }
        }
    }

    /// Both verifier blobs are attacker-controlled base64 and AES-CBC preserves length,
    /// so an under-length one used to panic on `&vi[..16]` / `&vh[..64]` — reached
    /// before any password work. The two lengths are now derived from the file's own
    /// `saltSize` / `blockSize` / `hashAlgorithm` rather than hardcoded, and checked on
    /// the ciphertext before either blob is decrypted.
    #[test]
    fn verifier_blobs_of_the_wrong_length_are_an_error_not_a_panic() {
        let h_final = PasswordDigest::new(vec![0u8; 64]);
        // 16/64 is the SHA-512 shape; each pair below is wrong in exactly one way.
        for (vi_len, vh_len) in [
            (0, 64),
            (16, 16),
            (0, 0),
            (16, 48),
            (16, 0),
            (32, 64),
            (16, 80),
        ] {
            let mut params = test_params(HashAlgorithm::Sha512, vec![0u8; 16]);
            params.encrypted_verifier_hash_input = vec![0u8; vi_len];
            params.encrypted_verifier_hash_value = vec![0u8; vh_len];
            assert!(
                matches!(
                    verify_password(&params, &h_final),
                    Err(Error::BadParameters(_))
                ),
                "input={vi_len} value={vh_len} must be an error"
            );
        }

        // The control: full-length blobs get past the guard and are refused by the
        // comparison instead, so the errors above come from the lengths.
        let params = test_params(HashAlgorithm::Sha512, vec![0u8; 16]);
        assert!(matches!(
            verify_password(&params, &h_final),
            Err(Error::WrongPassword)
        ));

        // ... and the required length really does follow the named hash: 48 bytes is
        // right for SHA-384 and wrong for SHA-512, from the same 16-byte input blob.
        let mut sha384 =
            password_params(HashAlgorithm::Sha512, HashAlgorithm::Sha384, vec![0u8; 16]);
        assert_eq!(sha384.encrypted_verifier_hash_value.len(), 48);
        assert!(matches!(
            verify_password(&sha384, &PasswordDigest::new(vec![0u8; 48])),
            Err(Error::WrongPassword)
        ));
        sha384.password_hash = HashAlgorithm::Sha512;
        assert!(matches!(
            verify_password(&sha384, &PasswordDigest::new(vec![0u8; 64])),
            Err(Error::BadParameters(_))
        ));
    }

    // ---- <keyEncryptors> holds more than one kind of <encryptedKey> ------------------

    /// A `<keyEncryptors>` sequence carrying the one `PasswordKeyEncryptor` and one
    /// `CertificateKeyEncryptor`, in the order and with the namespaces
    /// [MS-OFFCRYPTO] §2.3.4.10 defines. `cert_first` swaps them, because a parser that
    /// takes the last `encryptedKey` it sees and one that takes the first are both
    /// wrong and only one of them is caught by a single ordering.
    ///
    /// Every attribute on the certificate element is the schema's:
    /// `CT_CertificateKeyEncryptor` requires exactly `encryptedKeyValue`,
    /// `X509Certificate` and `certVerifier`, and shares the first of those three names
    /// with `CT_PasswordKeyEncryptor` — which is the whole of the problem.
    fn xml_with_certificate_key_encryptor(cert_first: bool) -> Vec<u8> {
        let zeros = |n: usize| BASE64.encode(vec![0u8; n]);
        // The certificate's wrapped key is the same length as the password's, so a
        // parser that picks the wrong one still passes every length check and fails
        // later, with a wrong password reported for a right one.
        let password_key_value = BASE64.encode([0xAAu8; 32]);
        let certificate_key_value = BASE64.encode([0x55u8; 32]);
        let password_encryptor = format!(
            r#"<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                 <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256"
                                 hashSize="64" cipherAlgorithm="AES"
                                 cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"
                                 saltValue="{s16}" encryptedVerifierHashInput="{s16}"
                                 encryptedVerifierHashValue="{s64}"
                                 encryptedKeyValue="{password_key_value}"/>
               </keyEncryptor>"#,
            s16 = zeros(16),
            s64 = zeros(64),
        );
        let certificate_encryptor = format!(
            r#"<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                 <c:encryptedKey encryptedKeyValue="{certificate_key_value}"
                                 X509Certificate="{cert}" certVerifier="{s64}"/>
               </keyEncryptor>"#,
            cert = zeros(48),
            s64 = zeros(64),
        );
        let (first, second) = if cert_first {
            (&certificate_encryptor, &password_encryptor)
        } else {
            (&password_encryptor, &certificate_encryptor)
        };
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
               <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                           xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
                           xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                 <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
                          cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                          hashAlgorithm="SHA512" saltValue="{s16}"/>
                 <keyEncryptors>{first}{second}</keyEncryptors>
               </encryption>"#,
            s16 = zeros(16),
        )
        .into_bytes()
    }

    /// A `CertificateKeyEncryptor` beside the `PasswordKeyEncryptor` must not supply the
    /// wrapped key — [MS-OFFCRYPTO] §2.3.4.10: "Exactly one PasswordKeyEncryptor MUST be
    /// present. Zero or more CertificateKeyEncryptor elements are contained within the
    /// KeyEncryptors element."
    ///
    /// Both elements are spelled `encryptedKey` and differ only by namespace, so a
    /// parser matching on the local name alone reads the certificate's
    /// `encryptedKeyValue` — the intermediate key wrapped with an RSA public key — over
    /// the password's. Nothing downstream can notice: the blob is the same length, it
    /// decrypts to 32 bytes of rubbish under the password-derived block key, and the
    /// caller is told their correct password is wrong.
    #[test]
    fn a_certificate_key_encryptor_does_not_displace_the_password_one() {
        for cert_first in [false, true] {
            let params = parse_encryption_info(&xml_with_certificate_key_encryptor(cert_first))
                .unwrap_or_else(|e| {
                    panic!("a file with both key encryptors is well-formed, got {e} (cert_first={cert_first})")
                });
            assert_eq!(
                params.encrypted_key_value,
                vec![0xAAu8; 32],
                "the password key encryptor's encryptedKeyValue must be the one parsed \
                 (cert_first={cert_first})"
            );
        }
    }

    /// The negative control for the test above: the certificate element is not merely
    /// ignored by accident of ordering — with no `PasswordKeyEncryptor` at all, the file
    /// is refused rather than parsed out of the certificate's attributes.
    #[test]
    fn a_certificate_key_encryptor_alone_is_refused_by_name() {
        let zeros = |n: usize| BASE64.encode(vec![0u8; n]);
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                           xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                 <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
                          cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                          hashAlgorithm="SHA512" saltValue="{s16}"/>
                 <keyEncryptors>
                   <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                     <c:encryptedKey encryptedKeyValue="{s32}" X509Certificate="{s48}"
                                     certVerifier="{s64}"/>
                   </keyEncryptor>
                 </keyEncryptors>
               </encryption>"#,
            s16 = zeros(16),
            s32 = zeros(32),
            s48 = zeros(48),
            s64 = zeros(64),
        )
        .into_bytes();
        let err = parse_encryption_info(&xml)
            .err()
            .expect("no password key encryptor is present");
        let text = err.to_string();
        assert!(
            matches!(err, Error::XmlParse(_)),
            "expected an XmlParse refusal, got {err:?}"
        );
        assert!(
            text.contains("PasswordKeyEncryptor"),
            "the refusal must say which encryptor is missing, got: {text}"
        );
    }

    /// "Exactly one PasswordKeyEncryptor MUST be present" (§2.3.4.10). Two of them is a
    /// file whose author had two answers to "which key wraps the session key", and
    /// picking either is a guess — so it is refused, the way a `<keyData>` /
    /// `<p:encryptedKey>` `keyBits` disagreement is.
    #[test]
    fn two_password_key_encryptors_are_refused_rather_than_resolved_by_order() {
        let zeros = |n: usize| BASE64.encode(vec![0u8; n]);
        let encryptor = |kv: &str| {
            format!(
                r#"<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256"
                                   hashSize="64" cipherAlgorithm="AES"
                                   cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"
                                   saltValue="{s16}" encryptedVerifierHashInput="{s16}"
                                   encryptedVerifierHashValue="{s64}" encryptedKeyValue="{kv}"/>"#,
                s16 = zeros(16),
                s64 = zeros(64),
            )
        };
        let xml = format!(
            r#"<encryption>
                 <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
                          cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                          hashAlgorithm="SHA512" saltValue="{s16}"/>
                 <keyEncryptors>{a}{b}</keyEncryptors>
               </encryption>"#,
            s16 = zeros(16),
            a = encryptor(&BASE64.encode([0xAAu8; 32])),
            b = encryptor(&BASE64.encode([0x55u8; 32])),
        )
        .into_bytes();
        let err = parse_encryption_info(&xml)
            .err()
            .expect("two password key encryptors");
        assert!(
            matches!(err, Error::BadParameters(_)),
            "expected BadParameters, got {err:?}"
        );
        assert!(
            err.to_string().contains("PasswordKeyEncryptor"),
            "the refusal must name the element, got: {err}"
        );
        // The control: one of them alone parses, so the refusal is about the pair.
        let one = format!(
            r#"<encryption>
                 <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
                          cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                          hashAlgorithm="SHA512" saltValue="{s16}"/>
                 <keyEncryptors>{a}</keyEncryptors>
               </encryption>"#,
            s16 = zeros(16),
            a = encryptor(&BASE64.encode([0xAAu8; 32])),
        )
        .into_bytes();
        assert!(parse_encryption_info(&one).is_ok());
    }

    #[test]
    fn test_aes256_cbc_roundtrip() {
        let key = vec![0u8; 32];
        let iv = vec![0u8; 16];
        let plaintext = b"Hello, Office!  "; // exactly 16 bytes
        let enc_buf = cbc_encrypt(plaintext, &key, &iv);
        let dec = aes_cbc_decrypt(&enc_buf, &key, &iv).unwrap();
        assert_eq!(&dec, plaintext);
    }
}
