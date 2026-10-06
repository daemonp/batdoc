//! Generate the agile password encryptor: salts, the session key and the three blobs.
//!
//! GH #6 step 4, the inverse of `agile::verify_password`. [MS-OFFCRYPTO] §2.3.4.11-13, and
//! herumi `include/encode.hpp:132-197` (`encode_in`), BSD-3 with attribution in `NOTICE`.
//!
//! ```text
//! password_salt  <- random(pw.saltSize)              public: p:encryptedKey/@saltValue
//! blob_iv         = fit_iv(password_salt, 16)        §2.3.4.12: pad 0x36 / truncate
//! H_final         = spin(H, password_salt, pw, spinCount)
//! skey1/2/3       = H(H_final || blockKey_{1,2,3})[..pw.keyBits/8]
//! verifier_input <- random(pw.saltSize), 0x00-padded to a block multiple    secret
//! encryptedVerifierHashInput = AES-CBC-Enc(verifier_input,       skey1, blob_iv)
//! encryptedVerifierHashValue = AES-CBC-Enc(H(verifier_input),    skey2, blob_iv)
//! session_key    <- random(keyData.keyBits/8)        secret
//! encryptedKeyValue          = AES-CBC-Enc(session_key ‖ 0x00…,  skey3, blob_iv)
//! key_data_salt  <- random(keyData.saltSize)         public: keyData/@saltValue
//! ```
//!
//! # Which element sizes what — and why crossing them is silent
//!
//! [`crate::EncryptParams`] carries **two** `keyBits` and **two** `saltSize` because
//! [MS-OFFCRYPTO] sizes the two halves from different elements, and a file written with
//! the wrong one of a pair is still a file this crate happily reads back:
//!
//! * `password_key_bits` / `password_salt_size` — the `<p:encryptedKey>` half. §2.3.4.11
//!   cuts the key-encrypting key from `H_final` at `PasswordKeyEncryptor.keyBits`, so the
//!   three block keys above take `params.password_key_bits`; the salt it spins over, and
//!   the blob IV derived from it, are `p:encryptedKey/@saltValue`.
//! * `key_data_key_bits` / `key_data_salt_size` — the `<keyData>` half. §2.3.4.13 step 1
//!   sizes the package key from `Encryptor.KeyData.keyBits`, so the **session key** is
//!   `params.key_data_key_bits / 8`; the salt that seeds every package segment IV and both
//!   `dataIntegrity` IVs is `keyData/@saltValue`.
//!
//! `hash` is one field, not two: §2.3.4.10 MUSTs the two `hashAlgorithm` attributes equal
//! (and the two `cipherAlgorithm`s with them).
//!
//! # The session key is 32 random bytes, drawn at full length
//!
//! **This module deliberately diverged from the reference here, and the reference has since
//! been corrected.** Until 2026-09-10 herumi drew the session key with
//! `FillRand(secretKey, encryptedKey.saltSize)` — the *salt* size, 16 — and then called
//! `normalizeKey(secretKey, encryptedKey.keyBits / 8)`, which is
//! `key.resize(keySize, char(0x36))` (`include/crypto_util.hpp:38-41`). On the AES-256 path
//! that produced a key whose top 16 bytes were the constant `36 36 … 36`: 128 bits of
//! entropy in a 256-bit key. The AES-128 path was unaffected, since `keyBits / 8` and
//! `saltSize` are both 16 there and the padding was a no-op.
//!
//! It read like `saltSize` written where `keyBits / 8` was meant. Nothing downstream
//! noticed — the file decrypts perfectly, in every reader, because the key is whatever the
//! writer says it is — so no round-trip test anywhere could find it. `ms-offcrypto-writer`
//! did not share it (`src/lib.rs:474`, `intermediate_key: [u8; 32]`), which is what made it
//! a bug rather than a reading of the format.
//!
//! **Reported privately to the maintainer and fixed upstream the same day**, in
//! herumi/msoffice commit `b5fed299`, which changes the draw to
//! `FillRand(secretKey, encryptedKey.keyBits / 8)` — the fix suggested in the report. He
//! confirmed the finding and gave permission to describe it. The history is kept because it
//! is why this module draws at full length and why the test below exists; anyone reading
//! today's upstream will find the two implementations agree.
//!
//! This crate draws `keyData.keyBits / 8` bytes. Porting the padding would have been the
//! easier read of `encode.hpp` and is the kind of thing § *Every input is hostile* exists
//! to catch on the way out as well as in.
//!
//! **`encryptedKeyValue`'s 0x00 tail is a different thing and must not be confused with
//! it.** The *key* is drawn at full length; the *blob* that wraps it is
//! `roundUp(keyBits / 8, blockSize)` and its tail is zero, because §2.3.4.13 pads the
//! three `p:encryptedKey` blobs with 0x00. At AES-256 the two coincide and there is no
//! tail at all; at AES-192 there are eight bytes of pad, and filling them with
//! `normalizeKey`'s `0x36` would be the bug above wearing this one's clothes.
//!
//! # Randomness is injected, never taken from a global — plan D3
//!
//! Every draw goes through the `rng` argument, so `generate` under a seeded
//! `chacha20::ChaCha12Rng` is byte-for-byte reproducible and can be committed as a golden.
//! [`crate::encrypt_ooxml`] is the production entry point and is a one-line wrapper over
//! `encrypt`, so the seeded path and the real path are the *same code* rather than two
//! that agree today.
//!
//! **The draw order is herumi's** — password salt, verifier input, session key, keyData
//! salt (`encode.hpp:151, 163, 177, 188`), with the `dataIntegrity` HMAC key drawn fifth
//! inside `integrity::generate`. Nothing in the format depends on it, and this module
//! could draw in any order; matching leaves the door open to diffing against a `SAME_KEY`
//! build of herumi later, and costs nothing to keep.
//!
//! It is also the **golden's** order. `agile_encrypt_tests` and `standard_encrypt_tests`
//! each pin a measured container digest under a fixed seed, so adding, removing or
//! reordering a draw moves both goldens even though no rule of the format was broken.
//! Parameterising the *lengths* does not: at the default tuple every draw above is the
//! length it was when the goldens were measured, which is exactly why the goldens are the
//! check that the default path did not move.
//!
//! # What is wrapped and what is not
//!
//! `session_key` and the intermediate verifier plaintext are key material and are wrapped.
//! The two salts and all three ciphertext blobs are **written into `EncryptionInfo` in the
//! clear**, so they are public by construction and wrapping them would be theatre — the
//! rule `sensitive.rs` states and `encryption_info::EncryptionInfoParams` relies on.

use crate::agile::{
    aes_cbc_encrypt, derive_block_key, spin_hash, BLOCK_KEY_VALUE, BLOCK_VERIFIER_HASH,
    BLOCK_VERIFIER_INPUT,
};
use crate::encryption_info::{self, EncryptionInfoParams};
use crate::error::Error;
use crate::hash::{fit_iv, HashAlgorithm};
use crate::segments::Segments;
use crate::sensitive::{SessionKey, VerifierPlaintext};
use crate::{dataspaces, integrity, limits, EncryptParams};
use rand::{TryCryptoRng, TryRng};
use secure_gate::RevealSecret;

/// The AES block size — the multiple every blob is padded up to before encryption, and
/// the `blockSize` both elements declare.
///
/// The one length here that is **not** a parameter: it is fixed by AES-CBC, not chosen,
/// which is why `EncryptParams` has no field for it.
const AES_BLOCK_LEN: usize = 16;

/// One generated password encryptor: what goes in the file, plus the key that does not.
pub(crate) struct AgileKeyMaterial {
    /// `keyData/@saltValue` — public. Seeds every package segment IV and both
    /// dataIntegrity IVs. `key_data_salt_size` bytes.
    pub(crate) key_data_salt: Vec<u8>,
    /// `p:encryptedKey/@saltValue` — public. The spin-hash salt, and the value
    /// §2.3.4.12 fits into the CBC IV for all three blobs below.
    /// `password_salt_size` bytes.
    pub(crate) password_salt: Vec<u8>,
    /// The key that actually encrypts the package. **The only secret here**, and the one
    /// value in this struct that never appears in the file except wrapped under
    /// `encrypted_key_value`.
    pub(crate) session_key: SessionKey,
    /// `p:encryptedKey/@encryptedVerifierHashInput` — public ciphertext.
    pub(crate) encrypted_verifier_hash_input: Vec<u8>,
    /// `p:encryptedKey/@encryptedVerifierHashValue` — public ciphertext.
    pub(crate) encrypted_verifier_hash_value: Vec<u8>,
    /// `p:encryptedKey/@encryptedKeyValue` — public ciphertext, the wrapped session key.
    pub(crate) encrypted_key_value: Vec<u8>,
}

/// Generate a password encryptor from an **injected** RNG, under the caller's tuple.
///
/// Deterministic given `rng`, which is the whole point (plan D3): a seeded
/// `chacha20::ChaCha12Rng` makes the output a committable golden, and the production path
/// is [`crate::encrypt_ooxml`] rather than a second copy of this function.
///
/// Every length below comes from `params`, and each from **its own element** — see the
/// module header for which field sizes which half, because getting that wrong produces a
/// file this crate still reads back and no other reader opens.
///
/// # Parameters this does not validate
///
/// [`encrypt`] runs [`EncryptParams::validate`] before calling this, and that is where the
/// caller's tuple is judged. Reached directly with an unjudged tuple — the golden tests and
/// `malformed_input`'s builders do — the failures are still errors rather than panics: the
/// only arithmetic here is [`round_up`], which saturates, and an out-of-range `keyBits`
/// comes back from `derive_block_key` or `aes_cbc_encrypt` as [`Error::BadParameters`] or
/// [`Error::CipherError`]. What it will not do is bound the *allocation* an absurd
/// `saltSize` asks for, which is `validate`'s job and one reason it runs first.
///
/// # Errors
///
/// [`Error::RandomSource`] if the RNG will not produce bytes, and
/// [`Error::BadParameters`] from the key derivation, when `params.password_key_bits`
/// asks for more key than `params.hash`'s digest can carry — `validate` refuses that
/// pair, so it is unreachable from [`encrypt`], and propagated rather than unwrapped
/// because this function is also reached without it.
pub(crate) fn generate<R: TryRng + TryCryptoRng>(
    password: &str,
    params: EncryptParams,
    rng: &mut R,
) -> Result<AgileKeyMaterial, Error> {
    // Draw order is herumi's; see the module header. Each step is numbered against
    // `include/encode.hpp` so the two can be read side by side.

    // encode.hpp:151 -- the password salt: `p:encryptedKey/@saltSize` bytes, because it is
    // `p:encryptedKey/@saltValue`. Nothing about `<keyData>` sizes it.
    let mut password_salt = vec![0u8; to_len(params.password_salt_size)];
    fill(rng, &mut password_salt)?;

    // §2.3.4.12 last step. The three blobs below use the salt itself as the IV -- "if a
    // blockKey is not provided" -- but that sentence's *third* step still applies: an IV
    // shorter than `blockSize` is padded with 0x36 and a longer one truncated. The reader
    // has always done this (`AgileParams::password_blob_iv` -> `hash::fit_iv`,
    // `agile.rs:217`); this side passed the raw salt, which agrees only because Office
    // writes `saltSize == blockSize == 16`. At any other salt size the writer and its own
    // reader disagreed -- a short salt failed inside AES as an IV-length mismatch, a long
    // one the same -- so this is the one place the two directions are joined rather than
    // two implementations of one sentence that happen to coincide at the default.
    let blob_iv = fit_iv(&password_salt, AES_BLOCK_LEN);

    // encode.hpp:156 -- the spin hash, under `p:encryptedKey/@hashAlgorithm`. One `hash`
    // field: §2.3.4.10 MUSTs `<keyData>`'s and `<p:encryptedKey>`'s equal.
    let h_final = spin_hash(params.hash, password, &password_salt, params.spin_count);

    // encode.hpp:158-160 -- the three block keys, cut to **`p:encryptedKey/@keyBits`**:
    // §2.3.4.11 sizes the key derived from `H_final` at `PasswordKeyEncryptor.keyBits`,
    // and `keyData/@keyBits` appears nowhere in that sentence. Same constants, same order
    // and the same `derive_block_key` the decrypt side uses; deriving them differently
    // here is exactly the drift that would produce a file only this crate could open.
    let pw_key_bits = params.password_key_bits;
    let skey_verifier_input =
        derive_block_key(params.hash, &h_final, &BLOCK_VERIFIER_INPUT, pw_key_bits)?;
    let skey_verifier_hash =
        derive_block_key(params.hash, &h_final, &BLOCK_VERIFIER_HASH, pw_key_bits)?;
    let skey_key_value = derive_block_key(params.hash, &h_final, &BLOCK_KEY_VALUE, pw_key_bits)?;

    // encode.hpp:163-168 -- the verifier input: `p:encryptedKey/@saltSize` random bytes in
    // a buffer already at the blob's full `roundUp(saltSize, blockSize)`, so the 0x00 tail
    // §2.3.4.13 requires is the untouched remainder rather than something a `resize`
    // writes. **Allocated at the final length on purpose**: drawing `saltSize` bytes and
    // growing would reallocate and abandon the drawn bytes unwiped, the defect
    // `standard_encrypt.rs:204-215` documents. `fill` rather than `new_with` because the
    // RNG can fail and `Dynamic::new_with`'s closure cannot carry a `Result` out.
    //
    // Wrapped: it is not a key, but hashing it is how a password guess is confirmed
    // offline, so it is worth no more exposure than the keys around it.
    let salt_len = password_salt.len();
    let mut verifier_input = vec![0u8; round_up(salt_len, AES_BLOCK_LEN)];
    fill(rng, &mut verifier_input[..salt_len])?;
    let verifier_input = VerifierPlaintext::new(verifier_input);

    let encrypted_verifier_hash_input = verifier_input
        .with_secret(|vi| skey_verifier_input.with_secret(|k| aes_cbc_encrypt(vi, k, &blob_iv)))?;

    // encode.hpp:171-172 -- H(verifier_input) in a buffer at `roundUp(hashSize,
    // blockSize)`. A no-op for SHA-512 (64 is a multiple of 16) and load-bearing for
    // SHA-1 (20 -> 32).
    //
    // **The hash is over the first `saltSize` bytes, not the padded buffer.**
    // §2.3.4.13's `encryptedVerifierHashValue` step 1 hashes "the random array of bytes
    // generated in step 1 of the steps for encryptedVerifierHashInput", and that array is
    // `saltSize` bytes; the `0x00` pad to a block multiple is added in step 3, when the
    // array is encrypted, after the hash.
    //
    // This hashed the padded buffer until 2026-09-21, and `agile::verify_password`
    // digested the whole decrypted blob to match -- two halves agreeing with each other
    // and disagreeing with the format. Invisible at every `saltSize` that is a multiple
    // of 16, which was every salt size that existed before it became a caller's choice,
    // and invisible to every round-trip test in this crate for the same reason. Real
    // Word 16 refused the result with `0x800A1520`.
    let salt_len = to_len(params.password_salt_size);
    let verifier_hash = verifier_input.with_secret(|vi| {
        VerifierPlaintext::new_with(round_up(params.hash.digest_len(), AES_BLOCK_LEN), |slot| {
            let digest = params.hash.digest(&vi[..salt_len]);
            // `digest.len() == params.hash.digest_len() <= slot.len()` by `round_up`.
            slot[..digest.len()].copy_from_slice(&digest);
        })
    });
    let encrypted_verifier_hash_value = verifier_hash
        .with_secret(|vh| skey_verifier_hash.with_secret(|k| aes_cbc_encrypt(vh, k, &blob_iv)))?;

    // encode.hpp:177 -- the session key: **`keyData/@keyBits / 8`** bytes, because
    // §2.3.4.13 step 1 sizes the package key from `Encryptor.KeyData.keyBits`. Drawn at
    // full length rather than drawn short and padded to fit; see the module header.
    let session_key_len = to_len(params.key_data_key_bits / 8);
    let session_key = SessionKey::from_rng(session_key_len, rng).map_err(random_source)?;

    // The plaintext of `encryptedKeyValue` is that key in a `roundUp(keyBits / 8,
    // blockSize)` buffer with a **zero** tail -- not the key alone. AES-192 is where the
    // two differ: a 24-byte key in a 32-byte blob, which `agile.rs:1160-1170` requires on
    // the way back in.
    //
    // Provenance, stated exactly, because the loose version of this sentence was written
    // here first and is the kind CLAUDE.md § *Evidence over intent* exists to stop:
    // `tests/fixtures/agile_aes192_sha384.docx` was **written by msoffcrypto-tool**
    // (`tools/gen_agile_fixtures.py:130-138`), not by Word. What real Word 16 did was
    // *open* it. Those are different claims and only the second is ours to make -- Office
    // 16 refuses to write any tuple but AES-256/SHA-512, which is
    // `docs/design/development-record.md` § 3.2 and the reason no Office-written fixture
    // exists for this case at all.
    //
    // Zero is the pad, measured rather than assumed: `tools/gen_agile_fixtures.py:63-78`
    // and the thirteen variants at `integrity.rs:453-466`.
    //
    // `SessionKey::new_with` and not a `Vec`: this buffer *is* the session key, so a bare
    // copy would be the whole secret unwrapped and unzeroized, and a `to_vec` grown to the
    // padded length would additionally abandon the first allocation.
    let key_value_plaintext = session_key.with_secret(|sk| {
        SessionKey::new_with(round_up(session_key_len, AES_BLOCK_LEN), |slot| {
            // `sk.len() == session_key_len <= slot.len()` by `round_up`.
            slot[..sk.len()].copy_from_slice(sk);
        })
    });
    let encrypted_key_value = key_value_plaintext
        .with_secret(|kv| skey_key_value.with_secret(|k| aes_cbc_encrypt(kv, k, &blob_iv)))?;

    // encode.hpp:188 -- the package salt, drawn last, at `keyData/@saltSize`.
    let mut key_data_salt = vec![0u8; to_len(params.key_data_salt_size)];
    fill(rng, &mut key_data_salt)?;

    Ok(AgileKeyMaterial {
        key_data_salt,
        password_salt,
        session_key,
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    })
}

/// Encrypt a package into the `EncryptedPackage` stream: the 8-byte little-endian
/// plaintext length, then every 4096-byte segment under its own IV.
///
/// Runs on [`Segments`], the same iterator `agile::decrypt_package` reads through, so the
/// two directions share one segmentation and one IV derivation (plan D4). The final
/// segment is zero-padded to a block multiple — herumi's `data.resize(RoundUp(size, 16))`
/// — and the reader truncates it away against the prefix.
///
/// `hash` is `<keyData>`'s. Production writes SHA-512; the parameter exists so the
/// mixed-hash test can write a file whose `<keyData>` names SHA-1. That file is
/// non-conforming — [MS-OFFCRYPTO] §2.3.4.10 tells a writer the two `hashAlgorithm`
/// attributes MUST match — and a reader must still honour each element.
pub(crate) fn encrypt_package(
    plaintext: &[u8],
    session_key: &SessionKey,
    key_data_salt: &[u8],
    hash: HashAlgorithm,
) -> Result<Vec<u8>, Error> {
    let mut stream = Vec::with_capacity(8 + plaintext.len() + AES_BLOCK_LEN);
    stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    for segment in Segments::new(plaintext, hash, key_data_salt, AES_BLOCK_LEN)? {
        let segment = segment?;
        let block =
            session_key.with_secret(|k| aes_cbc_encrypt(&segment.padded(), k, &segment.iv))?;
        stream.extend_from_slice(&block);
    }
    Ok(stream)
}

/// Encrypt an OOXML package with a password into a complete CFB container — the whole
/// agile write path, assembled, with the randomness injected.
///
/// The seeded entry point behind [`crate::encrypt_ooxml`], which is this plus the shape
/// guard and `rand::rngs::SysRng`. Everything a test can prove about production goes
/// through here: a committed golden under a seeded `chacha20::ChaCha12Rng`, and the
/// external readers in GH #8 opening what a seeded run wrote.
///
/// **This function encrypts whatever bytes it is handed.** The check that the input is a
/// plain OOXML package is [`crate::check_encryptable`]'s, and it lives in the public
/// wrapper rather than here — which is why this one is `pub(crate)`. Keeping it
/// unguarded is deliberate: it is the constructor the golden tests and the degenerate
/// payload tests drive, and those need to reach shapes the public entry point refuses.
///
/// ```text
/// params.validate()                                                    refuse first
/// material  = generate(password, params, rng)                          step 4
/// package   = LE64(len) || AES-CBC per 4096-byte segment, keyData IVs  step 2
/// integrity = integrity::generate(session_key, package, rng)           step 5
/// info      = encryption_info::write(params, material, integrity)      step 3
/// container = dataspaces::build_container(info, package)               step 1
/// ```
///
/// # The tuple is judged before anything else happens
///
/// [`EncryptParams::validate`] runs **first** — before the payload ceiling and before the
/// first RNG draw — so a refusal costs the caller nothing it cannot take back. It is here
/// and not only in [`crate::encrypt_ooxml`] because this function is the one the seeded
/// golden tests and the degenerate-payload tests drive directly; a guard only the public
/// wrapper ran would be a guard no test that reaches this code path exercises.
/// `encryption_info::write` validates the same value again at the end, which is one
/// function asked twice rather than two copies of a rule.
///
/// # Errors
///
/// [`Error::EncryptParams`] if `params` is not a tuple this crate will write;
/// [`Error::BadParameters`] if `package` exceeds
/// [`limits::PAYLOAD_CEILING`] — the same 1 GiB the decrypt side refuses, checked on the
/// input so that a file this crate writes is a file this crate can read back;
/// [`Error::RandomSource`] if `rng` will not produce bytes;
/// [`Error::Io`] if the in-memory container cannot be written.
pub(crate) fn encrypt<R: TryRng + TryCryptoRng>(
    package: &[u8],
    password: &str,
    params: EncryptParams,
    rng: &mut R,
) -> Result<Vec<u8>, Error> {
    params.validate()?;

    if package.len() > limits::PAYLOAD_CEILING {
        return Err(Error::BadParameters(format!(
            "the package is {} bytes; this crate encrypts at most {} (see \
             limits::PAYLOAD_CEILING), because that is what it will read back",
            package.len(),
            limits::PAYLOAD_CEILING
        )));
    }

    let material = generate(password, params, rng)?;
    // `params.hash` on both calls below is `<keyData>`'s: the package segment IVs and
    // the two `dataIntegrity` IVs are derived from `keyData/@saltValue` under
    // `keyData/@hashAlgorithm`. One field serves both elements because §2.3.4.10 MUSTs
    // them equal; the *salt* does not, which is why `key_data_salt` is passed and the
    // password salt is nowhere near this half.
    let encrypted_package = encrypt_package(
        package,
        &material.session_key,
        &material.key_data_salt,
        params.hash,
    )?;
    let blobs = integrity::generate(
        &material.session_key,
        params.hash,
        &material.key_data_salt,
        AES_BLOCK_LEN,
        &encrypted_package,
        rng,
    )?;
    let info = encryption_info::write(&EncryptionInfoParams {
        // The same value `generate` was handed, so the document cannot describe a tuple
        // other than the one that produced these blobs. That is the whole reason
        // `EncryptionInfoParams` carries an `EncryptParams` rather than loose numbers:
        // the disagreement is not expressible.
        params,
        key_data_salt: &material.key_data_salt,
        encrypted_hmac_key: &blobs.encrypted_hmac_key,
        encrypted_hmac_value: &blobs.encrypted_hmac_value,
        password_salt: &material.password_salt,
        encrypted_verifier_hash_input: &material.encrypted_verifier_hash_input,
        encrypted_verifier_hash_value: &material.encrypted_verifier_hash_value,
        encrypted_key_value: &material.encrypted_key_value,
    })?;
    dataspaces::build_container(&info, &encrypted_package)
}

/// `roundUp(n, multiple)` — the format's padding rule, in one place.
///
/// Saturating rather than wrapping, for `encryption_info::round_up_to_block`'s reason:
/// every input reaching it from [`encrypt`] has been through
/// [`EncryptParams::validate`], so the saturation is unreachable there and exists because
/// [`generate`] can also be called without it and an unreachable panic is still a panic.
fn round_up(n: usize, multiple: usize) -> usize {
    n.div_ceil(multiple).saturating_mul(multiple)
}

/// A `u32` parameter as a length, saturating for the reason [`round_up`] does.
///
/// [`EncryptParams::validate`] bounds `saltSize` at 65 536 and `keyBits` at 256, so on
/// every target this crate builds for the conversion is exact; the fallback is there so
/// that a 16-bit `usize` would produce a length the RNG fills rather than a panic.
fn to_len(n: u32) -> usize {
    usize::try_from(n).unwrap_or(usize::MAX)
}

/// Fill `dst` from the injected RNG, mapping the RNG's own error into this crate's.
pub(crate) fn fill<R: TryRng + TryCryptoRng>(rng: &mut R, dst: &mut [u8]) -> Result<(), Error> {
    rng.try_fill_bytes(dst).map_err(random_source)
}

/// The RNG's `Display`, which describes the *source* and never its output — on a failure
/// there is no output to describe.
///
/// **This is the one foreign `Display` this crate forwards on purpose**, and it is the
/// exception that `Error`'s own docs argue for rather than assume. `agile::xml_error`
/// classifies quick-xml's instead of forwarding it, because quick-xml's carries text drawn
/// from a document an attacker wrote. An RNG's does not: it names an environment failure —
/// no `getrandom` in the sandbox, an exhausted descriptor table — and that sentence is the
/// whole diagnostic value of the variant. Classifying it to "the random source failed"
/// would leave a caller with an unactionable error where the actionable one was free.
///
/// It is **truncated** all the same. The bound is not about this RNG, whose messages are a
/// short sentence; it is that the generic accepts any `Display`, so the length is a
/// property of whatever is passed rather than of anything checked here. Same reasoning as
/// `agile::unsupported_algorithm`'s 32-character cap on an attribute name: bounded at the
/// construction site, and the variant says so.
pub(crate) fn random_source<E: core::fmt::Display>(e: E) -> Error {
    /// Long enough for any real `io::Error` or `getrandom` sentence, short enough that the
    /// message cannot become a payload.
    const MAX: usize = 200;

    let text = e.to_string();
    let bounded = match text.char_indices().nth(MAX) {
        // `char_indices` keeps the cut on a boundary, so this cannot panic on a multi-byte
        // character the way `truncate(MAX)` would.
        Some((cut, _)) => format!("{}…", &text[..cut]),
        None => text,
    };
    Error::RandomSource(bounded)
}

#[cfg(test)]
#[path = "agile_encrypt_tests.rs"]
mod tests;
