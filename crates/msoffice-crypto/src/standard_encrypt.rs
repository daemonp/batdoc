//! Write ECMA-376 **standard** encryption — the Office 2007 format, AES-ECB at any of
//! the three key sizes [MS-OFFCRYPTO] §2.3.4.5 permits, under a SHA-1 key — as the
//! counterpart of `standard::decrypt`.
//!
//! GH #7 (plan S6). [MS-OFFCRYPTO] §2.3.4.5 (the `EncryptionInfo` layout), §2.3.3 (the
//! `EncryptionVerifier` structure and its construction steps), §2.3.4.4 (the
//! `EncryptedPackage` stream), §2.3.4.7 (the key derivation, shared with the reader as
//! `standard::derive_standard_key`), §2.3.4.8 (password verifier generation) and
//! §2.3.4.9 (verification, which is what the writer's tests drive).
//!
//! ```text
//! salt          <- random(16)                          public: EncryptionVerifier.Salt
//! key            = derive_standard_key(password, salt, KeySize/8)   secret, 50 000 SHA-1
//!                                                       rounds + ladder
//! verifier      <- random(16)                          secret, never stored in the clear
//! EncryptedVerifier     = AES-ECB(verifier,                       key)
//! EncryptedVerifierHash = AES-ECB(SHA1(verifier) || 0x00 * 12,    key)
//! EncryptedPackage      = LE64(len) || AES-ECB(package || 0x00 * pad, key)
//! ```
//!
//! # Provenance
//!
//! There is no permissive **writer** to port. herumi's `standard_encryption.hpp` is a
//! decoder — `EncryptionHeader::analyze` (`:18-79`), `EncryptionVerifier::analyze`
//! (`:87-118`) and `getEncryptionKey` (`:120-133`) — and this module writes the layout
//! those functions read; msoffcrypto-tool's `ecma376_standard.py` likewise reads only.
//! The layout itself is the spec's; two writers were read for behaviour, neither copied:
//!
//! * **LibreOffice** `oox/source/crypto/Standard2007Engine.cxx` (MPL-2.0, read only):
//!   `setupEncryption` `:190-196` sets `fAES | fCryptoAPI`, `AlgID 0x660E`,
//!   `AlgIDHash 0x8004`, `ProviderType 0x18`; `generateVerifier` `:58-60` pads the SHA-1
//!   to 32 bytes **with zeros** before encrypting it; `writeEncryptionInfo` `:213-241`
//!   writes the version, the `Flags` copy, the header size, the eight fields, the
//!   CSP name with its terminator, then the verifier; `encrypt` `:254-255, :270`
//!   writes the size as `u32` + `u32` reserved and rounds the tail up to a block.
//! * **excelize** `crypt.go` (BSD-3, fetched read-only): `standardKeyEncryption` writes
//!   the same `0x24` / `0x660E` / `0x8004` / `0x18` quadruple and the `(Prototype)`
//!   spelling of the CSP name, and its `Encrypt` writes the size prefix as a `u64`.
//!
//! # What the header says, and why each value
//!
//! * **`vMajor 4, vMinor 2`.** §2.3.4.5 admits `vMajor` 3 or 4; 4 is what
//!   `standard_encrypted.docx` carries, what LibreOffice names `VERSION_INFO_2007_FORMAT_SP2`
//!   (`mscodec.hxx:436`, constant only), and what this crate's own reader and classifier
//!   both accept. Real Word 16 opening the result is the check on the choice (recorded in
//!   `CHANGELOG.md`, GH #7).
//! * **`Flags = fCryptoAPI | fAES` (`0x24`), `AlgID` one of `0x660E` / `0x660F` /
//!   `0x6610`.** The conforming pairs. `standard_encrypted.docx` carries `0x36` and
//!   RC4's `0x6801` — a pair §2.3.2 forbids — and every reader forgives it; the writer
//!   does not repeat it.
//! * **`AlgIDHash 0x8004`, `KeySize` 128, 192 or 256 paired with that `AlgID`,
//!   `ProviderType 0x18`, `CSPName` "Microsoft Enhanced RSA and AES Cryptographic
//!   Provider".** All measured on the fixture, which is the AES-128 row; the CSP name is
//!   the one Office 2007 wrote (excelize adds " (Prototype)", the beta's spelling, which
//!   §2.3.2 also lists). `ProviderType` and `CSPName` do not move with the key size:
//!   §2.3.2 gives one `ProviderType` for AES and §2.3.4.5 one pair of CSP names for the
//!   whole format, and `PROV_RSA_AES` is the provider for all three.
//! * **The `EncryptedVerifierHash` tail is zero.** SHA-1 is 20 bytes; the AES blob is 32.
//!   GH #13 measured that Word compares the *whole* decrypted blob against its hash
//!   zero-extended, so a `0x36` pad — LibreOffice's agile choice — is read by Word as a
//!   wrong password. LibreOffice's standard writer already pads with zeros; so does this.
//!
//! # The one thing a caller may choose, and why it is a `u32` and not a struct
//!
//! **This format has exactly one writer-settable parameter, and it is the key size.**
//! Everything else the header carries is fixed by [MS-OFFCRYPTO] rather than by an
//! opinion, which is worth listing because it is the whole argument for the shape of
//! [`crate::encrypt_ooxml_standard_with_key_bits`]:
//!
//! | field | fixed by |
//! | --- | --- |
//! | `AlgIDHash` = SHA-1 | §2.3.4.5 "This value MUST be 0x00008004 (SHA-1)", and §2.3.4.7 fixes the KDF's hash outright |
//! | the 50 000 iterations | §2.3.4.7; **not a field at all** — the file does not carry it |
//! | `SaltSize` = 16, `VerifierHashSize` = 20, the 32-byte blob | §2.3.3 and §2.3.4.9 step 3 |
//! | `blockSize` | AES-ECB's block is 16 bytes; nobody chose it |
//! | `Flags`, `ProviderType`, `CSPName`, the version pair | §2.3.4.5's header table |
//! | `AlgID` | not independent: §2.3.2 pairs each `AlgID` with exactly one `KeySize`, so naming the key size names it |
//!
//! So the parameter space is one number with three values in it, and three shapes were
//! available for handing it in:
//!
//! * **Reuse [`crate::EncryptParams`]** — rejected. It is the *agile* tuple: `hash`, two
//!   `keyBits`, two `saltSize` and `spinCount`. Five of its six fields name nothing this
//!   format has as a file field (the table above), so a caller would be offered five
//!   knobs that cannot move — an affordance shaped by the other format, which is exactly
//!   the mistake the parameterisation plan warns about. It would also make
//!   `EncryptParams::validate` answer a question about a file it knows nothing of: its
//!   `keyBits` rules are `ST_KeyBits`', and `ST_KeyBits` does not govern this header.
//! * **A `StandardEncryptParams` struct of one field** — rejected as a shape borrowed
//!   from a format that needed it. It buys extensibility this format cannot use, since
//!   every other field is a MUST above, and it buys it at a cost: per the parameter
//!   plan's finding 1, such a struct must *not* be `#[non_exhaustive]` if
//!   `..Default::default()` is to work from another crate, so it is a public struct
//!   whose fields are all public from the first day. This crate is unpublished and
//!   breaking changes are free (CLAUDE.md § *Project Status*), so if the format ever
//!   grows a second knob, the entry point grows an argument or a sibling then.
//! * **A plain `key_bits: u32` on a second entry point** — chosen. It matches the wire
//!   field, which §2.3.2 defines as an unsigned integer *in bits*; it matches
//!   `EncryptParams::key_data_key_bits`, `classify`'s `AlgorithmParams::key_bits` and
//!   `limits::STANDARD_KEY_BITS_AES`, so one number means one thing across the crate;
//!   and it names the parameter in the function's own signature rather than behind a
//!   type a caller has to go and read.
//!
//! A `u32` can hold 200, and [`AesKeySize::new`] is what refuses it —
//! [`Error::EncryptParams`], before the password is touched. An enum of three variants
//! would have made that refusal unnecessary by construction; it was not chosen because
//! it puts the crate's *implementation* set into the type, and a caller whose key size
//! arrives at runtime (from a UI, or from [`crate::classify`] of another file) would then
//! have to write the mapping — and write the refusal — itself.
//!
//! # No integrity, by the format's own definition
//!
//! Standard encryption has no `dataIntegrity`, no HMAC, nothing that notices a modified
//! ciphertext: a flipped byte decrypts to a flipped block, and ECB means each block is
//! independent so the damage is local and silent. `decrypt_ooxml_with_policy` reports
//! `IntegrityOutcome::NotApplicable` for what this module writes, and a test in
//! `standard_encrypt_tests.rs` pins that a tampered file *decrypts* — the honest negative
//! control against the agile suite's `IntegrityCheckFailed`. This is why
//! [`crate::encrypt_ooxml`] is agile and this format is the explicitly-named alternative
//! for a consumer who needs Office 2007 to open the file.
//!
//! # Randomness is injected — plan D3
//!
//! Two draws, both through the `rng` argument: the salt (public, written to the header)
//! and the verifier (secret, held as `VerifierPlaintext` until it is encrypted). Under a
//! seeded `chacha20::ChaCha12Rng` the whole container is a committable golden;
//! [`crate::encrypt_ooxml_standard`] is one line over `encrypt` with `rand::rngs::SysRng`.
//!
//! # What is wrapped
//!
//! The derived key (`DerivedKey`, from `derive_standard_key` — the KDF also wraps its
//! `H_final` as `PasswordDigest`), the verifier and its hash (`VerifierPlaintext`). The
//! salt and both ciphertext blobs are written into the file in the clear and are not.

use crate::agile_encrypt::{fill, random_source};
use crate::error::{EncryptParam, EncryptParamProblem, Error};
use crate::sensitive::{DerivedKey, VerifierPlaintext};
use crate::standard::{
    aes_ecb_encrypt, derive_standard_key, AES128_KEY_LEN, ALG_ID_AES_128, ALG_ID_AES_192,
    ALG_ID_AES_256, FLAG_AES, HEADER_FIXED_LEN, SHA1_LEN,
};
use crate::{dataspaces, limits};
use rand::{TryCryptoRng, TryRng};
use secure_gate::RevealSecret;
use sha1::{Digest, Sha1};

/// `EncryptionInfo.vMajor` / `vMinor` — [MS-OFFCRYPTO] §2.3.4.5; see the module header.
pub(crate) const STANDARD_VERSION: (u16, u16) = (4, 2);

/// `EncryptionHeaderFlags.fCryptoAPI` — [MS-OFFCRYPTO] §2.3.1. "MUST be 1" for this
/// format; `fAES` is `standard::FLAG_AES`, the bit the reader dispatches on.
const FLAG_CRYPTO_API: u32 = 0x0000_0004;

/// `EncryptionHeader.Flags`, and the copy of it that precedes the header —
/// `fCryptoAPI | fAES`, `0x24`. `fExternal` and `fDocProps` clear, as §2.3.4.5 requires.
pub(crate) const FLAGS: u32 = FLAG_CRYPTO_API | FLAG_AES;

/// `AlgIDHash` — `CALG_SHA1`, [MS-OFFCRYPTO] §2.3.2. The only hash this format defines.
const ALG_ID_HASH_SHA1: u32 = 0x0000_8004;

/// `ProviderType` — `PROV_RSA_AES` (`0x18`), [MS-OFFCRYPTO] §2.3.2.
const PROVIDER_TYPE_AES: u32 = 0x0000_0018;

/// `CSPName`, as Office 2007 wrote it — measured on `standard_encrypted.docx`. Written as
/// UTF-16LE with a two-byte terminator; [`CSP_NAME_LEN`] is that on-disk length.
pub(crate) const CSP_NAME: &str = "Microsoft Enhanced RSA and AES Cryptographic Provider";

/// `CSPName` on disk: one UTF-16 code unit per character plus the terminator. The name
/// is ASCII, so that is one code unit per byte of the literal — asserted at compile time,
/// so a changed string cannot silently make [`ENCRYPTION_HEADER_SIZE`] wrong.
const CSP_NAME_LEN: usize = (CSP_NAME.len() + 1) * 2;
const _: () = assert!(CSP_NAME.is_ascii());

/// `EncryptionHeaderSize` — the header's fixed fields plus the CSP name: 140 on disk,
/// the same 140 the fixture declares.
const ENCRYPTION_HEADER_SIZE: u32 = 140;
const _: () = assert!(ENCRYPTION_HEADER_SIZE as usize == HEADER_FIXED_LEN + CSP_NAME_LEN);

/// `EncryptionHeader.KeySize` for [`crate::encrypt_ooxml_standard`] — AES-128, the
/// smallest of the three [MS-OFFCRYPTO] §2.3.4.5 defines, and what Office 2007 itself
/// writes.
///
/// It was the *only* key size this writer emitted until 2026-09-20, and before that the
/// only one the reader accepted either. The default does not move with the
/// parameterisation, and that is deliberate on two grounds: the committed golden in
/// `standard_encrypt_tests.rs` is the byte-for-byte evidence that this path is stable,
/// and the acceptance-gate verdicts recorded in `CHANGELOG.md` were all measured on
/// AES-128 artifacts. A default that moved would invalidate both at once.
pub(crate) const DEFAULT_KEY_BITS: u32 = limits::STANDARD_KEY_BITS_AES[0];
const _: () = assert!(DEFAULT_KEY_BITS as usize == AES128_KEY_LEN * 8);

/// The three `(KeySize, AlgID)` rows an ECMA-376 standard `EncryptionHeader` may carry,
/// ascending, in the order [`limits::STANDARD_KEY_BITS_AES`] gives them.
///
/// **Both halves of a row are one statement and neither may be written alone.**
/// [MS-OFFCRYPTO] §2.3.2's `AlgID` table reads `0x0000660E` "128-bit AES", `0x0000660F`
/// "192-bit AES", `0x00006610` "256-bit AES", and §2.3.4.5's header table requires
/// `AlgID` to be one of those three and `KeySize` to be `0x00000080` (AES-128),
/// `0x000000C0` (AES-192) or `0x00000100` (AES-256). A header pairing one row's `AlgID`
/// with another's `KeySize` describes no cipher, and this crate's own reader refuses it
/// by name (`standard::parse_encryption_info`, the `named_key_bits` check). Writing the
/// pair from a single table is what makes that unreachable from here.
///
/// The `AlgID`s are `standard`'s, not restated: the reader's list and the writer's list
/// are the same three numbers or the crate writes files it cannot open.
const AES_ROWS: [(u32, u32); 3] = [
    (limits::STANDARD_KEY_BITS_AES[0], ALG_ID_AES_128),
    (limits::STANDARD_KEY_BITS_AES[1], ALG_ID_AES_192),
    (limits::STANDARD_KEY_BITS_AES[2], ALG_ID_AES_256),
];

/// One of the three AES key sizes an ECMA-376 standard header may declare, checked at
/// the boundary and carried as a value the rest of this module cannot get wrong.
///
/// The type exists so that "a caller must not be able to express a header this crate
/// cannot write" is structural rather than remembered: [`Self::new`] is the only
/// constructor, it is the only place a `u32` from outside is judged, and every function
/// below takes this instead of a number — so `write_encryption_info` stays infallible,
/// and no second copy of the rule can drift from the first.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct AesKeySize {
    /// `EncryptionHeader.KeySize`, in bits — one of [`limits::STANDARD_KEY_BITS_AES`].
    bits: u32,
    /// `EncryptionHeader.AlgID`, the identifier §2.3.2 pairs with exactly that `KeySize`.
    alg_id: u32,
}

impl AesKeySize {
    /// Judge a caller's key size, before a password is touched or a byte is drawn.
    ///
    /// # Errors
    ///
    /// [`Error::EncryptParams`] with `param = `[`EncryptParam::KeySize`] and
    /// `problem = `[`EncryptParamProblem::OutsideSpecRange`] for anything that is not
    /// 128, 192 or 256.
    ///
    /// **Every refusal here is the format's, and [`EncryptParamProblem::UnsupportedByCipher`]
    /// has no producer on this path.** That is not an oversight and it is the one place
    /// this module's taxonomy departs from `encrypt_params`': there, `ST_KeyBits`
    /// (§2.3.4.10) is deliberately generic across cipher algorithms — `minInclusive="8"`,
    /// a multiple of 8, **no maximum** — so the format and AES are two authorities with
    /// different reach and a value can satisfy one and fail the other (512 is a legal
    /// `ST_KeyBits` and not an AES key). Here §2.3.4.5 names AES for this stream and then
    /// enumerates its three key sizes itself, so the format's set and the cipher's set are
    /// the *same* set and the format states it first. Reporting 64 as "permitted by
    /// [MS-OFFCRYPTO], not defined by AES" would tell a consumer this header may carry a
    /// 64-bit key; §2.3.4.5 says it may not. That is the typed lie
    /// [`EncryptParamProblem`]'s split exists to prevent, pointed the other way.
    ///
    /// §2.3.2's `KeySize` table does carry rows that are not AES sizes — `0x00000000`
    /// ("determined by `Flags`") and RC4's `0x00000028`–`0x00000080` — and they are not a
    /// counter-example: they are what that field may hold in *some* `EncryptionHeader`,
    /// and §2.3.4.5 is the clause that governs the one being written here.
    ///
    /// `min` and `max` carry the nearest accepted size and are equal, per
    /// [`Error::EncryptParams`]'s field docs: the accepted values are a set of three, and
    /// a span would name sizes between them that the next call refuses.
    pub(crate) fn new(key_bits: u32) -> Result<Self, Error> {
        for (bits, alg_id) in AES_ROWS {
            if bits == key_bits {
                return Ok(Self { bits, alg_id });
            }
        }
        let nearest = nearest_supported_key_bits(key_bits);
        Err(Error::EncryptParams {
            param: EncryptParam::KeySize,
            problem: EncryptParamProblem::OutsideSpecRange,
            got: key_bits,
            min: nearest,
            max: nearest,
        })
    }

    /// `EncryptionHeader.KeySize` — the number written into the header, in bits.
    pub(crate) fn bits(self) -> u32 {
        self.bits
    }

    /// `EncryptionHeader.AlgID` — the identifier §2.3.2 pairs with [`Self::bits`].
    pub(crate) fn alg_id(self) -> u32 {
        self.alg_id
    }

    /// The AES key length in bytes, which is what `derive_standard_key` cuts the ladder
    /// to and what `standard::aes_ecb_encrypt` dispatches the cipher on.
    ///
    /// Exact for every value [`Self::new`] admits: §2.3.2 requires `KeySize` to be a
    /// multiple of 8, and all three of ours are. No rounding, and no cast that can lose
    /// anything — 256 / 8 is 32 on every target.
    pub(crate) fn key_len(self) -> usize {
        (self.bits / 8) as usize
    }
}

const _: () = {
    assert!(AES_ROWS[0].0 == limits::STANDARD_KEY_BITS_AES[0]);
    assert!(AES_ROWS[1].0 == limits::STANDARD_KEY_BITS_AES[1]);
    assert!(AES_ROWS[2].0 == limits::STANDARD_KEY_BITS_AES[2]);
    assert!(AES_ROWS[0].1 == ALG_ID_AES_128);
    assert!(AES_ROWS[1].1 == ALG_ID_AES_192);
    assert!(AES_ROWS[2].1 == ALG_ID_AES_256);
    // The default is a row of the table rather than a fourth number beside it.
    assert!(AES_ROWS[0].0 == DEFAULT_KEY_BITS);
};

/// The accepted key size nearest what the caller asked for — the single value
/// [`Error::EncryptParams`]'s `min` and `max` both carry for a refusal here.
///
/// `encrypt_params::nearest_supported_key_bits` is the same function over
/// `limits::AGILE_KEY_BITS_ALLOWED`, and its doc comment carries the argument in full:
/// the accepted values are a **set**, so a `min < max` span would print numbers —
/// 200 among them — that the very next call refuses, and a refusal naming a value that
/// does not work is worse than a vague one. It is written twice rather than shared
/// because the two lists are two different facts with two different citations
/// (`ST_KeyBits` under §2.3.4.10 there, §2.3.4.5's header table here) that happen to hold
/// the same three numbers; `limits.rs` keeps them apart for that reason and sharing the
/// lookup would quietly join them.
///
/// The tie-break is `<=` over an ascending list, so 160 — equidistant from 128 and 192 —
/// resolves upwards to the stronger key, matching the sibling exactly.
fn nearest_supported_key_bits(requested: u32) -> u32 {
    let mut nearest = AES_ROWS[0].0;
    for (bits, _) in AES_ROWS {
        // `abs_diff`, not a subtraction: `requested` may be either side of a candidate,
        // and a wrapping subtraction on a caller's `u32` would name the wrong size.
        if bits.abs_diff(requested) <= nearest.abs_diff(requested) {
            nearest = bits;
        }
    }
    nearest
}

/// `EncryptionVerifier.SaltSize` and the salt's length — [MS-OFFCRYPTO] §2.3.3 fixes 16.
pub(crate) const SALT_LEN: usize = 16;

/// The random verifier's length — [MS-OFFCRYPTO] §2.3.3, 16 bytes; also one AES block,
/// so `EncryptedVerifier` needs no padding.
pub(crate) const VERIFIER_LEN: usize = 16;

/// `EncryptionVerifier.VerifierHashSize` — the SHA-1 digest length, 20.
const VERIFIER_HASH_SIZE: u32 = 20;
const _: () = assert!(VERIFIER_HASH_SIZE as usize == SHA1_LEN);

/// `EncryptedVerifierHash` — the 20-byte digest padded to 32 for AES, §2.3.3.
pub(crate) const ENCRYPTED_VERIFIER_HASH_LEN: usize = 32;
/// `generate` writes the digest into the first `SHA1_LEN` bytes of a slot this long, so a
/// digest that outgrew the blob would be a slice panic rather than a refusal. It cannot
/// for SHA-1, but "cannot" is an argument and this is a check: if either constant ever
/// moves — a hash generalisation, a format variant — this fails the build instead.
const _: () = assert!(SHA1_LEN <= ENCRYPTED_VERIFIER_HASH_LEN);

/// The AES block size — the multiple the package tail is padded to.
const AES_BLOCK_LEN: usize = 16;

/// The chunk the package is encrypted in. Under ECB the chunking changes nothing about
/// the output — every block is independent — so this is a working-set size, chosen to
/// match `standard::decrypt_package`'s so the two loops read alike.
const CHUNK_LEN: usize = 4096;

/// One generated standard encryptor: what goes in the header, plus the key that does not.
pub(crate) struct StandardKeyMaterial {
    /// The key size this material was generated for — what the header must declare, so
    /// that the `AlgID` and `KeySize` written into the file are the same statement the
    /// key was cut to and not a second one made beside it.
    pub(crate) key_size: AesKeySize,
    /// `EncryptionVerifier.Salt` — public, the KDF salt.
    pub(crate) salt: [u8; SALT_LEN],
    /// The AES key that encrypts the package and both verifier blobs, at
    /// [`AesKeySize::key_len`] bytes. **The only secret here**, and — unlike agile's
    /// session key — derived from the password rather than drawn, so it changes when the
    /// password does.
    pub(crate) key: DerivedKey,
    /// `EncryptionVerifier.EncryptedVerifier` — public ciphertext.
    pub(crate) encrypted_verifier: [u8; VERIFIER_LEN],
    /// `EncryptionVerifier.EncryptedVerifierHash` — public ciphertext.
    pub(crate) encrypted_verifier_hash: [u8; ENCRYPTED_VERIFIER_HASH_LEN],
}

/// Generate a standard encryptor from an **injected** RNG — [MS-OFFCRYPTO] §2.3.4.8
/// (`Password Verifier Generation (Standard Encryption)`) over the key of §2.3.4.7, and
/// the `EncryptionVerifier` construction steps of §2.3.3.
///
/// Deterministic given `rng`, which is the point (plan D3). The production path is
/// [`encrypt()`] under `rand::rngs::SysRng`, not a second copy of this function.
///
/// # Errors
///
/// [`Error::RandomSource`] if the RNG will not produce bytes;
/// [`Error::BadParameters`] and [`Error::CipherError`] are
/// propagated from the key derivation and the cipher rather than unwrapped, though both
/// are unreachable for every [`AesKeySize`]: §2.3.4.7's ladder yields 40 bytes and the
/// largest key it is cut to is AES-256's 32.
pub(crate) fn generate<R: TryRng + TryCryptoRng>(
    password: &str,
    key_size: AesKeySize,
    rng: &mut R,
) -> Result<StandardKeyMaterial, Error> {
    // The salt, public: it seeds the KDF and is written into the header. Its length is
    // 16 whatever the key size is — §2.3.3 fixes `SaltSize`, and the AES key length is
    // nowhere in that sentence — so the draw order and the draw *sizes* are the same for
    // all three rows, which is why the AES-128 golden is unmoved by this parameter
    // existing.
    let mut salt = [0u8; SALT_LEN];
    fill(rng, &mut salt)?;

    // The same derivation the reader runs — one KDF, not a writer's and a reader's — and
    // one KDF across the three key sizes as well: §2.3.4.7 branches on the length only in
    // its last step, "the first cbRequiredKeyLength bytes of X3" (see `standard`'s header).
    let key = derive_standard_key(password, &salt, key_size.key_len())?;

    // The verifier: 16 random bytes, held wrapped. Not a key, but its hash is how a
    // password guess is confirmed offline, so it gets no more exposure than the key.
    let verifier = VerifierPlaintext::from_rng(VERIFIER_LEN, rng).map_err(random_source)?;
    let encrypted_verifier =
        verifier.with_secret(|v| key.with_secret(|k| aes_ecb_encrypt(k, v)))?;

    // SHA1(verifier), zero-padded to the 32-byte AES blob. Zeros, not 0x36: Word compares
    // the whole decrypted blob (GH #13), and LibreOffice's standard writer pads the same
    // way (`Standard2007Engine.cxx:60`, behaviour only).
    //
    // The wrapping happens INSIDE the closure, and that placement is the point: built the
    // obvious way -- `VerifierPlaintext::new(with_secret(|v| Sha1::digest(v).to_vec()))` --
    // the digest is an unprotected `Vec` from the moment it is computed until the outer
    // constructor closes over it. The value is wrapped, so an audit asking "is this
    // wrapped?" sees a yes; the gap is *where*. Worse, `to_vec()` allocated exactly 20
    // bytes and `resize` to 32 could not fit, so it reallocated and freed the block
    // holding SHA1(verifier) without wiping it -- the same defect `rc4_cryptoapi.rs`
    // carried, measured here at 20 abandoned bytes.
    //
    // `new_with` closes both: one allocation at the final size, so nothing is abandoned,
    // and the tail is zero by secure-gate's guarantee rather than by a `resize` that
    // reallocates to write it.
    let verifier_hash = verifier.with_secret(|v| {
        VerifierPlaintext::new_with(ENCRYPTED_VERIFIER_HASH_LEN, |slot| {
            slot[..SHA1_LEN].copy_from_slice(&Sha1::digest(v));
        })
    });
    let encrypted_verifier_hash =
        verifier_hash.with_secret(|h| key.with_secret(|k| aes_ecb_encrypt(k, h)))?;

    Ok(StandardKeyMaterial {
        key_size,
        salt,
        key,
        // ECB preserves length, so these conversions cannot fail; they are mapped rather
        // than unwrapped because a cipher that returned the wrong length would be exactly
        // a cipher error, and this crate does not panic on a code path a caller reaches.
        encrypted_verifier: encrypted_verifier
            .try_into()
            .map_err(|_| Error::CipherError)?,
        encrypted_verifier_hash: encrypted_verifier_hash
            .try_into()
            .map_err(|_| Error::CipherError)?,
    })
}

/// Serialise the whole `\EncryptionInfo` stream, header included —
/// [MS-OFFCRYPTO] §2.3.4.5.
///
/// Infallible by construction: every length is fixed by the array types, every other
/// field is either a constant of this module or one of the two numbers an
/// [`AesKeySize`] carries — and that type cannot hold a pair §2.3.4.5 forbids — and
/// nothing here is attacker-chosen or secret. The layout, byte for byte:
///
/// ```text
/// u16 vMajor | u16 vMinor | u32 Flags (copy) | u32 EncryptionHeaderSize
/// u32 Flags | SizeExtra | AlgID | AlgIDHash | KeySize | ProviderType | Reserved1 | Reserved2
/// CSPName as UTF-16LE + 0x0000
/// u32 SaltSize | Salt[16] | EncryptedVerifier[16] | u32 VerifierHashSize | EncryptedVerifierHash[32]
/// ```
///
/// 224 bytes, the length of the fixture's stream; the test that pins the layout diffs
/// against the fixture with exactly the three non-conforming fields replaced.
pub(crate) fn write_encryption_info(
    key_size: AesKeySize,
    salt: &[u8; SALT_LEN],
    encrypted_verifier: &[u8; VERIFIER_LEN],
    encrypted_verifier_hash: &[u8; ENCRYPTED_VERIFIER_HASH_LEN],
) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 + 4 + HEADER_FIXED_LEN + CSP_NAME_LEN + 72);

    // EncryptionVersionInfo, then the Flags copy [MS-OFFCRYPTO] §2.3.4.5 puts ahead of
    // the header. The fixture carries zero here; every reader ignores it; the spec says
    // "a copy", so a copy is what is written.
    out.extend_from_slice(&STANDARD_VERSION.0.to_le_bytes());
    out.extend_from_slice(&STANDARD_VERSION.1.to_le_bytes());
    out.extend_from_slice(&FLAGS.to_le_bytes());
    out.extend_from_slice(&ENCRYPTION_HEADER_SIZE.to_le_bytes());

    // EncryptionHeader — [MS-OFFCRYPTO] §2.3.2. Reserved1 is "undefined and MUST be
    // ignored"; Office writes zero and so does this.
    for field in [
        FLAGS,
        0, // SizeExtra: MUST be 0
        // `AlgID` and `KeySize` from one `AesKeySize`, never assembled from two sources:
        // §2.3.2 pairs them and the reader cross-checks the pair.
        key_size.alg_id(),
        ALG_ID_HASH_SHA1,
        key_size.bits(),
        PROVIDER_TYPE_AES,
        0, // Reserved1
        0, // Reserved2: MUST be 0
    ] {
        out.extend_from_slice(&field.to_le_bytes());
    }
    for unit in CSP_NAME.encode_utf16().chain(std::iter::once(0u16)) {
        out.extend_from_slice(&unit.to_le_bytes());
    }

    // EncryptionVerifier — [MS-OFFCRYPTO] §2.3.3.
    out.extend_from_slice(&(SALT_LEN as u32).to_le_bytes());
    out.extend_from_slice(salt);
    out.extend_from_slice(encrypted_verifier);
    out.extend_from_slice(&VERIFIER_HASH_SIZE.to_le_bytes());
    out.extend_from_slice(encrypted_verifier_hash);
    out
}

/// Encrypt a package into the `EncryptedPackage` stream — [MS-OFFCRYPTO] §2.3.4.4: the
/// 8-byte little-endian `StreamSize`, then the plaintext under AES-ECB with its tail
/// zero-padded to a block. `StreamSize` is the *unencrypted* length, so "the actual size
/// of the \\EncryptedPackage stream can be larger than this value" — which is what the
/// reader's truncation undoes.
///
/// No segmentation and no IV: ECB has neither, which is what makes this the weaker
/// format. The reader (`standard::decrypt_package`) truncates the padding away against
/// the prefix, and refuses a prefix larger than the ciphertext — so the prefix is the
/// plaintext length exactly, never the padded length.
pub(crate) fn encrypt_package(plaintext: &[u8], key: &DerivedKey) -> Result<Vec<u8>, Error> {
    let mut stream = Vec::with_capacity(8 + plaintext.len() + AES_BLOCK_LEN);
    stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    for chunk in plaintext.chunks(CHUNK_LEN) {
        // Only the final chunk can be short of a block multiple; 4096 is one itself.
        let mut padded = chunk.to_vec();
        padded.resize(chunk.len().div_ceil(AES_BLOCK_LEN) * AES_BLOCK_LEN, 0);
        let block = key.with_secret(|k| aes_ecb_encrypt(k, &padded))?;
        stream.extend_from_slice(&block);
    }
    Ok(stream)
}

/// Encrypt an OOXML package with a password into a complete Office 2007 container — the
/// whole standard write path, assembled, with the randomness injected.
///
/// The seeded entry point behind [`crate::encrypt_ooxml_standard`], which is this plus
/// the shape guard and `rand::rngs::SysRng`. Everything a test can prove about
/// production goes through here: the committed golden under a seeded
/// `chacha20::ChaCha12Rng`, and the external readers opening what a seeded run wrote.
///
/// **This function encrypts whatever bytes it is handed.** The check that the input is a
/// plain OOXML package is [`crate::check_encryptable`]'s, and it lives in the public
/// wrapper rather than here — which is why this one is `pub(crate)`. See the agile
/// writer's note for why that placement is deliberate.
///
/// ```text
/// key_size  = AesKeySize::new(key_bits)                           §2.3.4.5, refuse first
/// material  = generate(password, key_size, rng)                   §2.3.4.7-8
/// package   = LE64(len) || AES-ECB(package || zero pad, key)      §2.3.4.4
/// info      = write_encryption_info(key_size, salt, verifiers)    §2.3.4.5
/// container = dataspaces::build_container(info, package)          §2.1
/// ```
///
/// # The key size is judged before anything else happens
///
/// `AesKeySize::new` runs before the payload ceiling, before the RNG is touched and
/// before the 50 000-round derivation — the `agile_encrypt::encrypt` ordering, for the
/// same reason: a caller that asked for a key size this format does not define should
/// not first pay for a password prompt, and a refusal that consumed a draw would make
/// the seeded goldens depend on which calls failed.
///
/// The `\x06DataSpaces` subtree is the agile one, unchanged: measured on
/// `standard_encrypted.docx` against `word16_agile.docx`, all four streams are
/// byte-identical between the two formats, so `build_container` is shared rather than
/// parameterised.
///
/// # Errors
///
/// [`Error::EncryptParams`] if `key_bits` is not one of the three key sizes §2.3.4.5
/// defines — see [`AesKeySize::new`];
/// [`Error::BadParameters`] if `package` exceeds
/// [`limits::PAYLOAD_CEILING`] — the same 1 GiB the decrypt side refuses, checked on the
/// input so that a file this crate writes is a file it reads back;
/// [`Error::RandomSource`] if `rng` will not produce bytes;
/// [`Error::Io`] if the in-memory container cannot be written.
pub(crate) fn encrypt<R: TryRng + TryCryptoRng>(
    package: &[u8],
    password: &str,
    key_bits: u32,
    rng: &mut R,
) -> Result<Vec<u8>, Error> {
    let key_size = AesKeySize::new(key_bits)?;

    if package.len() > limits::PAYLOAD_CEILING {
        return Err(Error::BadParameters(format!(
            "the package is {} bytes; this crate encrypts at most {} (see \
             limits::PAYLOAD_CEILING), because that is what it will read back",
            package.len(),
            limits::PAYLOAD_CEILING
        )));
    }

    let material = generate(password, key_size, rng)?;
    let encrypted_package = encrypt_package(package, &material.key)?;
    let info = write_encryption_info(
        material.key_size,
        &material.salt,
        &material.encrypted_verifier,
        &material.encrypted_verifier_hash,
    );
    dataspaces::build_container(&info, &encrypted_package)
}

#[cfg(test)]
#[path = "standard_encrypt_tests.rs"]
mod tests;
