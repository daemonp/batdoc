//! The encryption parameters a caller may choose, and the one place they are checked.
//!
//! [`EncryptParams`] is the agile write path's tuple — spin count, hash, the two key
//! sizes and the two salt sizes — as a value a caller can build, inspect and submit for
//! judgement *before* it pays for a password. Until it existed the tuple was six
//! constants in `encryption_info` and nothing about it was a caller's choice; the type is
//! the seam that turns "what this crate writes" into "what this call writes".
//!
//! # Why this is its own module
//!
//! It is the parameters and their validity, and nothing else. `agile_encrypt` performs
//! the key schedule, `encryption_info` serialises the result, and neither needs to own a
//! type whose whole job is to be handed in from outside — a module that both *decided*
//! the parameters and *used* them would be the "and" CLAUDE.md § *Layout* rules out, and
//! it would put the refusal a caller wants to ask for cheaply behind a function that has
//! already started encrypting.
//!
//! There is no `unwrap`, `expect`, `panic` or slice index over caller data here, so this
//! module carries no `#![deny]` header of its own — unlike `classify` and the legacy
//! walks, nothing in it parses a file. Every number it reads came from the caller's own
//! struct literal, and the only indexing it does is a constant index into a fixed-size
//! array in `limits`, which the compiler bounds-checks at compile time.
//!
//! # The spec sections this module encodes
//!
//! * §2.3.4.10 — the `CT_KeyData` and `CT_PasswordKeyEncryptor` schemas, and the two
//!   cross-element MUSTs (`hashAlgorithm`, `cipherAlgorithm`) that are the reason the
//!   hash is one field and the key and salt sizes are two each.
//! * §2.3.4.11 — the password-derived key, whose size is `PasswordKeyEncryptor.keyBits`.
//! * §2.3.4.13 step 1 — the package and intermediate key, whose size is
//!   `Encryptor.KeyData.keyBits`.

use crate::error::{EncryptParam, EncryptParamProblem, Error};
use crate::hash::HashAlgorithm;
use crate::limits;

/// `ST_KeyBits`'s floor: the format's own rule about `keyBits`, and nothing to do with AES.
///
/// [MS-OFFCRYPTO] §2.3.4.10 gives the simple type as
/// `<xs:restriction base="xs:unsignedInt"><xs:minInclusive value="8" /></xs:restriction>`
/// and the prose beneath the schema as "KeyBits: … It MUST be at least 8 and a multiple of
/// 8." Those are the whole of the format's opinion: **there is no `maxInclusive`**, because
/// the attribute is generic over cipher algorithms, so `keyBits="512"` is a conforming
/// value that AES simply has no key for.
///
/// It lives here rather than in `limits` only because this is the one place that asks the
/// question — `limits::crypto::AGILE_KEY_BITS_ALLOWED` is the *cipher* fact, which is a
/// different row of that module's provenance table and cannot stand in for this one.
const ST_KEY_BITS_MIN: u32 = 8;

/// The smallest key size AES defines; the seed and the fallback of
/// [`nearest_supported_key_bits`].
const KEY_BITS_SMALLEST: u32 = limits::AGILE_KEY_BITS_ALLOWED[0];

/// `AGILE_KEY_BITS_ALLOWED` is ascending, and has exactly three entries.
///
/// Both halves are load-bearing and both fail *here*, at compile time, rather than in a
/// message. Ascending is what makes [`KEY_BITS_SMALLEST`] the smallest rather than merely
/// the first, and it is what makes [`nearest_supported_key_bits`]'s tie-break resolve
/// upwards to the stronger key instead of to whichever entry happened to be written last.
/// The length pins the indexing to a list that has not grown a fourth AES key size behind
/// this module's back.
const _: () = {
    assert!(limits::AGILE_KEY_BITS_ALLOWED.len() == 3);
    assert!(limits::AGILE_KEY_BITS_ALLOWED[0] < limits::AGILE_KEY_BITS_ALLOWED[1]);
    assert!(limits::AGILE_KEY_BITS_ALLOWED[1] < limits::AGILE_KEY_BITS_ALLOWED[2]);
    assert!(limits::AGILE_KEY_BITS_ALLOWED[0] >= ST_KEY_BITS_MIN);
    assert!(limits::AGILE_KEY_BITS_ALLOWED[0] % 8 == 0);
    assert!(limits::AGILE_KEY_BITS_ALLOWED[1] % 8 == 0);
    assert!(limits::AGILE_KEY_BITS_ALLOWED[2] % 8 == 0);
};

/// The one accepted key size a `keyBits` refusal points the caller at.
///
/// # Why a refusal reports one value twice instead of the ends of the set
///
/// [`Error::EncryptParams`] renders its payload as `this crate accepts {min}..={max}`, and
/// [`limits::AGILE_KEY_BITS_ALLOWED`] is `[128, 192, 256]` — a **set of three**, not a
/// span. Filling that span with the ends of the set is what this module used to do, and it
/// made the message assert something false: 200 is inside `128..=256` and `validate`
/// refuses it. A refusal that names a value the very next call rejects is worse than a
/// vague one, because the caller acts on it.
///
/// So every `keyBits` refusal carries `min == max`: one size AES really does define, the
/// one nearest what was asked, ties resolving upwards to the stronger key. The span is
/// degenerate, so there is no "between the ends" left to be wrong about, and every number
/// the sentence prints is a number `validate` accepts.
///
/// What that gives up is completeness — the other two sizes are not in the payload — and
/// that is the right thing to give up here. A caller shown `128..=128` who sends 128
/// succeeds; a caller shown `128..=256` who sends 200 does not. The full set is
/// [`limits::AGILE_KEY_BITS_ALLOWED`], and it is named in [`EncryptParams::key_data_key_bits`],
/// in [`EncryptParams::password_key_bits`] and in `EncryptParam::KeyBits`'s own docs, which
/// is where a set belongs — in prose that can say "one of", rather than in two `u32`s that
/// can only say "from … to".
///
/// The cost is stated rather than hidden: [`Error::EncryptParams`]'s field docs call `min`
/// and `max` the smallest and largest value this crate accepts for the parameter, and for a
/// set-valued parameter that reading and the rendered sentence cannot both be satisfied by
/// one pair of numbers. This module resolves it in favour of the sentence, because the
/// sentence is what a consumer shows a human.
fn nearest_supported_key_bits(requested: u32) -> u32 {
    let mut nearest = KEY_BITS_SMALLEST;
    for candidate in limits::AGILE_KEY_BITS_ALLOWED {
        // `<=` with an ascending list is the tie-break: 160 is 32 from both 128 and 192,
        // and the later — larger — candidate wins. `abs_diff` rather than a subtraction
        // because `requested` may be either side of a candidate and a wrapping subtraction
        // on a caller's `u32` would silently name the wrong size.
        if candidate.abs_diff(requested) <= nearest.abs_diff(requested) {
            nearest = candidate;
        }
    }
    nearest
}

/// The parameters of an agile encryption, as the caller chooses them.
///
/// Every field is a value [MS-OFFCRYPTO] §2.3.4.10 puts on `<keyData>` or
/// `<p:encryptedKey>` and that a *writer* is free to pick. Attributes a writer is not
/// free to pick are deliberately absent, and their absence is a statement:
///
/// * `hashSize` MUST equal the named hash's digest length (§2.3.4.10), so it is derived
///   from [`Self::hash`] and never chosen — a field for it could only ever hold a value
///   that disagrees with the hash, i.e. an unwritable file.
/// * `blockSize` is 16 because this crate writes AES-CBC and AES's block is 16 bytes. It
///   is fixed by the cipher, not by an opinion.
/// * `cipherAlgorithm` and `cipherChaining` are `AES` / `ChainingModeCBC` for the same
///   reason, and §2.3.4.10 additionally requires `p:encryptedKey`'s to equal `keyData`'s.
///
/// # Why `keyBits` and `saltSize` are two fields each and `hash` is one
///
/// This is the shape of the spec, not a convenience. §2.3.4.10 states exactly two
/// cross-element equalities, both on `CT_PasswordKeyEncryptor`:
///
/// > hashAlgorithm: … The hashing algorithm specified MUST be the same as the hashing
/// > algorithm specified for the Encryption.keyData element.
///
/// > cipherAlgorithm: … The cipher algorithm specified MUST be the same as the cipher
/// > algorithm specified for the Encryption.keyData element.
///
/// It states **no such rule for `keyBits`, `saltSize` or `cipherChaining`** — and for
/// `keyBits` it goes further and names the two quantities apart, in the two sections that
/// consume them:
///
/// > §2.3.4.13 step 1: Generate a random array of bytes that is the same size as
/// > specified by the `Encryptor.KeyData.keyBits` attribute of the parent element.
///
/// > §2.3.4.11: If the size of the resulting `Hfinal` is smaller than that of
/// > `PasswordKeyEncryptor.keyBits`, the key MUST be padded …
///
/// Those are the package/intermediate key and the key-encrypting key that wraps it: two
/// keys, two sizes, and a file in which they differ is conforming. `saltSize` is
/// per-element in the same way — §2.3.4.10 binds each one only to *its own* element's
/// `saltValue`, whose decoded form "MUST be" that many bytes.
///
/// A single `key_bits` field would therefore encode a constraint the spec denies, and
/// would make an AES-256 KEK wrapping an AES-128 package key **inexpressible** through an
/// API whose whole purpose is to express what the format allows. A single `hash` field
/// encodes a constraint the spec asserts, which is the opposite thing: two hash fields
/// would let a caller ask for a non-conforming file this crate would then have to refuse.
/// The reader half already treats the two hashes separately (see `hash`'s module header —
/// a *reader* is handed bytes, not a promise), and that asymmetry between the two halves
/// is intentional.
///
/// # Not `#[non_exhaustive]`, deliberately
///
/// `#[non_exhaustive]` on a struct forbids struct-expression syntax from another crate —
/// and `..Default::default()` **is** struct-expression syntax. The two are mutually
/// exclusive, and the update syntax is the chosen ergonomics: a caller that wants Office
/// 16's tuple with one value moved writes
///
/// ```
/// use msoffice_crypto::{EncryptParams, HashAlgorithm};
///
/// let params = EncryptParams {
///     hash: HashAlgorithm::Sha256,
///     ..Default::default()
/// };
/// assert_eq!(params.spin_count, 100_000);
/// params.validate().expect("SHA-256 carries a 256-bit key");
/// ```
///
/// which is both the common case and the one that survives a field being added here.
/// Adding `#[non_exhaustive]` "for forward compatibility" would break exactly that line
/// in every consumer, which is why this paragraph exists: it is the kind of attribute
/// someone adds later as a tidy-up without noticing what it costs.
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{EncryptParam, EncryptParamProblem, EncryptParams, Error, HashAlgorithm};
///
/// // SHA-1's digest is 20 bytes; a 256-bit KEK wants 32 of them.
/// let err = EncryptParams {
///     hash: HashAlgorithm::Sha1,
///     ..Default::default()
/// }
/// .validate()
/// .unwrap_err();
/// assert!(matches!(
///     err,
///     Error::EncryptParams {
///         param: EncryptParam::KeyBitsWithHash,
///         problem: EncryptParamProblem::UnusableCombination,
///         ..
///     }
/// ));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg(feature = "crypto-ops")]
pub struct EncryptParams {
    /// `p:encryptedKey/@spinCount`: how many times the password hash is iterated
    /// ([MS-OFFCRYPTO] §2.3.4.11).
    ///
    /// **`p:encryptedKey` only.** `CT_KeyData` has no `spinCount` attribute at all —
    /// there is nothing to stretch on that element, whose key is random rather than
    /// password-derived — which is why this field is single where `keyBits` and
    /// `saltSize` are paired.
    ///
    /// `ST_SpinCount` is `0..=10000000` (§2.3.4.10). Both ends are the spec's; this crate
    /// adds nothing to either, and `limits::SPIN_COUNT_MAX` carries the measurement.
    pub spin_count: u32,

    /// `hashAlgorithm` **on both elements**, because §2.3.4.10 requires them equal — see
    /// the type's doc for the quoted MUST.
    ///
    /// §2.3.4.10 also says of this attribute that "values that are not defined MAY be
    /// used, and a compliant implementation is not required to support all defined
    /// values", so writing only the four [`HashAlgorithm`] carries is *explicitly
    /// conforming* rather than a gap. `hashSize` follows from this field and is not a
    /// parameter.
    pub hash: HashAlgorithm,

    /// `keyData/@keyBits`: the size of the **package key**, which is also the size of the
    /// intermediate key that key is.
    ///
    /// §2.3.4.13 step 1 sizes it: "Generate a random array of bytes that is the same size
    /// as specified by the `Encryptor.KeyData.keyBits` attribute of the parent element."
    /// This is the key the document body is encrypted under and the one that gets wrapped
    /// into `encryptedKeyValue`.
    ///
    /// `ST_KeyBits` sets `minInclusive="8"`, requires a multiple of 8 and states **no
    /// maximum**, being generic over cipher algorithms. [`EncryptParams::validate`] asks
    /// both questions and reports them apart: a value below 8 or not a multiple of 8 is
    /// [`EncryptParamProblem::OutsideSpecRange`], the format's own rule; anything else
    /// outside `limits::AGILE_KEY_BITS_ALLOWED` — one of 128, 192 and 256 — is
    /// [`EncryptParamProblem::UnsupportedByCipher`], which is AES's rule and not a claim
    /// about the format.
    ///
    /// Unlike [`Self::password_key_bits`] this value is **not** coupled to the hash: the
    /// package key is random bytes from the system RNG, not a slice of a digest.
    pub key_data_key_bits: u32,

    /// `p:encryptedKey/@keyBits`: the size of the **key-encrypting key** derived from the
    /// password, which wraps the package key.
    ///
    /// §2.3.4.11 sizes it, in the sentence that also explains the coupling
    /// [`EncryptParams::validate`] enforces: "If the size of the resulting `Hfinal` is
    /// smaller than that of `PasswordKeyEncryptor.keyBits`, the key MUST be padded by
    /// appending bytes with a value of 0x36." This crate declines to write a file that
    /// needs that pad — see `agile::derive_block_key` for the decrypt-side argument —
    /// so this field, and only this one, must fit inside [`Self::hash`]'s digest.
    ///
    /// `ST_KeyBits` and AES's three key sizes bind it exactly as they bind
    /// [`Self::key_data_key_bits`] — both checks run on both fields, and both run *before*
    /// the digest coupling, so a value that is not a legal `keyBits` is never reported as
    /// a problem with the pair.
    pub password_key_bits: u32,

    /// `keyData/@saltSize`: the length in bytes of `keyData/@saltValue`.
    ///
    /// §2.3.4.10 binds it to that one attribute — the decoded `saltValue` "MUST be"
    /// `saltSize` bytes — and to nothing on the other element. `ST_SaltSize` is
    /// `1..=65536` (`limits::AGILE_SALT_SIZE`), and this crate accepts the spec's whole
    /// range rather than a margin of its own.
    pub key_data_salt_size: u32,

    /// `p:encryptedKey/@saltSize`: the length in bytes of `p:encryptedKey/@saltValue`.
    ///
    /// Independent of [`Self::key_data_salt_size`] for the reason given there, and doing
    /// double duty in the algorithm: §2.3.4.12 makes this salt the IV for the three
    /// `p:encryptedKey` blobs, fitted to `blockSize` — padded with 0x36 if short,
    /// truncated if long. `ST_SaltSize` is `1..=65536` here as it is there, so the spec
    /// permits any of those lengths and has a rule for every one of them.
    ///
    /// **The writer applies that fit**, so the whole range is writable.
    /// `agile_encrypt::generate` puts this salt through `hash::fit_iv` once and uses the
    /// result as the IV for all three password blobs, which is what the decrypt side has
    /// always done (`agile::AgileParams::password_blob_iv`). For one release it did not:
    /// the raw salt went to `aes_cbc_encrypt`, and `agile::check_cbc_lengths` would have
    /// refused anything but 16 bytes. At 16 the fit is the identity, so the two paths
    /// agreed by coincidence and no test could see the difference until this field made
    /// the length a caller's choice.
    ///
    /// **Writable is not the same as opened**, and the distinction is the honest one
    /// here: no external reader has been measured on a salt size that is not a multiple
    /// of 16. It changes the `0x00` pad on `encryptedVerifierHashInput`
    /// (`roundUp(saltSize, blockSize)`), and Word is documented rejecting wrong pad
    /// *bytes* twice elsewhere in this format — `0x800A1520` on the verifier value,
    /// `0x800A1066` on the integrity blobs, thirteen measured variants deep. Until that
    /// measurement exists, a non-multiple is a conforming file this crate writes and
    /// nobody has confirmed anyone opens.
    pub password_salt_size: u32,
}

#[cfg(feature = "crypto-ops")]
impl Default for EncryptParams {
    /// The tuple **Office 16 writes**, measured rather than composed.
    ///
    /// Written out rather than `#[derive(Default)]`, which could not produce these
    /// numbers anyway (`u32::default()` is 0 and [`HashAlgorithm`] has no `Default`), but
    /// more to the point: each of these six values is a measurement, and a measurement
    /// wants its provenance next to it. Every one of them was read out of
    /// `tests/fixtures/{word16_agile.docx, excel16_agile.xlsx, powerpoint16_agile.pptx}`,
    /// whose `EncryptionInfo` streams are byte-identical at 1 289 bytes apart from the
    /// base64 values — see `encryption_info`'s module header for that measurement and
    /// `encryption_info_tests` for the byte-identity proof it feeds.
    ///
    /// * `spin_count`: 100 000, named once as `encryption_info::OFFICE_SPIN_COUNT` and
    ///   read from there rather than retyped. That constant is also what the writer's own
    ///   doc cites, so the two cannot drift apart into a claim and a number.
    /// * `hash`: SHA-512, and with it a 64-byte `hashSize`.
    /// * both `keyBits`: 256 — AES-256 for the package key and for the KEK alike. Office
    ///   writes them equal; the spec does not require it (see the type's doc), and the
    ///   two fields exist so that a caller can part them.
    /// * both `saltSize`: 16, which is also `blockSize`, so §2.3.4.12's fit of the
    ///   `p:encryptedKey` salt to the block length is the identity for this tuple.
    ///
    /// Changing any of these changes what this crate emits by default, which is what
    /// every acceptance-gate verdict recorded against real Word, real LibreOffice,
    /// `msoffcrypto-tool` and `office-crypto` was measured on — so a change here
    /// invalidates that evidence and needs its own gate run, not just a green test suite.
    fn default() -> Self {
        Self {
            spin_count: crate::encryption_info::OFFICE_SPIN_COUNT,
            hash: HashAlgorithm::Sha512,
            key_data_key_bits: 256,
            password_key_bits: 256,
            key_data_salt_size: 16,
            password_salt_size: 16,
        }
    }
}

#[cfg(feature = "crypto-ops")]
impl EncryptParams {
    /// Answer "would this be written?" without writing anything — and without asking for
    /// a password first.
    ///
    /// Public for the same reason [`crate::check_encryptable`] is: the alternative is a
    /// caller prompting for a new password, or unlocking a keychain, for a call that was
    /// always going to be refused. On an interactive path that question cannot be taken
    /// back once it has been asked.
    ///
    /// **[`crate::encrypt_ooxml_with_params`] is the public entry point that takes one**,
    /// and it does not maintain a second copy of these rules: the tuple travels down
    /// intact and is judged by this function. `agile_encrypt::encrypt` calls it before
    /// the payload ceiling and before the first RNG draw, `agile_encrypt::generate`
    /// sizes every draw from the same value, and `encryption_info::write` takes and
    /// re-checks it, so every length asserted and every number written comes from a
    /// tuple that has been through here. [`crate::encrypt_ooxml`] is that same call with
    /// [`EncryptParams::default`] supplied for the caller — one line, so the default
    /// path and the parameterised path are the same code rather than two that agree.
    ///
    /// So a caller running this early asks the identical function the identical
    /// question, which is the whole point of it being public: the answer it gets is the
    /// answer the encryption will give.
    ///
    /// An `Ok(())` is not a promise the encryption succeeds — the payload ceiling and the
    /// system RNG are still ahead — only that the parameters are not what stops it.
    ///
    /// # The order of the checks, and the rule behind it
    ///
    /// **A refusal that names a single field outranks any combination that field takes
    /// part in.** Reporting [`EncryptParam::KeyBitsWithHash`] for a `keyBits` that is
    /// invalid on its own would send the caller to reconsider a pair when one member was
    /// simply wrong, and might send it to change the *other* member — the one that was
    /// fine. So [`Self::hash`] and both `keyBits` are each settled alone before the pair
    /// is looked at. The two checks left over, `saltSize` and `spinCount`, then run in
    /// the order §2.3.4.10 lists the attributes in `CT_PasswordKeyEncryptor`.
    ///
    /// Within a pair of like fields — the two `keyBits`, the two `saltSize` — `keyData`'s
    /// is examined before `p:encryptedKey`'s, matching the document order of the elements
    /// themselves. [`Error::EncryptParams`] reports the offending *value* in `got` but has
    /// no field naming which of the two elements it came from; when both are invalid the
    /// first is reported and the second is found on the next call.
    ///
    /// # Errors
    ///
    /// - [`EncryptParam::KeyBits`] / [`EncryptParamProblem::OutsideSpecRange`] — below
    ///   `ST_KeyBits`' `minInclusive="8"`, or not a multiple of 8. No conforming file may
    ///   carry the value at all; `ST_KeyBits` states **no maximum**, so nothing is too
    ///   large for the format.
    /// - [`EncryptParam::KeyBits`] / [`EncryptParamProblem::UnsupportedByCipher`] — a key
    ///   size AES does not define, having passed `ST_KeyBits`. [MS-OFFCRYPTO] permits it;
    ///   AES has no such key. 64 and 512 are both this, not the case above.
    /// - [`EncryptParam::KeyBitsWithHash`] / [`EncryptParamProblem::UnusableCombination`]
    ///   — [`Self::password_key_bits`] exceeds [`Self::hash`]'s digest. Each is fine
    ///   alone.
    /// - [`EncryptParam::SaltSize`] / [`EncryptParamProblem::OutsideSpecRange`] — outside
    ///   `ST_SaltSize`'s `1..=65536`.
    /// - [`EncryptParam::SpinCount`] / [`EncryptParamProblem::OutsideSpecRange`] — above
    ///   `ST_SpinCount`'s `10000000`.
    ///
    /// # Examples
    ///
    /// ```
    /// use msoffice_crypto::{EncryptParams, HashAlgorithm};
    ///
    /// // An AES-256 KEK wrapping an AES-128 package key: no spec rule joins the two.
    /// EncryptParams {
    ///     key_data_key_bits: 128,
    ///     password_key_bits: 256,
    ///     ..Default::default()
    /// }
    /// .validate()
    /// .expect("the two keyBits are independent quantities");
    /// ```
    pub fn validate(&self) -> Result<(), Error> {
        // 1. The hash, matched exhaustively with no `_` arm.
        //
        // No refusal exists here today and none is expected: §2.3.4.10 says a compliant
        // implementation "is not required to support all defined values", so the four
        // this crate writes are conforming on their own. The match earns its place as a
        // tripwire rather than a check. `HashAlgorithm` is `#[non_exhaustive]`, which
        // binds other crates but not this one, so a fifth variant added to it is `E0004`
        // *here* — at the one place that has to decide whether the writer may emit it —
        // rather than a value that silently flows into a key schedule that cannot size
        // it. A `_` arm, or dropping the match as a no-op, discards exactly that.
        match self.hash {
            HashAlgorithm::Sha1
            | HashAlgorithm::Sha256
            | HashAlgorithm::Sha384
            | HashAlgorithm::Sha512 => {}
        }

        // 2. Each `keyBits`, alone and before the pair check — against **two** rules, in
        //    the order that keeps the verdict honest: the format's, then the cipher's.
        //
        // These are different facts with different authors, and the error type has a
        // variant for each precisely so a consumer can tell them apart:
        //
        // * `ST_KeyBits` (§2.3.4.10) is `minInclusive="8"` and the prose adds "MUST be at
        //   least 8 and a multiple of 8". There is **no `maxInclusive`** — the type is
        //   generic over cipher algorithms. Breaking it means no conforming file could
        //   carry the value: `OutsideSpecRange`.
        // * `AGILE_KEY_BITS_ALLOWED` is what AES defines. Breaking *that* while satisfying
        //   the schema means the format permits the value and this crate's cipher has no
        //   such key: `UnsupportedByCipher`. It is the same allowlist the reader enforces
        //   on both elements (`agile::check_key_bits`), so the writer cannot author a file
        //   its own parser refuses.
        //
        // The four cases the split exists for, and what each must report:
        //
        // | `keyBits` | why it fails | verdict |
        // | --- | --- | --- |
        // | 7   | below `ST_KeyBits`' floor of 8 | `(KeyBits, OutsideSpecRange)` |
        // | 12  | at least 8, not a multiple of 8 | `(KeyBits, OutsideSpecRange)` |
        // | 64  | legal `ST_KeyBits`; AES has no 64-bit key | `(KeyBits, UnsupportedByCipher)` |
        // | 512 | legal `ST_KeyBits` — it states no maximum | `(KeyBits, UnsupportedByCipher)` |
        //
        // A single membership test cannot make that distinction, and the version of this
        // module that used one labelled all four `UnsupportedByCipher`: it told a caller
        // that 7 and 12 were fine by the format and merely beyond AES, which is backwards.
        // Reporting a spec violation as a cipher limitation is a typed lie, and a typed
        // lie is the thing the fourth `EncryptParamProblem` variant was added to prevent —
        // a consumer renders a typed variant as authoritative and repeats it to its user.
        // Hence spec first: only a value the format admits can be judged by the cipher.
        for key_bits in [self.key_data_key_bits, self.password_key_bits] {
            if key_bits < ST_KEY_BITS_MIN || key_bits % 8 != 0 {
                // The bound this check applies is `ST_KEY_BITS_MIN`, and the payload
                // deliberately does not carry it: the message renders `min..=max` as what
                // *this crate accepts*, and this crate does not accept an 8-bit key. The
                // floor is stated by `problem` — `OutsideSpecRange`'s `Display` cites
                // §2.3.4.10 — and by this module's docs, while the numbers name a size the
                // caller can actually send. See `nearest_supported_key_bits`.
                return Err(Error::EncryptParams {
                    param: EncryptParam::KeyBits,
                    problem: EncryptParamProblem::OutsideSpecRange,
                    got: key_bits,
                    min: nearest_supported_key_bits(key_bits),
                    max: nearest_supported_key_bits(key_bits),
                });
            }
            if !limits::AGILE_KEY_BITS_ALLOWED.contains(&key_bits) {
                return Err(Error::EncryptParams {
                    param: EncryptParam::KeyBits,
                    problem: EncryptParamProblem::UnsupportedByCipher,
                    got: key_bits,
                    min: nearest_supported_key_bits(key_bits),
                    max: nearest_supported_key_bits(key_bits),
                });
            }
        }

        // 3. The one coupling between two fields, and it binds the **password** key bits.
        //
        // §2.3.4.11 is explicit about which quantity is cut from the digest: "If the size
        // of the resulting Hfinal is smaller than that of PasswordKeyEncryptor.keyBits,
        // the key MUST be padded by appending bytes with a value of 0x36." Hfinal is the
        // spin hash, so it is `p:encryptedKey/@hashAlgorithm`'s digest and
        // `p:encryptedKey/@keyBits` that meet — `keyData/@keyBits` is nowhere in that
        // sentence, and cannot be: §2.3.4.13 step 1 makes the package key from the RNG,
        // not from a hash, so no digest length constrains it. `agile.rs:1143` is the
        // decrypt side of the identical pairing, and it reads the value parsed at
        // `agile.rs:1064` from `p:encryptedKey/@keyBits` — verified against both before
        // this was written.
        //
        // This crate refuses rather than pads (`agile::derive_block_key` carries the
        // argument: the pad manufactures constant bytes of key no writer ever used), so
        // the refusal is `UnusableCombination` — a house rule about a legal-but-unreadable
        // file — and not a claim that the format forbids it.
        if !self.hash.can_carry_key_bits(self.password_key_bits) {
            // The largest accepted key size this hash *can* carry — one value that would
            // have worked, reported as the same degenerate span the two checks above use
            // and for the same reason: the accepted sizes are a set, so `min < max` here
            // would print a span containing values (200 among them) that `validate`
            // refuses. See `nearest_supported_key_bits` for the argument in full.
            //
            // At least one value always qualifies — SHA-1's 20 bytes carry AES-128's 16 —
            // so the fallback is unreachable, and is `KEY_BITS_SMALLEST` rather than a
            // panic because an unreachable branch that aborts is still an abort.
            let carryable_max = limits::AGILE_KEY_BITS_ALLOWED
                .into_iter()
                .filter(|bits| self.hash.can_carry_key_bits(*bits))
                .max()
                .unwrap_or(KEY_BITS_SMALLEST);
            return Err(Error::EncryptParams {
                param: EncryptParam::KeyBitsWithHash,
                problem: EncryptParamProblem::UnusableCombination,
                got: self.password_key_bits,
                min: carryable_max,
                max: carryable_max,
            });
        }

        // 4. Each `saltSize` against `ST_SaltSize`'s own range, which this crate adopts
        //    whole — so the refusal is `OutsideSpecRange` and names the format's rule
        //    honestly. Zero is the live end in practice: a 65 536-byte salt is absurd but
        //    conforming, and this crate does not add an opinion the spec does not have.
        for salt_size in [self.key_data_salt_size, self.password_salt_size] {
            if !limits::AGILE_SALT_SIZE.contains(&salt_size) {
                return Err(Error::EncryptParams {
                    param: EncryptParam::SaltSize,
                    problem: EncryptParamProblem::OutsideSpecRange,
                    got: salt_size,
                    min: *limits::AGILE_SALT_SIZE.start(),
                    max: *limits::AGILE_SALT_SIZE.end(),
                });
            }
        }

        // 5. `spinCount` against `ST_SpinCount`'s `0..=10000000`.
        //
        // Only the ceiling is testable: the field is a `u32` and the schema's floor is
        // `minInclusive="0"`, so the type already enforces the lower end and `min: 0`
        // below reports the spec's figure rather than a check that ran. `spinCount="0"`
        // makes a weak file and not a dangerous one, and this crate declines to have an
        // opinion about a writer's stretching — see `limits::SPIN_COUNT_MAX`, the same
        // ceiling the reader applies, which is what keeps the writer from emitting a file
        // it would then refuse to read.
        if self.spin_count > limits::SPIN_COUNT_MAX {
            return Err(Error::EncryptParams {
                param: EncryptParam::SpinCount,
                problem: EncryptParamProblem::OutsideSpecRange,
                got: self.spin_count,
                min: 0,
                max: limits::SPIN_COUNT_MAX,
            });
        }

        Ok(())
    }
}

#[cfg(test)]
#[path = "encrypt_params_tests.rs"]
mod tests;
