//! Detects, decrypts and encrypts Microsoft Office documents in the formats
//! [MS-OFFCRYPTO] defines.
//!
//! This crate's job is password-protected Word, Excel and PowerPoint files: the OOXML
//! packages (`.docx` / `.xlsx` / `.pptx`) Office wraps in a CFB container when a
//! password is set, and — behind the `legacy-binary` feature — the 97-2003 binaries
//! (`.doc` / `.xls` / `.ppt`). It is not a general CFB or ZIP library, not a reader of
//! certificate-based (`CertificateKeyEncryptor`) wrapping, and not affiliated with,
//! endorsed by, or sponsored by Microsoft; see [Trademarks](#trademarks).
//!
//! The public surface is re-exported at the crate root. Detection types
//! ([`Classification`], [`Family`], [`IntegrityDeclaration`]) compile into every build.
//! `decrypt_ooxml`, `encrypt_ooxml`, `encrypt_ooxml_with_params`,
//! `encrypt_ooxml_standard`, `encrypt_ooxml_standard_with_key_bits`,
//! `check_encryptable`,
//! `EncryptParams`, `Error`, `IntegrityPolicy` and `IntegrityOutcome` exist only under
//! `crypto-ops`; `decrypt_binary_office` exists only under `legacy-binary`. Those names
//! are not linked from this page because this crate-level document renders in the
//! detection-only build, where they are absent.
//!
//! - **Agile encryption** (Office 2010+, `vMajor=4` `vMinor=4`): AES-CBC in the files this
//!   crate reads and writes ([MS-OFFCRYPTO] §2.3.4.10 also names CFB and other ciphers;
//!   those are refused as unimplemented). The key derivation runs whichever of SHA-1 /
//!   SHA-256 / SHA-384 / SHA-512 the file names — read separately from `<keyData>` and
//!   `<p:encryptedKey>`, so that a file whose two elements disagree still decrypts under
//!   the hash each half was written with. A writer MUST make them match (§2.3.4.10); a
//!   reader that assumes they do silently uses the wrong hash on whichever half it
//!   guessed. SHA-512 is what Office 16 writes; a file naming another is not a wrong
//!   password.
//! - **Standard encryption** (Office 2007, `vMajor` 2/3/4 with `vMinor=2` and `fAES` set):
//!   AES in ECB under a SHA-1 KDF of 50,000 iterations ([MS-OFFCRYPTO] §2.3.4.7). The
//!   header (§2.3.4.5) declares AES-128, AES-192 or AES-256; this crate **reads all
//!   three** and writes AES-128, which is Office's default. SHA-1, ECB and the 50 000
//!   iterations are fixed by the format (§2.3.4.7) and are not file fields; the key
//!   length is the one parameter the header chooses, and the derivation does not branch
//!   on it — all three keys are prefixes of the same 40-byte ladder output.
//!
//! Both OOXML formats use a CFB container (magic: `D0 CF 11 E0 A1 B1 1A E1`). The
//! decrypted output of the modern path is the original OOXML ZIP.
//!
//! # Getting started
//!
//! [`classify()`] is the first call on bytes that have not been vetted. It never panics
//! and never returns an error: every unreadable shape collapses to [`Family::Unknown`].
//! [`is_cfb_office`] is the cheaper CFB-magic check when that is all that is needed.
//!
//! Decrypting and encrypting need the `crypto-ops` feature. The 97-2003 binaries need
//! `legacy-binary`, a superset of `crypto-ops`.
//!
//! # Examples
//!
//! ```
//! use msoffice_crypto::{classify, Container, Family, IntegrityDeclaration};
//!
//! let class = classify(include_bytes!("../tests/fixtures/agile_encrypted.docx"));
//! assert_eq!(class.container, Container::Cfb);
//! assert_eq!(class.family, Family::Agile);
//! assert_eq!(class.data_integrity, IntegrityDeclaration::Declared);
//! assert!(class.is_encrypted());
//! assert!(class.is_supported());
//! ```
//!
//! Decrypting needs `crypto-ops`. The example compiles as an empty test without that
//! feature, and runs against the shipped agile fixture with it:
//!
//! ```
//! # #[cfg(feature = "crypto-ops")]
//! # fn doctest() -> Result<(), msoffice_crypto::Error> {
//! use msoffice_crypto::decrypt_ooxml;
//!
//! let package = decrypt_ooxml(
//!     include_bytes!("../tests/fixtures/agile_encrypted.docx"),
//!     "testpass",
//! )?;
//! assert!(package.starts_with(b"PK\x03\x04"));
//! # Ok(())
//! # }
//! # #[cfg(feature = "crypto-ops")]
//! # doctest().unwrap();
//! ```
//!
//! # Two builds
//!
//! **Detection is the default build.** [`classify()`] and [`is_cfb_office`] answer what a
//! file is — container shape, encryption family, the algorithm tuple it declares, and
//! whether it carries a `dataIntegrity` element — with no cryptographic dependency at
//! all. `cargo add msoffice-crypto` installs that and nothing more.
//!
//! **Decryption and encryption are the `crypto-ops` feature.** `decrypt_ooxml`,
//! `decrypt_ooxml_with_policy`, `encrypt_ooxml`, `encrypt_ooxml_with_params`,
//! `encrypt_ooxml_standard`, `encrypt_ooxml_standard_with_key_bits`,
//! `check_encryptable`, the `Decrypted` and `EncryptParams`
//! structs and the `IntegrityPolicy` / `IntegrityOutcome` enums live behind it,
//! together with `aes`, `cbc`, `ecb`, `sha1`,
//! `sha2`, `hmac`, `base64` and `rand`:
//!
//! ```toml
//! msoffice-crypto = { version = "0.1.0-rc.5", features = ["crypto-ops"] }
//! ```
//!
//! <div class="warning">
//!
//! **If you see** `cannot find type Error in crate msoffice_crypto`, this is the
//! feature you are missing.
//!
//! To be exact about what is gated, because the imprecise version misleads: the error
//! *type* compiles in every configuration — every variant payload is a `&'static str`,
//! `String`, `u16`, `std::io::Error` or a fieldless `Copy` enum from the detection half
//! (`Family`, `Document`), so it costs the detection build nothing but `thiserror` — and
//! it is only the `pub use` that `crypto-ops` gates. The reason is not
//! that the type needs a cipher. It is that once the crypto-only variants are gated, an
//! ungated re-export would be a public type whose *shape* changes with a feature the
//! consumer cannot see from the name, in a build where `classify()` is infallible and
//! nothing produces it.
//!
//! The first crate to wire against this reported that the failure surfaces on a function
//! signature naming `Error`, which reads as a plumbing mistake rather than a missing
//! feature. The diagnostic cannot be improved from inside the crate: a `compile_error!`
//! on "no cryptography features" would break the detection-only build, which is a
//! supported configuration and the default one.
//!
//! </div>
//!
//! `secure-gate` — this crate's zeroizing primitive — rides on `crypto-ops` as well: the
//! detection build holds no key material. No secure-gate type crosses the public API,
//! by design; the password is `&str` and the plaintext is `Vec<u8>`. Key material derived
//! inside the crate is zeroized on drop; the password you pass and the plaintext you
//! receive are yours to wipe.
//!
//! **The 97-2003 binary formats are the `legacy-binary` feature**, a superset of
//! `crypto-ops`. `decrypt_binary_office` reads a `.doc`, `.xls` or `.ppt` protected
//! with RC4 CryptoAPI ([MS-OFFCRYPTO] §2.3.5), Office 97/2000 RC4 (§2.3.6) or, for a
//! workbook, XOR obfuscation (§2.3.7), and adds `rc4` and `md-5` — the two primitives
//! nothing modern needs — to that build alone:
//!
//! ```toml
//! msoffice-crypto = { version = "0.1.0-rc.5", features = ["legacy-binary"] }
//! ```
//!
//! # Bounds on untrusted input
//!
//! Every number below is read from a file an attacker may have written, so each one is
//! capped. A declared size is an allocation request and an iteration count is a promise
//! of work; neither is believed. The values are listed because a consumer sizing its own
//! limits, or deciding whether this crate can be handed a particular document, should not
//! have to ask. **They are internal constants (`src/limits.rs`, all `pub(crate)`), not
//! public API** — they are quoted here for discoverability and may tighten in any
//! release.
//!
//! Compiled into **every build**, because `classify` reads them before anything is
//! authenticated: the `EncryptionInfo` stream is read to at most 1 MiB, a 97-2003 binary
//! header to 512 bytes, the `.xls` BIFF scan to 1 MiB, `/Current User` to 256 bytes, and
//! the PowerPoint persist directory to 8 MiB across at most 2^20 objects — that last
//! figure being the spec's own, since `persistId` is 20 bits ([MS-PPT] § 2.3.5), not a
//! margin this crate chose.
//!
//! Under **`crypto-ops`**: a package is at most **1 GiB** in either direction — the same
//! ceiling refuses an oversized `EncryptedPackage` on read and an oversized `package` on
//! write, so a file this crate writes is a file it can read back. `spinCount` is capped
//! at **10 000 000**, `ST_SpinCount`'s own `maxInclusive` ([MS-OFFCRYPTO] §2.3.4.10):
//! about seven seconds of one core, against the fifty minutes an uncapped `u32::MAX`
//! buys, which no `Result` can report. A caller wanting a tighter rule than the
//! format's has what it needs before a single round runs — [`classify()`] reports the
//! declared spin count unvalidated — and that is where such a rule belongs, because a
//! ceiling imposed here is one the caller cannot loosen for a document its owner
//! already holds. Agile key sizes are 128, 192 or 256 bits, salts 1..=65536 bytes, and
//! the standard path takes the same three key sizes — [MS-OFFCRYPTO] §2.3.4.5's
//! `0x00000080` / `0x000000C0` / `0x00000100` — with the further requirement that the
//! value agree with the `AlgID` beside it. That path refused AES-192 and AES-256 by
//! name until 2026-09-20, which was this crate's decision and not the format's.
//!
//! Under **`legacy-binary`**: RC4 key sizes 40..=128 bits (the spec's own range,
//! [MS-OFFCRYPTO] § 2.3.5.1) and an XOR-obfuscation password of at most 15 characters,
//! which is structural rather than a margin — the `InitialCode` table has exactly 15
//! entries.
//!
//! # Trademarks
//!
//! Microsoft, Microsoft Office, Word, Excel and PowerPoint are trademarks of Microsoft
//! Corporation. This crate is not affiliated with, endorsed by, or sponsored by
//! Microsoft; it is an independent implementation of the [MS-OFFCRYPTO] formats and uses
//! those names only to describe what it reads and writes.
//!
//! # A note on `GH #N` in the source
//!
//! Comments throughout this crate cite issues as `GH #4`, `GH #13` and so on. Those
//! numbers belong to the **private repository this crate was developed in**, which is
//! archived and is not this one — they are not issues in the published repository and
//! will not resolve there.
//!
//! They are kept because the surrounding sentence usually needs them to make sense: the
//! reason a guard is shaped the way it is, or the change that made a hardcoded value
//! configurable, is often the only record of why an obvious-looking simplification is
//! wrong. `docs/design/development-record.md` in the repository is the index that maps every one
//! of them to what it was, alongside the decisions that were reversed and the negative
//! results worth not repeating. Plan slice identifiers (`S1`–`S12`) appear beside many of
//! them and are defined in the same place.

// Feature badges on docs.rs. Two thirds of this crate's public surface is behind
// `crypto-ops` or `legacy-binary`, and without this rustdoc renders `decrypt_ooxml`
// beside `classify` with nothing to say one needs a feature and the other does not — a
// reader takes the default build and gets a compile error the docs did not predict.
// § *Two builds* above says it in prose; this says it per item, which is where a reader
// actually looks.
//
// `doc_cfg` in its auto mode rather than annotating each item with `doc(cfg(..))`: the badge is
// derived from the `#[cfg]` that is already there, so an item that changes features
// cannot keep a stale badge, and adding a gated item cannot forget one. Nightly-only and
// inert without `--cfg docsrs`, which only docs.rs and the `docsrs` CI job pass — stable
// builds, the MSRV job and `cargo test` never see it.
#![cfg_attr(docsrs, feature(doc_cfg))]
// The crate has never contained an `unsafe` block, and this is what turns that from a
// fact about today's source into a property a reviewer does not have to re-check. It is
// the one claim in § *Every input is hostile* that the compiler can enforce on its own:
// the parsers are handed bytes an attacker chose, so memory unsafety here is reachable by
// anyone who can hand a caller a file. `forbid` rather than `deny` on purpose -- `deny` is
// overridable by an inner `#[allow]`, which is exactly the edit that would need to be
// noticed, and `forbid` makes that edit a compile error instead of a diff to catch.
//
// It binds the whole crate, dependencies excluded: `cfb`, `quick-xml` and the RustCrypto
// stack each answer for their own, and `cargo deny` is what watches them.
#![forbid(unsafe_code)]

mod binary_office;
mod cfb_reader;
mod classify;
mod error;
mod limits;

// The 97-2003 binary formats, GH #4: the container they are decrypted inside, the
// three cipher families, and one module per application format. A `//` comment, not a
// `///` one: this describes the eight modules below, and as rustdoc it would become
// `excel97`'s alone and be prepended to that module's own header.
#[cfg(feature = "legacy-binary")]
mod excel97;
#[cfg(feature = "legacy-binary")]
mod legacy_container;
#[cfg(feature = "legacy-binary")]
mod powerpoint97;
#[cfg(feature = "legacy-binary")]
mod rc4;
#[cfg(feature = "legacy-binary")]
mod rc4_cryptoapi;
#[cfg(feature = "legacy-binary")]
mod rc4_office97;
#[cfg(feature = "legacy-binary")]
mod word97;
#[cfg(feature = "legacy-binary")]
mod xor_obfuscation;

#[cfg(feature = "crypto-ops")]
mod agile;
/// The agile write path: the password encryptor (the inverse of
/// `agile::verify_password`), the package encryptor, and the assembly behind
/// [`encrypt_ooxml`]. Where `secure-gate`'s `rand` turns on.
#[cfg(feature = "crypto-ops")]
mod agile_encrypt;
/// The `\x06DataSpaces` subtree and the CFB container that carries it — the write half
/// of the format, promoted out of `tests/` in GH #6 step 1 so the byte-identity proof
/// covers the code the encrypt path will call rather than a copy of it.
///
/// `cfg(test)` until GH #6 step 6 made [`encrypt_ooxml`] its production caller — the
/// gate was the reminder that the flip was due, and it fired on schedule.
#[cfg(feature = "crypto-ops")]
mod dataspaces;
/// The encryption parameters a caller may choose, and the one function that judges them.
/// Public through the [`EncryptParams`] re-export below — the type is an input to the
/// encrypt path rather than a decision this crate makes alone, which is why it is a
/// module of its own and not a private struct inside `agile_encrypt`.
///
/// The threading is complete: [`encrypt_ooxml_with_params`] takes one from the caller,
/// `agile_encrypt::generate` sizes every draw from it, and `encryption_info::write`
/// re-checks the same value and writes it out — so the document cannot describe a tuple
/// other than the one that produced its blobs. [`encrypt_ooxml`] is that same path with
/// [`EncryptParams::default`] supplied for the caller.
#[cfg(feature = "crypto-ops")]
mod encrypt_params;
/// Serialise the agile `EncryptionInfo` stream — the inverse of `agile`'s parser, and the
/// half GH #6 step 3 added. Shaped byte-for-byte on what Word 16 writes.
#[cfg(feature = "crypto-ops")]
mod encryption_info;
/// The four hash algorithms an agile file may name, as operations rather than as a
/// label. The enum itself lives in `classify`, which must report the hash in a build
/// with no cipher crate at all; this is the `crypto-ops` half.
///
/// Private, but not entirely internal: two of the inherent methods it hangs on
/// [`HashAlgorithm`] — `digest_len` and `name` — are `pub`, and reach a consumer through
/// the enum's own re-export below rather than through this module, which is why there is
/// no `pub use` here to add. They are facts about SHA rather than about this crate, and a
/// caller choosing encryption parameters needs them to predict the `keyBits <= digest`
/// coupling the agile path imposes. That makes them `crypto-ops`-only API on a type the
/// detection build also exports — deliberate, and the same shape as [`Error`], whose
/// re-export is gated for the same reason.
#[cfg(feature = "crypto-ops")]
mod hash;
#[cfg(feature = "crypto-ops")]
mod integrity;
/// The 4096-byte segment layout the agile package is encrypted in, with each segment's
/// IV. One segmentation shared by decrypt and, from GH #6 step 4, encrypt — plan D4.
#[cfg(feature = "crypto-ops")]
mod segments;
#[cfg(feature = "crypto-ops")]
mod sensitive;
#[cfg(feature = "crypto-ops")]
mod standard;
/// The standard (Office 2007) write path: the salt, the derived key, the two verifier
/// blobs, the binary header, and the assembly behind [`encrypt_ooxml_standard`]. Runs on
/// `standard`'s own KDF and ECB helper, so the two directions share one derivation.
///
/// It also holds the one parameter this format has — the key size, judged by a private
/// `AesKeySize` and reaching a caller through
/// [`encrypt_ooxml_standard_with_key_bits`]'s plain `u32`. No type of its own is
/// exported for it: [MS-OFFCRYPTO] fixes every other field of the header, so there is no
/// tuple to name, and that module's own header argues the choice against the two
/// alternatives.
#[cfg(feature = "crypto-ops")]
mod standard_encrypt;

/// Malformed-parameter tests: every number a file declares about itself must produce an
/// error rather than a panic or a hang. Test-only, and separate from the fixture tests
/// below because it needs a synthetic container builder those have no use for.
#[cfg(all(test, feature = "crypto-ops"))]
mod malformed_input;

/// The same discipline for the binary formats: every offset and length a `.doc`, `.xls`
/// or `.ppt` declares, poisoned one at a time, with a control beside each.
#[cfg(all(test, feature = "legacy-binary"))]
mod legacy_malformed;

pub use classify::{
    classify, AlgorithmParams, CipherAlgorithm, Classification, Container, ContainerRead, Document,
    Family, HashAlgorithm, IntegrityDeclaration,
};
/// The parameters of an agile encryption, as a caller chooses them — and
/// `EncryptParams::validate`, which judges them before a password is asked for.
///
/// Gated with the encrypt path it parameterises. Exported from the crate root rather
/// than from a module of its own for the same reason every other public item here is:
/// this crate has one public surface, and `msoffice_crypto::encrypt_params::EncryptParams`
/// would be a second path to the same type.
#[cfg(feature = "crypto-ops")]
pub use encrypt_params::EncryptParams;
/// Gated with the functions that return it. In a detection-only build no public
/// function returns a `Result`, so an ungated re-export was a public type nothing
/// produced — and, once its crypto-only variants were gated, a type whose public shape
/// depended on a feature the consumer could not see from the name. See the enum's doc.
#[cfg(feature = "crypto-ops")]
pub use error::Error;
/// The two halves of [`Error::EncryptParams`]'s payload, gated with the variant that
/// carries them. They are deliberately enums rather than strings so that a caller can
/// `match` a rejected encryption parameter — which is unreachable if the types are not
/// exported, so this re-export is part of that variant, not a convenience.
#[cfg(feature = "crypto-ops")]
pub use error::{EncryptParam, EncryptParamProblem};
#[cfg(feature = "crypto-ops")]
pub use integrity::{IntegrityOutcome, IntegrityPolicy};

/// What [`decrypt_ooxml_with_policy`] returns: the package, and what was established
/// about it.
///
/// A struct rather than the `(Vec<u8>, IntegrityOutcome)` this returned until the first
/// consumer wired against it, for two reasons that only showed up at a real call site.
///
/// **Arity.** A tuple freezes the number of facts at publication. There are two here and
/// a plausible third — which cipher a file actually used, which spec branch it parsed as
/// — and adding one to a tuple is a breaking change for every caller, while adding a
/// field to a `#[non_exhaustive]` struct is not. `#[non_exhaustive]` is free before the
/// first publish and impossible to add afterwards without the same break, which is why
/// this is the shape that ships.
///
/// **Prominence.** [`Self::integrity`] is not a detail attached to the bytes; for a
/// caller that stores what it decrypts it is the predicate deciding whether the bytes may
/// be kept at all. `let (package, _) = ...` is eight characters and reads as idiom rather
/// than as a decision, and this crate's own [`decrypt_ooxml`] is the demonstration — it is
/// the one place that consumes this and it discards the outcome. That is correct *there*,
/// because that wrapper exists to be the "I do not need to ask" path, which is exactly
/// what made it the wrong default shape for everyone else. As a field the check reads as
/// `decrypted.integrity`, named at every call site and in every review diff.
///
/// There is deliberately no `require_verified()` helper. [`IntegrityPolicy::Require`]
/// already refuses unauthenticated plaintext *before* any work is done, and a second
/// gate after the fact would have to invent an error variant for "the policy allowed this
/// but I changed my mind", which is not a fact about the file.
#[cfg(feature = "crypto-ops")]
#[derive(Debug)]
#[non_exhaustive]
pub struct Decrypted {
    /// The decrypted package: the plain `.docx` / `.xlsx` / `.pptx` ZIP.
    ///
    /// **Not zeroized on drop**, deliberately and unlike this crate's key material. It is
    /// a document, its size is the file's, and a caller that needs it wiped knows that
    /// better than this crate does — wrap it at the boundary. No `secure-gate` type
    /// crosses this API by design; see the crate docs.
    pub package: Vec<u8>,

    /// What was established about [`Self::package`]'s integrity.
    ///
    /// Ask [`IntegrityOutcome::is_authenticated`] rather than matching: the enum is
    /// `#[non_exhaustive]`, so a wildcard arm in a caller's `match` cannot know which side
    /// a future variant belongs on, and that predicate can.
    pub integrity: IntegrityOutcome,
}

/// Returns `true` if `data` begins with the CFB magic `D0 CF 11 E0 A1 B1 1A E1`.
///
/// A Word / Excel / PowerPoint file encrypted via File → Protect → Encrypt with Password
/// becomes a CFB container; the ZIP magic (`PK\x03\x04`) is replaced by the CFB magic.
/// This is only a prefix check: a truncated or corrupt container that still starts with
/// those eight bytes returns `true`, and [`classify()`] is what reports whether the rest
/// can be read.
///
/// # Examples
///
/// ```
/// use msoffice_crypto::is_cfb_office;
///
/// assert!(is_cfb_office(include_bytes!("../tests/fixtures/agile_encrypted.docx")));
/// assert!(!is_cfb_office(include_bytes!("../tests/fixtures/plain.docx")));
/// assert!(!is_cfb_office(b"PK\x03\x04"));
/// ```
///
/// # See Also
///
/// [`classify()`] reports container, family and integrity declaration together.
pub fn is_cfb_office(data: &[u8]) -> bool {
    data.len() >= 8 && data[..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
}

/// Decrypt a password-protected OOXML file to the original ZIP package.
///
/// Supports agile encryption (Office 2010+) and standard encryption (Office 2007).
/// Returns the decrypted OOXML ZIP bytes (a valid `.docx` / `.xlsx` / `.pptx` package).
///
/// Integrity is checked under [`IntegrityPolicy::RequireWhereDefined`], the default, and
/// it **fails closed**: an agile file must carry a `dataIntegrity` tag and it must match,
/// or the file is refused. ECMA-376 standard encryption defines no such element and still
/// decrypts. Use [`decrypt_ooxml_with_policy`] to demand a tag of every format, to accept
/// an agile file that lacks one, to skip the check, or to learn which of those happened.
///
/// The password is `&str` and the returned plaintext is a plain `Vec<u8>`, by design: no
/// `secure-gate` type crosses this boundary. Key material derived inside the crate is
/// zeroized on drop; the password you pass and the package you receive are yours to wipe.
///
/// On error, `data` is left untouched and no plaintext is produced. The agile HMAC covers
/// ciphertext and is checked before decryption, so a failed integrity check never yields
/// unauthenticated bytes.
///
/// # Errors
///
/// - [`Error::NotACfbFile`] — `data` is not a CFB container
/// - [`Error::MissingStream`] — `EncryptionInfo` or `EncryptedPackage` is
///   absent, unreadable, or shorter than its header
/// - [`Error::XmlParse`] — the agile `EncryptionInfo` XML is malformed, a
///   required attribute is missing, or the file carries only `CertificateKeyEncryptor`
///   elements and no `PasswordKeyEncryptor` (this crate opens password-protected
///   documents only)
/// - [`Error::BadParameters`] — a declared length, spin count, reserved word
///   or sibling field is out of range or inconsistent
/// - [`Error::WrongPassword`] — password verification failed
/// - [`Error::IntegrityCheckFailed`] — the package does not match its HMAC
/// - [`Error::IntegrityElementMissing`] — the file declares agile encryption
///   but carries no `dataIntegrity` element to check it against
/// - [`Error::UnsupportedAlgorithm`] — the file names a cipher or hash this
///   crate does not implement. Never reported as a wrong password: the password may be
///   correct and simply unusable
/// - [`Error::UnsupportedEncryptionVersion`] — a version pair this crate does
///   not implement
/// - [`Error::CipherError`] — an AES operation rejected a block (a length that
///   is not a block multiple, typically a truncated stream)
/// - [`Error::Io`] — reading the in-memory container failed
///
/// [`Error::IntegrityUnavailable`] is not reachable here: that variant is
/// [`IntegrityPolicy::Require`] on a format that defines no tag, and this function uses
/// the default policy.
///
/// # Examples
///
/// ```
/// use msoffice_crypto::decrypt_ooxml;
///
/// let package = decrypt_ooxml(
///     include_bytes!("../tests/fixtures/agile_encrypted.docx"),
///     "testpass",
/// )?;
/// assert!(package.starts_with(b"PK\x03\x04"));
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// A wrong password is a distinct variant from a tampered package:
///
/// ```
/// use msoffice_crypto::{decrypt_ooxml, Error};
///
/// let err = decrypt_ooxml(
///     include_bytes!("../tests/fixtures/agile_encrypted.docx"),
///     "wrongpass",
/// )
/// .unwrap_err();
/// assert!(matches!(err, Error::WrongPassword));
/// ```
///
/// # See Also
///
/// [`decrypt_ooxml_with_policy`] chooses the HMAC policy and reports
/// [`IntegrityOutcome`]. [`encrypt_ooxml`] is the inverse for the agile tuple Office 16
/// writes. [`classify()`] is the pre-flight that does not decrypt.
#[cfg(feature = "crypto-ops")]
pub fn decrypt_ooxml(data: &[u8], password: &str) -> Result<Vec<u8>, Error> {
    decrypt_ooxml_with_policy(data, password, IntegrityPolicy::default()).map(|d| d.package)
}

/// Refuse, before a password is asked for, anything [`encrypt_ooxml`] and
/// [`encrypt_ooxml_standard`] will refuse.
///
/// **This is the same function those two call**, not a second copy that agrees with them:
/// each opens with `check_encryptable(package)?`. A caller that runs it and gets `Ok(())`
/// is not promised the encryption will succeed — the payload ceiling and the system RNG
/// are still ahead — but it is promised that the *shape* of the input will not be what
/// stops it.
///
/// It exists because the alternative is asking for a password first. Prompting for a new
/// password, or reading one from a keychain, for a file that is about to be refused is a
/// question the user answers for nothing, and on an interactive path it is the part of a
/// refusal that cannot be taken back. This crate's own CLI calls it in that position; so
/// did a consumer that had to write its own copy before this existed.
///
/// **Exactly one thing is encryptable: a plain OOXML package**, which [`classify()`]
/// reports as [`Container::Zip`] for every `PK` signature it knows. Four bytes of magic
/// are the whole test — no entry is read and no content type is examined, so a `.vsdx`, a
/// `.jar` and a backup archive all pass here (see [`Document::ZipArchive`]). That is the
/// honest limit of what was checked, and encrypting a ZIP that is not an Office package
/// harms nobody: the result is a container whose payload the caller chose. What is
/// refused is everything that would produce a *misleading* artifact — above all a second
/// wrap around a file that is already encrypted.
///
/// # Errors
///
/// - [`Error::AlreadyEncrypted`] — a CFB container that already carries a
///   password-to-open. The payload names the family and document kind, for a caller
///   choosing its own words
/// - [`Error::NotAPlainPackage`] — a CFB container that does not: a 97-2003 binary
///   document, or one this crate could not read
/// - [`Error::UnknownContainer`] — neither a ZIP package nor a CFB container
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{check_encryptable, encrypt_ooxml, Error};
///
/// // Ask before the prompt, not after.
/// let package = include_bytes!("../tests/fixtures/plain.docx");
/// check_encryptable(package)?;
/// let password = "correct horse battery staple"; // …whatever asking cost you
/// let sealed = encrypt_ooxml(package, password)?;
///
/// // And the answer for bytes that were never worth asking about:
/// assert!(matches!(
///     check_encryptable(&sealed),
///     Err(Error::AlreadyEncrypted { .. })
/// ));
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`classify()`] is the full pre-flight; this is the one question the encrypt path asks
/// of it. Every public encrypt entry point this crate gains must call this.
// No `#[must_use]`: `Result` already carries it, and adding a second fires
// `clippy::double_must_use`, which is `-D warnings` in all five feature columns.
#[cfg(feature = "crypto-ops")]
pub fn check_encryptable(package: &[u8]) -> Result<(), Error> {
    let class = classify(package);
    // Exhaustive without a `_` arm: `Container` is `#[non_exhaustive]` only to other
    // crates, so a variant added later is `E0004` right here and someone has to decide
    // what it is, rather than it defaulting into "encryptable" or into one refusal.
    match class.container {
        Container::Zip => Ok(()),
        Container::Cfb if class.is_encrypted() => Err(Error::AlreadyEncrypted {
            family: class.family,
            document: class.document,
        }),
        Container::Cfb => Err(Error::NotAPlainPackage),
        Container::Unknown => Err(Error::UnknownContainer),
    }
}

/// Encrypt an OOXML package with a password, producing the CFB container Office writes.
///
/// `package` is the plain `.docx` / `.xlsx` / `.pptx` ZIP. The result is ECMA-376 agile
/// encryption in the one tuple Office 16 itself writes — AES-256-CBC, SHA-512, a
/// 100 000-round spin count — with a `dataIntegrity` HMAC over the `EncryptedPackage`
/// ciphertext, including the 8-byte size prefix ([MS-OFFCRYPTO] §2.3.4.14), and it
/// is what [`decrypt_ooxml`] reads back under its fail-closed default. Every step of the
/// write path has been checked byte for byte against real Office output where Office's
/// own random inputs could be recovered: the `EncryptionInfo` document and the two
/// `dataIntegrity` blobs reproduce Word's, Excel's and PowerPoint's exactly.
///
/// **This function's profile is fixed.** The spin count is 100 000 and the rest of the
/// tuple is Office 16's, with no parameter here to change any of it — no builder, no
/// environment variable — because the one tuple Office writes is the whole point of
/// *this* entry point, and a caller reading this signature can predict the bytes without
/// tracing a configuration. That is a statement about this function and not about the
/// crate: [`encrypt_ooxml_with_params`] takes an [`EncryptParams`], and this function is
/// one line delegating to it with [`EncryptParams::default`], so "the default path and
/// the parameterised path are the same code" is a fact rather than a claim.
///
/// **A `<dataIntegrity>` element is written unconditionally, and that guarantee is *not*
/// scoped to this function.** Every agile artifact this crate produces declares one —
/// from here or from [`encrypt_ooxml_with_params`], under every `EncryptParams` tuple —
/// so a consumer checking [`Classification::data_integrity`] on agile output from this
/// crate may assert [`IntegrityDeclaration::Declared`] and know the assertion cannot
/// fail. It does not extend to [`encrypt_ooxml_standard`], whose format defines no such
/// element.
///
/// The session key, block keys and spin hash are held in `secure-gate` wrappers and
/// zeroized on drop; the password is `&str` and the input and output are plain bytes, by
/// design. Randomness comes from the operating system's CSPRNG through the same function
/// the seeded golden tests drive, so those tests are about this code path.
///
/// On error, nothing is written: the function returns without a container.
///
/// # Errors
///
/// - [`Error::AlreadyEncrypted`], [`Error::NotAPlainPackage`],
///   [`Error::UnknownContainer`] — `package` is not a plain OOXML package. Checked by
///   [`check_encryptable`] before anything else, so a caller can ask the same question
///   before it pays for a password
/// - [`Error::BadParameters`] — `package` is over the 1 GiB this crate would
///   read back
/// - [`Error::RandomSource`] — the system RNG would not produce bytes
/// - [`Error::CipherError`] — an AES-CBC step rejected a block (an internal
///   length invariant, not a property of a well-formed `package`)
/// - [`Error::Io`] — the in-memory container could not be written
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{decrypt_ooxml, encrypt_ooxml, is_cfb_office};
///
/// let package = include_bytes!("../tests/fixtures/plain.docx");
/// let sealed = encrypt_ooxml(package, "testpass")?;
/// assert!(is_cfb_office(&sealed));
/// assert_eq!(decrypt_ooxml(&sealed, "testpass")?, package);
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`encrypt_ooxml_with_params`] is the same write path with the tuple chosen by the
/// caller. [`encrypt_ooxml_standard`] writes the Office 2007 format for a reader that
/// cannot open agile files. Prefer this function unless one of those constraints
/// applies.
#[cfg(feature = "crypto-ops")]
pub fn encrypt_ooxml(package: &[u8], password: &str) -> Result<Vec<u8>, Error> {
    // One line, deliberately. The measured Office 16 tuple is `EncryptParams::default`,
    // and delegating rather than repeating the call is what makes "the default path is
    // the parameterised path" checkable by reading one line instead of by comparing two
    // argument lists that could drift.
    encrypt_ooxml_with_params(package, password, EncryptParams::default())
}

/// Encrypt an OOXML package with a password and an [`EncryptParams`] tuple of the
/// caller's choosing.
///
/// [`encrypt_ooxml`] is this function with [`EncryptParams::default`] — literally: that
/// function's body is one call to this one. Everything [`encrypt_ooxml`] documents about
/// the container, the `secure-gate` wrapping of the key schedule, the system CSPRNG and
/// writing nothing on error holds here unchanged, and so does the **unconditional
/// `<dataIntegrity>` element**: it is written for every tuple, so a consumer may assert
/// [`IntegrityDeclaration::Declared`] on agile output from this function exactly as it
/// may on output from that one.
///
/// What changes is the six values in the `EncryptionInfo` document — the spin count, the
/// hash, the two `keyBits` and the two `saltSize` — and the lengths of the blobs those
/// imply. [`EncryptParams`] documents each field against the [MS-OFFCRYPTO] §2.3.4.10
/// attribute it is, including why `keyBits` and `saltSize` are two fields each while the
/// hash is one; `hashSize` and `blockSize` are absent from the type because the format
/// derives them.
///
/// **Ask first if the answer is expensive.** [`EncryptParams::validate`] is this
/// function's own parameter check, callable on its own, and [`check_encryptable`] is the
/// same for `package`. Between them a caller learns that a call would be refused without
/// having prompted for a password. This function runs both itself — `package` first,
/// then the parameters, both before the first byte is drawn from the RNG.
///
/// A `password_salt_size` other than 16 is **written**, not refused. §2.3.4.12 fits that
/// salt to the block length before using it as the three password blobs' IV — padding a
/// short one with `0x36`, truncating a long one — and this writer applies that fit, as
/// the reader always has. At 16 the fit is the identity, which is why nothing caught its
/// absence until the salt size became a caller's choice.
///
/// What [`EncryptParams::validate`] accepts is therefore what this function writes,
/// across the spec's whole `1..=65536`. **What no external reader has been measured on
/// is a salt size that is not a multiple of 16**, which changes the pad on
/// `encryptedVerifierHashInput`; Word is documented rejecting wrong pad bytes elsewhere
/// in this format. That is an evidence gap, not a refusal, and it is recorded as one.
///
/// # Errors
///
/// Every error [`encrypt_ooxml`] returns, plus:
///
/// - [`Error::EncryptParams`] — `params` is not a tuple this crate will write. The
///   payload names the parameter ([`EncryptParam`]), whose rule it broke
///   ([`EncryptParamProblem`] — the format's, the cipher's, or this crate's) and one
///   value that would have been accepted
///
/// # Examples
///
/// A tuple that is not the default, round-tripped through the ordinary decrypt path:
///
/// ```
/// use msoffice_crypto::{
///     classify, decrypt_ooxml, encrypt_ooxml_with_params, EncryptParams, HashAlgorithm,
///     IntegrityDeclaration,
/// };
///
/// let package = include_bytes!("../tests/fixtures/plain.docx");
/// let sealed = encrypt_ooxml_with_params(
///     package,
///     "testpass",
///     EncryptParams {
///         hash: HashAlgorithm::Sha384,
///         key_data_key_bits: 192,
///         spin_count: 1_000,
///         ..Default::default()
///     },
/// )?;
///
/// // The tuple reached the file — `classify` reads the numbers back out of it — and
/// // the integrity guarantee is not tuple-dependent.
/// let class = classify(&sealed);
/// let key_data = class.key_data.expect("agile files declare <keyData>");
/// assert_eq!(key_data.key_bits, Some(192));
/// assert_eq!(key_data.hash, Some(HashAlgorithm::Sha384));
/// assert_eq!(
///     class.password_key.and_then(|p| p.spin_count),
///     Some(1_000)
/// );
/// assert_eq!(class.data_integrity, IntegrityDeclaration::Declared);
/// assert_eq!(decrypt_ooxml(&sealed, "testpass")?, package);
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`encrypt_ooxml`] for the tuple Office 16 writes, which is what any given reader is
/// most likely to have been tested against. [`EncryptParams::validate`] to ask before
/// paying for a password.
#[cfg(feature = "crypto-ops")]
pub fn encrypt_ooxml_with_params(
    package: &[u8],
    password: &str,
    params: EncryptParams,
) -> Result<Vec<u8>, Error> {
    // `params` by value, not by reference, on the `decrypt_ooxml_with_policy` precedent
    // above: six `Copy` scalars are no larger than the pointer to them, and a caller
    // that built the tuple inline has nothing left to borrow it from.
    check_encryptable(package)?;
    // `agile_encrypt::encrypt` calls `params.validate()` itself, before the payload
    // ceiling and before the first draw. Not repeated here: two copies of one question
    // are two places for the answer to change.
    agile_encrypt::encrypt(package, password, params, &mut rand::rngs::SysRng)
}

/// Encrypt an OOXML package in the Office 2007 format, ECMA-376 standard encryption.
///
/// AES-128-ECB under a SHA-1-derived key, for a reader that predates agile encryption.
/// [`encrypt_ooxml_standard_with_key_bits`] writes the other two key sizes
/// [MS-OFFCRYPTO] §2.3.4.5 defines; this function is that one with AES-128 supplied, and
/// AES-128 is what Office 2007 itself wrote.
///
/// **Prefer [`encrypt_ooxml`].** Standard encryption defines no integrity element: a
/// modified ciphertext decrypts, silently, to a modified document, and ECB leaks equal
/// plaintext blocks as equal ciphertext blocks. This entry point exists because Office
/// 2007 cannot open an agile file and some tooling still targets it; it is named for the
/// format so that choosing it is a decision rather than a default. [`decrypt_ooxml`]
/// reads the result back under its fail-closed default, [`decrypt_ooxml_with_policy`]
/// reports [`IntegrityOutcome::NotApplicable`] for it, and
/// [`IntegrityPolicy::Require`] refuses it by name.
///
/// The header is [MS-OFFCRYPTO] §2.3.4.5's, written with the conforming
/// `fCryptoAPI | fAES` / `AlgID 0x660E` pair and the CSP name Office 2007 wrote; the
/// verifier hash is zero-padded to its 32-byte blob, the pad real Word compares. The
/// derived key, the password digest it comes from and the verifier plaintexts are held
/// in `secure-gate` wrappers and zeroized on drop; the password is `&str` and the input
/// and output are plain bytes, by design. Randomness comes from the operating system's
/// CSPRNG through the same function the seeded golden test drives.
///
/// On error, nothing is written: the function returns without a container.
///
/// # Errors
///
/// - [`Error::AlreadyEncrypted`], [`Error::NotAPlainPackage`],
///   [`Error::UnknownContainer`] — `package` is not a plain OOXML package. Checked by
///   [`check_encryptable`] before anything else, so a caller can ask the same question
///   before it pays for a password
/// - [`Error::BadParameters`] — `package` is over the 1 GiB this crate would
///   read back
/// - [`Error::RandomSource`] — the system RNG would not produce bytes
/// - [`Error::CipherError`] — an AES-ECB step returned a blob of the wrong
///   length (unreachable for the AES-128 key size this function writes)
/// - [`Error::Io`] — the in-memory container could not be written
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{
///     decrypt_ooxml_with_policy, encrypt_ooxml_standard, IntegrityOutcome,
///     IntegrityPolicy,
/// };
///
/// let package = include_bytes!("../tests/fixtures/plain.docx");
/// let sealed = encrypt_ooxml_standard(package, "testpass")?;
/// let decrypted = decrypt_ooxml_with_policy(
///     &sealed,
///     "testpass",
///     IntegrityPolicy::RequireWhereDefined,
/// )?;
/// assert_eq!(decrypted.package, package);
/// assert_eq!(decrypted.integrity, IntegrityOutcome::NotApplicable);
/// assert!(!decrypted.integrity.is_authenticated());
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`encrypt_ooxml`] writes agile encryption with a `dataIntegrity` HMAC, which is what
/// Office 16 writes and what this crate recommends.
/// [`encrypt_ooxml_standard_with_key_bits`] writes AES-192 and AES-256 in this same
/// format.
#[cfg(feature = "crypto-ops")]
pub fn encrypt_ooxml_standard(package: &[u8], password: &str) -> Result<Vec<u8>, Error> {
    // One line over the parameterised entry point, for the reason `encrypt_ooxml` is one
    // line over `encrypt_ooxml_with_params`: "the default path and the parameterised
    // path are the same code" is then a fact anyone can check by reading it, rather than
    // two argument lists that agree today.
    encrypt_ooxml_standard_with_key_bits(package, password, standard_encrypt::DEFAULT_KEY_BITS)
}

/// Encrypt an OOXML package in the Office 2007 format at a key size of the caller's
/// choosing — AES-128, AES-192 or AES-256.
///
/// [`encrypt_ooxml_standard`] is this function with 128 supplied, and everything it
/// documents holds here unchanged: the container, the conforming
/// `fCryptoAPI | fAES` header, the zero-padded verifier hash blob, the `secure-gate`
/// wrapping, the system CSPRNG, the absence of any integrity element, and writing
/// nothing on error. What changes is two fields of the `EncryptionHeader` and the length
/// of the derived key.
///
/// **`key_bits` MUST be 128, 192 or 256** — [MS-OFFCRYPTO] §2.3.4.5 says of this header's
/// `KeySize` that "This value MUST be 0x00000080 (AES-128), 0x000000C0 (AES-192), or
/// 0x00000100 (AES-256)", and of its `AlgID` that it MUST be the matching one of
/// `0x0000660E` / `0x0000660F` / `0x00006610`. The two are one statement: this function
/// writes the pair from a single table, so the mismatched header §2.3.2 forbids — and
/// that [`decrypt_ooxml`] refuses by name — is not expressible through it. Anything else
/// is [`Error::EncryptParams`], raised before the password is used for anything.
///
/// The key size is the **only** parameter this format has. `AlgIDHash` is SHA-1 by
/// §2.3.4.5, the 50 000 iterations are fixed by §2.3.4.7 and are not a field in the file
/// at all, and the salt and verifier lengths are fixed by §2.3.3 — so there is no tuple
/// here and no [`EncryptParams`]: that type is the agile format's, where five of its six
/// fields name attributes this format does not have. `src/standard_encrypt.rs`'s header
/// carries that argument in full, including why this is a plain `u32` and not a struct.
///
/// # Interoperability is measured on AES-128 only
///
/// The four-reader acceptance gate's verdicts in `CHANGELOG.md`, and the committed
/// byte-for-byte golden, are all AES-128 artifacts, because that is the tuple Office 2007
/// wrote and therefore the only one a fixture can exist for. AES-192 and AES-256 are
/// proved here against this crate's own reader and against the spec clauses above; what
/// external readers do with them is an evidence gap, recorded as one rather than
/// implied away. Prefer [`encrypt_ooxml_standard`] — or, better, [`encrypt_ooxml`] —
/// unless you have a reason to write a wider key.
///
/// # Errors
///
/// Every error [`encrypt_ooxml_standard`] returns, plus:
///
/// - [`Error::EncryptParams`] — `key_bits` is not one of the three sizes §2.3.4.5
///   defines. The payload names the field ([`EncryptParam::KeySize`]), whose rule it
///   broke ([`EncryptParamProblem::OutsideSpecRange`] — the format's, always, on this
///   path) and, in `min` and `max` alike, the nearest size that would have been accepted
///
/// # Examples
///
/// ```
/// use msoffice_crypto::{
///     classify, decrypt_ooxml, encrypt_ooxml_standard_with_key_bits, CipherAlgorithm,
///     EncryptParam, EncryptParamProblem, Error, Family,
/// };
///
/// let package = include_bytes!("../tests/fixtures/plain.docx");
/// let sealed = encrypt_ooxml_standard_with_key_bits(package, "testpass", 256)?;
///
/// // The header declares the key size, and the round trip is this crate's own reader.
/// let class = classify(&sealed);
/// assert_eq!(class.family, Family::Standard);
/// let header = class.key_data.expect("the EncryptionHeader is present");
/// assert_eq!(header.key_bits, Some(256));
/// assert_eq!(header.cipher, Some(CipherAlgorithm::Aes));
/// assert_eq!(decrypt_ooxml(&sealed, "testpass")?, package);
///
/// // A size the format does not define is refused, and named.
/// let err = encrypt_ooxml_standard_with_key_bits(package, "testpass", 64).unwrap_err();
/// assert!(matches!(
///     err,
///     Error::EncryptParams {
///         param: EncryptParam::KeySize,
///         problem: EncryptParamProblem::OutsideSpecRange,
///         got: 64,
///         min: 128,
///         max: 128,
///     }
/// ));
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`encrypt_ooxml_with_params`] is the agile format's parameterised entry point, and
/// the one with an integrity element.
#[cfg(feature = "crypto-ops")]
pub fn encrypt_ooxml_standard_with_key_bits(
    package: &[u8],
    password: &str,
    key_bits: u32,
) -> Result<Vec<u8>, Error> {
    check_encryptable(package)?;
    // `standard_encrypt::encrypt` judges `key_bits` itself, before the payload ceiling
    // and before the first draw. Not repeated here, for the reason
    // `encrypt_ooxml_with_params` does not repeat `EncryptParams::validate`: two copies
    // of one question are two places for the answer to change.
    standard_encrypt::encrypt(package, password, key_bits, &mut rand::rngs::SysRng)
}

/// Decrypt a Word 97-2003, Excel 97-2003 or PowerPoint 97-2003 document in place.
///
/// The three binary formats keep their encryption inside their own records rather than
/// in an `EncryptionInfo` stream, and their plaintext is not a package but the same CFB
/// container with its encrypted streams replaced: the `WordDocument`, table and `Data`
/// streams of a `.doc`; the `Workbook` stream of a `.xls`; the `PowerPoint Document`
/// stream of a `.ppt`. The result is what `msoffcrypto-tool -d` writes, byte for byte on
/// every fixture in `tests/fixtures/`, and what Word, Excel and PowerPoint open without a
/// password.
///
/// Which format is decided by the streams the container holds ([`classify()`] reports the
/// same answer in [`Document`]), and which scheme by the format's own marker: the FIB's
/// `fEncrypted` bit and the version at the top of the table stream, the `FILEPASS`
/// record, the `UserEditAtom`'s `encryptSessionPersistIdRef`. Three schemes are read:
/// RC4 CryptoAPI ([MS-OFFCRYPTO] §2.3.5, all three formats), Office 97/2000 RC4 (§2.3.6,
/// Word and Excel) and XOR obfuscation (§2.3.7, Excel only — Word's variant is refused
/// by name). None of them defines an integrity element, so there is no policy to
/// choose and nothing to report: RC4 is a stream cipher and XOR is a transformation,
/// and a modified file decrypts to modified bytes with no error. That is the format's
/// limit, not this crate's, and the reason the modern formats exist.
///
/// The password hash, every block key and the XOR array are held in `secure-gate`
/// wrappers and zeroized on drop, and the RC4 key schedule is wiped with them.
///
/// On error, `data` is left untouched and no decrypted container is produced.
///
/// # Errors
///
/// - [`Error::NotACfbFile`] — `data` is not a CFB container
/// - [`Error::MissingStream`] — the container carries none of the three
///   formats' streams, or a stream is shorter than its header
/// - [`Error::NotEncrypted`] — the document carries no password-to-open
/// - [`Error::WrongPassword`] — the verifier did not match. Every scheme has
///   one: the RC4 families' encrypted verifier, XOR's 16-bit `verificationBytes`
/// - [`Error::UnsupportedAlgorithm`] — XOR obfuscation of a `.doc`, a BIFF5
///   workbook, a header naming `fExternal` or `fAES`, or an `AlgID` / `AlgIDHash` naming
///   anything but RC4 and SHA-1
/// - [`Error::UnsupportedEncryptionVersion`] — a version pair the format's own
///   walk does not implement: none of `1.1`, `2.2`, `3.2`, `4.2` for a `.doc` or `.xls`,
///   and none of `2.2`, `3.2`, `4.2` for a `.ppt`, whose `CryptSession10Container`
///   defines RC4 CryptoAPI only ([MS-OFFCRYPTO] §2.3.5)
/// - [`Error::BadParameters`] — a length or offset the file declares is out of
///   range: an `lKey` past the table stream, a record length past the workbook, a persist
///   offset past the presentation, a `KeySize` outside 40..=128, a `FILEPASS` out of order
/// - [`Error::Io`] — rewriting a stream inside the in-memory container failed
///
/// # Examples
///
/// A non-CFB input is refused before any format walk:
///
/// ```
/// use msoffice_crypto::{decrypt_binary_office, Error};
///
/// let err = decrypt_binary_office(b"PK\x03\x04", "testpass").unwrap_err();
/// assert!(matches!(err, Error::NotACfbFile));
/// ```
///
/// A password-protected `.doc` / `.xls` / `.ppt` decrypts in place. The crate tarball
/// does not ship a binary fixture, so this example is not executed:
///
/// ```no_run
/// use msoffice_crypto::decrypt_binary_office;
///
/// let data = std::fs::read("protected.doc")?;
/// let plain = decrypt_binary_office(&data, "testpass")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
///
/// # See Also
///
/// [`classify()`] reports [`Document::WordBinary`], [`Document::ExcelBinary`] or
/// [`Document::PowerPointBinary`] for these files. [`decrypt_ooxml`] is the modern-OOXML
/// path and does not read a `.doc`.
#[cfg(feature = "legacy-binary")]
pub fn decrypt_binary_office(data: &[u8], password: &str) -> Result<Vec<u8>, Error> {
    if !is_cfb_office(data) {
        return Err(Error::NotACfbFile);
    }
    let mut container = legacy_container::LegacyContainer::open(data)?;
    match container.format() {
        Some(binary_office::BinaryFormat::Word) => word97::decrypt(&mut container, password)?,
        Some(binary_office::BinaryFormat::Excel) => excel97::decrypt(&mut container, password)?,
        Some(binary_office::BinaryFormat::PowerPoint) => {
            powerpoint97::decrypt(&mut container, password)?
        }
        None => {
            return Err(Error::MissingStream(
                "WordDocument, Workbook or PowerPoint Document",
            ))
        }
    }
    container.into_bytes()
}

/// `EncryptionInfo.Reserved` for agile encryption — [MS-OFFCRYPTO] §2.3.4.10 requires
/// exactly this value in bytes 4..8 of the stream.
///
/// Deliberately **not** in [`limits`]: that module's own doc comment excludes "lengths
/// fixed by the on-disk format rather than by a field in it", and this is a fixed word
/// in a header rather than a bound on anything a file chooses. It lives beside the code
/// that reads the offsets it describes, next to the `vMajor` / `vMinor` pair it follows.
///
/// `classify` reads the same header and does not check this, which is correct rather
/// than an oversight: a classifier reports what a file claims to be and must never fail
/// loudly (see [`classify()`]'s contract), so a wrong Reserved word there belongs in the
/// verdict, not in a refusal. Acting on the file is where the refusal belongs.
#[cfg(feature = "crypto-ops")]
const AGILE_ENCRYPTION_RESERVED: u32 = 0x0000_0040;

/// Decrypt a password-protected OOXML file, choosing what happens about the package HMAC.
///
/// Returns the decrypted OOXML ZIP bytes together with the [`IntegrityOutcome`] that
/// actually applied — so a caller can tell "verified" from "this format has no tag to
/// verify" without inspecting the file itself.
///
/// The check runs **before** the package is decrypted (the HMAC covers ciphertext), so
/// a failure means no plaintext was produced, not that plaintext was produced and
/// withheld. On any error, `data` is left untouched.
///
/// # Errors
///
/// Every variant [`decrypt_ooxml`] documents, plus:
///
/// - [`Error::IntegrityUnavailable`] — [`IntegrityPolicy::Require`] on
///   ECMA-376 standard encryption, which defines no integrity element at all. An *agile*
///   file that omits `<dataIntegrity>` is [`Error::IntegrityElementMissing`]
///   instead, under `Require` and under the default alike.
///
/// [`Error::BadParameters`] also covers a declared tag whose parameters this
/// crate cannot use: a `hashSize` that contradicts the hash `<keyData>` names, a
/// `blockSize` that is not the AES block, or a blob that is not a block multiple or is
/// shorter than the digest. A `hashAlgorithm` this crate does not implement is
/// [`Error::UnsupportedAlgorithm`], never `BadParameters`: `agile::resolve_hash`
/// refuses the name before `integrity::verify` is reached, and the file is well formed —
/// the password may be exactly right.
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
/// assert!(decrypted.integrity.is_authenticated());
/// assert!(decrypted.package.starts_with(b"PK\x03\x04"));
/// # Ok::<(), msoffice_crypto::Error>(())
/// ```
///
/// # See Also
///
/// [`decrypt_ooxml`] is this function under [`IntegrityPolicy::RequireWhereDefined`],
/// discarding the outcome. [`IntegrityPolicy`] documents each choice; [`classify()`]
/// reports [`IntegrityDeclaration`] before any decrypt.
#[cfg(feature = "crypto-ops")]
pub fn decrypt_ooxml_with_policy(
    data: &[u8],
    password: &str,
    policy: IntegrityPolicy,
) -> Result<Decrypted, Error> {
    if !is_cfb_office(data) {
        return Err(Error::NotACfbFile);
    }

    let streams = cfb_reader::read_cfb_streams(data)?;

    // The version header is eight bytes. Copying it into an array is how a short stream
    // becomes [`Error::MissingStream`] rather than a panic on `try_into` or
    // on indexing past the end — both of which clippy::missing_panics_doc would then
    // demand a `# Panics` section for, on a path that is an error, not a panic.
    let header: [u8; 8] = streams
        .encryption_info
        .get(..8)
        .and_then(|s| s.try_into().ok())
        .ok_or(Error::MissingStream("EncryptionInfo too short"))?;

    // Bytes 0-1: vMajor (LE u16), bytes 2-3: vMinor (LE u16)
    let v_major = u16::from_le_bytes([header[0], header[1]]);
    let v_minor = u16::from_le_bytes([header[2], header[3]]);

    match (v_major, v_minor) {
        // Agile Encryption — XML starts after the 8-byte header (4-byte version + 4-byte reserved)
        (4, 4) => {
            // [MS-OFFCRYPTO] §2.3.4.10: the four bytes after `EncryptionVersionInfo` are
            // `Reserved` and MUST be `0x00000040` — 0x40, not zero, which is the one
            // thing about this field that is easy to get backwards. A structural check
            // on attacker-supplied bytes that runs *before* the XML parser is handed
            // anything, for the cost of one comparison.
            //
            // Gated to this arm on purpose: for `vMinor = 2` (standard encryption) the
            // same four bytes are `EncryptionHeader.Flags`, a different field entirely —
            // `standard_encrypted.docx` carries `00 00 00 00` there — so a check hoisted
            // above the `match` would reject every Office 2007 file in existence.
            //
            // herumi checks it in the same shape this crate does — read the four bytes as
            // a little-endian `u32` and compare: `include/crypto_util.hpp:322-323`
            // (`const uint32_t reserved = cybozu::Get32bitAsLE(p + 4);
            // MS_ASSERT_EQUAL(reserved, 0x40u);`), immediately before its XML parse, with
            // the encode side writing `0x40` at `:350-353`. That is BSD-3, so the fact is
            // available from a permissive source, exactly as `limits::AGILE_SALT_SIZE`
            // records for its own range. LibreOffice checks it too
            // (`AgileEngine.cxx:522-530`, against `msfilter::AGILE_ENCRYPTION_RESERVED`
            // at `mscodec.hxx:441`; behaviour and constant only, nothing copied), and its
            // comparison *shape* is the one deliberately not followed: `readBytes`
            // resizes to what it actually read and `std::equal` then walks that range, so
            // a stream truncated inside the field compares fewer elements and passes.
            // `header` is `[u8; 8]`, so these four bytes exist.
            let reserved = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
            if reserved != AGILE_ENCRYPTION_RESERVED {
                return Err(Error::BadParameters(format!(
                    "agile EncryptionInfo Reserved is {reserved:#010x}; [MS-OFFCRYPTO] \
                     2.3.4.10 requires {AGILE_ENCRYPTION_RESERVED:#010x}"
                )));
            }
            agile::decrypt(
                &streams.encryption_info[8..],
                &streams.encrypted_package,
                password,
                policy,
            )
            .map(|(package, integrity)| Decrypted { package, integrity })
        }
        // Standard Encryption — binary EncryptionHeader starts after the 8-byte header
        // vMajor 2/3/4 all indicate Standard Encryption per MS-OFFCRYPTO spec.
        //
        // [MS-OFFCRYPTO] §2.3.4.5 defines no integrity element for this format: there is
        // no dataIntegrity, no HMAC, nothing to check. That is reported as
        // `IntegrityOutcome::NotApplicable` rather than as a failure — absence here is
        // the spec, not a defect in the file. `Require` is the one exception, and only
        // because the caller asked for a guarantee the format cannot give; it is
        // refused before any work is done rather than after.
        //
        // The standard half of `IntegrityPolicy`'s contract; the agile half is
        // `agile::check_integrity`. Matched exhaustively rather than written as
        // `if policy == Require`, so that this arm's tolerance of a new policy variant
        // is a decision someone makes rather than a fall-through they inherit — the
        // fail-closed default of GH #12 deliberately stops short of here.
        (2, 2) | (3, 2) | (4, 2) => {
            match policy {
                IntegrityPolicy::Require => {
                    return Err(Error::IntegrityUnavailable(
                        "ECMA-376 standard encryption (Office 2007) defines no integrity \
                         element. Re-saving the file with Office 2013 or later writes agile \
                         encryption, which does; IntegrityPolicy::RequireWhereDefined opens \
                         this one as it is, unauthenticated",
                    ))
                }
                IntegrityPolicy::RequireWhereDefined
                | IntegrityPolicy::VerifyIfPresent
                | IntegrityPolicy::Skip => {}
            }
            let package = standard::decrypt(
                &streams.encryption_info[8..],
                &streams.encrypted_package,
                password,
            )?;
            Ok(Decrypted {
                package,
                integrity: IntegrityOutcome::NotApplicable,
            })
        }
        _ => Err(Error::UnsupportedEncryptionVersion(v_major, v_minor)),
    }
}

#[cfg(test)]
#[path = "lib_tests.rs"]
mod tests;

// The README's examples are compiled and run by `cargo test --doc`, so they cannot silently
// rot. They are the only examples in this repository that nothing else checks: rustdoc never
// sees `README.md` otherwise, and a wrong field name in one of them was caught by hand twice
// before this existed.
//
// `cfg(doctest)` is set only while rustdoc *collects* doctests, never while it *builds*
// documentation, so this item reaches neither docs.rs nor `cargo doc` output and the README
// is not duplicated onto the crate page.
//
// Gated on `legacy-binary` as well, because that is the superset under which every README
// example compiles: the decrypt/encrypt pair needs `crypto-ops` and `decrypt_binary_office`
// needs `legacy-binary`. Without the gate, CI's `--no-default-features` doctest run would
// fail on imports the README does not feature-gate.
//
// It lives at the end of the file deliberately. The technique is `secure-gate`'s
// (`~/Projects/secure-gate-workspace/src/lib.rs`), which documents the same reasoning.
#[cfg(all(doctest, feature = "legacy-binary"))]
#[doc = include_str!("../README.md")]
pub struct ReadmeDoctests;
