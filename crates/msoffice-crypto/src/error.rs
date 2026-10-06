//! Failures from decryption and encryption.
//!
//! [`Error`] is re-exported at the crate root under `crypto-ops`. In a
//! detection-only build no public function returns a [`Result`], so the type is not
//! public there — see the enum's own docs.

/// Every reason this crate refuses a file.
///
/// `#[non_exhaustive]`: variants are still arriving — RC4 CryptoAPI decryption (GH #4),
/// the encrypt path (GH #6) and the encrypt guard have each added their own — and this
/// crate's own testing rule makes consumers match on variants by name, so a wildcard arm
/// has to be theirs to write.
///
/// Adding a variant therefore never fails a consumer's build. That is the hazard, not the
/// safety: their wildcard arm silently absorbs the new case, so a variant that splits a
/// fact they were handling changes what their code does with nothing red anywhere. The
/// crate is published on a release-candidate line where that is permitted (see CLAUDE.md
/// § *Project Status*), and the obligation it leaves is to name the split in the changelog
/// as a fact.
///
/// Thirteen of the eighteen variants exist only under `crypto-ops`, one only under
/// `legacy-binary`, and so does the public re-export of the type itself. In a
/// detection-only build no public function returns a `Result` — [`crate::classify()`]
/// answers every input and [`crate::is_cfb_office()`] is a `bool` — so the type would be a
/// public name nothing
/// produces. The four ungated variants are the ones the container reader constructs on
/// the way to `classify`, which swallows them.
///
/// Messages never carry key material. Where a message quotes the file —
/// [`Error::UnsupportedAlgorithm`]'s `name`, the lengths and declared values
/// interpolated into [`Error::BadParameters`] — it is bounded and truncated at
/// the construction site, and the variant says so. Match with a `_` arm: the enum does
/// not implement `PartialEq`.
///
/// [`Error::EncryptParams`] quotes nothing either, and for a reason none of the others
/// can claim: it is raised before any file is opened at all. Its three numbers are the
/// caller's own arguments and this crate's own bounds, and its two enum payloads are
/// fieldless.
///
/// The encrypt guard's three variants quote nothing at all.
/// [`Error::AlreadyEncrypted`] carries [`crate::Family`] and [`crate::Document`], which
/// are fieldless `Copy` enums from the detection half; the other two carry no payload.
/// None of the three has a byte of the file in it, and none interpolates its payload into
/// its message.
///
/// # Who the messages are written for
///
/// **A programmer choosing a policy, not the person holding the file.** Several name a
/// [`crate::IntegrityPolicy`] variant, which is a Rust path and useful only to whoever
/// writes the call. A consumer that renders a failure to an end user should **match on
/// the variant and write its own copy**, not forward `Display`.
///
/// That is a real hazard rather than a style note, and it was found by a consumer reading
/// its own test output. The three integrity failures are *not* symmetric in what their
/// opt-out costs:
///
/// - [`Error::IntegrityCheckFailed`] offers none, because there is none.
/// - [`Error::IntegrityUnavailable`]'s is a property of the **format** — a 2007 file has
///   no tag to check, and opening it anyway concedes only what that format never offered.
/// - [`Error::IntegrityElementMissing`]'s is "decrypt evidence of tampering anyway", on a
///   file whose missing element is, per that variant's docs, never innocent.
///
/// Forwarded uniformly, the third reaches the person holding the file as instructions for
/// opening the document an attacker prepared for them. The message says what the opt-out
/// concedes, so that forwarding it is at worst unhelpful rather than misleading — but the
/// fix is to map the variant, and this paragraph exists so that nobody has to discover
/// that by reading test output a second time.
///
/// # What `source()` returns, and why it is mostly `None`
///
/// `None` for every variant except [`Error::Io`] — deliberately, and not for want of
/// `thiserror`. The variants that wrap a foreign failure reduce it to a string instead of
/// holding it behind `#[source]`. Holding it would reopen, in two places rather than one,
/// the conduit that reduction exists to close: `source()` would hand a caller the
/// dependency's `Display`, and this enum's derived `Debug` would print it. A walkable
/// error chain is worth less here than a message this crate can characterise completely —
/// this type's whole input is a document an attacker wrote.
///
/// The three foreign failures are not treated alike, and the difference is the point:
///
/// | Variant | Foreign `Display` | Why |
/// | --- | --- | --- |
/// | [`Error::XmlParse`] | **never** — classified by an exhaustive match | quick-xml quotes text drawn from the document |
/// | [`Error::RandomSource`] | forwarded, truncated to 200 characters | names an environment failure; on a failure nothing was generated |
/// | [`Error::Io`] | forwarded | `cfb` 0.14.0 audited: lengths and fixed strings only |
///
/// [`Error::Io`] is the one `#[from]`, and it does forward `std::io::Error`'s `Display`.
/// That is audited rather than assumed: on these paths the only producers are `cfb`,
/// whose `invalid_data!` messages interpolate lengths and fixed strings and never a
/// stream name or a file byte (`direntry.rs:118-140` in 0.14.0), and `std::io::Cursor`,
/// which fails only on allocation. **Re-check it on a `cfb` bump** — that is the standing
/// cost of a `#[from]` on a foreign error type, and the reason there is only one.
// `unreachable_pub` fires on this type in the **detection** build and only there: the
// re-export at `lib.rs` is `#[cfg(feature = "crypto-ops")]`, so without that feature this
// is a `pub` item in a private module that nothing re-exports. That is the design the doc
// above describes, not an oversight — the type is the crate's public error under
// `crypto-ops` and `legacy-binary`, and `pub(crate)` would delete it from the public API.
//
// `allow` rather than `expect`: the lint fires in one of the three configurations, so an
// `expect` would itself go unfulfilled in the other two.
#[allow(unreachable_pub)]
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The bytes are not a CFB container: they lack the eight-byte magic, or `cfb`
    /// refused to open them.
    #[error("not a CFB (Compound Binary File) container")]
    NotACfbFile,

    /// A binary document handed to [`crate::decrypt_binary_office()`] carries no password-to-open
    /// at all: the FIB's `fEncrypted` bit is clear, the workbook has no `FILEPASS`
    /// record, the presentation's `UserEditAtom` is the unencrypted shape.
    ///
    /// Its own variant, per CLAUDE.md § *Cryptographic Rules*: "there is nothing to
    /// decrypt" is a fact a caller acts on by using the bytes it already has, which is
    /// the opposite of what it does for every other variant here. [`crate::classify()`]
    /// reports the same file as [`crate::Family::Unencrypted`] wherever its own probe
    /// reached the marker; a workbook whose record walk ran out before proving anything is
    /// [`crate::Family::Unknown`] there — `Unencrypted` is a proof, not a default — while
    /// this path still answers `NotEncrypted`.
    #[error("the document is not encrypted; there is nothing to decrypt")]
    #[cfg(feature = "legacy-binary")]
    NotEncrypted,

    /// A required CFB stream is missing, unreadable, shorter than its header — or
    /// carries a fixed structural field this crate cannot proceed past.
    ///
    /// The payload names the stream or the structure (`"EncryptionInfo"`, `"Workbook"`,
    /// `"EncryptionVerifier missing or truncated"`, …). It is always a fixed
    /// `&'static str`, never attacker-chosen file content.
    #[error("required CFB stream missing or unreadable: {0}")]
    MissingStream(&'static str),

    /// The `EncryptionInfo` version pair names a format this crate does not implement.
    ///
    /// The message says only that, because that is all the pair tells us. It used to
    /// assert "Office XP/2003 RC4 encryption is not supported — re-save the file with
    /// Office 2007 or later" for *every* pair that reached it, including `vMinor = 3`
    /// (extensible encryption) and, until `standard::aes_key_bits` landed (as
    /// `require_aes_128`, which is what it was called then), every
    /// Office 2007 file whose `AlgID` was the conforming `0x660E`. A cipher this crate
    /// has not implemented is [`Error::UnsupportedAlgorithm`]; this variant is
    /// for the version pair alone.
    #[error(
        "unsupported Office encryption version {0}.{1} \
        (this crate implements ECMA-376 standard encryption, vMinor 2, \
        and agile encryption, 4.4)"
    )]
    #[cfg(feature = "crypto-ops")]
    UnsupportedEncryptionVersion(u16, u16),

    /// The agile `EncryptionInfo` XML could not be parsed, or a required attribute was
    /// absent.
    ///
    /// Distinct from [`Self::BadParameters`]: that one is a value that parsed and is
    /// out of range; this one is XML that did not yield a value at all.
    ///
    /// **The string is always one this crate wrote.** No dependency's `Display` reaches
    /// it: `agile::xml_error` and `agile::base64_error` classify quick-xml and base64
    /// failures into fixed descriptions, and both match exhaustively, so a variant added
    /// upstream is a compile error here rather than a silent forward.
    ///
    /// That is a correction and not a description of how it always was. Until
    /// 2026-09-11 the quick-xml `Display` was forwarded verbatim by four call sites, three
    /// of which run on attribute *values* — `saltValue`, `encryptedKeyValue` and the two
    /// verifier blobs. quick-xml quotes the text between `&` and the next `;` of whatever
    /// it is unescaping, so a crafted `encryptedKeyValue` put its own text here, bounded
    /// only by `limits::ENCRYPTION_INFO_READ_CAP` — 1 MiB. Not key material, since the
    /// attacker supplies the text, but unbounded attacker-chosen content in an error
    /// string is the thing [`Error::UnsupportedAlgorithm`] truncates to 32 characters and
    /// documents doing. Found by the downstream consumer, reading this file.
    ///
    /// The guard is
    /// `malformed_input::a_hostile_entity_in_an_attribute_value_never_reaches_the_error_message`,
    /// with a well-formed-value control beside it; reverting `xml_error` to
    /// `e.to_string()` fails it with a 4609-byte message quoting 4608 bytes of the file.
    #[error("EncryptionInfo XML parse error: {0}")]
    #[cfg(feature = "crypto-ops")]
    XmlParse(String),

    /// The password verifier did not match.
    ///
    /// Distinct from [`Self::UnsupportedAlgorithm`] on purpose: that one is a file this
    /// crate cannot act on, and reporting it as a wrong password sends the user looking
    /// for a typo in a password that was right.
    #[error("wrong password")]
    #[cfg(feature = "crypto-ops")]
    WrongPassword,

    /// An AES operation rejected its input.
    ///
    /// Typically a ciphertext length that is not a block multiple — a truncated stream,
    /// not a wrong password. The encrypt path reaches it only if an internal length
    /// invariant fails.
    #[error("cipher operation failed")]
    #[cfg(feature = "crypto-ops")]
    CipherError,

    /// A parameter the file declares is out of range, or is inconsistent with a sibling
    /// parameter.
    ///
    /// The message describes the *shape* of the problem — a length, an attribute name,
    /// an algorithm name. It never carries key material or file content.
    ///
    /// Contrast [`Error::UnsupportedAlgorithm`], which is the file being
    /// internally consistent about something this crate has not implemented.
    #[error("unsupported or inconsistent encryption parameter: {0}")]
    BadParameters(String),

    /// The file names an algorithm this crate does not implement.
    ///
    /// **Distinct from [`Error::WrongPassword`] on purpose.** The password
    /// may be perfectly correct and simply unusable: before this variant existed, an
    /// agile file declaring `hashAlgorithm="SHA256"` derived a SHA-512 spin hash, failed
    /// the verifier comparison, and was reported as a wrong password — telling the user
    /// the one thing that was not true. `office-crypto` already returned a dedicated
    /// `Unimplemented` here (`src/lib.rs:29`), so silently misdiagnosing was strictly
    /// worse than the alternative this crate exists to improve on.
    ///
    /// `what` names the attribute — a fixed set of `&'static str`, never formatted, so
    /// this cannot become a second free-text channel. `name` is the file's own spelling,
    /// truncated: it is attacker-chosen XML attribute text, bounded only by
    /// `limits::ENCRYPTION_INFO_READ_CAP` (1 MiB) before it reaches here.
    ///
    /// # This refusal is authorised by the format, which is a third thing
    ///
    /// Not a spec violation and not merely an implementation limit. [MS-OFFCRYPTO]
    /// §2.3.4.10, immediately after the `HashAlgorithm` table that lists MD5, MD4, MD2,
    /// RIPEMD-128, RIPEMD-160 and WHIRLPOOL alongside the SHA family: **"Values that are
    /// not defined MAY be used, and a compliant implementation is not required to support
    /// all defined values."**
    ///
    /// So a document naming WHIRLPOOL is conforming, this crate refusing it is
    /// conforming, and both are true at once. A consumer rendering this has three
    /// sentences available and should not collapse them:
    ///
    /// | | what the format says |
    /// | --- | --- |
    /// | a spec violation | the file is wrong |
    /// | an implementation limit | the format permits it; *we could* and chose not to, so an override is coherent |
    /// | **this** | the format permits the value **and permits us not to support it** |
    ///
    /// The third is not a weaker form of the second. An override is not a coherent thing
    /// to offer, because there is nothing to override — the format contemplated the
    /// refusal.
    ///
    /// Worth contrasting with the `spinCount` ceiling, which looked like the same shape
    /// and was not: the spec states a maximum there and carries no clause permitting a
    /// narrower one, so this crate's tighter bound was a defect and was removed. Here the
    /// clause exists. Whether to add the older hashes is a support decision; it is not a
    /// conformance one, and the difference is this sentence.
    #[error("{what} names an algorithm this crate does not implement: {name}")]
    #[cfg(feature = "crypto-ops")]
    UnsupportedAlgorithm {
        /// Which attribute named it, e.g. `p:encryptedKey/@hashAlgorithm`.
        what: &'static str,
        /// Either the algorithm name as the file spells it, truncated to 32 characters
        /// by `agile::unsupported_algorithm` — the one path where this is
        /// attacker-chosen text — or a fixed description of what was refused, sometimes a
        /// formatted `AlgID`. Never key material, and never unbounded file text.
        name: String,
    },

    /// The `dataIntegrity` HMAC did not match the `EncryptedPackage` stream.
    ///
    /// The password is verified before this check runs, so a failure here means the
    /// package was corrupted or modified — by someone who did *not* hold the password,
    /// since producing a matching tag requires the session key.
    #[error(
        "package integrity check failed: the encrypted package does not match its \
        dataIntegrity HMAC (the file is corrupt or was modified after encryption)"
    )]
    #[cfg(feature = "crypto-ops")]
    IntegrityCheckFailed,

    /// A file declaring **agile** encryption carries no `<dataIntegrity>` element, under
    /// a policy that requires one — which since GH #12 includes the default policy.
    ///
    /// Distinct from [`Self::IntegrityUnavailable`] on purpose, per CLAUDE.md
    /// § *Cryptographic Rules*: that one is a caller asking a format for a guarantee the
    /// format does not define, a property of the **request**. This one is a file that
    /// should carry a tag and does not, a property of the **file** — and deleting the
    /// element is the cheapest tamper there is, needing no password. A caller acts on
    /// the two differently, so they are different variants.
    ///
    /// The message names the opt-out, because at least one consumer flattens this error
    /// to its `Display` string and a dead end is a worse answer than a signposted one.
    /// The `Display` text names [`crate::IntegrityPolicy`]'s opt-out in plain words
    /// because a `#[error]` string cannot carry an intra-doc link.
    ///
    /// **This is the one opt-out in the enum that concedes something a caller should not
    /// concede lightly**, and the message therefore carries its cost rather than only its
    /// name. [`Self::IntegrityCheckFailed`] offers no way out because none exists, and
    /// [`Self::IntegrityUnavailable`]'s is a limit of the format; this one is "accept a
    /// file whose tamper-evidence was deleted", on a file where — see above — the deletion
    /// is never innocent. A consumer that forwards `Display` to an end user forwards that
    /// too. See the enum's own docs: render integrity failures by matching the variant.
    #[error(
        "this file declares agile encryption but carries no <dataIntegrity> element: \
        every known agile writer emits one, so it was removed or the file is malformed. \
        IntegrityPolicy::VerifyIfPresent decrypts it unverified, which accepts a file \
        whose tamper-evidence is absent"
    )]
    #[cfg(feature = "crypto-ops")]
    IntegrityElementMissing,

    /// An integrity guarantee was demanded of a format that defines no integrity tag at
    /// all — `IntegrityPolicy::Require` on ECMA-376 standard encryption (Office 2007).
    ///
    /// An *agile* file missing its element is [`Self::IntegrityElementMissing`] instead:
    /// that is a defect in the file, this is a limit of the format. The two want different
    /// words in front of a user — "this format cannot prove it was not modified" against
    /// "this file's proof was removed" — which is the whole reason they are two variants.
    ///
    /// Like that sibling, the message names the way out, because a consumer flattens this
    /// to its `Display`. Measured at one: the string reaches a user as *"integrity
    /// verification was required, but ECMA-376 standard encryption (Office 2007) defines
    /// no integrity element…"*, and a message that states the problem without the remedy
    /// is a dead end where a signposted one costs nothing.
    ///
    /// Both this variant and [`crate::IntegrityPolicy`] exist only under `crypto-ops`,
    /// so the link resolves in every configuration that renders it.
    #[error("integrity verification was required, but {0}")]
    #[cfg(feature = "crypto-ops")]
    IntegrityUnavailable(&'static str),

    /// The random source failed while generating key material.
    ///
    /// Its own variant rather than folding into [`Self::CipherError`], per CLAUDE.md
    /// § *Cryptographic Rules*: a CSPRNG that will not produce bytes is an environment
    /// failure a caller may be able to act on — a sandbox with no `getrandom`, an
    /// exhausted file-descriptor table — and it is emphatically **not** the same fact as
    /// "the cipher rejected this input". Encryption is the only path that reaches it;
    /// decryption generates nothing.
    ///
    /// The string is the RNG's own `Display`, which describes *the source* — never its
    /// output. No generated byte can reach here: on failure there is nothing generated.
    /// Truncated to 200 characters at the construction site, `agile_encrypt::random_source`,
    /// which explains why this one foreign `Display` is forwarded where
    /// [`Self::XmlParse`]'s is not.
    #[error("the random source failed while generating key material: {0}")]
    #[cfg(feature = "crypto-ops")]
    RandomSource(String),

    /// A caller asked an encrypt entry point for encryption parameters this crate will
    /// not write.
    ///
    /// **This is a caller error, raised before a single byte of any document has been
    /// examined.** Nothing in it describes the input, so a consumer that renders it as
    /// "your document is damaged" tells the holder of a perfectly good file that their
    /// file is broken. Often there is no file yet: the parameters are validated at the
    /// top of the call, ahead of the container sniff that produces
    /// [`Self::AlreadyEncrypted`] and its siblings. The remedy is always to change the
    /// argument, never to change the document — which is why this variant maps to the
    /// CLI's usage exit code and not to any of its file codes.
    ///
    /// Distinct from [`Self::BadParameters`], and the distinction is *whose number it
    /// is*. `BadParameters` is a value the **file** declares about itself, arriving out
    /// of an attacker-supplied `EncryptionInfo` stream; this is a value the **caller**
    /// passed in. They are raised at opposite ends of the crate and acted on by different
    /// people — one by whoever chose the file, one by whoever wrote the call — so
    /// collapsing them would hand a programmer's mistake to an end user as a verdict on
    /// their document.
    ///
    /// **No file byte can reach this variant.** `got`, `min` and `max` are `u32`s the
    /// caller supplied or that this crate's own `limits` module holds, and `param` and
    /// `problem` are fieldless `Copy` enums. There is deliberately no `String` field: the
    /// free-text channel `BadParameters` needs is exactly what makes it the wrong variant
    /// for a caller error, and opening one here would invite the same
    /// file-content-in-a-message defect [`Self::XmlParse`] documents having had.
    ///
    /// # Why the problem is a second enum rather than part of the first
    ///
    /// [`EncryptParamProblem`] is `src/limits.rs`'s own **spec / cipher / margin**
    /// provenance taxonomy made matchable — see the table in that module's header. A
    /// bound this crate enforces is one of three quite different kinds of fact, and until
    /// they were labelled a consumer had to read each doc comment and infer, which one
    /// did, got partly wrong, and had to ask about. Reporting a *margin* as though it
    /// were a spec violation is a typed lie, and a typed lie is worse than a vague
    /// message: a consumer renders a typed fact as authoritative and tells its user the
    /// format forbids something the format permits.
    /// [`EncryptParamProblem::ExceedsImplementationLimit`] exists so that never has to
    /// happen.
    #[error(
        "encryption parameter rejected before any file was read: {param} {problem} \
        (requested {got}; this crate accepts {min}..={max})"
    )]
    #[cfg(feature = "crypto-ops")]
    EncryptParams {
        /// Which `EncryptionInfo` attribute the rejected request was setting.
        param: EncryptParam,
        /// Whose rule the value broke — the format's, the cipher's, or this crate's own.
        problem: EncryptParamProblem,
        /// The value the caller asked for. Caller-supplied; never read from a file.
        got: u32,
        /// The low end of an accepted value for `param`.
        ///
        /// For a range-valued parameter — `spinCount`, `saltSize` — this is the bound's
        /// real floor and `min..=max` is the whole accepted set. For a **set**-valued one
        /// it cannot be: `keyBits` accepts `{128, 192, 256}`, and printing that as
        /// `128..=256` would name 200 as acceptable when the next call refuses it. Those
        /// refusals therefore carry `min == max`, the single nearest accepted size, so
        /// that every number in the message is one the caller may actually use.
        ///
        /// The cost, stated rather than hidden: for `keyBits` the pair is no longer the
        /// *whole* accepted set, only a correct member of it. Completeness is the right
        /// thing to lose — a caller who follows `192..=192` succeeds, and a caller who
        /// follows `128..=256` may not.
        min: u32,
        /// The high end of an accepted value for `param`. Equal to `min` where the
        /// accepted values are a set rather than a range — see `min`.
        max: u32,
    },

    /// The bytes handed to an encrypt entry point already carry a password-to-open.
    ///
    /// **This crate used to encrypt them anyway.** The result was a CFB container wrapped
    /// in a second CFB container, indistinguishable from a single wrap without decrypting
    /// it, and openable only by decrypting twice with two passwords the holder believed
    /// was one. It was found by a consumer, which had to reimplement this crate's own CLI
    /// guard to avoid it — the definition of a check that was in the wrong place.
    ///
    /// Distinct from [`Self::NotAPlainPackage`], and the distinction is the remedy: this
    /// file can be decrypted and re-encrypted, and that one cannot be encrypted at all. A
    /// renderer that collapses the two tells the holder of a `.doc` to decrypt it first,
    /// which this crate will do and then refuse to undo.
    ///
    /// The partition is [`crate::Classification::is_encrypted`], so a CFB whose container
    /// could not be read ([`crate::ContainerRead::Unreadable`]) reports
    /// [`Self::NotAPlainPackage`] rather than this. Both are refusals, so nothing is
    /// admitted because a family could not be determined.
    ///
    /// `family` and `document` are [`crate::classify()`]'s verdict on the same bytes:
    /// `Copy` enums from the detection half, never text, so no part of the file reaches
    /// this variant. They are payload rather than message because the wording a caller
    /// wants depends on them — "decrypt it first" is a remedy for an encrypted package and
    /// a dead end for an encrypted `.doc`, which has no writer in any build — and because
    /// making a caller re-run `classify()` to learn a fact this refusal already
    /// established is what the note above tells consumers they should not have to do.
    /// Neither is interpolated into the message: neither has a `Display` impl,
    /// deliberately, because naming them is a renderer's job.
    #[error(
        "the input is already encrypted: this crate encrypts a plain OOXML package, and \
        these bytes are an Office-encrypted CFB container"
    )]
    #[cfg(feature = "crypto-ops")]
    AlreadyEncrypted {
        /// The encryption family [`crate::classify()`] found.
        family: crate::Family,
        /// The document kind it found, or [`crate::Document::Unknown`] where it could not
        /// tell — eight bytes of CFB magic are `Unknown`, and the remedy differs by kind.
        document: crate::Document,
    },

    /// The bytes are a CFB container that is not an encrypted package: a 97-2003 binary
    /// document, or a container this crate could not read far enough to say.
    ///
    /// Distinct from [`Self::AlreadyEncrypted`] because there is no remedy. This crate has
    /// no writer for the 97-2003 binary formats, so "decrypt it first" is advice that ends
    /// in a second refusal in every build.
    ///
    /// **Deliberately carries no `Document`.** The commonest input here — a CFB whose
    /// directory is unreachable — is [`crate::Document::Unknown`], and a message built
    /// from it would name a document kind [`crate::classify()`] refused to name. The
    /// variant has no field to build one from, so that cannot be written by accident.
    #[error(
        "the input is a CFB container, not a plain OOXML package: this crate writes \
        encryption around a .docx/.xlsx/.pptx ZIP, and there is no writer for the \
        97-2003 binary formats"
    )]
    #[cfg(feature = "crypto-ops")]
    NotAPlainPackage,

    /// The bytes are neither a ZIP package nor a CFB container.
    ///
    /// Its own variant rather than a case of [`Self::NotAPlainPackage`] because a caller
    /// acts on it differently, and this crate's CLI already proves it: an unrecognised
    /// container is "you handed me the wrong file" (exit 3) and a CFB is "I recognise this
    /// and will not write into it" (exit 5). Two facts, two numbers, two variants.
    ///
    /// Not [`Self::NotACfbFile`], which is the decrypt side's refusal and means the
    /// opposite thing: there a CFB is what was wanted, and here it is one of the two
    /// shapes being refused.
    #[error(
        "the input is neither an OOXML package nor a CFB container: there is nothing \
        here to encrypt"
    )]
    #[cfg(feature = "crypto-ops")]
    UnknownContainer,

    /// Reading or writing the in-memory CFB container failed.
    ///
    /// The inner [`std::io::Error`] describes the operation — a flush, a stream write —
    /// never key material or file content.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

/// Which `EncryptionInfo` attribute a rejected encryption request named.
///
/// Half the payload of [`Error::EncryptParams`]; the other half is
/// [`EncryptParamProblem`]. A fieldless `Copy` enum rather than the `&'static str` the
/// neighbouring [`Error::UnsupportedAlgorithm`] uses, because this one is a caller's
/// mistake and a caller acts on it in code: a string forces a string comparison, which
/// is a match nobody can make exhaustive and nobody notices going stale.
///
/// `Display` gives the attribute's own spelling, as [MS-OFFCRYPTO] §2.3.4.10 spells it
/// in `CT_KeyData` and `CT_PasswordKeyEncryptor`, so the `#[error]` string interpolates
/// it directly and a consumer with no match arm still reads the real attribute name
/// rather than a Rust identifier.
///
/// `#[non_exhaustive]`: `cipherAlgorithm`, `cipherChaining` and `blockSize` are absent
/// because this crate writes exactly one value for each and so has no caller parameter
/// to reject — `blockSize` is 16 because AES-CBC fixes it, not because anyone chose it.
/// If one of them ever becomes a caller's choice, it arrives here as a variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg(feature = "crypto-ops")]
#[non_exhaustive]
pub enum EncryptParam {
    /// `spinCount`: the iteration count of the password derivation ([MS-OFFCRYPTO]
    /// §2.3.4.11).
    ///
    /// `ST_SpinCount` is `0..=10000000` (§2.3.4.10) — `minInclusive="0"`, and the prose
    /// states only "It MUST NOT be greater than 10,000,000". **The floor is zero, and
    /// only one end is live.**
    ///
    /// The single refusal is [`EncryptParamProblem::OutsideSpecRange`], for a value above
    /// 10 000 000. Below that nothing rejects: the field is a `u32`, which already
    /// enforces `minInclusive="0"`, so the floor needs no check and cannot be tripped,
    /// and `limits::SPIN_COUNT_MAX` is the spec's own figure rather than a reduced one,
    /// so there is **no margin between the two and no second refusal to report**.
    ///
    /// In particular this crate does **not** refuse a weak spin count.
    /// `spinCount="0"` is conforming, this writer emits it on request, and an earlier
    /// version of this doc claimed an
    /// [`EncryptParamProblem::ExceedsImplementationLimit`] refusal for values "too weak
    /// to author" that no code path has ever produced. Declining to have an opinion about
    /// a writer's stretching is the position `EncryptParams::validate` actually takes.
    SpinCount,

    /// `hashAlgorithm`: the hash driving the key derivation and the verifier.
    ///
    /// **A hash algorithm is not defined by a cipher**, so a refusal here is *not*
    /// [`EncryptParamProblem::UnsupportedByCipher`], whatever an earlier version of this
    /// doc said. AES fixes key sizes and a block length; it has no opinion about SHA-512,
    /// and citing it as the authority that refused a hash names an authority with no say.
    ///
    /// The four hashes this crate writes — SHA-1, SHA-256, SHA-384, SHA-512 — are a
    /// **crate choice the format explicitly licenses**. [MS-OFFCRYPTO] §2.3.4.10, in the
    /// sentence immediately after the `HashAlgorithm` table (the table that also lists
    /// MD5, MD4, MD2, RIPEMD-128, RIPEMD-160 and WHIRLPOOL), says:
    ///
    /// > Values that are not defined MAY be used, and a compliant implementation is not
    /// > required to support all defined values.
    ///
    /// **That is a third shape, distinct from both of the others**, and the distinction
    /// is what a consumer needs to write refusal copy. A spec violation says the caller
    /// asked for a file no conforming reader need accept; an implementation margin says
    /// the format and the cipher both allow it and this crate declines. This is neither:
    /// the format permits the value **and** permits an implementation not to support it,
    /// so the honest sentence is "that hash is conforming, and this writer does not emit
    /// it" — no fault on either side.
    ///
    /// No refusal carries this variant today. `EncryptParams::validate`'s exhaustive
    /// match over [`crate::HashAlgorithm`] rejects nothing; it is a tripwire that turns a
    /// fifth variant added to that enum into an `E0004` at the one place that must decide
    /// whether the writer may emit it. Should such a refusal ever be wanted, it arrives
    /// with an [`EncryptParamProblem`] variant that states the shape above — not by
    /// borrowing `UnsupportedByCipher`, and not by borrowing
    /// [`EncryptParamProblem::OutsideSpecRange`], which would assert that the format
    /// forbids what the clause quoted here expressly allows.
    HashAlgorithm,

    /// `keyBits`: the size of a key, in bits.
    ///
    /// **The format has two of these and they are independent quantities.** §2.3.4.13
    /// step 1 sizes the package and intermediate key from `Encryptor.KeyData.keyBits`,
    /// while §2.3.4.11 sizes the password-derived key-encrypting key from
    /// `PasswordKeyEncryptor.keyBits`. §2.3.4.10 requires `p:encryptedKey` to match
    /// `keyData` on `hashAlgorithm` and on `cipherAlgorithm` — and on nothing else, so
    /// no rule makes these two equal. This variant names whichever of them the rejected
    /// call was setting.
    ///
    /// **Two rules refuse a `keyBits`, and they are not the same authority.**
    /// `ST_KeyBits` sets `minInclusive="8"`, requires a multiple of 8, and states **no
    /// maximum**, being generic across cipher algorithms. AES then defines exactly three
    /// sizes inside that. So:
    ///
    /// | value | refused by | `problem` |
    /// | --- | --- | --- |
    /// | 7, 12 | `ST_KeyBits` — below 8, or not a multiple of 8 | [`EncryptParamProblem::OutsideSpecRange`] |
    /// | 64, 512 | AES — legal `ST_KeyBits`, no such key size | [`EncryptParamProblem::UnsupportedByCipher`] |
    ///
    /// The spec test runs first, because "AES has no 7-bit key" is true and beside the
    /// point: 7 is not a legal `keyBits` for any cipher, and naming AES would credit a
    /// refusal to an authority that never had jurisdiction.
    ///
    /// This doc said for one commit that every refusal here was
    /// `UnsupportedByCipher` — written the same day the variant that exists to prevent
    /// exactly that mislabelling was added, and caught by an adversarial audit rather
    /// than by a test. The prose form of the defect is the harder one to see, because
    /// nothing compiles it.
    ///
    /// **This variant is the agile format's XML attribute only.** ECMA-376 *standard*
    /// encryption states its key size in a different field, in a different structure,
    /// under a different clause; it is [`Self::KeySize`], and the two are not
    /// interchangeable.
    KeyBits,

    /// `EncryptionHeader.KeySize`: the key size of an ECMA-376 **standard** (Office
    /// 2007) file, in bits ([MS-OFFCRYPTO] §2.3.2).
    ///
    /// Separate from [`Self::KeyBits`] rather than folded into it, because the two name
    /// different fields under different rules and a consumer acts on the difference:
    ///
    /// | | [`Self::KeyBits`] | this variant |
    /// | --- | --- | --- |
    /// | field | `keyBits`, an XML attribute of `<keyData>` / `<p:encryptedKey>` | `KeySize`, a `u32` in the binary `EncryptionHeader` |
    /// | typed by | `ST_KeyBits` (§2.3.4.10) — `minInclusive="8"`, a multiple of 8, **no maximum** | §2.3.4.5's header table — `0x00000080`, `0x000000C0` or `0x00000100` |
    /// | written by | [`crate::encrypt_ooxml_with_params`] | [`crate::encrypt_ooxml_standard_with_key_bits`] |
    ///
    /// `Display` prints `EncryptionHeader.KeySize`, so a consumer forwarding the sentence
    /// names the field its user would find in the file rather than an attribute that is
    /// not in it.
    ///
    /// **Its only `problem` is [`EncryptParamProblem::OutsideSpecRange`], and
    /// [`EncryptParamProblem::UnsupportedByCipher`] is unreachable here** — the opposite
    /// of [`Self::KeyBits`], where both fire. `ST_KeyBits` is generic across cipher
    /// algorithms, so there the format and AES are two authorities with different reach.
    /// §2.3.4.5 names AES for this stream and then enumerates AES's three key sizes
    /// itself, so here the format's set and the cipher's set are one set and the format
    /// states it first. Reporting 64 as a cipher limitation would assert that this header
    /// may carry a 64-bit key — §2.3.4.5 says it may not — which is the typed lie
    /// [`EncryptParamProblem`]'s split exists to prevent. `AesKeySize::new` in
    /// `src/standard_encrypt.rs` carries the argument in full, including why §2.3.2's
    /// RC4 and `0x00000000` rows are not a counter-example.
    KeySize,

    /// `saltSize`: the length in bytes of a salt.
    ///
    /// Per-element in the same way `keyBits` is: §2.3.4.10 binds each `saltSize` only to
    /// its own element's `saltValue`, whose decoded form "MUST be" that many bytes, and
    /// imposes no equality between `keyData`'s and `p:encryptedKey`'s.
    ///
    /// `ST_SaltSize` is `1..=65536`, and **this crate adopts that range whole**, so the
    /// single refusal is [`EncryptParamProblem::OutsideSpecRange`]: zero at one end, and
    /// anything above 65 536 at the other. A 65 536-byte salt is absurd and conforming,
    /// and `EncryptParams::validate` writes it — an earlier version of this doc called
    /// that an [`EncryptParamProblem::ExceedsImplementationLimit`] refusal, which no code
    /// path produced, because `limits::AGILE_SALT_SIZE` is the schema's range and not a
    /// reduced one.
    SaltSize,

    /// keyBits AND hashAlgorithm together — neither wrong alone.
    ///
    /// The case the single-attribute variants cannot express: both values are acceptable
    /// on their own and the pair is not, so naming either one alone would send the caller
    /// to change the parameter that was fine.
    ///
    /// The concrete pair is a digest shorter than the key asked for — SHA-1's 20 bytes
    /// with `keyBits="256"`, which wants 32. [MS-OFFCRYPTO] §2.3.4.11 says such a key
    /// "MUST be padded by appending bytes with a value of 0x36"; `agile::derive_block_key`
    /// refuses to do it, because the pad manufactures twelve constant bytes of key that
    /// no writer ever used. That refusal is this crate's decrypt-side property, and this
    /// variant is the encrypt side declining to author the file that would trip it.
    KeyBitsWithHash,
}

/// Whose rule a rejected encryption parameter broke.
///
/// The other half of [`Error::EncryptParams`]'s payload, and the reason that variant is
/// shaped as two enums rather than one. It is `src/limits.rs`'s **spec / cipher /
/// margin** provenance taxonomy — see the table in that module's header — turned into
/// something a consumer can `match` on, plus a fourth case for two values that are each
/// fine alone.
///
/// | variant | label in `limits.rs` | was the request non-conforming? |
/// | --- | --- | --- |
/// | [`Self::OutsideSpecRange`] | **spec** | yes |
/// | [`Self::UnsupportedByCipher`] | **cipher** | no — AES does not define it |
/// | [`Self::UnusableCombination`] | — | no, not value by value |
/// | [`Self::ExceedsImplementationLimit`] | **margin** | no — this crate declines |
///
/// Every `Display` string names whose rule it was, so that a consumer which forwards the
/// message instead of matching on it still says something true. That is the whole point
/// of the fourth variant: a margin reported as a spec violation is a typed lie, and a
/// typed lie is worse than a vague one, because a consumer renders a typed fact as
/// authoritative.
///
/// The fourth is also the one nothing currently produces, because this crate currently
/// imposes no margin on a caller's parameter. [`Self::ExceedsImplementationLimit`]'s own
/// docs say why it is kept anyway.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg(feature = "crypto-ops")]
#[non_exhaustive]
pub enum EncryptParamProblem {
    /// Outside the range [MS-OFFCRYPTO] states for the attribute.
    ///
    /// The **spec** row of `limits.rs`'s table: the number the format itself states —
    /// §2.3.4.10's simple types for the agile attributes, §2.3.4.5's header table for
    /// [`EncryptParam::KeySize`] — so a request that trips it would produce a file no
    /// conforming reader need accept. The only one of the four that says the *caller*
    /// asked for something the format forbids.
    OutsideSpecRange,

    /// Inside the spec, outside AES as implemented here.
    ///
    /// The **cipher** row: fixed by the algorithm rather than by the format. `ST_KeyBits`
    /// states no maximum precisely because it is generic across ciphers; AES defines
    /// three sizes, and this crate writes AES. A request refused for this reason is not
    /// non-conforming, and a message implying it was would be wrong.
    UnsupportedByCipher,

    /// Legal apart, unreadable together.
    ///
    /// No row in `limits.rs`'s table, because a range check cannot see it: each value is
    /// inside its own bound and the combination is still not writable. Paired with
    /// [`EncryptParam::KeyBitsWithHash`], whose docs give the case.
    UnusableCombination,

    /// Spec-legal, cipher-fine, and this crate declines anyway.
    ///
    /// The **margin** row: this crate's own defence, as the read caps and
    /// `limits::crypto::PAYLOAD_CEILING` are. A file built to the requested value could
    /// be perfectly valid and perfectly implementable, and this crate is choosing not to
    /// write it. **It does not mean the format forbids the value**, and its `Display`
    /// string says so in as many words — that distinction is why this variant exists
    /// separately from [`Self::OutsideSpecRange`] instead of being folded into it.
    ///
    /// # Nothing produces it, and that is the design
    ///
    /// **No code path in this crate constructs this variant**, and the docs that once
    /// said otherwise — a spin count "too weak to author", a 65 536-byte salt "this crate
    /// declines to write" — described refusals `EncryptParams::validate` has never
    /// emitted. They are corrected at [`EncryptParam::SpinCount`] and
    /// [`EncryptParam::SaltSize`]. The only constructions left are the two exit-code
    /// tables' synthetic payloads, in `src/error.rs`'s own `exit_code_canary` and in
    /// `src/bin/msoffice-crypto_tests.rs`, which pin a number rather than report a
    /// refusal.
    ///
    /// The reason is simply that **this crate presently imposes no margin on any
    /// caller-settable parameter**. `limits::SPIN_COUNT_MAX` is the spec's own
    /// 10 000 000, `limits::AGILE_SALT_SIZE` is `ST_SaltSize` whole, and both `keyBits`
    /// refusals belong to the format or to AES. The margins that do exist — the read
    /// caps, `PAYLOAD_CEILING` — bound what a *file* declares or how large a payload may
    /// be, and are reported by [`Error::BadParameters`] and its neighbours, never by
    /// [`Error::EncryptParams`].
    ///
    /// It is kept unproduced **deliberately**, as a design commitment rather than dead
    /// code, and the commitment is this: the first margin imposed on a caller's parameter
    /// arrives wearing this variant. Without it the cheap thing to do is to reach for
    /// [`Self::OutsideSpecRange`], which would tell a consumer the format forbids a value
    /// the format permits — a typed lie, and the precise failure the fourth variant was
    /// added to prevent. An unproduced variant with a stated reason costs a consumer one
    /// match arm that never fires; the alternative costs its users a false statement
    /// about [MS-OFFCRYPTO], rendered as authoritative.
    ExceedsImplementationLimit,
}

#[cfg(feature = "crypto-ops")]
impl core::fmt::Display for EncryptParam {
    /// The field's own spelling in [MS-OFFCRYPTO] — §2.3.4.10's attribute names for the
    /// agile parameters, §2.3.2's field name for the standard header's `KeySize` — never
    /// the Rust identifier.
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        // Exhaustive, with no `_` arm, so that a variant added above is a compile error
        // here rather than a parameter that silently prints as something else.
        f.write_str(match self {
            Self::SpinCount => "spinCount",
            Self::HashAlgorithm => "hashAlgorithm",
            Self::KeyBits => "keyBits",
            Self::SaltSize => "saltSize",
            Self::KeyBitsWithHash => "keyBits together with hashAlgorithm",
            // Qualified by its structure, because `keyBits` is printed above and a bare
            // "KeySize" beside it would read as a spelling variant of the same field
            // rather than as a different field in a different format.
            Self::KeySize => "EncryptionHeader.KeySize",
        })
    }
}

#[cfg(feature = "crypto-ops")]
impl core::fmt::Display for EncryptParamProblem {
    /// A predicate whose subject is the [`EncryptParam`] printed before it, naming
    /// **whose** rule was broken.
    ///
    /// A consumer with no match arm forwards this sentence, so each one has to be honest
    /// standing alone: the spec case names [MS-OFFCRYPTO] as the authority, and the
    /// margin case says this crate declines without implying the format agrees.
    ///
    /// **The spec case names no section number, deliberately.** Which clause states the
    /// range depends on the parameter printed before it — §2.3.4.10's simple types for
    /// the agile attributes, §2.3.4.5's header table for
    /// [`EncryptParam::KeySize`] — so a fixed citation here would be right for one and
    /// wrong for the other. It said "§2.3.4.10" until the standard writer learned its
    /// second and third key sizes, at which point the sentence would have cited the
    /// agile schema at a caller who never touched it. The section numbers live in the
    /// [`EncryptParam`] variants' docs, where they can be per-parameter.
    ///
    /// Each is phrased "was given …" rather than "is …" so that it reads correctly after
    /// a compound subject too — [`EncryptParam::KeyBitsWithHash`] names two attributes,
    /// and "keyBits together with hashAlgorithm is legal on its own" is not a sentence.
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        f.write_str(match self {
            Self::OutsideSpecRange => {
                "was given a value outside the range [MS-OFFCRYPTO] states for it, so \
                 the file would not conform"
            }
            Self::UnsupportedByCipher => {
                "was given a value [MS-OFFCRYPTO] permits but AES, the only cipher this \
                 crate writes, does not define"
            }
            Self::UnusableCombination => {
                "was given values that are legal apart and unusable together; the file \
                 they describe could not be read back"
            }
            Self::ExceedsImplementationLimit => {
                "was given a value [MS-OFFCRYPTO] permits and AES can use, which this \
                 crate declines as a limit of its own — the format does not forbid it"
            }
        })
    }
}

/// The coverage half of the CLI's exit-code proof.
///
/// `src/bin/msoffice-crypto.rs` maps every [`Error`] to an exit code, and its `match`
/// needs a `_` arm: the enum is `#[non_exhaustive]` and the binary is a separate crate.
/// A `_` arm produces no diagnostic when a variant is added, so a nineteenth variant
/// would silently become exit 7 with nothing red anywhere.
///
/// This table duplicates that one **without** a `_` arm, inside the defining crate,
/// where `#[non_exhaustive]` does not apply. Adding a variant to [`Error`] and not
/// deciding its exit code is then `E0004` here, naming the variant. It is deliberately
/// a duplicate: the alternative — a public `Error::exit_code()` — would make an API
/// promise at publish for a binary's benefit.
///
/// The values are checked against the binary's own table by
/// `exit_codes_map_every_error_class` in `src/bin/msoffice-crypto_tests.rs`. This test
/// is the coverage evidence; that one is the value evidence. Both are required.
#[cfg(test)]
mod exit_code_canary {
    use super::Error;

    fn exit_code(e: &Error) -> u8 {
        match e {
            Error::NotACfbFile => 3,
            Error::MissingStream(_) => 6,
            Error::BadParameters(_) => 6,
            Error::Io(_) => 2,
            #[cfg(feature = "legacy-binary")]
            Error::NotEncrypted => 5,
            #[cfg(feature = "crypto-ops")]
            Error::UnsupportedEncryptionVersion(_, _) => 9,
            #[cfg(feature = "crypto-ops")]
            Error::XmlParse(_) => 6,
            #[cfg(feature = "crypto-ops")]
            Error::WrongPassword => 4,
            #[cfg(feature = "crypto-ops")]
            Error::CipherError => 6,
            #[cfg(feature = "crypto-ops")]
            Error::UnsupportedAlgorithm { .. } => 9,
            #[cfg(feature = "crypto-ops")]
            Error::IntegrityCheckFailed => 8,
            #[cfg(feature = "crypto-ops")]
            Error::IntegrityElementMissing => 8,
            #[cfg(feature = "crypto-ops")]
            Error::IntegrityUnavailable(_) => 8,
            #[cfg(feature = "crypto-ops")]
            Error::RandomSource(_) => 7,
            #[cfg(feature = "crypto-ops")]
            Error::EncryptParams { .. } => 1,
            #[cfg(feature = "crypto-ops")]
            Error::AlreadyEncrypted { .. } => 5,
            #[cfg(feature = "crypto-ops")]
            Error::NotAPlainPackage => 5,
            #[cfg(feature = "crypto-ops")]
            Error::UnknownContainer => 3,
        }
    }

    #[test]
    fn every_error_variant_is_named_in_the_exit_code_table() {
        assert_eq!(exit_code(&Error::NotACfbFile), 3);
        assert_eq!(exit_code(&Error::Io(std::io::Error::other("x"))), 2);
        #[cfg(feature = "crypto-ops")]
        assert_eq!(exit_code(&Error::IntegrityElementMissing), 8);
        #[cfg(feature = "legacy-binary")]
        assert_eq!(exit_code(&Error::NotEncrypted), 5);
        // The two encrypt-guard refusals that share a code, and the one that does not:
        // 5 is "I recognise this and will not write into it", 3 is "wrong file".
        #[cfg(feature = "crypto-ops")]
        assert_eq!(exit_code(&Error::NotAPlainPackage), 5);
        #[cfg(feature = "crypto-ops")]
        assert_eq!(exit_code(&Error::UnknownContainer), 3);
        // The one caller error in the table: EX_USAGE (1), which every other variant
        // avoids because every other variant is a verdict on a file. By the time this
        // one is raised there is often no file at all.
        #[cfg(feature = "crypto-ops")]
        assert_eq!(
            exit_code(&Error::EncryptParams {
                param: super::EncryptParam::SpinCount,
                problem: super::EncryptParamProblem::ExceedsImplementationLimit,
                got: 0,
                min: 100_000,
                max: 10_000_000,
            }),
            1
        );
    }
}
