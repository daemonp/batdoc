//! Numeric bounds on the parameters a file declares about itself.
//!
//! Everything here arrives as a bare `u32` out of an attacker-supplied
//! `EncryptionInfo` stream, and each one is used as a loop count or a slice length
//! before anything else in the crate has had a chance to reject the file. The sibling
//! crate `odf-crypto` keeps the same list for the same reason (`odf-crypto/src/limits.rs`
//! — `PBKDF2_MAX_ITER`, `DERIVED_KEY_MIN_LEN`/`MAX_LEN`); the *shapes* match
//! deliberately — one module, every bound labelled by where it comes from. **No number
//! crosses between the two crates.** [MS-OFFCRYPTO] and ODF's manifest schema are
//! different documents with different facets, so where the two files hold the same
//! figure (`PAYLOAD_CEILING`, `1 << 30` in both) that is either crate's own arithmetic
//! landing on the same number, not one borrowed from the other — see that constant's
//! doc comment for the derivation this file is required to give on its own terms.
//!
//! # Where each bound comes from
//!
//! A consumer rendering a refusal needs to know whether the file broke the format or
//! merely asked for more than this crate does, because those are different sentences and
//! only one of them implies the file is at fault. Until 2026-09-20 that was answerable
//! only by reading each doc comment and inferring, which a downstream consumer did, got
//! partly wrong, and had to ask about. So each bound is now labelled, and the three
//! labels mean different things:
//!
//! | label | meaning | if a file trips it |
//! | --- | --- | --- |
//! | **spec** | the number [MS-OFFCRYPTO] itself states | the file is non-conforming |
//! | **cipher** | fixed by the algorithm, not by the format | the file may be valid; this crate will not do it |
//! | **margin** | this crate's own defence against unbounded work | the file may be valid *and* implementable; this crate declines |
//!
//! **spec** — [`crypto::SPIN_COUNT_MAX`] (§2.3.4.10 `ST_SpinCount`, `0..=10000000`),
//! [`crypto::AGILE_SALT_SIZE`] (§2.3.4.10 `ST_SaltSize`, `1..=65536`),
//! [`crypto::STANDARD_KEY_BITS_AES`] (§2.3.2 and §2.3.4.5 both **enumerate** the AES
//! `KeySize` values as 0x80 / 0xC0 / 0x100),
//! [`PPT_PERSIST_OBJECTS_MAX`] (§2.3.5, `persistId` is 20 bits),
//! [`legacy::RC4_KEY_BITS`] and [`legacy::RC4_KEY_BITS_DEFAULT`] (§2.3.5.1),
//! [`legacy::XOR_PASSWORD_MAX_LEN`] (§2.3.7.2, structural — the `InitialCode` table has
//! exactly 15 entries).
//!
//! **cipher** — [`crypto::AGILE_KEY_BITS_ALLOWED`] (what AES defines; the schema's
//! `ST_KeyBits` sets `minInclusive="8"` and *no maximum*, being generic across cipher
//! algorithms). It holds the same three numbers as `STANDARD_KEY_BITS_AES` one row up
//! and is labelled differently on purpose: there the format states the set, here AES
//! does.
//!
//! **margin** — every read cap ([`ENCRYPTION_INFO_READ_CAP`], [`BINARY_HEADER_READ_CAP`],
//! [`BIFF_SCAN_CAP`], [`CURRENT_USER_READ_CAP`], [`PPT_PERSIST_DIRECTORY_READ_CAP`],
//! [`RC4_ENCRYPTION_HEADER_SIZE_MAX`], [`ENCRYPTION_HEADER_STRUCTURE_MAX`]),
//! [`crypto::PAYLOAD_CEILING`] and its two aliases. Each names its margin over the
//! largest real value in its own comment.
//!
//! Two bounds the schema states are **not** enforced as ranges here, because the check
//! beside the code is strictly tighter: `ST_BlockSize` (`2..=4096`) is required to equal
//! the 16-byte AES block by `agile::check_block_size`, and `ST_HashSize` (`1..=65536`) is
//! required to equal the *named* algorithm's digest length by `agile::check_hash_size`.
//! A schema-range check would accept values both of those reject, so adding one would
//! weaken the parser, not strengthen it.
//!
//! What is deliberately *not* here: lengths fixed by the on-disk format rather than by
//! a field in it — the 72-byte `EncryptionVerifier`, the 16-byte AES block. Those are
//! layout facts and live beside the code that reads the layout, so a reader checking
//! "is this the number the spec gives?" finds them next to the offsets they describe.
//!
//! The bounds `classify` reads are compiled in every configuration — the read caps and
//! structure bounds above the gate, which is what keeps detection safe on an untrusted
//! upload without a cipher in the graph. The rest exist solely for a `crypto-ops` path, [`PAYLOAD_CEILING`]
//! included since `classify` never opens the payload, and live in [`crypto`] behind one
//! gate, so
//! `dead_code` stays live everywhere rather than being silenced by a module-wide `allow`
//! that would hide a genuinely unused bound as readily as an expected one. That split is
//! `odf-crypto/src/limits.rs`'s shape, deliberately — the gating structure, not any figure
//! inside it.

/// Ceiling on the `\EncryptionInfo` stream this crate will read into memory.
///
/// The stream is a version header plus either a ~1.5 KB XML document (agile — the fixture
/// here is 1 441 bytes) or a ~220-byte binary header (standard). 1 MiB is roughly 700x the
/// larger of those, so no writer this crate expects to meet comes near it — and it is read
/// *before* anything has authenticated the file, which is the whole point: `classify` reads
/// it on an untrusted upload by definition.
///
/// The `\EncryptedPackage` stream's counterpart is [`PAYLOAD_CEILING`], which is a
/// `crypto-ops` bound rather than one of these: `classify` never opens the payload, so the
/// "detection is free" property does not depend on it.
pub(crate) const ENCRYPTION_INFO_READ_CAP: usize = 1 << 20;

/// Bytes of a legacy binary format's header stream that [`crate::binary_office`] reads.
///
/// Word's FIB puts its flags at offset 0x0A and the table stream puts the whole
/// `EncryptionHeader` in its first ~200 bytes, so 512 covers both with margin. Read
/// before anything has authenticated the file, like everything else in `classify`.
pub(crate) const BINARY_HEADER_READ_CAP: usize = 512;

/// Ceiling on the BIFF stream `classify` will walk looking for a `FILEPASS` record.
///
/// `FILEPASS` is mandated to appear in the globals substream near the top of the workbook
/// — the fixture here has it at offset 0x14 — so a megabyte is roughly five orders of
/// magnitude of margin over where it legitimately occurs.
///
/// One use site, `binary_office`'s workbook read. It used to bound PowerPoint's
/// attacker-chosen `offsetToCurrentEdit` as well, and no longer does: that walk reads at
/// the offset rather than up to it, so there is no buffer to cap — the bound that
/// replaced it is the container's own extent.
///
/// **A bound on hostile input, not on how large a legitimate workbook may be.** The walk
/// never needs to reach the end of the stream to say "not encrypted": [MS-XLS] permits
/// only six records in the clear ahead of `FILEPASS`, so the first record outside that set
/// proves there is none, and a real plain workbook is decided at its third record. A
/// stream that exhausts this cap before any record decides it is reported as unknown —
/// never as unencrypted — so the figure has no margin to justify against the size of a
/// globals substream, which can run to megabytes of shared strings.
pub(crate) const BIFF_SCAN_CAP: usize = 1 << 20;

/// Bytes of PowerPoint's `Current User` stream that are read.
///
/// `CurrentUserAtom` is 20 bytes plus an 8-byte record header, and the fixture's whole
/// stream is 32. 256 is generous and still refuses to materialise a hostile stream that
/// merely claims the name.
pub(crate) const CURRENT_USER_READ_CAP: usize = 256;

/// Ceiling on `EncryptionHeaderSize` — the same `EncryptionHeader` of [MS-OFFCRYPTO]
/// §2.3.2, used by RC4 CryptoAPI (§2.3.5.1) and by ECMA-376 standard encryption (§2.3.4.5).
///
/// The header is 32 bytes of fixed fields followed by `CSPName`, a null-terminated
/// UTF-16LE cryptographic-service-provider name. The longest name Windows ships is
/// "Microsoft Enhanced RSA and AES Cryptographic Provider" — 53 characters, 108 bytes
/// with its terminator — and every legacy fixture here carries the 47-character Enhanced
/// provider (a 126-byte header). 1 024 is roughly seven times the longest real header.
/// The field is a `u32` the file declares, and it is the distance from the header to the
/// `EncryptionVerifier`; LibreOffice bounds it only from below
/// (`sw/source/filter/ww8/ww8par.cxx:5617-5621`, behaviour only), which is enough for a
/// seek and not for a slice. Compiled in every configuration because
/// [`ENCRYPTION_HEADER_STRUCTURE_MAX`] is derived from it, and that derived bound is what
/// `classify` applies to a `.ppt`'s `CryptSession10Container`. The standard decrypt path
/// uses the same number for the same field.
pub(crate) const RC4_ENCRYPTION_HEADER_SIZE_MAX: usize = 1024;

/// Ceiling on a whole RC4 CryptoAPI encryption header *structure*: the 12-byte prefix,
/// a header at [`RC4_ENCRYPTION_HEADER_SIZE_MAX`], and the 60-byte RC4-shaped verifier.
///
/// Two file fields are checked against this before they slice: `FibBase.lKey` in a `.doc`
/// ([MS-DOC] §2.5.2, `word97.rs`) and the `CryptSession10Container`'s `recLen` in a `.ppt`
/// ([MS-PPT] §2.3.7, `binary_office.rs`). The fixtures' structures are 198 bytes.
///
/// The `.xls` leg is **not** one of them, and the reason is not the one this comment used
/// to give. A `FILEPASS` record length is a `u16`, which is 65 535 and therefore *not*
/// inherently under 1 096; what bounds it is that the record body is sliced out of a
/// buffer already capped at [`BIFF_SCAN_CAP`], so an over-long declaration runs out of
/// bytes rather than out of memory.
pub(crate) const ENCRYPTION_HEADER_STRUCTURE_MAX: usize = 12 + RC4_ENCRYPTION_HEADER_SIZE_MAX + 60;

/// The most persist objects a PowerPoint persist directory can name.
///
/// Not a margin: `PersistDirectoryEntry.persistId` is a 20-bit field and "MUST be less
/// than or equal to 0xFFFFE" ([MS-PPT] §2.3.5), so a directory naming more identifiers
/// than that names one twice. The count bounds the work of walking the directory and,
/// on the decrypt path, the number of persist objects rekeyed.
pub(crate) const PPT_PERSIST_OBJECTS_MAX: usize = 1 << 20;

/// Ceiling on a `PersistDirectoryAtom`'s `recLen` — the bytes read to walk it.
///
/// Derived from [`PPT_PERSIST_OBJECTS_MAX`]: each identifier costs at most 8 bytes, a
/// 4-byte offset plus, in the least compact packing of one identifier per entry, a
/// 4-byte entry word. The fixture's atom is 64 bytes for 15 objects. `classify` reads the
/// atom on a hostile upload, so the bound is compiled into every configuration.
pub(crate) const PPT_PERSIST_DIRECTORY_READ_CAP: usize = PPT_PERSIST_OBJECTS_MAX * 8;

#[cfg(feature = "crypto-ops")]
pub(crate) use crypto::*;

#[cfg(feature = "legacy-binary")]
pub(crate) use legacy::*;

/// Bounds only the `legacy-binary` decrypt paths reach. Gated as a module for the reason
/// [`crypto`] is.
#[cfg(feature = "legacy-binary")]
mod legacy {
    /// The read cap on each stream a binary document's decrypt holds in memory —
    /// `WordDocument`, the table stream and `Data`; `Workbook`; `PowerPoint Document` —
    /// aliased to [`super::PAYLOAD_CEILING`] rather than given its own number.
    ///
    /// The same memory bound as `EncryptedPackage`'s, for the same reason: a decrypt holds
    /// a stream and its plaintext at once, and `cfb` already refuses a forged length, so
    /// what remains to bound is the allocation. A `.doc` cannot legitimately be this
    /// large — its FIB offsets are 32-bit and Word's own limit is 32 MB — and the
    /// number is deliberately not tightened per format: one ceiling, one place, one
    /// `const` assertion tying it to the shared figure (GH #10's patterns 3 and 5).
    pub(crate) const LEGACY_STREAM_READ_CAP: usize = super::PAYLOAD_CEILING;

    const _: () = assert!(LEGACY_STREAM_READ_CAP == super::PAYLOAD_CEILING);

    /// `EncryptionHeader.KeySize` for RC4 CryptoAPI, in bits — [MS-OFFCRYPTO] §2.3.5.1:
    /// "greater than or equal to 0x00000028 bits and less than or equal to 0x00000080
    /// bits, in increments of 8 bits". The spec's own range, not a margin. It sizes the
    /// key every block is decrypted under, and `rc4::Keystream::new` dispatches a cipher
    /// on the byte length it produces, so it is checked at parse and again where the key
    /// is derived.
    pub(crate) const RC4_KEY_BITS: core::ops::RangeInclusive<u32> = 40..=128;

    /// What a `KeySize` of zero means — §2.3.5.1: "If set to 0x00000000, it MUST be
    /// interpreted as 0x00000028 bits."
    pub(crate) const RC4_KEY_BITS_DEFAULT: u32 = 40;

    /// The longest password XOR obfuscation can take — [MS-OFFCRYPTO] §2.3.7.2:
    /// "Password MUST NOT be longer than 15 characters". A structural fact, not a margin:
    /// `InitialCode` has 15 entries indexed by length minus one, and `XorMatrix` is walked
    /// seven entries per character from index 0x68, which 15 characters exhaust exactly.
    /// A longer password indexes past both tables, so it is refused as a password Excel
    /// could not have set rather than sliced.
    pub(crate) const XOR_PASSWORD_MAX_LEN: usize = 15;
}

/// Bounds no detection-only build can reach. One gate on the module covers all of them:
/// `crypto-ops` is a single feature, so there is no configuration in which some of these
/// are live and others are not.
#[cfg(feature = "crypto-ops")]
mod crypto {
    /// **1 GiB** — the largest `EncryptedPackage` payload this crate will hold in memory.
    ///
    /// This is the bound GH #10 asked for and could not honestly pick, deferred to plan D4
    /// and landed with the segment iterator (GH #6 step 2). Two things had to be true first,
    /// and now are: the segmentation is explicit rather than an inline `chunks(4096)`, so
    /// there is a stated working set to reason about; and the figure has somewhere to be
    /// shared from, rather than being repeated at each `Vec` that grows.
    ///
    /// **1 GiB, argued on this crate's own numbers — not borrowed from the sibling.**
    /// [MS-OFFCRYPTO] §2.3.4.4 declares `StreamSize` (the length prefix on
    /// `\EncryptedPackage`) an 8-byte unsigned integer and states no ceiling on it at all,
    /// so nothing here caps what the format permits; every payload this figure refuses is
    /// refused on this crate's own grounds, not the spec's. Those grounds are the memory
    /// argument above (ciphertext plus plaintext held at once, so 1 GiB peaks near 2 GiB)
    /// together with real-world headroom: Office's own guidance keeps documents far below
    /// this, a `.pptx` in the hundreds of megabytes — the size GH #10's deferral worried
    /// about — has 2-10x of room under it, and the per-file limits these documents travel
    /// under (OneDrive/SharePoint uploads, mail gateways) have historically sat at or under
    /// 2 GB, so this ceiling refuses only what those transports would already have refused.
    ///
    /// **`odf-crypto`'s `PAYLOAD_CEILING` lands on the identical `1 << 30`**
    /// (`odf-crypto/src/limits.rs`), and that is worth naming precisely because it is not
    /// the reason for this one: both crates hold a whole file's ciphertext and plaintext in
    /// memory at once, so the same arithmetic on the same class of machine produces the
    /// same number for two different formats. Presenting the match as the justification
    /// would be exactly the inversion `docs/plans/msoffice-crypto-encrypt-params-2026-09-20.md`
    /// § *Governing principle* rules out — a sibling's figure is at most a tiebreak where
    /// the spec is silent, never the default a different number would need a reason to
    /// beat. Here there is nothing to tie-break: the number above is derived fresh, and the
    /// sibling agreeing is corroboration, not provenance.
    ///
    /// **What it is not.** It is not an anti-amplification bound: `cfb` 0.14 bounds every
    /// read by the directory entry's `stream_len` *and* by the real FAT chain, so a forged
    /// length fails the read instead of allocating, and a file claiming a gigabyte must
    /// actually be a gigabyte. It is a memory bound, argued above.
    ///
    /// **If this ever refuses a real document, the fix is the streaming API, not a bigger
    /// number.** Raising it buys one more document and moves the same wall; `Read + Write +
    /// Seek` over [`crate::segments`] removes the wall, which is what D4 is for.
    pub(crate) const PAYLOAD_CEILING: usize = 1 << 30;

    /// The read cap on `\EncryptedPackage`, aliased to [`PAYLOAD_CEILING`] rather than
    /// given its own number.
    ///
    /// GH #10's pattern 3, from `odf-crypto`: one shared named ceiling per use site, so a
    /// hostile size cannot allocate past it on one path while another still would. The
    /// stream read and the plaintext buffer are the two allocation paths here, and they are
    /// the same figure by construction.
    pub(crate) const ENCRYPTED_PACKAGE_READ_CAP: usize = PAYLOAD_CEILING;

    /// GH #10's pattern 5 — `const`-assert that related constants agree.
    ///
    /// A compile error rather than a test failure, deliberately: "the payload read cap is
    /// the shared ceiling" is a statement about two constants, and there is no input that
    /// makes it more or less true. The runtime tests in `cfb_reader` prove the *mechanism*
    /// refuses and that the payload cap reaches the payload stream; this proves the number
    /// they cannot afford to exercise at size.
    ///
    /// The second assertion is the one that catches a plausible edit: a ceiling that
    /// drifted down to the 1 MiB header figure would refuse every real document, and every
    /// test in the suite would still pass, because no fixture here is over a megabyte.
    const _: () = {
        assert!(ENCRYPTED_PACKAGE_READ_CAP == PAYLOAD_CEILING);
        assert!(PAYLOAD_CEILING > super::ENCRYPTION_INFO_READ_CAP);
        assert!(PAYLOAD_CEILING == 1 << 30);
    };

    /// Inclusive ceiling on agile `p:encryptedKey/@spinCount`.
    ///
    /// `spin_hash` runs this many SHA-512 rounds *before* the password check, before the
    /// integrity check, before anything that could reject the file — so the cost is paid by
    /// any caller that merely attempts to open the document. Unbounded, a 41 KB file buys
    /// roughly 50 minutes of one core (`u32::MAX` rounds, measured in release), and there is
    /// no in-process defence against it: it is a hang, not a panic, so a caller's
    /// `catch_unwind` never fires and a Rust worker thread cannot be interrupted.
    ///
    /// **Spec-derived, exactly.** [MS-OFFCRYPTO] §2.3.4.10 declares
    /// `ST_SpinCount` as `<xs:restriction base="xs:unsignedInt">` with
    /// `minInclusive="0"` and `maxInclusive="10000000"`, and the prose beside it says
    /// "It MUST NOT be greater than 10,000,000." This constant is that number and not a
    /// margin this crate picked.
    ///
    /// **It was `1 << 21` until 2026-09-20, and that was a defect on the read path.**
    /// 2,097,152 is 4.8x under the spec ceiling, so a `.docx` declaring `spinCount`
    /// anywhere between them was conforming, opened in Word, and was refused here. The
    /// margin was defended on the grounds that the spec ceiling "permits a denial of
    /// service" — measured against the `u32::MAX` anchor above, 10,000,000 rounds is
    /// about **7 seconds** of one core, which is a real cost to a user and not a denial
    /// of service. Three figures in the superseded comment were wrong in the direction
    /// that made the margin look necessary; the arithmetic is recorded in
    /// `CHANGELOG.md` rather than repeated here.
    ///
    /// **Where the denial-of-service policy actually belongs: the caller, who already has
    /// what it needs.** [`crate::classify()`] reports `spin_count` *unbounded and
    /// unvalidated* (see its own doc) before a single round runs, so a caller that wants
    /// Office's 100 000, or `1 << 21`, or any other threshold enforces it in one
    /// comparison on data this crate hands it for free. A ceiling this crate imposes
    /// instead is a policy the caller cannot loosen — and the file it refuses belongs to
    /// the person being refused. CLAUDE.md § *Every input is hostile* is about not
    /// trusting the file; it is not licence to lock an owner out of their own document.
    ///
    /// The write path uses the same constant, which keeps `encryption_info::write`'s
    /// "could not be read back" claim true — and that now matters, because the writer no
    /// longer emits one hardcoded value. `EncryptParams::spin_count` is a caller's
    /// choice across this whole range, defaulting to
    /// `encryption_info::OFFICE_SPIN_COUNT` (100 000, the measured Office figure), so
    /// this ceiling is the only thing standing between a caller and a file this crate
    /// could not read back. Standard encryption needs no entry here: its spin count is
    /// the hardcoded `standard::SPIN_COUNT`, fixed by §2.3.4.7 and not a file field.
    ///
    /// **No floor, and that is also the spec's answer** — `minInclusive="0"`.
    /// `spinCount="0"` makes a weak file, not a dangerous one: the stretching its writer
    /// chose is the writer's decision, and refusing to decrypt a document its owner
    /// already holds buys this crate nothing.
    pub(crate) const SPIN_COUNT_MAX: u32 = 10_000_000;

    /// Pins the figure itself, for the same reason `PAYLOAD_CEILING` is pinned one screen
    /// away — and it was not pinned until the 2026-09-10 pre-publish audit demonstrated
    /// why. **Every test of this bound is written relative to the constant**
    /// (`SPIN_COUNT_MAX + 1`, `SPIN_COUNT_MAX.to_string()`), so the guards move with it:
    /// raising this to `u32::MAX - 1` leaves all 205 tests green while the hang guard it
    /// exists to be is gone. A ceiling that its own tests cannot see move is not a ceiling.
    ///
    /// Both directions matter now, and for different reasons. A *lower* value refuses
    /// files this crate could have opened and that the format permits — the compatibility
    /// bug this constant used to be. A *higher* one accepts what the spec forbids and
    /// reintroduces the unbounded-work hazard the `u32::MAX` measurement describes. The
    /// equality assertion pins both at once, which is why it names the spec's figure
    /// rather than a range.
    const _: () = {
        assert!(SPIN_COUNT_MAX == 10_000_000); // [MS-OFFCRYPTO] §2.3.4.10 ST_SpinCount
        assert!(SPIN_COUNT_MAX > 100_000); // must not refuse what Office itself writes
    };

    /// The values of agile `keyBits` this crate will act on — on **either** element.
    ///
    /// `<keyData>` and `<p:encryptedKey>` each declare one, about two different keys:
    /// the package key and the key that wraps it. Only the second was bounded here at
    /// first, which left `<keyData>`'s unread entirely — see
    /// `agile::check_key_bits`, which asks the question of both, and
    /// `AgileParams::key_data_key_bits` for what the unread one cost.
    ///
    /// ECMA-376 admits exactly these three for AES, and LibreOffice's reader accepts the
    /// same set (`AgileEngine.cxx:574-612`, behaviour only). `keyBits / 8` is used verbatim
    /// as a truncation length inside `agile::derive_block_key`, reachable before any
    /// password is checked, so an unbounded value slices off the end of the digest.
    ///
    /// **This bound is necessary and no longer sufficient for that slice.** It was written
    /// when the digest was always SHA-512's 64 bytes, which made "≤ 512 bits" the whole
    /// story. Since the password path honours `p:encryptedKey/@hashAlgorithm` (issue #11)
    /// the digest may be 20 bytes, and 192 and 256 are both longer than that — so the
    /// slice's safety comes from `agile`'s explicit `keyBits / 8 ≤ digest_len` check
    /// against the *named* hash, and this list only says which key sizes AES defines. Do
    /// not read the two as one guard again.
    ///
    /// All three open. Until GH #13 the cipher was AES-256 only and 128 and 192 were
    /// accepted here and refused one frame later; now `agile::aes_cbc_decrypt` dispatches
    /// on the key length, so this list is the format's and the cipher's at once.
    pub(crate) const AGILE_KEY_BITS_ALLOWED: [u32; 3] = [128, 192, 256];

    /// The range agile `saltSize` may declare, on either `<keyData>` or
    /// `<p:encryptedKey>`.
    ///
    /// [MS-OFFCRYPTO] §2.3.4.10 states it "MUST be at least 1 and no greater than 65,536".
    /// This is the spec's own number rather than a margin this crate picked — LibreOffice
    /// quotes the same sentence at `AgileEngine.cxx:557` and herumi enforces the identical
    /// range at `include/crypto_util.hpp:138-140`, which makes the figure available from a
    /// BSD-3 source and not only from the MPL one.
    ///
    /// It bounds real work: `roundUp(saltSize, blockSize)` is how many bytes of decrypted
    /// `encryptedVerifierHashInput` `agile::verify_password` hashes, and the salt itself is
    /// used verbatim as a CBC IV. The tighter constraint in practice is the cross-check
    /// beside it — `saltSize` must equal the decoded `saltValue`'s length — which pins the
    /// value to whatever base64 the file actually carries.
    pub(crate) const AGILE_SALT_SIZE: std::ops::RangeInclusive<u32> = 1..=65536;

    /// The `EncryptionHeader.KeySize` values this crate's standard-encryption path
    /// accepts — **a set of three, not a single value**.
    ///
    /// [MS-OFFCRYPTO] §2.3.2's `KeySize` table enumerates them for AES outright —
    /// "0x00000080, 0x000000C0, 0x00000100 … 128-bit, 192-bit, or 256-bit" — and
    /// §2.3.4.5 repeats the same three against the three AlgIDs the same header may
    /// declare: "This value MUST be 0x00000080 (AES-128), 0x000000C0 (AES-192), or
    /// 0x00000100 (AES-256)."
    ///
    /// **That enumeration is why this is a *spec* bound while
    /// [`AGILE_KEY_BITS_ALLOWED`] is a *cipher* one.** Agile's `ST_KeyBits` sets
    /// `minInclusive="8"` and no maximum, so there the set of three comes from AES and
    /// not from the format; here the format states it. Same three numbers, different
    /// provenance, and the distinction is what a consumer needs to say whether a file
    /// that trips it is non-conforming (here: yes) or merely beyond what this crate
    /// implements (agile: also yes, but for the cipher's reason).
    ///
    /// Until 2026-09-20 this was a single `STANDARD_KEY_BITS_AES128 = 128` and
    /// `standard::require_aes_128` refused the other two AlgIDs **by name**, so the
    /// crate could not open a conforming Office 2007 document that used them — the
    /// refusal, not the format, was the constraint, and the refusal was then cited as
    /// evidence the format offered no choice.
    ///
    /// The field is an unconstrained `u32` in the file and it sizes two things:
    /// `KeySize / 8` truncates the 40-byte SHA-1 XOR-ladder buffer (§2.3.4.7 step 1 caps
    /// `cbRequiredKeyLength` at 40, and AES-256's 32 sits inside that), and the result is
    /// handed to AES, which accepts exactly 16, 24 or 32 bytes and panics otherwise.
    /// `standard::parse_encryption_info` additionally requires the value to **agree with
    /// the AlgID**, which is the tighter of the two checks.
    pub(crate) const STANDARD_KEY_BITS_AES: [u32; 3] = [128, 192, 256];

    /// Pins the three values and their order. `standard_encrypt` takes `[0]` as the
    /// AES-128 it writes, so ascending order is load-bearing there and not only here,
    /// and `standard::aes_key_bits` maps the three AlgIDs onto these three positions.
    const _: () = {
        assert!(STANDARD_KEY_BITS_AES[0] == 128); // §2.3.2 0x00000080, AlgID 0x0000660E
        assert!(STANDARD_KEY_BITS_AES[1] == 192); // §2.3.2 0x000000C0, AlgID 0x0000660F
        assert!(STANDARD_KEY_BITS_AES[2] == 256); // §2.3.2 0x00000100, AlgID 0x00006610
    };
}
