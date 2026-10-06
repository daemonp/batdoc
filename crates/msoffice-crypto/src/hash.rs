//! Everything this crate computes with the hash algorithm a file names.
//!
//! An agile file names *two* hashes and they drive disjoint halves of the algorithm:
//! `keyData/@hashAlgorithm` governs the per-segment package IVs, the two `dataIntegrity`
//! IVs and the package HMAC, while `p:encryptedKey/@hashAlgorithm` governs the spin hash,
//! the block-key derivation and the password verifier.
//!
//! [MS-OFFCRYPTO] §2.3.4.10 tells a *writer* to make them equal — the
//! `PasswordKeyEncryptor`'s "hashing algorithm specified MUST be the same as the hashing
//! algorithm specified for the Encryption.keyData element" — and Office obliges, which is
//! what hides a crossed-element bug in any round trip against your own writer (herumi
//! `include/encode.hpp:146-147` sets both from one pair). A *reader* is handed bytes, not
//! a promise: a file whose two elements disagree is non-conforming and still has to
//! decrypt under the hash each half was actually written with, because assuming they
//! match means silently using the wrong one on whichever half was guessed. The two
//! callers therefore pass their own algorithm in; nothing here decides which is in force,
//! and nothing here refuses the disagreement.
//!
//! Sources for the derivations: herumi/msoffice (BSD-3)
//! `include/crypto_util.hpp:432-452`, `include/decode.hpp:88-101`; msoffcrypto-tool (MIT)
//! `msoffcrypto/method/ecma376_agile.py:169-201`. Behaviour cross-checked against
//! LibreOffice `oox/source/crypto/AgileEngine.cxx:194-232, 252-262` (MPL-2.0, read-only —
//! no expression from it is reproduced here).
//!
//! **Why this is its own module.** The four operations below used to be private to
//! `integrity`, whose one-sentence purpose is verifying the `dataIntegrity` HMAC. Once
//! the password path needed the same dispatch (issue #11), leaving them there would have
//! given that module a second job and forced `agile` to import its hashing from the
//! integrity checker. The enum itself stays in `classify`, which must *report* the hash
//! in a build with no cipher crate at all; this module is the `crypto-ops` half.

use crate::error::Error;
use hmac::{Hmac, Mac};
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha384, Sha512};

/// The hash named by `keyData/@hashAlgorithm` or `p:encryptedKey/@hashAlgorithm`.
///
/// Declared in `classify` — the detection build has to *report* it, and a second copy
/// here would be a second spelling table to keep in step. Re-exported rather than
/// re-declared so every module keeps one import path.
pub(crate) use crate::classify::HashAlgorithm;

impl HashAlgorithm {
    /// The unhyphenated ECMA-376 spelling — what Office writes, and what this crate
    /// prints in errors. `HashAlgorithm::parse` also accepts the hyphenated form
    /// (`SHA-1`), so a file that used it is reported back in the canonical spelling
    /// rather than in its own.
    ///
    /// `pub` rather than `pub(crate)` for the same reason as [`Self::digest_len`]: the
    /// spelling is the one a caller has to put back into `hashAlgorithm` when it
    /// describes its own parameters, and it is not derivable from `Debug`, which prints
    /// the Rust variant name (`Sha512`) and not the attribute value (`SHA512`).
    pub fn name(self) -> &'static str {
        match self {
            Self::Sha1 => "SHA1",
            Self::Sha256 => "SHA256",
            Self::Sha384 => "SHA384",
            Self::Sha512 => "SHA512",
        }
    }

    /// Digest length in bytes. A `hashSize` attribute is checked against this before it
    /// is trusted as a truncation length, and `keyBits / 8` before it is trusted as one.
    ///
    /// **Why this is `pub` when everything else on this impl is internal.** It is a fact
    /// about SHA, not about this crate — FIPS 180-4 fixes all four numbers, and no
    /// decision of ours can move them. What is ours is the *coupling* it feeds: a caller
    /// choosing encryption parameters meets a refusal when `keyBits / 8` exceeds the
    /// digest of the hash it named (`can_carry_key_bits`, internal), and without this
    /// number it cannot reason about that rule before it trips over it — it would have
    /// to hardcode the same four constants to predict our own error. `hashSize` is not
    /// a companion to this: [MS-OFFCRYPTO] §2.3.4.10 requires it to *equal* the named
    /// hash's digest length, so it is derived from this value rather than chosen beside
    /// it, which is why no setter for it exists.
    pub fn digest_len(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Sha256 => 32,
            Self::Sha384 => 48,
            Self::Sha512 => 64,
        }
    }

    /// Can a block key of `key_bits` bits be cut from this hash's digest without padding
    /// it?
    ///
    /// The one predicate behind three refusals, so that they cannot drift: the
    /// `keyBits`/`hashAlgorithm` pair check in `agile::parse_encryption_info`, the
    /// re-check inside `agile::derive_block_key` (which is reachable from three call
    /// sites and declines to trust the parser), and the message
    /// `agile::unusable_key_bits` builds for both. The expression was written out twice
    /// before this function existed, and two copies of a bound are one bound and one
    /// latent divergence.
    ///
    /// The bound is *not* in [MS-OFFCRYPTO]. §2.3.4.11 ends by telling an implementation
    /// to 0x36-pad a digest shorter than `PasswordKeyEncryptor.keyBits` and to truncate
    /// a longer one, and `ST_KeyBits` has a minimum of 8 and a multiple-of-8 constraint
    /// but **no maximum at all**, so the combination is legal on the wire. This crate
    /// refuses the padding half deliberately — see `agile::derive_block_key` for the
    /// argument and for what the four references each do instead. This predicate is
    /// therefore a house rule with a spec-shaped name, and says so rather than implying
    /// the spec forbids what it rejects.
    ///
    /// The truncating divide is the right one and not a rounding hazard. Every value
    /// that reaches here has passed `agile::check_key_bits`, whose allowlist holds only
    /// multiples of 8, so `key_bits / 8` is exact today; and were a non-multiple ever to
    /// arrive, `key_bits / 8` is *the same expression* the truncation downstream uses to
    /// cut the digest, so the predicate stays a true statement about that cut rather
    /// than an independent calculation that could disagree with it.
    pub(crate) fn can_carry_key_bits(self, key_bits: u32) -> bool {
        (key_bits / 8) as usize <= self.digest_len()
    }

    /// `H(data)`.
    ///
    /// One operand, used by the password verifier over the *whole* decrypted
    /// `encryptedVerifierHashInput` (herumi `include/decode.hpp:97-98`).
    pub(crate) fn digest(self, data: &[u8]) -> Vec<u8> {
        self.digest_two(data, &[])
    }

    /// `H(first || second)` — order matters, and reversing the two is a silently-wrong
    /// key or IV rather than an error (herumi `include/crypto_util.hpp:438` for the IV, `:449` for the key of the same shape).
    ///
    /// Every hash this format performs has this shape, which is why it takes two slices
    /// rather than an 8-byte block key: the two `dataIntegrity` IVs and the three block
    /// keys append an 8-byte constant, the package IVs a 4-byte little-endian segment
    /// index, the spin hash's first round the UTF-16LE password and its later rounds the
    /// previous digest (the spin hash through [`Self::digest_two_into`], since rc.5).
    pub(crate) fn digest_two(self, first: &[u8], second: &[u8]) -> Vec<u8> {
        macro_rules! run {
            ($d:ty) => {{
                let mut h = <$d>::new();
                h.update(first);
                h.update(second);
                h.finalize().to_vec()
            }};
        }
        match self {
            Self::Sha1 => run!(Sha1),
            Self::Sha256 => run!(Sha256),
            Self::Sha384 => run!(Sha384),
            Self::Sha512 => run!(Sha512),
        }
    }

    /// `H(first || second)`, truncated to `out.len()` and written into `out` — the same
    /// bytes [`Self::digest_two`] returns, without the `Vec`.
    ///
    /// **Why the truncating form is the one that takes a slot.** Its caller,
    /// `agile::derive_block_key`, wants `key_len` bytes of a digest that is longer, and
    /// the shortest way to say that in Rust builds the whole digest on the heap and then
    /// copies a prefix out of it — two allocations, one of them holding the complete
    /// digest of `H_final`, both freed without being wiped. Writing into a
    /// `Dynamic::new_with` slot makes the wrapper the only place the bytes ever land.
    /// See `docs/design/heap-residue.md` for why the abandoned block is beyond the
    /// wrapper's reach.
    ///
    /// **`digest_two` is not reimplemented in terms of this, deliberately.** Its
    /// remaining callers want an owned digest of the natural length, and routing that
    /// through a slot they would then have to allocate anyway buys nothing. The two
    /// share the format's `H(a || b)` and nothing else, and a test below pins them to
    /// the same bytes.
    ///
    /// `spin_hash` was one of those callers until rc.5 and is now this function's
    /// second: its loop, which runs `spinCount` — up to 10 000 000 — times per decrypt,
    /// alternates between two wrapped slots of exactly `digest_len` bytes, so there the
    /// slot is not allocated per call at all, and every round lands somewhere wiped.
    ///
    /// # Contract
    ///
    /// `out.len()` must not exceed [`Self::digest_len`]. It is the caller's check, not
    /// this function's, because the caller is the one with an [`Error`] to return and a
    /// message worth reading: `derive_block_key` asks [`Self::can_carry_key_bits`] two
    /// lines earlier and refuses with `agile::unusable_key_bits`. A longer `out` is
    /// filled as far as the digest goes and its tail left as the caller supplied it —
    /// which from `Dynamic::new_with` is zero — rather than padded with `0x36` the way
    /// [`fit_iv`] pads an IV. Neither is defensible as a *key*; the refusal upstream is,
    /// and that is why there is one.
    pub(crate) fn digest_two_into(self, first: &[u8], second: &[u8], out: &mut [u8]) {
        // The contract above, made loud where it is stated. Without this, violating it
        // is *silent*: `out` keeps whatever tail the caller allocated, which from
        // `Dynamic::new_with` is zeros, and a short digest becomes a zero-padded key
        // rather than an error — precisely what `derive_block_key` refuses to do one
        // frame up, arrived at by omission instead of by decision.
        //
        // `debug_assert` rather than a `Result`: the caller already owns the refusal and
        // its message, so returning a second error here would duplicate the decision
        // without improving it. This catches a wrong call in every test run and compiles
        // out of release, so it adds no panic to a shipped parse path.
        debug_assert!(
            out.len() <= self.digest_len(),
            "digest_two_into: out is {} bytes, {} produces {} — the caller must refuse \
             this before asking (see can_carry_key_bits)",
            out.len(),
            self.name(),
            self.digest_len(),
        );
        macro_rules! run {
            ($d:ty) => {{
                let mut h = <$d>::new();
                h.update(first);
                h.update(second);
                let digest = h.finalize();
                let n = out.len().min(digest.len());
                out[..n].copy_from_slice(&digest[..n]);
            }};
        }
        match self {
            Self::Sha1 => run!(Sha1),
            Self::Sha256 => run!(Sha256),
            Self::Sha384 => run!(Sha384),
            Self::Sha512 => run!(Sha512),
        }
    }

    pub(crate) fn hmac(self, key: &[u8], message: &[u8]) -> Vec<u8> {
        macro_rules! run {
            ($d:ty) => {{
                // HMAC accepts a key of any length, so `new_from_slice` cannot fail
                // here; the `expect` documents that rather than hiding a real case.
                let mut mac =
                    <Hmac<$d>>::new_from_slice(key).expect("HMAC accepts keys of any length");
                mac.update(message);
                mac.finalize().into_bytes().to_vec()
            }};
        }
        match self {
            Self::Sha1 => run!(Sha1),
            Self::Sha256 => run!(Sha256),
            Self::Sha384 => run!(Sha384),
            Self::Sha512 => run!(Sha512),
        }
    }
}

/// `H(keyData.saltValue || suffix)` truncated to `keyData/@blockSize`.
///
/// Every IV in the agile format that is seeded by `<keyData>` comes from this one
/// derivation, and both of its consumers live outside this module: the two
/// `dataIntegrity` blob IVs (suffix = an 8-byte block key) and [`crate::segments`]'s
/// per-segment package IVs (suffix = the little-endian segment index), which both the
/// agile decrypt and encrypt paths run through since plan D4. They are the same shape and
/// take their hash from the same file attribute, so they share the function rather than
/// each spelling it out — a package IV that hardcoded SHA-512 while the integrity path
/// honoured the file is precisely how a mixed-algorithm document could be reported
/// `Verified` and then decrypted with the wrong IVs.
///
/// Truncation only. herumi's shared `normalizeKey` would 0x36-pad a digest shorter than
/// `block_size` (`include/crypto_util.hpp:38-41`), but every supported hash produces at
/// least 20 bytes and `block_size` is pinned to 16 by both callers, so that branch is
/// unreachable here — and a file that reached it would be one we should reject rather
/// than pad.
pub(crate) fn derive_iv(
    hash: HashAlgorithm,
    salt: &[u8],
    suffix: &[u8],
    block_size: usize,
) -> Result<Vec<u8>, Error> {
    let digest = hash.digest_two(salt, suffix);
    if digest.len() < block_size {
        return Err(Error::BadParameters(format!(
            "keyData blockSize {} exceeds the {} digest length {}",
            block_size,
            hash.name(),
            digest.len()
        )));
    }
    Ok(fit_iv(&digest, block_size))
}

/// [MS-OFFCRYPTO] §2.3.4.12 last step: an IV shorter than `blockSize` is padded with
/// `0x36`; a longer one is truncated.
///
/// Package and `dataIntegrity` IVs go through [`derive_iv`] first, whose digest is
/// always longer than the AES block, so they only ever truncate. The password-encryptor
/// blobs use the salt itself as the IV ("if a blockKey is not provided", same section),
/// and `saltSize` is 1..=65 536 — a short salt used to fail inside AES as a 16-byte-IV
/// mismatch rather than being padded the way the spec writes.
pub(crate) fn fit_iv(bytes: &[u8], block_size: usize) -> Vec<u8> {
    if bytes.len() >= block_size {
        bytes[..block_size].to_vec()
    } else {
        let mut iv = Vec::with_capacity(block_size);
        iv.extend_from_slice(bytes);
        iv.resize(block_size, 0x36);
        iv
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Each variant must actually run the algorithm it names. A `match` arm that fell
    /// through to the wrong hasher would be invisible to every other test in the crate
    /// as long as writer and reader agreed on it — which, this crate being both, they
    /// would.
    #[test]
    fn each_variant_runs_the_algorithm_it_names() {
        // Known-answer vectors for the empty string, from FIPS 180-4.
        for (hash, want) in [
            (
                HashAlgorithm::Sha1,
                "da39a3ee5e6b4b0d3255bfef95601890afd80709",
            ),
            (
                HashAlgorithm::Sha256,
                "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
            ),
            (
                HashAlgorithm::Sha384,
                "38b060a751ac96384cd9327eb1b1e36a21fdb71114be07434c0cc7bf63f6e1da\
                 274edebfe76f65fbd51ad2f14898b95b",
            ),
            (
                HashAlgorithm::Sha512,
                "cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce\
                 47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e",
            ),
        ] {
            let got: String = hash
                .digest(b"")
                .iter()
                .map(|b| format!("{b:02x}"))
                .collect();
            assert_eq!(got, want.replace(' ', ""), "{hash:?}");
            assert_eq!(hash.digest_len(), got.len() / 2);
        }
    }

    /// `digest` is `digest_two` with an empty tail, and the split must not change the
    /// bytes: the verifier hashes one operand, every other site hashes two.
    #[test]
    fn digest_is_digest_two_with_nothing_appended() {
        for hash in [
            HashAlgorithm::Sha1,
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ] {
            assert_eq!(hash.digest(b"abc"), hash.digest_two(b"a", b"bc"));
            assert_ne!(hash.digest_two(b"a", b"bc"), hash.digest_two(b"bc", b"a"));
        }
    }

    /// `digest_two_into` must be `digest_two` truncated and nothing else.
    ///
    /// Two functions computing the same format step is exactly the drift the crate
    /// avoids elsewhere by sharing one expression, and here it cannot: one returns an
    /// owned `Vec` and the other writes a slot, which is the whole point of having both.
    /// So the agreement is pinned by test instead — at *every* truncation length, not
    /// only at the four `keyBits / 8` values in the wild, because a `copy_from_slice`
    /// with the wrong endpoint is a silently wrong key rather than an error.
    #[test]
    fn digest_two_into_is_digest_two_truncated() {
        for hash in [
            HashAlgorithm::Sha1,
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ] {
            let want = hash.digest_two(b"H_final stand-in", b"\x14\x6e\x0b\xe7\xab\xac\xd0\xd6");
            assert_eq!(want.len(), hash.digest_len());
            for n in 0..=hash.digest_len() {
                let mut got = vec![0xAAu8; n];
                hash.digest_two_into(
                    b"H_final stand-in",
                    b"\x14\x6e\x0b\xe7\xab\xac\xd0\xd6",
                    &mut got,
                );
                assert_eq!(got, want[..n], "{hash:?} truncated to {n}");
            }
        }
    }

    /// Asking for more bytes than the digest has is a **contract violation**, not a
    /// behaviour.
    ///
    /// This replaced a test that pinned what the over-long case *does* — fill as far as
    /// the digest goes, leave the caller's tail, which from `Dynamic::new_with` is zeros.
    /// That was written to stop someone "improving" it into a `0x36` pad, which is a real
    /// hazard and the right thing to want. But pinning the behaviour of a call the
    /// contract forbids documents it as available, and the behaviour it documented is a
    /// silently zero-padded key — exactly what `agile::derive_block_key` refuses to
    /// produce one frame up, reached here by omission rather than by decision.
    ///
    /// So the `debug_assert` in `digest_two_into` is the guard and this is its proof. The
    /// `0x36` hazard is covered better than before: a pad of *any* byte is now
    /// unreachable, rather than merely a different constant from the one nobody wants.
    ///
    /// `cfg(debug_assertions)` because that is when the assert exists — it compiles out
    /// of release so a shipped parse path gains no panic, which means a release test run
    /// would see the old silent behaviour and this `should_panic` would fail.
    #[test]
    #[should_panic(expected = "the caller must refuse this before asking")]
    #[cfg(debug_assertions)]
    fn digest_two_into_refuses_more_bytes_than_the_digest_holds() {
        let mut over = vec![0u8; HashAlgorithm::Sha1.digest_len() + 1];
        HashAlgorithm::Sha1.digest_two_into(b"H_final stand-in", b"block-key", &mut over);
    }

    /// [MS-OFFCRYPTO] §2.3.4.12 step 3, both directions: short pads with `0x36`, long
    /// truncates, exact is left alone.
    ///
    /// The pad byte is `0x36` and not zero. A zero pad is the same length and the same
    /// shape, decrypts to rubbish, and — on a file with no `<dataIntegrity>` element —
    /// produces no error at all, so the constant is the whole content of this test.
    #[test]
    fn fit_iv_pads_short_with_0x36_and_truncates_long() {
        // Short: eight salt bytes then eight pad bytes.
        let mut want = vec![0x11u8; 8];
        want.resize(16, 0x36);
        assert_eq!(fit_iv(&[0x11u8; 8], 16), want);

        // Long: the first blockSize bytes, nothing appended.
        assert_eq!(fit_iv(&[0x22u8; 32], 16), vec![0x22u8; 16]);

        // Exact: identity, which is every file Office writes.
        assert_eq!(fit_iv(&[0x33u8; 16], 16), vec![0x33u8; 16]);

        // A single byte is still a whole IV, and every byte after the first is the pad.
        assert_eq!(
            fit_iv(&[0x44u8], 16),
            [&[0x44u8][..], &[0x36u8; 15][..]].concat()
        );
    }

    /// `derive_iv` truncates and never reaches [`fit_iv`]'s pad branch: every hash this
    /// crate accepts produces at least 20 bytes and `blockSize` is pinned to 16, so a
    /// `blockSize` past the digest is refused above rather than padded. Stated as a test
    /// so that wiring `fit_iv` into the salt-as-IV path cannot quietly turn this one into
    /// a pad as well — the two cases are the same spec sentence and different code.
    #[test]
    fn derive_iv_truncates_and_refuses_rather_than_padding() {
        for hash in [
            HashAlgorithm::Sha1,
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ] {
            let iv = derive_iv(hash, b"salt", b"\x00\x00\x00\x00", 16).expect("16 <= 20");
            assert_eq!(iv, hash.digest_two(b"salt", b"\x00\x00\x00\x00")[..16]);
            // One byte past the digest is an error, not a 0x36 tail.
            let err = derive_iv(hash, b"salt", b"", hash.digest_len() + 1)
                .expect_err("blockSize past the digest must be refused");
            assert!(matches!(err, Error::BadParameters(_)));
        }
    }
}
