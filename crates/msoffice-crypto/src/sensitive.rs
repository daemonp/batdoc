//! Key material wrapped with `secure-gate`.
//!
//! Internal only. The public [`crate::decrypt_ooxml`] API still takes
//! `password: &str` and returns a plain `Vec<u8>` — the caller already owns
//! the password, and the decrypted OOXML package *is* the return value, so
//! wrapping either end would be ceremony. Everything between them is wrapped:
//! the crate's own UTF-16LE re-encoding of the password, the spin hash, every
//! block key derived from it, the session key, and the verifier plaintexts that
//! reveal whether the password was right.
//!
//! The first of those is not a contradiction of the sentence above. The `&str`
//! the caller passes is the caller's, and their problem; the UTF-16LE buffer the
//! KDF hashes is a *copy this crate makes on this crate's heap*, and outlives the
//! call only if we let it. Whose allocation it is, is the whole distinction.
//!
//! This is the property that distinguishes this crate from `office-crypto`,
//! `ms-offcrypto-writer`, `msoffcrypto-tool` and `herumi/msoffice`, all of
//! which hold the same values in bare containers. It is not a claim that they
//! are wrong — a CLI that exits after one file has little to gain — but a
//! long-lived process that decrypts documents alongside other secrets does.
//!
//! Every alias whose length the *file* decides is `Dynamic<Vec<u8>>` rather than
//! `Fixed<[u8; N]>` — the five `Fixed` aliases (`XorObfuscationArray`, and the KDF scratch
//! buffers `SpinState`, `PadBlock`, `Office97Buffer` and `AnsiPassword`) are the exceptions
//! and say so themselves, each length being the spec's rather than the file's.
//!
//! **That split is two rules, not one, and they carry different weight.** In the
//! file-decided direction it is not a choice at all: `Fixed<[u8; N]>` needs `N` at compile
//! time, and the agile `keyBits` attribute (16/24/32) is read from the file's own
//! `EncryptionInfo` XML, so nothing but `Dynamic` can even name the type — the same way
//! `Vec<u8>` rather than `[u8; N]` is forced anywhere else in the crate a length comes from
//! an attacker. In the *other* direction it is house style, not a correctness claim:
//! `XorObfuscationArray`'s 16 bytes could sit in a `Dynamic<Vec<u8>>` just as safely,
//! allocation and all, and nothing would misbehave. `Fixed` is chosen there because it
//! makes "this length is the spec's, permanently, not the file's" a fact the type states
//! rather than one a reader has to take from a comment. Credited to odf-crypto because the
//! convention is shared house style, same shape for the same reason — not because either
//! crate's file-decided direction needed the other's permission to be forced.

//! These were `dynamic_alias!` / `fixed_alias!` invocations until secure-gate 0.9.0-rc.10
//! deleted all four alias macros. They only ever expanded to a `type` alias, so the change
//! below is spelling: not one call site moved, across ~46 `with_secret` closures, six
//! `ct_eq` comparisons and three `from_rng` draws.
//!
//! They remain aliases rather than `fixed_newtype!` / `dynamic_newtype!` newtypes, and that
//! is deliberate. A newtype would make `SessionKey` and `DerivedKey` distinct types the
//! compiler could keep apart — today they are the same nominal type and nothing stops one
//! being passed where the other is meant. The separation here buys greppable names, not type
//! safety, and the rustdoc on the deleted macros said so. Taking the newtypes is a change
//! worth arguing on its own evidence, not one to smuggle in under a dependency bump.
//!
//! **Revisited at twelve.** The secure-gate skill set the ninth alias as the point to
//! settle this; the rc.5 residue fixes added four at once. They do not move the answer,
//! because they are not the kind the argument is about: a `Fixed<[u8; N]>` of one size is
//! a different type from every other size and from every `Dynamic`, so passing a
//! `SpinState` where a `PadBlock` or a `DerivedKey` is meant is already a compile error.
//! The one same-shape pair is `AnsiPassword` and `XorObfuscationArray`, both 16 bytes,
//! both confined to `xor_obfuscation.rs`. The seven interchangeable `Dynamic` roles are
//! where the question still lives, unchanged.

use secure_gate::Dynamic;

// Ungated since the standard KDF's `SpinState` and `PadBlock` below: `crypto-ops` now
// declares `Fixed` aliases of its own, not only `legacy-binary`. This module is compiled
// only with `crypto-ops`, so the detection build still never sees the import.
use secure_gate::Fixed;

/// The password itself, encoded UTF-16LE — the input every KDF in these formats hashes
/// first: agile's `H_0 = H(salt || password_utf16le)` and standard encryption's SHA-1
/// equivalent.
///
/// It is the one value here worth more than a key. A [`SessionKey`] opens one document;
/// this opens every document its owner ever protected, and every account that reuses it.
/// The public API takes `password: &str` unwrapped by design — the caller already owns
/// it (see this module's header) — but the crate's own *re-encoding* of it is a second
/// copy this crate made, on this crate's heap, and that copy is ours to wipe.
///
/// `Dynamic`, not `Fixed`, and for once the length is decided neither by us nor by the
/// file but by the caller's own string: `2 * password.encode_utf16().count()`.
pub(crate) type Utf16Password = Dynamic<Vec<u8>>;

/// Encode `password` as UTF-16LE into a [`Utf16Password`], in exactly one allocation.
///
/// **Why this function exists at all.** The obvious spelling is
/// `password.encode_utf16().flat_map(|c| c.to_le_bytes()).collect::<Vec<u8>>()`, and it
/// was what `agile::spin_hash` and `standard::derive_standard_key` both did — and, until
/// rc.5, the two RC4 key schedules under `legacy-binary` as well. `collect`
/// sizes the `Vec` from the iterator's *lower* `size_hint`, and for this iterator that
/// bound is a fraction of the truth — `Chars` can promise only `len.div_ceil(3)` UTF-16
/// units for `len` UTF-8 bytes, because a three-byte char yields one unit — so the `Vec`
/// grows by doubling and every reallocation frees a block holding a prefix of the
/// password **unwiped**. Measured on rustc 1.97.0 for the fixture password `testpass`:
/// two allocation events and 8 bytes handed back to the allocator, i.e. one abandoned
/// block holding `t\0e\0s\0t\0`. A 28-character passphrase abandons 60 bytes across two.
/// That is the defect class of `docs/design/heap-residue.md`, on the live default
/// decrypt path, and it is the same shape as the two instances recorded there.
///
/// `new_with` closes it the way it closed those: one allocation at the final length, a
/// slot that is a slice so growth is not expressible, and the wrapper owning the bytes
/// from before they are written rather than from after.
///
/// **Why it lives in `sensitive` and not in `hash`.** Its two callers are in different
/// modules and neither owns the other; `hash`'s one job is what this crate computes with
/// the hash algorithm a *file* names, and a text encoding names none. What this is, is a
/// constructor for the alias directly above — so it sits with it, where the doc comment
/// explaining why the buffer is wrapped is the doc comment explaining how it is built.
///
/// No overflow in `* 2`: `count()` is at most `password.len()`, and no Rust allocation
/// exceeds `isize::MAX` bytes, so the product is at most `usize::MAX - 1`.
pub(crate) fn utf16le_password(password: &str) -> Utf16Password {
    let len = password.encode_utf16().count() * 2;
    Utf16Password::new_with(len, |slot| {
        // `len` is `2 * count`, so the zip is exact in both directions: every unit gets a
        // chunk and every chunk gets a unit. `chunks_exact_mut(2)` rather than indexing
        // keeps that a property of the iterator rather than of arithmetic that could be
        // got wrong, and `copy_from_slice` is then a 2-into-2 that cannot panic.
        for (unit, out) in password.encode_utf16().zip(slot.chunks_exact_mut(2)) {
            out.copy_from_slice(&unit.to_le_bytes());
        }
    })
}

/// The password hash every key in a file derives from — agile's `H_final` after
/// `spinCount` rounds of SHA-512, standard encryption's 50 000-round SHA-1 digest,
/// RC4 CryptoAPI's un-iterated `SHA1(salt || password)` (`legacy-binary`), or the
/// five bytes of Office 97/2000 RC4's second MD5 that its block keys are made from.
/// It is the single most valuable intermediate in the crate. Length follows the
/// hash algorithm named in the file — 20, 32, 48 or 64 bytes — plus the five Office
/// 97/2000 RC4 keeps from its second MD5, hence `Vec<u8>`.
pub(crate) type PasswordDigest = Dynamic<Vec<u8>>;

/// A block key: `Hp(H_final || block_key)` — under the hash
/// `p:encryptedKey/@hashAlgorithm` names — truncated to `keyBits / 8` for
/// agile — one per purpose: verifier input, verifier hash, key value — or, for the
/// RC4 families, `H(H_0 || block_number)` cut to the header's `KeySize` (and zero-padded
/// to 128 bits at exactly 40), one per 512- or 1024-byte block of a stream or per
/// PowerPoint persist object. Length comes from the file's `keyBits` or `KeySize`, so
/// `Vec<u8>` rather than a fixed array.
pub(crate) type DerivedKey = Dynamic<Vec<u8>>;

/// The 16-byte XOR obfuscation array of \[MS-OFFCRYPTO\] §2.3.7.2: the password's
/// bytes, padded, XORed with the 16-bit key and rotated. Not a key in any
/// cryptographic sense — it is the password with a fixed transformation applied,
/// which is exactly why it is held wrapped for the few lines it exists. `Fixed`, not
/// `Dynamic`: the length is the spec's, not the file's.
#[cfg(feature = "legacy-binary")]
pub(crate) type XorObfuscationArray = Fixed<[u8; 16]>;

/// Standard encryption's iterated password hash `H_i = SHA1(LE32(i) || H_{i-1})`
/// ([MS-OFFCRYPTO] §2.3.4.7), held across all 50 000 rounds in one reused slot.
///
/// **Every round is key-equivalent, not merely an intermediate.** Hashing runs forward
/// with public inputs, so any `H_i` yields `H_final` — and from it the key — for the
/// cost of the rounds still to go; the last one is a single SHA-1 away. The spin count
/// protects the *password* from inversion, not the key from a state captured late.
/// `Fixed`, and 20 bytes, because SHA-1 is the only hash standard encryption names.
pub(crate) type SpinState = Fixed<[u8; 20]>;

/// One of standard encryption's two 64-byte blocks, `H_final` XOR 0x36 or 0x5C
/// ([MS-OFFCRYPTO] §2.3.4.7). Hashing it is the whole of what remains between it and
/// the key, so it is key-equivalent. The 64 is SHA-1's block size — the spec's, not
/// the file's.
pub(crate) type PadBlock = Fixed<[u8; 64]>;

/// Office 97/2000 RC4's 336-byte buffer, `TruncatedHash || salt` sixteen times over
/// ([MS-OFFCRYPTO] §2.3.6.2): 16 × (5 + 16). `TruncatedHash` is the first five bytes
/// of `MD5(password)`, so this buffer is the password's hash repeated; and `H0`
/// itself is written into its tail before being overwritten, so the full MD5 never
/// sits anywhere this crate cannot wipe. Both lengths are the spec's.
#[cfg(feature = "legacy-binary")]
pub(crate) type Office97Buffer = Fixed<[u8; 336]>;

/// The XOR obfuscation password as the bytes §2.3.7.2 hashes — the password itself,
/// one byte per character. At most 15 (`XOR_PASSWORD_MAX_LEN`), carried with a length
/// beside it; the sixteenth byte keeps it the same size as [`XorObfuscationArray`]
/// and is never read. `Fixed` for the reason that alias gives: the bound is the spec's.
#[cfg(feature = "legacy-binary")]
pub(crate) type AnsiPassword = Fixed<[u8; 16]>;

/// The session encryption key that actually decrypts `EncryptedPackage`, recovered
/// by decrypting `encryptedKeyValue` under a [`DerivedKey`]. This is the key an
/// attacker wants: it is independent of the password and unlocks the document on
/// its own. Length follows `keyBits`.
pub(crate) type SessionKey = Dynamic<Vec<u8>>;

/// A decrypted `encryptedVerifierHashInput` / `encryptedVerifierHashValue`. Not a
/// key, but derived from the password and sufficient to confirm a password guess
/// offline, so it is held wrapped for the few lines it exists.
pub(crate) type VerifierPlaintext = Dynamic<Vec<u8>>;

/// The HMAC key recovered from `dataIntegrity/@encryptedHmacKey`. It is not a block
/// key — the two dataIntegrity block constants derive the *IVs*, and the AES key that
/// unwraps this blob is the [`SessionKey`]. Holding it is equivalent to being able to
/// forge an integrity tag for any package the session key encrypts, so it is wrapped
/// for the few lines between its decryption and the HMAC. Length is
/// `keyData/@hashSize` (20/32/48/64), read from the file's own XML, hence `Vec<u8>`
/// rather than a fixed array.
pub(crate) type IntegrityKey = Dynamic<Vec<u8>>;

/// A package HMAC — both the expected value decrypted from
/// `dataIntegrity/@encryptedHmacValue` and the one this crate computes over the
/// `EncryptedPackage` stream. A MAC tag is not a key and not secret against someone
/// holding the file, but it is wrapped so the comparison happens inside nested
/// `with_secret` closures like every other comparison in the crate, and so neither
/// operand can be printed by accident. Length follows `keyData/@hashSize`, which the
/// file declares, hence `Vec<u8>`.
pub(crate) type IntegrityTag = Dynamic<Vec<u8>>;

#[cfg(test)]
#[path = "sensitive_tests.rs"]
mod tests;
