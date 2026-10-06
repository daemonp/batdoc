#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! RC4 CryptoAPI encryption — [MS-OFFCRYPTO] §2.3.5, the "password to open" Office XP
//! and 2003 wrote into `.doc`, `.xls` and `.ppt`, and the one Office 16 still writes
//! into them today (every legacy fixture in `tests/fixtures/` is this, at 128 bits).
//!
//! Three pieces: the header structure (§2.3.5.1, over the shared `EncryptionHeader` of
//! §2.3.2 and `EncryptionVerifier` of §2.3.3), key generation (§2.3.5.2) and the password
//! check (§2.3.5.6). Where the header lives — the first `lKey` bytes of a Word table
//! stream, a `FILEPASS` record body, a `CryptSession10Container` — is each format's own
//! business and is not known here.
//!
//! Behaviour ported from office-crypto `src/method/rc4.rs` and `src/format/doc97.rs`
//! (MIT) and msoffcrypto-tool `msoffcrypto/method/rc4_cryptoapi.py` and
//! `msoffcrypto/format/common.py` (MIT); see NOTICE. LibreOffice's reader of the same
//! structure (`sw/source/filter/ww8/ww8par.cxx:5604-5656`,
//! `sc/source/filter/excel/xicontent.cxx:1157-1220`, MPL, behaviour only) refuses the
//! same things this one does, `fExternal` included, which neither can serve.

use crate::error::Error;
use crate::limits::{RC4_ENCRYPTION_HEADER_SIZE_MAX, RC4_KEY_BITS, RC4_KEY_BITS_DEFAULT};
use crate::rc4::{self, BlockKeySchedule};
use crate::sensitive::{utf16le_password, DerivedKey, PasswordDigest};
use secure_gate::RevealSecret;
use sha1::{Digest, Sha1};

/// `AlgID` for RC4 — [MS-OFFCRYPTO] §2.3.2; `0` means "determined by Flags", which with
/// `fAES` clear is RC4 too.
/// The 5 bytes of `Hfinal` a 40-bit key takes — [MS-OFFCRYPTO] §2.3.5.2.
const RC4_DEFAULT_KEY_LEN: usize = 5;

/// …and the 16 bytes it becomes once zero-padded, "creating a 128-bit key" (§2.3.5.2).
/// The padding is a key-schedule input, not storage: RC4 keyed with 5 bytes and RC4 keyed
/// with those 5 bytes plus 11 zeros are different ciphers.
const RC4_PADDED_KEY_LEN: usize = 16;

const ALG_ID_RC4: u32 = 0x0000_6801;
/// `AlgIDHash` for SHA-1 — §2.3.2; `0` with `fExternal` clear means SHA-1 as well.
const ALG_ID_HASH_SHA1: u32 = 0x0000_8004;

/// `EncryptionHeaderFlags` — [MS-OFFCRYPTO] §2.3.1.
const FLAG_CRYPTO_API: u32 = 0x0000_0004;
const FLAG_DOC_PROPS: u32 = 0x0000_0008;
const FLAG_EXTERNAL: u32 = 0x0000_0010;
const FLAG_AES: u32 = 0x0000_0020;

/// Layout facts, beside the code that reads them: the structure opens with
/// `EncryptionVersionInfo`(4) `Flags`(4) `EncryptionHeaderSize`(4); the header's fixed
/// part is 32 bytes before `CSPName`; the RC4-shaped `EncryptionVerifier` is
/// `SaltSize`(4) `Salt`(16) `EncryptedVerifier`(16) `VerifierHashSize`(4)
/// `EncryptedVerifierHash`(20) = 60 bytes (§2.3.3, §2.3.5.6 step 3).
const PREFIX_LEN: usize = 12;
const HEADER_FIXED_LEN: usize = 32;
const VERIFIER_LEN: usize = 60;
const SALT_LEN: usize = 16;
const VERIFIER_HASH_LEN: usize = 20;

/// What a parsed RC4 CryptoAPI header structure yields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct CryptoApiHeader {
    /// `EncryptionHeader.KeySize`, with `0` already read as 40 (§2.3.5.1).
    pub(crate) key_bits: u32,
    /// `EncryptionHeader.Flags`.
    pub(crate) flags: u32,
    pub(crate) salt: [u8; SALT_LEN],
    pub(crate) encrypted_verifier: [u8; 16],
    pub(crate) encrypted_verifier_hash: [u8; VERIFIER_HASH_LEN],
}

impl CryptoApiHeader {
    /// `fDocProps` set: the document properties are **not** encrypted. When it is clear
    /// the properties travel in an `EncryptedSummary` stream (§2.3.5.4) that this crate
    /// does not decrypt — neither does msoffcrypto-tool, whose output this one matches.
    #[allow(dead_code)]
    pub(crate) fn doc_props_in_the_clear(&self) -> bool {
        self.flags & FLAG_DOC_PROPS != 0
    }
}

/// Parse an RC4 CryptoAPI encryption header structure — [MS-OFFCRYPTO] §2.3.5.1 —
/// starting at its `EncryptionVersionInfo`.
///
/// Every number the file declares is checked before it is used as a length, and every
/// refusal names the field. The verifier is located from `EncryptionHeaderSize` rather
/// than by scanning `CSPName` for its terminator, which is what the structure's own
/// size field is for and how every reference reader finds it.
pub(crate) fn parse(structure: &[u8]) -> Result<CryptoApiHeader, Error> {
    let le16 = |at: usize| crate::binary_office::le16(structure, at);
    let le32 = |at: usize| crate::binary_office::le32(structure, at);

    let (Some(major), Some(minor), Some(header_size)) = (le16(0), le16(2), le32(8)) else {
        return Err(Error::BadParameters(format!(
            "the RC4 CryptoAPI encryption header structure is {} bytes; its version, \
             flags and size prefix alone take {PREFIX_LEN}",
            structure.len()
        )));
    };
    // §2.3.5.1: vMajor 2, 3 or 4 with vMinor 2. Version 1.1 is the MD5 family,
    // `rc4_office97`; a caller dispatches on the pair before arriving here, so a
    // mismatch is the same fact the version variant already names.
    if !matches!(major, 2..=4) || minor != 2 {
        return Err(Error::UnsupportedEncryptionVersion(major, minor));
    }

    let header_size = usize::try_from(header_size)
        .map_err(|_| Error::BadParameters("EncryptionHeaderSize does not fit usize".to_string()))?;
    if header_size < HEADER_FIXED_LEN {
        return Err(Error::BadParameters(format!(
            "EncryptionHeaderSize is {header_size}; the header's fixed fields alone take \
             {HEADER_FIXED_LEN}"
        )));
    }
    if header_size > RC4_ENCRYPTION_HEADER_SIZE_MAX {
        return Err(Error::BadParameters(format!(
            "EncryptionHeaderSize is {header_size}, over the {RC4_ENCRYPTION_HEADER_SIZE_MAX} \
             this crate allows for 32 fixed bytes and a CSP name"
        )));
    }
    let header_end = PREFIX_LEN + header_size;
    let Some(header) = structure.get(PREFIX_LEN..header_end) else {
        return Err(Error::BadParameters(format!(
            "EncryptionHeaderSize is {header_size} but only {} bytes follow the prefix",
            structure.len().saturating_sub(PREFIX_LEN)
        )));
    };

    let h32 = |at: usize| crate::binary_office::le32(header, at);
    let (Some(flags), Some(alg_id), Some(alg_id_hash), Some(key_size)) =
        (h32(0), h32(8), h32(12), h32(16))
    else {
        // Unreachable with header_size >= 32, but a slice-length assumption is not an
        // argument this crate accepts from itself either.
        return Err(Error::BadParameters(
            "EncryptionHeader fixed fields unreadable".to_string(),
        ));
    };

    // Refuse by name what this path cannot serve, before touching a key size. The
    // precedence is §2.3.2's: with fAES set the AlgID does not get to name RC4, and with
    // fExternal set nothing here is defined at all.
    const WHAT_FLAGS: &str = "EncryptionHeader.Flags";
    if flags & FLAG_EXTERNAL != 0 {
        return Err(Error::UnsupportedAlgorithm {
            what: WHAT_FLAGS,
            name: "fExternal (application-defined encryption)".to_string(),
        });
    }
    if flags & FLAG_AES != 0 {
        return Err(Error::UnsupportedAlgorithm {
            what: WHAT_FLAGS,
            name: "fAES in a binary document".to_string(),
        });
    }
    if flags & FLAG_CRYPTO_API == 0 {
        return Err(Error::BadParameters(
            "EncryptionHeader.Flags has fCryptoAPI clear; [MS-OFFCRYPTO] 2.3.5.1 requires it"
                .to_string(),
        ));
    }
    if alg_id != 0 && alg_id != ALG_ID_RC4 {
        return Err(Error::UnsupportedAlgorithm {
            what: "EncryptionHeader.AlgID",
            name: format!("{alg_id:#010x}"),
        });
    }
    if alg_id_hash != 0 && alg_id_hash != ALG_ID_HASH_SHA1 {
        return Err(Error::UnsupportedAlgorithm {
            what: "EncryptionHeader.AlgIDHash",
            name: format!("{alg_id_hash:#010x}"),
        });
    }

    // §2.3.5.1: 40..=128 bits in 8-bit increments; 0 MUST be read as 40. The value sizes
    // the key every block is decrypted under, and `Keystream::new` dispatches on it.
    let key_bits = if key_size == 0 {
        RC4_KEY_BITS_DEFAULT
    } else {
        key_size
    };
    if !RC4_KEY_BITS.contains(&key_bits) || key_bits % 8 != 0 {
        return Err(Error::BadParameters(format!(
            "EncryptionHeader.KeySize is {key_size}; RC4 CryptoAPI allows {}..={} bits in \
             steps of 8",
            RC4_KEY_BITS.start(),
            RC4_KEY_BITS.end()
        )));
    }

    let Some(verifier) = structure.get(header_end..header_end + VERIFIER_LEN) else {
        return Err(Error::BadParameters(format!(
            "EncryptionVerifier missing or truncated: {} bytes follow the header, \
             {VERIFIER_LEN} are needed",
            structure.len().saturating_sub(header_end)
        )));
    };
    let v32 = |at: usize| crate::binary_office::le32(verifier, at);
    let (Some(salt_size), Some(verifier_hash_size)) = (v32(0), v32(36)) else {
        return Err(Error::BadParameters(
            "EncryptionVerifier fields unreadable".to_string(),
        ));
    };
    if salt_size != SALT_LEN as u32 {
        return Err(Error::BadParameters(format!(
            "EncryptionVerifier.SaltSize is {salt_size}; [MS-OFFCRYPTO] 2.3.5.2 requires 16"
        )));
    }
    if verifier_hash_size != VERIFIER_HASH_LEN as u32 {
        return Err(Error::BadParameters(format!(
            "EncryptionVerifier.VerifierHashSize is {verifier_hash_size}; SHA-1 gives 20"
        )));
    }

    let mut salt = [0u8; SALT_LEN];
    let mut encrypted_verifier = [0u8; 16];
    let mut encrypted_verifier_hash = [0u8; VERIFIER_HASH_LEN];
    for (dst, range) in [
        (&mut salt[..], 4..20),
        (&mut encrypted_verifier[..], 20..36),
        (&mut encrypted_verifier_hash[..], 40..60),
    ] {
        match verifier.get(range) {
            Some(src) if src.len() == dst.len() => dst.copy_from_slice(src),
            _ => {
                return Err(Error::BadParameters(
                    "EncryptionVerifier fields unreadable".to_string(),
                ))
            }
        }
    }

    Ok(CryptoApiHeader {
        key_bits,
        flags,
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

/// The password hash every block key derives from — [MS-OFFCRYPTO] §2.3.5.2.
///
/// `H0 = SHA1(salt + password)`, the password as UTF-16LE, **not iterated**: the spec
/// says so in as many words, and it is what makes this family cheap to attack and this
/// hash worth wrapping. It is the single value from which every key in the document
/// follows.
pub(crate) struct CryptoApiKeySchedule {
    h0: PasswordDigest,
    key_bits: u32,
}

impl CryptoApiKeySchedule {
    /// `key_bits` is the header's `KeySize` after `parse` has bounded it; it is bounded
    /// again here so that a caller skipping `parse` cannot make `block_key` slice past
    /// a SHA-1 digest.
    pub(crate) fn new(password: &str, salt: &[u8; SALT_LEN], key_bits: u32) -> Result<Self, Error> {
        if !RC4_KEY_BITS.contains(&key_bits) || key_bits % 8 != 0 {
            return Err(Error::BadParameters(format!(
                "an RC4 CryptoAPI key of {key_bits} bits is outside {}..={}",
                RC4_KEY_BITS.start(),
                RC4_KEY_BITS.end()
            )));
        }
        // The UTF-16LE re-encoding is wrapped and scoped to this statement, as in
        // `standard::derive_standard_key`: see `sensitive::utf16le_password` for the
        // `collect()` it replaced, which abandoned unwiped prefixes of the password.
        let h0 = utf16le_password(password).with_secret(|pw| {
            let mut hasher = Sha1::new();
            hasher.update(salt);
            hasher.update(pw);
            PasswordDigest::new(hasher.finalize().to_vec())
        });
        Ok(Self { h0, key_bits })
    }

    /// The password check for this header — [MS-OFFCRYPTO] §2.3.5.6.
    pub(crate) fn verify(&self, header: &CryptoApiHeader) -> Result<(), Error> {
        rc4::verify_password(
            self,
            &header.encrypted_verifier,
            &header.encrypted_verifier_hash,
            |v| Sha1::digest(v).to_vec(),
        )
    }
}

impl BlockKeySchedule for CryptoApiKeySchedule {
    /// `Hfinal = SHA1(H0 + block)`, the block as a 32-bit little-endian integer; the key
    /// is the first `keyLength` bits — except at exactly 40 bits, where it is those 5
    /// bytes followed by 11 zero bytes, "creating a 128-bit key" (§2.3.5.2). The
    /// padding is a real key-schedule input, not a storage convention: RC4 keyed with
    /// 5 bytes and RC4 keyed with those 5 bytes plus 11 zeros are different ciphers.
    fn block_key(&self, block: u32) -> DerivedKey {
        self.h0.with_secret(|h0| {
            let mut hasher = Sha1::new();
            hasher.update(h0);
            hasher.update(block.to_le_bytes());
            let hfinal = hasher.finalize();
            if self.key_bits == RC4_KEY_BITS_DEFAULT {
                // The slot is exactly `RC4_PADDED_KEY_LEN` bytes and every one of them is
                // zero -- both guaranteed by `Dynamic::new_with` since secure-gate
                // 0.9.0-rc.12, which is why there is no `reserve_exact` and no trailing
                // `resize` here any more. The eleven zero bytes past the copy are the
                // §2.3.5.2 padding, and the guarantee is what lets us rely on them rather
                // than writing them.
                //
                // Do not fill this slot from a `Vec` grown inside this closure: the
                // reallocation the sized slot exists to prevent would simply move to the
                // source buffer. `hfinal` is a fixed-size digest, so the copy below is the
                // only write.
                DerivedKey::new_with(RC4_PADDED_KEY_LEN, |k| {
                    k[..RC4_DEFAULT_KEY_LEN].copy_from_slice(&hfinal[..RC4_DEFAULT_KEY_LEN]);
                })
            } else {
                let len = usize::try_from(self.key_bits / 8)
                    .unwrap_or(0)
                    .min(hfinal.len());
                DerivedKey::new(hfinal[..len].to_vec())
            }
        })
    }
}

#[cfg(test)]
mod tests {
    // The lint header this file opens with is about the parser, not the tests that
    // drive it: a test that cannot unwrap its own fixture is unreadable, and a panic
    // here is a failing test rather than a reachable one. `classify_tests.rs` says the
    // same thing at file scope for the same reason.
    #![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
    use super::*;

    /// The `word97_password.doc` fixture's own header, as `tools`-side inspection reads
    /// it: salt, key size, the two encrypted blobs. Public by construction — all of it is
    /// in the file in the clear.
    const DOC_SALT: [u8; 16] = [
        0x8a, 0x4a, 0x96, 0x3f, 0x5c, 0xd7, 0xa6, 0xb8, 0x54, 0x15, 0x7f, 0xe6, 0x48, 0x9f, 0x06,
        0x21,
    ];
    const DOC_ENCRYPTED_VERIFIER: [u8; 16] = [
        0x80, 0x8b, 0x8e, 0x03, 0x7e, 0x11, 0x21, 0x14, 0xbf, 0x79, 0x54, 0xc0, 0x66, 0x96, 0xa9,
        0xcd,
    ];
    const DOC_ENCRYPTED_VERIFIER_HASH: [u8; 20] = [
        0xcd, 0x73, 0xee, 0xa7, 0x4f, 0x5b, 0xe5, 0xa5, 0x20, 0x0a, 0x04, 0x4d, 0x5d, 0x8f, 0x96,
        0x27, 0x6c, 0xcb, 0xce, 0x89,
    ];

    fn hex(bytes: &[u8]) -> String {
        bytes.iter().map(|b| format!("{b:02x}")).collect()
    }

    /// Known answers from msoffcrypto-tool's `_makekey` (`method/rc4_cryptoapi.py`), the
    /// independent implementation, for the fixture's salt and the password `testpass`:
    /// three key sizes, three block numbers. The 40-bit rows are the zero-padding rule
    /// of §2.3.5.2; 56 is a length no fixture can exercise, because Office 16 writes
    /// 128 whatever `SetPasswordEncryptionOptions` asks for (measured 2026-09-05).
    #[test]
    fn block_keys_match_msoffcrypto_for_every_key_size() {
        let cases: [(u32, u32, &str); 9] = [
            (40, 0, "1c58e4e72f0000000000000000000000"),
            (40, 1, "22186d9b510000000000000000000000"),
            (40, 7, "09c0aaa4260000000000000000000000"),
            (56, 0, "1c58e4e72f2c12"),
            (56, 1, "22186d9b518ca1"),
            (56, 7, "09c0aaa42644aa"),
            (128, 0, "1c58e4e72f2c12f5f36491b984fdd742"),
            (128, 1, "22186d9b518ca1237ad94a8bd5c857c6"),
            (128, 7, "09c0aaa42644aa20b21c658c3974ff91"),
        ];
        for (bits, block, expected) in cases {
            let schedule = CryptoApiKeySchedule::new("testpass", &DOC_SALT, bits).unwrap();
            let got = schedule.block_key(block).with_secret(|k| hex(k));
            assert_eq!(got, expected, "bits={bits} block={block}");
        }
    }

    /// The fixture's own verifier: the right password passes, a wrong one is
    /// `WrongPassword` and nothing else. Both answers were confirmed with msoffcrypto's
    /// `DocumentRC4CryptoAPI.verifypw` on the same bytes.
    #[test]
    fn the_fixture_header_verifies_testpass_and_refuses_another() {
        let header = CryptoApiHeader {
            key_bits: 128,
            flags: FLAG_CRYPTO_API | FLAG_DOC_PROPS,
            salt: DOC_SALT,
            encrypted_verifier: DOC_ENCRYPTED_VERIFIER,
            encrypted_verifier_hash: DOC_ENCRYPTED_VERIFIER_HASH,
        };
        assert!(CryptoApiKeySchedule::new("testpass", &DOC_SALT, 128)
            .unwrap()
            .verify(&header)
            .is_ok());
        assert!(matches!(
            CryptoApiKeySchedule::new("wrongpass", &DOC_SALT, 128)
                .unwrap()
                .verify(&header),
            Err(Error::WrongPassword)
        ));
        // A key size the header did not declare derives a different key, and the
        // verifier catches that too: the KeySize field is authenticated by the check.
        assert!(matches!(
            CryptoApiKeySchedule::new("testpass", &DOC_SALT, 40)
                .unwrap()
                .verify(&header),
            Err(Error::WrongPassword)
        ));
    }

    /// A header structure with every field under the caller's control, in the layout
    /// the fixtures carry: version 4.2, a 126-byte header naming the Enhanced provider,
    /// then the 60-byte verifier. Nine positional fields, deliberately: each test below
    /// varies one of them against the others held at the fixture's values, and a struct
    /// would hide which one in a sea of `..Default`.
    #[allow(clippy::too_many_arguments)]
    fn structure(
        major: u16,
        minor: u16,
        flags: u32,
        alg_id: u32,
        alg_id_hash: u32,
        key_size: u32,
        header_size: u32,
        salt_size: u32,
        hash_size: u32,
    ) -> Vec<u8> {
        let csp: Vec<u8> = "Microsoft Enhanced Cryptographic Provider v1.0\0"
            .encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect();
        let mut s = Vec::new();
        s.extend_from_slice(&major.to_le_bytes());
        s.extend_from_slice(&minor.to_le_bytes());
        s.extend_from_slice(&flags.to_le_bytes());
        s.extend_from_slice(&header_size.to_le_bytes());
        // EncryptionHeader: Flags SizeExtra AlgID AlgIDHash KeySize ProviderType R1 R2 CSPName
        for v in [flags, 0, alg_id, alg_id_hash, key_size, 1, 0, 0] {
            s.extend_from_slice(&v.to_le_bytes());
        }
        s.extend_from_slice(&csp);
        s.extend_from_slice(&salt_size.to_le_bytes());
        s.extend_from_slice(&DOC_SALT);
        s.extend_from_slice(&DOC_ENCRYPTED_VERIFIER);
        s.extend_from_slice(&hash_size.to_le_bytes());
        s.extend_from_slice(&DOC_ENCRYPTED_VERIFIER_HASH);
        s
    }

    const GOOD_FLAGS: u32 = FLAG_CRYPTO_API | FLAG_DOC_PROPS;
    const CSP_HEADER_SIZE: u32 = 32 + 47 * 2;

    #[test]
    fn a_fixture_shaped_structure_parses_to_its_fields() {
        let s = structure(
            4,
            2,
            GOOD_FLAGS,
            ALG_ID_RC4,
            ALG_ID_HASH_SHA1,
            128,
            CSP_HEADER_SIZE,
            16,
            20,
        );
        let h = parse(&s).unwrap();
        assert_eq!(h.key_bits, 128);
        assert_eq!(h.salt, DOC_SALT);
        assert_eq!(h.encrypted_verifier, DOC_ENCRYPTED_VERIFIER);
        assert_eq!(h.encrypted_verifier_hash, DOC_ENCRYPTED_VERIFIER_HASH);
        assert!(h.doc_props_in_the_clear());
        // The two "determined by Flags" zeros are RC4 and SHA-1 (§2.3.2).
        assert_eq!(
            parse(&structure(
                3,
                2,
                GOOD_FLAGS,
                0,
                0,
                128,
                CSP_HEADER_SIZE,
                16,
                20
            ))
            .unwrap()
            .key_bits,
            128
        );
    }

    /// `KeySize` 0 reads as 40 (§2.3.5.1); every value outside 40..=128 or off the 8-bit
    /// grid is refused before it can size a key.
    #[test]
    fn key_size_is_bounded_and_zero_means_forty() {
        assert_eq!(
            parse(&structure(
                4,
                2,
                GOOD_FLAGS,
                ALG_ID_RC4,
                ALG_ID_HASH_SHA1,
                0,
                CSP_HEADER_SIZE,
                16,
                20
            ))
            .unwrap()
            .key_bits,
            40
        );
        for bits in [40u32, 48, 56, 64, 128] {
            assert_eq!(
                parse(&structure(
                    4,
                    2,
                    GOOD_FLAGS,
                    ALG_ID_RC4,
                    ALG_ID_HASH_SHA1,
                    bits,
                    CSP_HEADER_SIZE,
                    16,
                    20
                ))
                .unwrap()
                .key_bits,
                bits
            );
        }
        for bits in [8u32, 32, 39, 41, 100, 136, 256, u32::MAX] {
            let got = parse(&structure(
                4,
                2,
                GOOD_FLAGS,
                ALG_ID_RC4,
                ALG_ID_HASH_SHA1,
                bits,
                CSP_HEADER_SIZE,
                16,
                20,
            ));
            assert!(
                matches!(&got, Err(Error::BadParameters(m)) if m.contains("KeySize")),
                "KeySize {bits}: {got:?}"
            );
        }
        // The schedule's own guard, for a caller that skips `parse`.
        assert!(matches!(
            CryptoApiKeySchedule::new("pw", &DOC_SALT, 256),
            Err(Error::BadParameters(_))
        ));
    }

    /// `EncryptionHeaderSize` is a length the file declares; it is bounded above, below,
    /// and against the bytes actually present — and the verifier after it must be whole.
    #[test]
    fn declared_lengths_are_checked_against_the_structure() {
        let good = structure(
            4,
            2,
            GOOD_FLAGS,
            ALG_ID_RC4,
            ALG_ID_HASH_SHA1,
            128,
            CSP_HEADER_SIZE,
            16,
            20,
        );
        // Too small for the fixed fields.
        assert!(
            matches!(parse(&structure(4, 2, GOOD_FLAGS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, 31, 16, 20)), Err(Error::BadParameters(m)) if m.contains("EncryptionHeaderSize"))
        );
        // Over the cap.
        assert!(
            matches!(parse(&structure(4, 2, GOOD_FLAGS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, RC4_ENCRYPTION_HEADER_SIZE_MAX as u32 + 1, 16, 20)), Err(Error::BadParameters(m)) if m.contains("over the"))
        );
        // Under the cap but past the bytes present: the header would swallow the verifier
        // and run off the end.
        assert!(
            matches!(parse(&structure(4, 2, GOOD_FLAGS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, RC4_ENCRYPTION_HEADER_SIZE_MAX as u32, 16, 20)), Err(Error::BadParameters(m)) if m.contains("follow the prefix"))
        );
        // Truncations: the prefix, the header, the verifier.
        for cut in [0usize, 11, 12, 40, good.len() - 60, good.len() - 1] {
            assert!(
                matches!(parse(&good[..cut]), Err(Error::BadParameters(_))),
                "cut at {cut}"
            );
        }
        // And the whole thing is accepted, so the refusals are attributable to the cut.
        assert!(parse(&good).is_ok());
    }

    /// The version pair, the flags and the algorithm identifiers are refused **by name**
    /// — never as a wrong password.
    #[test]
    fn unsupported_headers_are_named_not_misdiagnosed() {
        for (major, minor) in [(1u16, 1u16), (4, 4), (5, 2), (2, 3)] {
            assert!(matches!(
                parse(&structure(major, minor, GOOD_FLAGS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, CSP_HEADER_SIZE, 16, 20)),
                Err(Error::UnsupportedEncryptionVersion(m, n)) if (m, n) == (major, minor)
            ));
        }
        assert!(matches!(
            parse(&structure(4, 2, GOOD_FLAGS | FLAG_EXTERNAL, 0, 0, 128, CSP_HEADER_SIZE, 16, 20)),
            Err(Error::UnsupportedAlgorithm { what: "EncryptionHeader.Flags", name }) if name.contains("fExternal")
        ));
        assert!(matches!(
            parse(&structure(4, 2, GOOD_FLAGS | FLAG_AES, 0x660E, ALG_ID_HASH_SHA1, 128, CSP_HEADER_SIZE, 16, 20)),
            Err(Error::UnsupportedAlgorithm { what: "EncryptionHeader.Flags", name }) if name.contains("fAES")
        ));
        assert!(matches!(
            parse(&structure(4, 2, GOOD_FLAGS, 0x660E, ALG_ID_HASH_SHA1, 128, CSP_HEADER_SIZE, 16, 20)),
            Err(Error::UnsupportedAlgorithm { what: "EncryptionHeader.AlgID", name }) if name == "0x0000660e"
        ));
        assert!(matches!(
            parse(&structure(4, 2, GOOD_FLAGS, ALG_ID_RC4, 0x800C, 128, CSP_HEADER_SIZE, 16, 20)),
            Err(Error::UnsupportedAlgorithm { what: "EncryptionHeader.AlgIDHash", name }) if name == "0x0000800c"
        ));
        assert!(matches!(
            parse(&structure(4, 2, FLAG_DOC_PROPS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, CSP_HEADER_SIZE, 16, 20)),
            Err(Error::BadParameters(m)) if m.contains("fCryptoAPI")
        ));
        // The verifier's own two sizes.
        assert!(
            matches!(parse(&structure(4, 2, GOOD_FLAGS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, CSP_HEADER_SIZE, 32, 20)), Err(Error::BadParameters(m)) if m.contains("SaltSize"))
        );
        assert!(
            matches!(parse(&structure(4, 2, GOOD_FLAGS, ALG_ID_RC4, ALG_ID_HASH_SHA1, 128, CSP_HEADER_SIZE, 16, 32)), Err(Error::BadParameters(m)) if m.contains("VerifierHashSize"))
        );
    }
}
