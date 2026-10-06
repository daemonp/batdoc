#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! Office binary document RC4 encryption — [MS-OFFCRYPTO] §2.3.6: Office 97 and 2000's
//! own "password to open", which Office XP and 2003 kept offering as "Office 97/2000
//! Compatible". Version 1.1, MD5 throughout, and a key that is 128 bits wide but derived
//! from 40 bits of hash.
//!
//! **No fixture exercises this family end to end, and the reason is recorded rather than
//! assumed.** Office 16 refuses to write it: Word's
//! `SetPasswordEncryptionOptions("Office97/2000Compatible", ...)` fails with
//! "Insufficient memory" and Excel's with "Value does not fall within the expected
//! range" (COM, 2026-09-05), and LibreOffice — whose Calc export writes exactly this
//! codec (`sc/source/filter/excel/xestream.cxx`, `MSCodec_Std97`) — cannot be driven
//! headless with a password on this machine (handoff §4). What pins the module is
//! msoffcrypto-tool's own doctest vector for the key and the verifier
//! (`msoffcrypto/method/rc4.py`), and the synthetic Word container in
//! `legacy_malformed` built with a test-side writer over this same schedule.
//!
//! Behaviour ported from office-crypto `src/method/rc4.rs` (`makekey_rc4`,
//! `DocumentRC4`, MIT) and msoffcrypto-tool `msoffcrypto/method/rc4.py` (MIT); see
//! NOTICE.

use crate::error::Error;
use crate::rc4::{self, BlockKeySchedule};
use crate::sensitive::{utf16le_password, DerivedKey, Office97Buffer, PasswordDigest};
use md5::digest::Output;
use md5::{Digest, Md5};
use secure_gate::RevealSecret;

/// The header structure — §2.3.6.1: `EncryptionVersionInfo`(4, `1.1`), `Salt`(16),
/// `EncryptedVerifier`(16), `EncryptedVerifierHash`(16). 52 bytes, no size field.
const HEADER_LEN: usize = 52;
const SALT_LEN: usize = 16;

/// `TruncatedHash` — the first 40 bits of `H0` (§2.3.6.2).
const TRUNCATED_LEN: usize = 5;
/// One `TruncatedHash || salt` repetition.
const REPEAT_LEN: usize = TRUNCATED_LEN + SALT_LEN;
/// Sixteen repetitions: the 336 bytes `H1` is the MD5 of.
const BUFFER_LEN: usize = 16 * REPEAT_LEN;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Office97Header {
    pub(crate) salt: [u8; SALT_LEN],
    pub(crate) encrypted_verifier: [u8; 16],
    /// MD5, so 16 bytes — not CryptoAPI's 20.
    pub(crate) encrypted_verifier_hash: [u8; 16],
}

/// Parse an RC4 encryption header — §2.3.6.1 — starting at its `EncryptionVersionInfo`.
pub(crate) fn parse(structure: &[u8]) -> Result<Office97Header, Error> {
    let (Some(major), Some(minor)) = (
        crate::binary_office::le16(structure, 0),
        crate::binary_office::le16(structure, 2),
    ) else {
        return Err(Error::BadParameters(
            "the RC4 encryption header is shorter than its 4-byte version".to_string(),
        ));
    };
    if (major, minor) != (1, 1) {
        return Err(Error::UnsupportedEncryptionVersion(major, minor));
    }
    let Some(body) = structure.get(4..HEADER_LEN) else {
        return Err(Error::BadParameters(format!(
            "the RC4 encryption header is {} bytes; [MS-OFFCRYPTO] 2.3.6.1 lays out \
             {HEADER_LEN}",
            structure.len()
        )));
    };
    let mut header = Office97Header {
        salt: [0; SALT_LEN],
        encrypted_verifier: [0; 16],
        encrypted_verifier_hash: [0; 16],
    };
    for (dst, range) in [
        (&mut header.salt[..], 0..16),
        (&mut header.encrypted_verifier[..], 16..32),
        (&mut header.encrypted_verifier_hash[..], 32..48),
    ] {
        match body.get(range) {
            Some(src) if src.len() == dst.len() => dst.copy_from_slice(src),
            _ => {
                return Err(Error::BadParameters(
                    "RC4 encryption header fields unreadable".to_string(),
                ))
            }
        }
    }
    Ok(header)
}

/// The password hash every block key derives from — §2.3.6.2.
///
/// `H0 = MD5(password)` (UTF-16LE); `TruncatedHash` its first 5 bytes; a 336-byte
/// buffer of `TruncatedHash + salt` repeated 16 times; `H1 = MD5(buffer)`. What is kept
/// is `H1`'s first 5 bytes — the whole of what the block keys are made from, which is
/// why the family is called 40-bit whatever the key's width.
pub(crate) struct Office97KeySchedule {
    truncated_h1: PasswordDigest,
}

impl Office97KeySchedule {
    pub(crate) fn new(password: &str, salt: &[u8; SALT_LEN]) -> Self {
        // The whole §2.3.6.2 buffer is built inside one wrapped slot, and `H0` with it:
        // it is finalized into the slot's last 16 bytes, copied out five bytes at a time
        // into every repetition, and finally overwritten by the last repetition's salt.
        // So neither the UTF-16LE password (see `sensitive::utf16le_password` for the
        // `collect()` this replaced) nor `MD5(password)` nor the repeated
        // `TruncatedHash` is ever in storage this crate does not wipe. The typed closure
        // argument is what ties `BUFFER_LEN` to `Office97Buffer`: a mismatch is a
        // compile error, not a short copy.
        let buffer = Office97Buffer::new_with(|b: &mut [u8; BUFFER_LEN]| {
            let (body, last) = b.split_at_mut(BUFFER_LEN - REPEAT_LEN);
            let (_, h0) = last.split_at_mut(TRUNCATED_LEN);
            utf16le_password(password).with_secret(|pw| {
                Md5::new()
                    .chain_update(pw)
                    .finalize_into(Output::<Md5>::from_mut_slice(h0));
            });
            // `last` is `[5 unwritten | H0]`; `H0[..5]` is `last[5..10]`.
            let truncated = TRUNCATED_LEN..2 * TRUNCATED_LEN;
            for chunk in body.chunks_exact_mut(REPEAT_LEN) {
                let (t, s) = chunk.split_at_mut(TRUNCATED_LEN);
                t.copy_from_slice(&last[truncated.clone()]);
                s.copy_from_slice(salt);
            }
            // Overlapping, and `copy_within` is memmove: the last repetition's own five
            // bytes come from inside it, then its salt overwrites the rest of `H0`.
            last.copy_within(truncated, 0);
            last[TRUNCATED_LEN..].copy_from_slice(salt);
        });
        let truncated_h1 =
            buffer.with_secret(|b| PasswordDigest::new(Md5::digest(b)[..TRUNCATED_LEN].to_vec()));
        Self { truncated_h1 }
    }

    /// The password check — §2.3.6.4: the same shape as CryptoAPI's, with MD5.
    pub(crate) fn verify(&self, header: &Office97Header) -> Result<(), Error> {
        rc4::verify_password(
            self,
            &header.encrypted_verifier,
            &header.encrypted_verifier_hash,
            |v| Md5::digest(v).to_vec(),
        )
    }
}

impl BlockKeySchedule for Office97KeySchedule {
    /// `Hfinal = MD5(TruncatedHash + block)`, the block little-endian; all 128 bits of it
    /// are the key (§2.3.6.2, "the first 128 bits of Hfinal").
    fn block_key(&self, block: u32) -> DerivedKey {
        self.truncated_h1.with_secret(|t| {
            let mut hasher = Md5::new();
            hasher.update(t);
            hasher.update(block.to_le_bytes());
            DerivedKey::new(hasher.finalize().to_vec())
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

    fn hex(bytes: &[u8]) -> String {
        bytes.iter().map(|b| format!("{b:02x}")).collect()
    }

    /// msoffcrypto-tool's doctest vector (`method/rc4.py`, `_makekey` and
    /// `DocumentRC4.verifypw`): password `password1`, a salt, a block-0 key, and an
    /// encrypted verifier pair that passes. Block 1's key is the same tool run on the
    /// same inputs (`scratch kat.py`, 2026-09-05).
    const SALT: [u8; 16] = [
        0xe8, 0x77, 0x2c, 0x1d, 0x91, 0xc5, 0x6a, 0x37, 0x96, 0x47, 0x61, 0xb2, 0x80, 0x18, 0x32,
        0x17,
    ];
    const ENCRYPTED_VERIFIER: [u8; 16] = [
        0xc9, 0xe9, 0x97, 0xd4, 0x54, 0x97, 0x3d, 0x31, 0x0b, 0xb1, 0xba, 0x70, 0x14, 0x26, 0x83,
        0x7e,
    ];
    const ENCRYPTED_VERIFIER_HASH: [u8; 16] = [
        0xb1, 0xde, 0x17, 0x8f, 0x07, 0xe9, 0x89, 0xc4, 0x4d, 0xae, 0x5e, 0x4c, 0xf9, 0x6a, 0xc4,
        0x07,
    ];

    #[test]
    fn block_keys_match_msoffcrypto_doctest_vector() {
        let schedule = Office97KeySchedule::new("password1", &SALT);
        assert_eq!(
            schedule.block_key(0).with_secret(|k| hex(k)),
            "20bf32ddf540858c513744af0f24e03c"
        );
        assert_eq!(
            schedule.block_key(1).with_secret(|k| hex(k)),
            "8d6bfc4838ece4392b11f8f6c4d2843a"
        );
        // A second password/salt pair, from the same tool.
        let salt: [u8; 16] = core::array::from_fn(|i| i as u8);
        assert_eq!(
            Office97KeySchedule::new("testpass", &salt)
                .block_key(0)
                .with_secret(|k| hex(k)),
            "69620d508c5a8ee437f6432344b33db8"
        );
    }

    #[test]
    fn the_doctest_verifier_passes_and_a_wrong_password_is_named() {
        let header = Office97Header {
            salt: SALT,
            encrypted_verifier: ENCRYPTED_VERIFIER,
            encrypted_verifier_hash: ENCRYPTED_VERIFIER_HASH,
        };
        assert!(Office97KeySchedule::new("password1", &SALT)
            .verify(&header)
            .is_ok());
        assert!(matches!(
            Office97KeySchedule::new("password2", &SALT).verify(&header),
            Err(Error::WrongPassword)
        ));
    }

    #[test]
    fn the_header_is_parsed_and_bounded() {
        let mut s = vec![1u8, 0, 1, 0];
        s.extend_from_slice(&SALT);
        s.extend_from_slice(&ENCRYPTED_VERIFIER);
        s.extend_from_slice(&ENCRYPTED_VERIFIER_HASH);
        let h = parse(&s).unwrap();
        assert_eq!(h.salt, SALT);
        assert_eq!(h.encrypted_verifier_hash, ENCRYPTED_VERIFIER_HASH);
        for cut in [0usize, 3, 4, 51] {
            assert!(
                matches!(parse(&s[..cut]), Err(Error::BadParameters(_))),
                "cut {cut}"
            );
        }
        let mut wrong_version = s.clone();
        wrong_version[0] = 2;
        assert!(matches!(
            parse(&wrong_version),
            Err(Error::UnsupportedEncryptionVersion(2, 1))
        ));
    }
}
