#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! XOR obfuscation, Method 1 — [MS-OFFCRYPTO] §2.3.7.1 to §2.3.7.3, as Excel applies it
//! to a BIFF workbook ([MS-XLS] §2.2.10). Not encryption in any sense the rest of this
//! crate means: a 16-byte array derived from the password, XORed into each record body
//! with a bit rotation, and a 16-bit checksum for a verifier. It is here because the
//! files exist and a caller holding the password deserves the bytes back.
//!
//! Word's Method 2 (§2.3.7.4 to §2.3.7.6) is a different transformation over the
//! `WordDocument` stream and is **not** implemented; msoffcrypto-tool does not implement
//! it either, and a `.doc` with `fObfuscated` set is refused by name.
//!
//! Every routine below is written from the spec's pseudocode. The three tables are the
//! spec's; the array initialisation is stated as the closed form the pseudocode
//! computes — position `p` of the array is the password byte, or the pad byte past the
//! password, XORed with alternating halves of the 16-bit key and rotated right once —
//! which msoffcrypto (`method/xor_obfuscation.py`, MIT) and LibreOffice
//! (`filter/source/msfilter/mscodec.cxx:141-166`, MPL, behaviour only) both compute the
//! long way, and which the known-answer tests below check against msoffcrypto's output
//! for four passwords.

use crate::error::Error;
use crate::limits::XOR_PASSWORD_MAX_LEN;
use crate::sensitive::{AnsiPassword, XorObfuscationArray};
use secure_gate::{ConstantTimeEq, RevealSecret};

/// `PadArray` — §2.3.7.2.
const PAD_ARRAY: [u8; 15] = [
    0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00,
];

/// `InitialCode` — §2.3.7.2, indexed by password length minus one.
const INITIAL_CODE: [u16; 15] = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9,
    0x280C, 0xA96A, 0x4EC3,
];

/// `XorMatrix` — §2.3.7.2, 105 entries walked from 0x68 downwards, seven per character.
const XOR_MATRIX: [u16; 105] = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7,
    0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4,
    0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, 0x6F45,
    0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC,
    0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD,
    0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483,
    0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
    0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373,
    0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4,
];

/// `CreatePasswordVerifier_Method1` — §2.3.7.1. The 16-bit verifier a workbook stores in
/// `FILEPASS.verificationBytes` ([MS-XLS] `XORObfuscation`), and the same function
/// [MS-XLS] uses for sheet and workbook protection hashes.
pub(crate) fn password_verifier(password: &[u8]) -> u16 {
    let mut verifier: u16 = 0;
    // The length byte is prepended, then the whole array is walked in reverse.
    let len_byte = u8::try_from(password.len()).unwrap_or(u8::MAX);
    for &byte in password.iter().rev().chain(core::iter::once(&len_byte)) {
        let intermediate1 = u16::from(verifier & 0x4000 != 0);
        let intermediate2 = verifier.wrapping_mul(2) & 0x7FFF;
        verifier = (intermediate1 | intermediate2) ^ u16::from(byte);
    }
    verifier ^ 0xCE4B
}

/// `CreateXorKey_Method1` — §2.3.7.2. `password` is 1..=15 bytes; the caller checks.
fn xor_key(password: &[u8]) -> u16 {
    let mut key = INITIAL_CODE
        .get(password.len().wrapping_sub(1))
        .copied()
        .unwrap_or(0);
    let mut element: usize = 0x68;
    for &ch in password.iter().rev() {
        let mut ch = ch;
        for _ in 0..7 {
            if ch & 0x40 != 0 {
                key ^= XOR_MATRIX.get(element).copied().unwrap_or(0);
            }
            ch = ch.wrapping_mul(2);
            element = element.wrapping_sub(1);
        }
    }
    key
}

/// `CreateXorArray_Method1` — §2.3.7.2, in the closed form the module docs describe:
/// position `p` holds `XorRor(source[p], key half)` where `source` is the password
/// followed by `PadArray`, and the key half is the high byte at odd positions and the
/// low byte at even ones.
fn xor_array(password: &[u8], key: u16) -> XorObfuscationArray {
    let [key_lo, key_hi] = key.to_le_bytes();
    XorObfuscationArray::new_with(|array| {
        for (p, slot) in array.iter_mut().enumerate() {
            let source = password
                .get(p)
                .or_else(|| PAD_ARRAY.get(p.wrapping_sub(password.len())))
                .copied()
                .unwrap_or(0);
            let half = if p % 2 == 1 { key_hi } else { key_lo };
            *slot = (source ^ half).rotate_right(1);
        }
    })
}

/// A password's XOR obfuscation state: the array, the key it stores in the file, and
/// the verifier it stores beside it.
pub(crate) struct XorObfuscator {
    array: XorObfuscationArray,
    key: u16,
    verifier: u16,
}

impl XorObfuscator {
    /// Derive everything from `password`.
    ///
    /// §2.3.7.2 takes an ASCII string of at most 15 characters, and [MS-XLS] §2.2.10 has
    /// Excel convert the user's Unicode password to the system ANSI code page first. A
    /// password no Excel could have used — longer than 15 characters, or holding a
    /// character above U+00FF — cannot be the one that produced any file, so it is
    /// reported as [`Error::WrongPassword`] rather than as a parameter
    /// problem: the file is fine. Characters U+0080..=U+00FF are taken as their code
    /// unit, which is the Windows-1252 byte for the Latin-1 range and is what
    /// msoffcrypto's `ord(ch)` does.
    pub(crate) fn new(password: &str) -> Result<Self, Error> {
        // Validated before anything is written, so that the bytes -- the password
        // itself -- go straight into a wrapped fixed slot rather than a `collect()`ed
        // `Vec` that grows, abandons its earlier blocks and is never wiped.
        let len = password.chars().count();
        if !(1..=XOR_PASSWORD_MAX_LEN).contains(&len)
            || !password.chars().all(|c| u8::try_from(u32::from(c)).is_ok())
        {
            return Err(Error::WrongPassword);
        }
        let bytes = AnsiPassword::new_with(|slot| {
            for (out, c) in slot.iter_mut().zip(password.chars()) {
                // Checked above; the fallback is unreachable, not a substitution.
                *out = u8::try_from(u32::from(c)).unwrap_or(0);
            }
        });
        Ok(bytes.with_secret(|b| {
            // `len` is at most 15 and the slot 16, so this never falls back.
            let b = b.get(..len).unwrap_or_default();
            let key = xor_key(b);
            Self {
                array: xor_array(b, key),
                key,
                verifier: password_verifier(b),
            }
        }))
    }

    /// The password check — §2.3.7.7 — against the `key` and `verificationBytes` a
    /// `FILEPASS` record carries ([MS-XLS] `XORObfuscation`).
    ///
    /// Both words are compared, as LibreOffice's `MSCodec_Xor95::VerifyKey` does
    /// (`mscodec.cxx:196-199`, behaviour only); msoffcrypto checks the verifier alone. A
    /// verifier that matches beside a key that does not is a malformed file or a wrong
    /// password, and decrypting under a key the file disagrees with would produce
    /// garbage with no error. Constant-time, as every comparison in this crate.
    pub(crate) fn verify(&self, key: u16, verification_bytes: u16) -> Result<(), Error> {
        let key_ok = self.key.to_le_bytes().ct_eq(&key.to_le_bytes());
        let verifier_ok = self
            .verifier
            .to_le_bytes()
            .ct_eq(&verification_bytes.to_le_bytes());
        if key_ok & verifier_ok {
            Ok(())
        } else {
            Err(Error::WrongPassword)
        }
    }

    /// `DecryptData_Method1` — §2.3.7.3: XOR with the array from `xor_array_index`,
    /// wrapping at 16, then rotate each byte right by 5.
    pub(crate) fn decrypt(&self, data: &mut [u8], xor_array_index: usize) {
        self.array.with_secret(|array| {
            for (i, byte) in data.iter_mut().enumerate() {
                let index = xor_array_index.wrapping_add(i) % array.len();
                *byte = (*byte ^ array[index]).rotate_right(5);
            }
        });
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

    /// Known answers from msoffcrypto-tool's `DocumentXOR` (`create_xor_key_method1`,
    /// `create_xor_array_method1`, `verifypw`) for four passwords spanning the length
    /// range, plus the verifier the spec's own pseudocode yields. `testpass`'s key and
    /// verifier are also what **Excel 16** wrote into the FILEPASS of an Excel 5.0/95
    /// save with that password (`0xA6CE`, `0x9727`, measured 2026-09-05), so the two
    /// derivations agree with Microsoft's writer as well as with the oracle.
    #[test]
    fn key_array_and_verifier_match_msoffcrypto_and_excel() {
        let cases: [(&str, u16, u16, &str); 4] = [
            (
                "testpass",
                0xA6CE,
                0x9727,
                "5de1de695fe3deeabaac980e98acbb13",
            ),
            (
                "VelvetSweatshop",
                0xB359,
                0x9A0A,
                "876b9ae21ee305621e699660986e9404",
            ),
            ("a", 0x9D77, 0xCE88, "0b134431e6314412fbcee449bb113cce"),
            (
                "123456789012345",
                0x8265,
                0x9EB1,
                "2a582b5b285a295d2e592a582b5b289c",
            ),
        ];
        for (password, key, verifier, array) in cases {
            let o = XorObfuscator::new(password).unwrap();
            assert_eq!(o.key, key, "{password}: key");
            assert_eq!(o.verifier, verifier, "{password}: verifier");
            assert_eq!(o.array.with_secret(|a| hex(a)), array, "{password}: array");
            assert!(o.verify(key, verifier).is_ok(), "{password}");
        }
    }

    /// msoffcrypto's own doctest: `VelvetSweatshop` — Excel's default password for a
    /// file saved with none, which is why it appears in every implementation — verifies
    /// against `0x9A0A`.
    #[test]
    fn velvet_sweatshop_verifies_against_the_documented_word() {
        let o = XorObfuscator::new("VelvetSweatshop").unwrap();
        assert!(o.verify(0xB359, 0x9A0A).is_ok());
        assert!(matches!(
            o.verify(0xB359, 0x9A0B),
            Err(Error::WrongPassword)
        ));
        // The key is checked as well as the verifier.
        assert!(matches!(
            o.verify(0xB358, 0x9A0A),
            Err(Error::WrongPassword)
        ));
    }

    /// The transformation inverts §2.3.7.3's `EncryptData_Method1`, at every starting
    /// index, and the index wraps at 16.
    #[test]
    fn decrypt_inverts_the_spec_encrypt_transformation() {
        let o = XorObfuscator::new("testpass").unwrap();
        let plain: Vec<u8> = (0..70u8).collect();
        for start in [0usize, 1, 7, 15, 16, 33] {
            let mut data = plain.clone();
            // EncryptData_Method1: rotate left 5, then XOR with the array.
            o.array.with_secret(|a| {
                for (i, b) in data.iter_mut().enumerate() {
                    *b = b.rotate_left(5) ^ a[(start + i) % 16];
                }
            });
            assert_ne!(data, plain);
            o.decrypt(&mut data, start);
            assert_eq!(data, plain, "start {start}");
        }
    }

    /// A password Excel could not have used is a wrong password, never a panic: the
    /// spec's 15-character ceiling, an empty string, and non-Latin-1 characters.
    #[test]
    fn impossible_passwords_are_wrong_not_panics() {
        for password in [
            "",
            "1234567890123456",
            "sixteen chars!!!",
            "pässwörd☃",
            "日本語",
        ] {
            assert!(
                matches!(XorObfuscator::new(password), Err(Error::WrongPassword)),
                "{password:?}"
            );
        }
        assert!(
            XorObfuscator::new("pässwörd").is_ok(),
            "Latin-1 is a code-page byte"
        );
    }
}
