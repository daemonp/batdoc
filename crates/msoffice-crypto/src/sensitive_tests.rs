//! Tests for the secure-gate aliases and the UTF-16 password constructor.
//!
//! Split out for the reason `lib_tests.rs`'s header gives: an inline `#[cfg(test)]`
//! module in a production file cannot be excluded from CodeQL by path, and three
//! `rust/cleartext-logging` alerts pointed at `assert_eq!(.., "{password:?}")` here --
//! a panic message, interpolating a test constant.

use super::*;
use secure_gate::RevealSecret;

/// The replacement must produce the bytes the `collect()` it replaced produced.
///
/// This is the whole correctness argument for the change: the buffer is hashed, so
/// one byte of difference is a different `H_0`, a different `H_final` and a file
/// this crate can no longer open — and the two committed encrypt goldens would move.
/// The old expression is written out here rather than referenced because it no
/// longer exists in `src/`; it is the oracle, not a call.
///
/// The three character classes are the ones where UTF-8 length and UTF-16 length
/// disagree differently: ASCII (1 byte → 1 unit), a three-byte BMP character
/// (3 bytes → 1 unit, the worst case for `collect`'s size hint) and a non-BMP
/// character (4 bytes → a surrogate *pair*, 2 units). A `password.len() * 2` sizing
/// would be right for the first, too long for the second and right again for the
/// third, so only the second and third can catch it.
#[test]
fn utf16le_password_is_the_collect_it_replaced() {
    for password in [
        "",
        "a",
        "testpass",
        "correct horse battery staple",
        "pässwörd",
        "パスワード",
        "𝄞𝄞𝄞 music",
        "mixed ä 漢 𝄞 tail",
    ] {
        let want: Vec<u8> = password
            .encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect();
        utf16le_password(password).with_secret(|got| {
            assert_eq!(got, &want, "{password:?}");
        });
    }
}

/// One allocation at the exact length, which is the point of the function.
///
/// `capacity == len` is the observable half of "growth is not expressible": a `Vec`
/// that had grown into this size would carry slack from the doubling, and slack is
/// the signature of the reallocation that abandoned the earlier block. Assert it
/// rather than the allocation count, which needs a global allocator to see and is
/// not a property of the source — see `docs/design/heap-residue.md` on the 116
/// allocations that stop happening at `--release`.
#[test]
fn utf16le_password_allocates_exactly_its_length() {
    for password in ["a", "testpass", "correct horse battery staple", "漢字漢字"] {
        let expected = password.encode_utf16().count() * 2;
        utf16le_password(password).with_secret(|pw| {
            assert_eq!(pw.len(), expected, "{password:?}");
            assert_eq!(pw.capacity(), expected, "{password:?} carries slack");
        });
    }
}

/// The RC4 key schedules re-encode the password only through [`utf16le_password`].
///
/// Both carried the `encode_utf16().flat_map(..).collect()` spelling that function
/// replaced, until rc.5, a release after agile and standard stopped. Nothing about
/// the old line fails a test that checks bytes: it produces the same bytes, and only
/// the abandoned heap blocks differ. So this reads the source instead. It passes on
/// the fix and fails on the old line, which is the only thing a guard for a residue
/// fix can be made to observe.
///
/// Only the production half of each file is scanned: `rc4_cryptoapi.rs`'s own tests
/// encode a CSP name with `encode_utf16()`, which is not a password. And the scan
/// also asserts the replacement is *there*, so an unreadable or truncated file, or
/// a rewrite that stops hashing the password at all, cannot pass it by absence.
///
/// `include_str!` rather than a runtime read: it resolves beside this file in every
/// configuration, including a packaged crate, and runs without `legacy-binary`.
#[test]
fn rc4_schedules_encode_the_password_only_through_utf16le_password() {
    for (name, source) in [
        ("rc4_office97.rs", include_str!("rc4_office97.rs")),
        ("rc4_cryptoapi.rs", include_str!("rc4_cryptoapi.rs")),
    ] {
        let production = source
            .split_once("#[cfg(test)]")
            .map_or(source, |(before, _)| before);
        assert!(
            !production.contains(".encode_utf16()"),
            "{name} encodes the password itself; use sensitive::utf16le_password"
        );
        assert!(
            production.contains("utf16le_password(password)"),
            "{name} no longer builds its password through sensitive::utf16le_password"
        );
    }
}
