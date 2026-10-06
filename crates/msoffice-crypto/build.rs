//! Detects whether the fixture corpus is present, and tells the test suite.
//!
//! # The problem this solves
//!
//! `include` in `Cargo.toml` ships `src/**/*.rs`, which carries this crate's `#[cfg(test)]`
//! modules — `classify_tests.rs`, `legacy_malformed.rs`, the `tests` modules inside
//! `lib.rs`, `integrity.rs`, `encryption_info.rs` and `standard_encrypt.rs`. Since
//! 2026-09-05 it ships **2 of the 19 fixtures**, because the other 17 were two thirds of
//! the `.crate` and the repository is where the evidence is checked.
//!
//! Those two facts together meant `cargo test` on the published crate reported **30
//! failures**, every one of them `fixture ... must be present -- these tests are not
//! optional`. Nothing was broken: that is the *fixtures-are-not-optional* rule working
//! exactly as designed, against a tarball that deliberately does not carry the corpus. But
//! a red suite on a cryptography crate is a bad signal to the one audience that runs it —
//! distribution packagers and `cargo vendor` reviewers — and "read the release notes" is
//! not a fix.
//!
//! # Why a `cfg` and not a runtime skip
//!
//! The obvious repair is an early `return` when the file is missing. That is precisely the
//! anti-pattern this crate removed: four tests once read `let Ok(data) = fs::read(..) else
//! { return }` and a missing fixture produced a green suite that tested nothing. A test
//! that reports `ok` without exercising anything is, in CLAUDE.md's words, decoration.
//!
//! `#[ignore]` is the honest mechanism — the test is reported as **not run**, with a
//! reason, in cargo's normal output — and it is an attribute, so the decision has to be
//! made at compile time. That is what this script is for.
//!
//! # What it does NOT do
//!
//! **A partially present corpus counts as present.** If any fixture is found, the cfg is
//! set and every test runs; one that is individually missing then fails with its own
//! `must be present` panic, exactly as before. So this script can only ever disable the
//! corpus tests where there is *no* corpus at all, and a deleted fixture in a real checkout
//! still fails loudly. It is deliberately not a validity check, and it never fails a build.
//!
//! # Cost
//!
//! One build script, no build-dependencies, a few `Path::exists` calls. It adds nothing to
//! the dependency graph, so the detection build's "no cipher, hash, MAC, RNG or
//! key-wrapping crate" property is untouched.

use std::path::Path;

/// The fixtures the `include` allowlist does **not** ship, and which the corpus-dependent
/// tests read. `plain.docx` and `agile_encrypted.docx` are omitted on purpose: they ship in
/// every tarball (`classify`'s doc examples pull them in with `include_bytes!`, so without
/// them the crate does not *compile*), which means their presence says nothing about
/// whether the corpus is here.
const WITHHELD_FIXTURES: &[&str] = &[
    "agile_aes128_sha1.docx",
    "agile_aes128_sha384.docx",
    "agile_aes192_sha384.docx",
    "agile_aes256_sha256.docx",
    "agile_aes256_sha384.docx",
    "excel16_agile.xlsx",
    "excel97_password.xls",
    "excel97_plain.xls",
    "excel97_xor.xls",
    "plain_content.txt",
    "powerpoint16_agile.pptx",
    "powerpoint97_password.ppt",
    "powerpoint97_plain.ppt",
    "standard_encrypted.docx",
    "word16_agile.docx",
    "word97_password.doc",
    "word97_plain.doc",
];

fn main() {
    // The directory, not each file: a fixture appearing or disappearing has to re-run this,
    // and naming the directory covers a fixture that is added later without editing the
    // list above to match.
    println!("cargo::rerun-if-changed=tests/fixtures");
    println!("cargo::rerun-if-changed=build.rs");
    println!("cargo::rustc-check-cfg=cfg(fixture_corpus)");

    let root = std::env::var("CARGO_MANIFEST_DIR").unwrap_or_else(|_| ".".to_owned());
    let fixtures = Path::new(&root).join("tests").join("fixtures");

    // `any`, not `all` — see the module comment. Present-but-incomplete must behave like
    // present, so that the missing one fails its own test by name.
    let any_present = WITHHELD_FIXTURES
        .iter()
        .any(|name| fixtures.join(name).exists());

    if any_present {
        println!("cargo::rustc-cfg=fixture_corpus");
    }
}
