# msoffice-crypto

[![crates.io](https://img.shields.io/crates/v/msoffice-crypto.svg?include_prereleases)](https://crates.io/crates/msoffice-crypto)
[![docs.rs](https://img.shields.io/docsrs/msoffice-crypto)](https://docs.rs/msoffice-crypto)
[![CI](https://github.com/Slurp9187/msoffice-crypto/actions/workflows/ci.yml/badge.svg)](https://github.com/Slurp9187/msoffice-crypto/actions/workflows/ci.yml)
[![MSRV](https://img.shields.io/crates/msrv/msoffice-crypto)](https://github.com/Slurp9187/msoffice-crypto#msrv)
[![License](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue.svg)](#license)

[MS-OFFCRYPTO]-faithful Microsoft Office encryption: detect it, decrypt it, write it.

Every encryption format the specification defines — OOXML and the 97-2003 binary
documents alike — with **key material that zeroizes on drop**. What it writes is a subset
of what it reads, and the table below says which.

> **Status: release candidates on crates.io, no stable release.** This tree is
> `0.1.0-rc.5` and is not released; `0.1.0-rc.4` is the newest on the registry. The API is
> still free to break between candidates.

## What it does

| Format | Detect | Decrypt | Encrypt | Needs |
| --- | :-: | --- | --- | --- |
| **ECMA-376 Agile** (Office 2010+) | ✅ | ✅ AES-128/192/256 × SHA-1/256/384/512 | ✅ the same set (SHA-1 at 128 only); default AES-256/SHA-512 | `crypto-ops` |
| **ECMA-376 Standard** (Office 2007) | ✅ | ✅ AES-128/192/256 | ✅ AES-128/192/256; default AES-128 | `crypto-ops` |
| **RC4 CryptoAPI** (Office XP–2003, and what Office 16 writes into a binary file) | ✅ | ✅ `.doc` `.xls` `.ppt` | ❌ | `legacy-binary` |
| **Office 97/2000 RC4** (MD5) | ✅ | ✅ `.doc` `.xls` | ❌ | `legacy-binary` |
| **XOR obfuscation** (Method 1) | ✅ | ✅ `.xls` | ❌ | `legacy-binary` |
| **Unencrypted** OOXML or binary | ✅ | n/a | this is what encrypt takes in | *(default)* |

**Writable is not the same as externally accepted.** Every encrypt cell above is what this
crate emits. What other readers do with it is measured, not claimed: the default agile tuple
and standard AES-128 open in all four readers of the [acceptance gate](#how-its-verified);
the rest are recorded reader by reader, cell by cell, in [`CHANGELOG.md`][changelog].

**The three ❌ are a refusal, not a backlog.** RC4 and XOR obfuscation are broken by any
modern standard — XOR is not encryption at all. **Reading them is not optional**: a 2026
Office build still writes RC4 CryptoAPI when you save as *97-2003*, so these are current
files, not museum pieces. Writing them *is* optional, and this crate declines — nothing
obliges a writer to emit every format a specification defines, and emitting one would
produce a newly weak document. See [Security](#security).

**Out of scope, permanently:** password recovery and cracking.

## Why this one

Aimed at callers who handle documents they did not create — vaults, backup and archival
tools, mail and upload gateways, DLP scanners, indexing pipelines that keep hitting files
they cannot open.

1. **One API across detect, decrypt and encrypt.** `office-crypto` decrypts but does not
   encrypt; `ms-offcrypto-writer` encrypts agile but does not decrypt; `msoffcrypto-tool` is
   Python and `herumi/msoffice` is C++. A complete picture otherwise means several
   dependencies with several postures toward the same key bytes.
2. **Key material is wrapped.** Those four all hold the spin hash, block keys and session key
   in bare `Vec<u8>` / `std::string`. This crate wraps them in
   [`secure-gate`](https://crates.io/crates/secure-gate): zeroized on drop, `[REDACTED]` in
   `Debug`, reachable only inside a `with_secret` closure, and comparable *only* in constant
   time — no `PartialEq` at all, so a verifier or HMAC cannot be compared with `==` by
   accident. Scope is honest: this is the key material *this crate holds*.
   [`SECURITY.md`](SECURITY.md) says which dependency state is wiped and which is not.
3. **Detection costs nothing.** *What is this file, is it encrypted, how strongly* is the
   common case, and it pulls in no cryptography. `cargo tree --no-default-features` is the
   claim.
4. **It fails closed, and says which failure happened.** The agile `dataIntegrity` HMAC is
   computed over ciphertext, so it is checked *before* any plaintext is returned;
   `IntegrityPolicy::Require` is the default and `Skip` comes back labelled. `WrongPassword`,
   `IntegrityCheckFailed` and `UnsupportedAlgorithm` are three variants, because telling
   someone their password is wrong when the file was modified is actively misleading.
5. **The claims are checkable rather than assertable.** Fixtures ship with their generators,
   so evidence is re-run rather than taken on faith. Where an encrypt path exists the bar is
   that real Word opens what it wrote and that independent implementations recover the same
   bytes — not that it round-trips against itself. [`CHANGELOG.md`][changelog] records what
   was run and what each reader answered, against the artifact's hash.

## Install

```toml
[dependencies]
# Detection only — no cryptographic dependency.
msoffice-crypto = "0.1.0-rc.5"

# Detection, decryption and encryption.
msoffice-crypto = { version = "0.1.0-rc.5", features = ["crypto-ops"] }
```

Pre-release versions are not matched by ordinary requirements — name the full version as
above; `"0.1"` will not resolve to it.

The CLI is a separate install; see [Command line](#command-line).

## Usage

Three builds, and you pay only for the one you take.

**Detection is free.** `classify` answers what a file is — container shape, encryption
family, the algorithm tuple it declares, whether it carries a `dataIntegrity` element — with
no cipher, hash, MAC, RNG or key-wrapping crate in the graph at all, a property CI asserts on
every push. It returns no `Result` and never panics: an unreadable file classifies as
`Unknown`, and a malformed one is never called unencrypted.

```rust,no_run
use msoffice_crypto::{classify, Family, IntegrityDeclaration};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let class = classify(&std::fs::read("protected.docx")?);
    if class.family == Family::Agile && class.data_integrity == IntegrityDeclaration::Declared {
        println!("agile, and its package HMAC can be verified");
    }
    Ok(())
}
```

**`crypto-ops`** adds the ciphers, hashes, `secure-gate` and a CSPRNG:

```rust,no_run
use msoffice_crypto::{decrypt_ooxml, encrypt_ooxml};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // the original OOXML package — a valid .docx/.xlsx/.pptx
    let zip = decrypt_ooxml(&std::fs::read("protected.docx")?, "correct horse battery staple")?;
    std::fs::write("plain.docx", &zip)?;

    // agile, AES-256/SHA-512, 100 000 rounds, dataIntegrity HMAC — what Office 16 writes
    let sealed = encrypt_ooxml(&zip, "correct horse battery staple")?;
    std::fs::write("resealed.docx", sealed)?;
    Ok(())
}
```

That default is a default, not a limit. `encrypt_ooxml_with_params` takes an `EncryptParams`
— spin count, hash, and `keyBits`/`saltSize` for the package key and the password key
separately, as [MS-OFFCRYPTO] treats them — and `validate()` refuses only what the format or
AES actually forbids, saying which of the two it was:

```rust,no_run
use msoffice_crypto::{encrypt_ooxml_with_params, EncryptParams, HashAlgorithm};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let package = std::fs::read("report.docx")?;
    let sealed = encrypt_ooxml_with_params(&package, "correct horse battery staple",
        EncryptParams {
            hash: HashAlgorithm::Sha384,
            key_data_key_bits: 192,
            password_key_bits: 192,
            ..Default::default()
        })?;
    std::fs::write("report-protected.docx", sealed)?;
    Ok(())
}
```

Ten `(hash, keyBits)` combinations are writable: SHA-1 carries 128 only, because a key cannot
be longer than the digest it is derived from, and SHA-256/384/512 each carry all three.
`EncryptParams::default()` is byte-for-byte what `encrypt_ooxml` writes.

For a reader that predates agile encryption, `encrypt_ooxml_standard` writes the Office 2007
format — AES-ECB under a SHA-1-derived key, and **no integrity element at all**, so a
modified ciphertext decrypts silently to a modified document. It is named for the format so
that choosing it is a decision.

**`legacy-binary`** is a superset of `crypto-ops` adding RC4 and MD5, and decrypts a
97-2003 document *in place* — the same CFB container with its encrypted streams replaced,
which is what Word, Excel and PowerPoint open:

```rust,no_run
use msoffice_crypto::{classify, decrypt_binary_office, Document};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let data = std::fs::read("protected.doc")?;
    if classify(&data).document == Document::WordBinary {
        // byte for byte what msoffcrypto-tool writes
        let doc = decrypt_binary_office(&data, "correct horse battery staple")?;
        std::fs::write("unlocked.doc", doc)?;
    }
    Ok(())
}
```

## Command line

Behind the opt-in `cli` feature, so no library consumer builds an argument parser to ask
whether a file is encrypted.

```text
cargo install msoffice-crypto --features cli
cargo install msoffice-crypto --features cli,legacy-binary   # also opens 97-2003 files
```

**`classify` cannot fail.** An unencrypted package, sixteen bytes of junk and a container
this crate has never seen all exit 0, because all three are answers; only a file that cannot
be *read* exits 2. `--json` prints one object with the same key set for every input — an
unencrypted file carries `key_data` and `password_key` as `null` rather than dropping them,
so a script can index the result without checking first.

```text
$ msoffice-crypto classify protected.docx
container:      cfb
container-read: opened
document:       ooxml-package
version:        4.4
family:         agile
encrypted:      yes
supported:      yes
integrity:      declared
key-cipher:     AES          key-hash:  SHA-512   key-bits:  256
pw-cipher:      AES          pw-hash:   SHA-512   pw-bits:   256
pw-spin:        100000
```

**`decrypt` says what it verified.** The output path is derived from the input unless `-o`
gives one, an existing file is never overwritten without `--force`, and the write is a
temporary file renamed over the target, so an interrupted run cannot leave a half-written
`.docx` that looks complete. The `integrity:` line goes to stderr after *every* successful
decrypt — including after `--integrity skip` — so `-o -` into a pipe is unaffected and
nobody holds unauthenticated bytes without being told.

```text
$ msoffice-crypto decrypt protected.docx --password-env MSOFFICE_CRYPTO_PASSWORD
msoffice-crypto: wrote protected.decrypted.docx
integrity: verified
```

`--integrity` takes `require`, `require-where-defined`, `verify-if-present` or `skip`; the
default is rendered from the library's own `IntegrityPolicy::default()` rather than typed
into the help text, because it has already moved once. A file with nothing to decrypt is
refused rather than copied, with exit 5.

**`encrypt`** defaults to `--format agile`. `--format standard` writes the Office 2007
format and says what that costs rather than leaving it to folklore:

```text
$ msoffice-crypto encrypt report.docx --format standard --password-env MSOFFICE_CRYPTO_PASSWORD
msoffice-crypto: wrote report.encrypted.docx
integrity: not-applicable
msoffice-crypto: Office 2007 standard encryption defines no dataIntegrity element: a file modified after it was encrypted decrypts without complaint. `--format agile` writes one.
```

### Passwords never come from `argv`

`argv` is world-readable in a process listing for the lifetime of the run — `ps aux`,
Task Manager. **There is deliberately no `--password VALUE` flag, and adding one later would
be a regression rather than a feature.** `--password` is registered hidden, so reaching for
it is answered with that reason instead of clap's generic "unexpected argument". Exactly one
source may be given; two is a usage error rather than a silent precedence win.

| Flag | Source | For |
| --- | --- | --- |
| `--password-env NAME` | that environment variable | scripts, CI |
| `--password-file PATH` | first line, one trailing CR-LF or LF stripped | secret managers, `/dev/stdin` |
| `--password-stdin` | one line from stdin | pipelines |
| *(none)* | a non-echoing terminal prompt | interactive use |

`--password-env` takes the variable's **name**: naming it is the only way this tool reads it,
because a password that applies without being asked for is how the wrong file gets decrypted
in a loop.

### Exit codes

A CLI that returns 1 for everything cannot be scripted.

| Code | Meaning | Maps from |
| --- | --- | --- |
| 0 | success | — |
| 1 | usage error | bad flags, missing operand, two password sources, output exists without `--force`; also `Error::EncryptParams`, which no CLI flag can produce today |
| 2 | I/O error | unreadable input, unwritable output, `Error::Io` |
| 3 | not a Microsoft Office file | `Error::NotACfbFile` |
| 4 | wrong password | `Error::WrongPassword` |
| 5 | refused: nothing to do | decrypt of an unencrypted file, encrypt of an already-encrypted one, `Error::NotEncrypted` |
| 6 | malformed or hostile file | `MissingStream`, `XmlParse`, `BadParameters`, `CipherError` |
| 7 | internal invariant violated | `RandomSource` |
| 8 | **integrity** | `IntegrityCheckFailed`, `IntegrityElementMissing`, `IntegrityUnavailable` |
| 9 | **unsupported encryption** | `UnsupportedEncryptionVersion`, `UnsupportedAlgorithm`, `Family::Unsupported`, a legacy family without `legacy-binary` |

8 is not folded into 4 or 6, and 9 is not folded into 5, for the reason the library keeps
`WrongPassword` and `IntegrityCheckFailed` apart: at a process boundary the number is all a
script gets, and "try again", "this file was changed after it was encrypted" and "rebuild
with `--features cli,legacy-binary`" are three different next steps. One is a retry, one is a
support ticket, one is an incident.

## Features

| Feature | Adds crypto? | Enables | Dependencies |
| --- | :-: | --- | --- |
| *(none — default)* | no | `classify`, `is_cfb_office` | `cfb`, `quick-xml`, `thiserror` — **13 crates** |
| `crypto-ops` | yes | `decrypt_ooxml`, `decrypt_ooxml_with_policy`, `encrypt_ooxml`, `encrypt_ooxml_with_params`, `encrypt_ooxml_standard`, `encrypt_ooxml_standard_with_key_bits`, `check_encryptable`, `EncryptParams`, `IntegrityPolicy` / `IntegrityOutcome` | + `aes`, `cbc`, `ecb`, `sha1`, `sha2`, `hmac`, `base64`, `rand`, `secure-gate` — **36 crates** |
| `legacy-binary` | superset of `crypto-ops` | `decrypt_binary_office` | + `rc4`, `md-5` — **38 crates** |
| `cli` | via `crypto-ops` | the `msoffice-crypto` binary | + `clap`, `serde_json`, `rpassword`/`rtoolbox` (**Apache-2.0-only**) — **49 crates** |

Everything reachable from a library build is dual MIT/Apache-2.0 or more permissive.
`rpassword` and `rtoolbox` are **Apache-2.0-only** and arrive with the CLI's non-echoing
prompt under `cli`; CI fails if either reaches the default graph. It matters because this
crate is offered as `MIT OR Apache-2.0` and the point of a dual offer is that you may take
*either*. `deny.toml` records the same finding beside the allow-list.

## How it's verified

A round-trip through this crate's own `decrypt` proves only that it agrees with itself.
What it writes is held to four readers instead — two shipping products, two implementations
sharing no code with this one:

| Reader | Driver | Bar |
| --- | --- | --- |
| Word / Excel / PowerPoint 16 | COM, `tools/office_com_check.ps1` | opens, content matches, wrong password refused for the password reason (`0x800A1520`) |
| LibreOffice 26.2 | UNO, `tools/libreoffice_uno_check.py` | opens, content matches, wrong password refused by its verifier |
| `msoffcrypto-tool` | CLI, and the library for the HMAC | plaintext **byte-identical**; `dataIntegrity` verifies; wrong password refused |
| `office-crypto` | `examples/office_crypto_check.rs` | plaintext **byte-identical**; wrong password does not yield it |

`tools/acceptance_gate.py` runs all four. The two implementations run in CI on every push
with two mutation runs beside them — `--tamper`, one flipped ciphertext bit, and
`--corrupt-integrity` — which must each produce `GATE: FAIL`, because a gate that cannot fail
is not a gate. The two applications need an interactive Windows desktop, so they are the
local gate, run before a change to the encrypt path merges and recorded in
[`CHANGELOG.md`][changelog] against the artifact's SHA-256.

That gate reads `MSOFFICE_CRYPTO_ARTIFACT_DIR`, and it must be set to a private
directory. Every `cargo test` writes the artifact under one fixed name, so a gate
pointed at shared system temp can measure whichever build wrote it last — which has
happened, and is why the evidence is recorded against a hash rather than a path.

## Security

Report privately through **Report a vulnerability** on this repository's Security tab.
[`SECURITY.md`](SECURITY.md) is the policy, and its useful half is what does *not* count:
RC4 and XOR obfuscation are broken by any modern standard — XOR is not encryption at all.
They live behind the off-by-default `legacy-binary` feature so those files can be
**read** — including ones a current Office writes when you choose the 97-2003 format —
and nothing here writes them.

Two properties are enforced rather than claimed: no `unsafe` (`#![forbid(unsafe_code)]` at
the crate root, which also refuses an `#[allow]` override), and `cargo deny` over licences,
advisories, bans and sources on every push. The first thing that job found was a reachable
quadratic-time denial of service in this crate's XML parser — `RUSTSEC-2026-0194`,
remediated by the `quick-xml` floor in `Cargo.toml`.

## MSRV

Rust **1.85**, checked by a dedicated CI job rather than declared and hoped for.

## Sibling crate

[`odf-crypto`](https://github.com/Slurp9187/odf-crypto) does for OpenDocument what this
crate does for Microsoft Office — same method, same author, different format family.
**Compatibility:** this crate tracks `secure-gate` 0.9.x, and **no secure-gate type crosses
its public API** — the password is `&str` and the plaintext is `Vec<u8>`, by design. A plan
to put its types on the boundary was withdrawn as a design error, so a consumer on a
different secure-gate version resolves both side by side without conflict.

## Trademarks

Microsoft, Microsoft Office, Word, Excel and PowerPoint are trademarks of Microsoft
Corporation. This project is not affiliated with, endorsed by, or sponsored by Microsoft. It
is an independent implementation of the [MS-OFFCRYPTO] file formats — published by Microsoft
under the Open Specification Promise, which is a patent promise and grants no trademark
rights — and uses those names only to describe the formats it reads and writes.

## Acknowledgements

None of this is legally required beyond what [`NOTICE`](NOTICE) records. It is here because
the work would have been substantially harder without them.

- **[herumi/msoffice](https://github.com/herumi/msoffice)** (BSD-3, Cybozu Labs) — the
  encryption reference; `resource.hpp` is the entire `\x06DataSpaces` subtree in four
  constants.
- **[office-crypto](https://github.com/Udbhav-Muthakana/office-crypto)** (MIT) — the
  differential oracle in dev-dependencies, and the source the RC4 families and the Word 97
  walk were ported from.
- **[ms-offcrypto-writer](https://github.com/42triangles/ms-offcrypto-writer)** (MIT/Apache)
  — proof that the `cfb` crate builds containers Office accepts.
- **[msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool)** (MIT) — fixtures and
  known-answer vectors that caught two real bugs in this crate's standard-encryption key
  derivation, and the source of the Excel 97 and PowerPoint 97 walks.
- **[LibreOffice](https://www.libreoffice.org/)** (MPL-2.0) and
  **[excelize](https://github.com/qax-os/excelize)** (BSD-3) — behavioural references,
  consulted and cited, never copied.
- **[cfb](https://github.com/mdsteele/rust-cfb)** (MIT) — the container layer, so this crate
  does not hand-roll a third implementation of FAT chains and red-black directory trees.

The arc of the work is in
[`docs/plans/msoffice-crypto-foundation-2026-09-04.md`](https://github.com/Slurp9187/msoffice-crypto/blob/main/docs/plans/msoffice-crypto-foundation-2026-09-04.md)
and the issues it indexes.

## License

`MIT OR Apache-2.0`, at your option. See [`LICENSE-MIT`](LICENSE-MIT),
[`LICENSE-APACHE`](LICENSE-APACHE) and [`NOTICE`](NOTICE).

[MS-OFFCRYPTO]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
[changelog]: https://github.com/Slurp9187/msoffice-crypto/blob/main/CHANGELOG.md
