# Security policy

## Reporting a vulnerability

Report privately through GitHub's **[Report a vulnerability]** button on this repository's
Security tab. That opens a draft advisory visible only to you and the maintainer; it is
the whole intake process, and it is deliberately the only one, so that no address needs to
be published or kept working.

Please do not open a public issue for anything in the next section. For anything in the
section after that, a public issue is exactly right.

Expect an acknowledgement within a week. There is no bounty.

## What counts as a vulnerability here

This crate's entire job is parsing files whose bytes an attacker chose. A caller opens an
untrusted `.docx` and hands it to `classify` or `decrypt_ooxml` **before anything else has
validated it**. That makes several things vulnerabilities that would be ordinary bugs in
another crate:

- **A panic on any input.** Slice indexing, `unwrap`, `expect`, an integer cast or a
  subtraction on a file-derived number — any of them reached by a crafted file is a denial
  of service for every caller not catching unwind. There is no such thing as a
  "malformed input" panic that is merely a bug. Return an error instead.
- **`classify` returning anything but a classification.** It is the first call made on an
  unknown file. It has no `Result` and must never panic; an unreadable file is `Unknown`.
- **Unbounded work or unbounded allocation.** `spinCount` is a file-controlled iteration
  count and a declared plaintext size is a file-controlled allocation. A file that makes
  the crate hang or exhaust memory is a denial of service that no `Result` can report.
  Every such parameter is bounded in `src/limits.rs`; a reachable path around one of those
  bounds is a vulnerability.
- **Key material escaping its wrapper.** The spin hash, the block keys and the session key
  are held in `secure-gate` types so they zeroize on drop. Any path that copies one into an
  error, a log line, a panic message, a `format!`, or a container that outlives the
  wrapper defeats the property this crate exists for.
- **A non-constant-time comparison on secret-derived data.** The `dataIntegrity` check
  compares a MAC this crate computes against a value the *file* supplies, which is the
  textbook MAC-forgery oracle. A short-circuiting comparison there leaks how many leading
  bytes of a forged tag were correct.
- **Returning plaintext that failed verification.** The agile package HMAC is computed over
  ciphertext and is checked before decryption. Bytes that failed that check must not reach
  the caller under the fail-closed `IntegrityPolicy`.
- **Reporting the wrong failure.** "Wrong password" and "file tampered" are different facts
  about an untrusted file. Telling a user their password is wrong when the file was
  modified is actively misleading, and the error variants are separate for that reason.

## What is not a vulnerability

- **The weakness of the formats themselves.** RC4 CryptoAPI ([MS-OFFCRYPTO] §2.3.5),
  Office 97/2000 RC4 (§2.3.6) and XOR obfuscation (§2.3.7) are broken by any modern
  standard — XOR is not encryption at all. They are implemented behind the off-by-default
  `legacy-binary` feature so that existing files can be *read*, and nothing in this crate
  writes them. A report that RC4 is weak is a report about 1997.
- **Low `spinCount`, weak passwords, or Office's own key derivation.** This crate
  implements the specified derivations. It cannot make a four-character password strong.
- **The password being `&str` at the public boundary, and the plaintext being `Vec<u8>`.**
  Neither is zeroized on the caller's copy. This is a documented design decision, not an
  oversight: no `secure-gate` type crosses the public API, which is what lets a consumer on
  a different `secure-gate` version link this crate at all. Wrapping happens the moment the
  password is used. If you want the caller's copy zeroized, that is the caller's
  `Zeroizing<String>` to hold.
- **Key-derived state held inside a dependency, where no wrapper here can reach it.** Three
  kinds, and they are different in what can be done about them. The cipher key schedules
  *are* wiped, because `aes`, `cbc` and `rc4` each have a `zeroize` feature and this crate
  enables all three — a CI invariant keeps them on, since losing one is silent. The
  `HmacSha*` opad/ipad state, derived from the integrity key, is **not**: `hmac` 0.12 has no
  `Drop` and no feature to enable, so it moves only on a dependency bump. Hasher buffers
  (`sha1`, `sha2`, `md-5` at 0.10) are likewise not wipeable and hold password bytes until
  `finalize`. Reports here are welcome and will be forwarded upstream, but this crate cannot
  fix them in place.
- **Anything requiring the attacker to already control the calling process.**
- **Findings in a dependency** with an advisory already published. Those belong upstream;
  `cargo deny check advisories` runs here on every push and will surface them.

## Scope

| | |
| --- | --- |
| In scope | `src/**` in every one of the five feature configurations — the default detection build, `crypto-ops`, `legacy-binary`, `cli`, and `cli,legacy-binary`. `src/bin/msoffice-crypto.rs` is `src/**`: the CLI reads files nothing has validated, so a panic there on a crafted input is a vulnerability on the same terms as a panic in the library. |
| In scope | The published crate on crates.io, at the latest version. |
| Out of scope | `tools/**` — fixture generators and the acceptance-gate drivers. They run on a maintainer's machine against files the maintainer chose, and are not shipped in the crate. |
| Out of scope | The test fixtures, which are deliberately hostile inputs. |

## Supported versions

Pre-1.0. **Only the most recently published version is supported.** A fix ships as a new
release rather than a backport.

## Design notes a reporter may find useful

- The **default build links no cryptography at all** — no cipher, hash, MAC, RNG or
  key-wrapping crate. CI asserts that property on every push. A consumer who only calls
  `classify` has none of the attack surface below it.
- The crate contains **no `unsafe`**, enforced by `#![forbid(unsafe_code)]` at the crate
  root, so memory-safety findings would have to originate in a dependency. That attribute
  binds `src/` — the code that ships. One test harness in `tests/`, which are separate
  crates and reach no consumer, does use `unsafe` to install a `#[global_allocator]`:
  `legacy_allocation.rs`, which measures peak allocation. It is not in the published
  tarball, and a `grep` over a clone finding `unsafe` there is not a contradiction of the
  line above.
- `classify`, the binary-format prober and every 97-2003 parser (`word97`, `excel97`,
  `powerpoint97`, the three RC4 families, `xor_obfuscation`, `legacy_container`) deny
  `clippy::unwrap_used`, `expect_used` and `panic` on themselves, and CI runs clippy under
  `-D warnings`. **The modern-format parsers do not carry that header today** — `agile`,
  `standard`, `integrity`, `cfb_reader`, `hash`, `segments` and `dataspaces` are held to the
  same rule by review and by the malformed-input suite, but not by the compiler. Treat a
  panic reachable in those as in scope exactly as if the lint were on; the missing header is
  a gap in enforcement, not a relaxation of the standard.
- **`src/bin/msoffice-crypto.rs` does not carry that header either**, and the scope row above
  puts it in scope regardless. It parses no bytes of its own: every one goes to `classify`,
  `decrypt_ooxml_with_policy`, `decrypt_binary_office` or an encrypt entry point, so what
  reaches the binary from a hostile file is the bounded attacker-chosen text inside
  `XmlParse`, `BadParameters` and `UnsupportedAlgorithm { name }`, written to stderr as-is and
  never re-interpolated into a `--json` field, a filename or a shell-quoted string. Its four
  `expect` calls are on `clap` values a `required(true)` or a `default_value` guarantees.
  Report a panic there as you would one in a parser.

None of these are guarantees against a logic flaw, which is what the list above is about.
