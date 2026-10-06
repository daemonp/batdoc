//! Unit tests for the CLI's own logic — the parts that are not the library and not clap.
//!
//! **Fixture-free, and that is a packaging constraint rather than a preference.**
//! `Cargo.toml`'s `include` allowlist ships `src/**/*.rs`, so this module is published,
//! while seventeen of the nineteen fixtures are not. A single `include_bytes!` here
//! would turn the published crate's `cargo test` red over a file the tarball
//! deliberately does not ship. Nothing below reads a fixture — not even one of the two
//! that do ship.
//!
//! End-to-end behaviour (exit codes as a process, argv handling, file side effects,
//! every fixture-driven decrypt) is `tests/cli.rs`'s; that file never reaches the
//! tarball and may read all nineteen. The two file-system tests below write only to a
//! directory under `std::env::temp_dir()` that they create and remove themselves.

use super::*;
/// The two halves of [`Error::EncryptParams`]'s payload. Imported here rather than at
/// the top of the binary because the CLI itself never names them: it cannot raise that
/// variant today (no `--spin-count`/`--key-bits` flags), and `exit_code` matches it with
/// `..`. Only the exit-code assertion below has to build one.
use msoffice_crypto::{EncryptParam, EncryptParamProblem};

/// `AlgorithmParams` is `#[non_exhaustive]`, so from this crate it can be neither
/// struct-literalled nor built with `..Default::default()`. Field assignment after
/// `Default::default()` is the only route, and routing it through a closure keeps
/// `clippy::field_reassign_with_default` off the call sites.
fn params(fill: impl FnOnce(&mut AlgorithmParams)) -> AlgorithmParams {
    let mut p = AlgorithmParams::default();
    fill(&mut p);
    p
}

#[test]
fn the_command_definition_is_internally_consistent() {
    // clap's own audit: a duplicate id, a group naming an argument that does not exist,
    // a short flag used twice -- and a `default_value` outside its own value_parser's
    // possible values, which is how a hard-coded `--integrity` default would surface.
    cli().debug_assert();
}

#[test]
fn help_names_all_three_subcommands() {
    let text = cli().render_long_help().to_string();
    for name in ["classify", "decrypt", "encrypt"] {
        assert!(text.contains(name), "top-level help must name {name}");
    }
}

#[test]
fn the_exit_code_table_in_the_help_carries_all_ten_codes() {
    // The sibling's block stops at 7. Ported unchanged it would leave the two codes
    // this crate added undocumented in the one place a user looks for them.
    for text in [AFTER_HELP, PASSWORD_AFTER_HELP] {
        for token in [
            "0 ok",
            "1 usage",
            "2 io",
            "3 not-office",
            "4 wrong-password",
            "5 refused",
            "6 malformed",
            "7 internal",
            "8 integrity",
            "9 unsupported",
        ] {
            assert!(text.contains(token), "exit-code help must carry {token:?}");
        }
    }
}

#[test]
fn derived_output_inserts_before_the_extension() {
    assert_eq!(
        derived_output(Path::new("report.docx"), "decrypted"),
        PathBuf::from("report.decrypted.docx")
    );
    assert_eq!(
        derived_output(Path::new("/tmp/a/book.xlsx"), "encrypted"),
        PathBuf::from("/tmp/a/book.encrypted.xlsx")
    );
    // The legacy extensions survive too: `decrypt_binary_office` returns a rewritten
    // CFB, not a package, so the extension must not be normalised to anything.
    assert_eq!(
        derived_output(Path::new("memo.doc"), "decrypted"),
        PathBuf::from("memo.decrypted.doc")
    );
}

#[test]
fn derived_output_appends_when_there_is_no_extension() {
    assert_eq!(
        derived_output(Path::new("report"), "decrypted"),
        PathBuf::from("report.decrypted")
    );
}

#[test]
fn derived_output_keeps_a_dotted_stem() {
    // `a.b.docx` has stem `a.b`; the suffix goes before the real extension only.
    assert_eq!(
        derived_output(Path::new("a.b.docx"), "decrypted"),
        PathBuf::from("a.b.decrypted.docx")
    );
}

#[test]
fn first_line_strips_one_trailing_newline_and_no_more() {
    assert_eq!(first_line("testpass\n"), "testpass");
    assert_eq!(first_line("testpass\r\n"), "testpass");
    assert_eq!(first_line("testpass"), "testpass");
    // Only the first line: a file with more in it is not a multi-line password.
    assert_eq!(first_line("testpass\nignored\n"), "testpass");
    // Trailing spaces can be deliberate, so they survive.
    assert_eq!(first_line("testpass  \n"), "testpass  ");
    assert_eq!(first_line(""), "");
}

#[test]
fn exit_codes_map_every_error_class() {
    // `Error` has no `PartialEq` (deliberately), so every assertion here is on the u8
    // the mapping returned, never on the error value.
    assert_eq!(exit_code(&Error::NotACfbFile), EX_NOT_OFFICE);
    assert_eq!(
        exit_code(&Error::MissingStream("EncryptionInfo")),
        EX_MALFORMED
    );
    assert_eq!(
        exit_code(&Error::BadParameters(String::new())),
        EX_MALFORMED
    );
    assert_eq!(exit_code(&Error::Io(std::io::Error::other("x"))), EX_IO);
    assert_eq!(
        exit_code(&Error::UnsupportedEncryptionVersion(4, 3)),
        EX_UNSUPPORTED
    );
    assert_eq!(exit_code(&Error::XmlParse(String::new())), EX_MALFORMED);
    assert_eq!(exit_code(&Error::WrongPassword), EX_WRONG_PASSWORD);
    assert_eq!(exit_code(&Error::CipherError), EX_MALFORMED);
    assert_eq!(
        exit_code(&Error::UnsupportedAlgorithm {
            what: "p:encryptedKey/@hashAlgorithm",
            name: String::new()
        }),
        EX_UNSUPPORTED
    );
    assert_eq!(exit_code(&Error::IntegrityCheckFailed), EX_INTEGRITY);
    assert_eq!(exit_code(&Error::IntegrityElementMissing), EX_INTEGRITY);
    assert_eq!(exit_code(&Error::IntegrityUnavailable("x")), EX_INTEGRITY);
    assert_eq!(exit_code(&Error::RandomSource(String::new())), EX_INTERNAL);
    assert_eq!(
        exit_code(&Error::AlreadyEncrypted {
            family: Family::Agile,
            document: Document::OoxmlPackage
        }),
        EX_REFUSED
    );
    assert_eq!(exit_code(&Error::NotAPlainPackage), EX_REFUSED);
    assert_eq!(exit_code(&Error::UnknownContainer), EX_NOT_OFFICE);
    // The one caller error in the table. It is here although today's CLI cannot raise
    // it -- there is no `--spin-count` flag -- because the number is the contract the
    // library's callers see through this binary's exit status, and the canary in
    // `src/error.rs` pins the same 1 from the other side. EX_USAGE, not EX_MALFORMED:
    // there is often no file to call malformed. The payload is a *margin* rejection
    // deliberately, so the assertion also fails if the variant loses its shape.
    assert_eq!(
        exit_code(&Error::EncryptParams {
            param: EncryptParam::SpinCount,
            problem: EncryptParamProblem::ExceedsImplementationLimit,
            got: 20_000_000,
            min: 100_000,
            max: 10_000_000,
        }),
        EX_USAGE
    );

    // Gated with the variant itself: `cli` does not enable `legacy-binary`, and
    // `Error::NotEncrypted` does not exist in that column.
    #[cfg(feature = "legacy-binary")]
    assert_eq!(exit_code(&Error::NotEncrypted), EX_REFUSED);
}

#[test]
fn the_four_facts_the_table_exists_for_stay_distinct() {
    // Collapsing any pair of these re-creates, at the process boundary, the bug
    // CLAUDE.md's cryptographic rules forbid inside the library.
    let codes = [
        EX_WRONG_PASSWORD,
        EX_REFUSED,
        EX_INTEGRITY,
        EX_UNSUPPORTED,
        EX_MALFORMED,
    ];
    for (i, a) in codes.iter().enumerate() {
        for b in &codes[i + 1..] {
            assert_ne!(a, b, "two distinct facts share one exit code");
        }
    }
}

#[test]
fn the_integrity_default_is_the_librarys_own_default() {
    // Not "the default is require-where-defined": that flip has already happened once
    // (GH #12), and a CLI with the old name typed in would have survived it looking
    // correct. What is pinned is that the flag's default IS whatever the library
    // defaults to, rendered through the table clap validates against.
    let rendered = policy_name(IntegrityPolicy::default());
    assert!(
        POLICY_NAMES.contains(&rendered),
        "the library default renders as {rendered:?}, which --integrity would reject"
    );
    let m = cli()
        .try_get_matches_from(["msoffice-crypto", "decrypt", "f.docx"])
        .expect("decrypt must parse with no --integrity");
    let sub = m.subcommand_matches("decrypt").expect("decrypt");
    assert_eq!(
        sub.get_one::<String>("integrity").map(String::as_str),
        Some(rendered)
    );
}

#[test]
fn the_equals_form_parses() {
    // The GNU `--flag=value` form the sibling's hand-rolled parser rejected outright.
    let m = cli()
        .try_get_matches_from(["msoffice-crypto", "decrypt", "f.docx", "--output=x.docx"])
        .expect("--flag=value must parse");
    let sub = m.subcommand_matches("decrypt").expect("decrypt");
    assert_eq!(
        sub.get_one::<String>("output").map(String::as_str),
        Some("x.docx")
    );
}

#[test]
fn exactly_one_password_source_is_accepted() {
    // clap's ArgGroup enforces it, so this pins the wiring rather than the rule.
    let two = cli().try_get_matches_from([
        "msoffice-crypto",
        "decrypt",
        "f.docx",
        "--password-stdin",
        "--password-env",
        "PW",
    ]);
    assert!(two.is_err(), "two password sources must be rejected");

    for one in [
        vec!["msoffice-crypto", "decrypt", "f.docx", "--password-stdin"],
        vec![
            "msoffice-crypto",
            "decrypt",
            "f.docx",
            "--password-env",
            "PW",
        ],
        vec![
            "msoffice-crypto",
            "decrypt",
            "f.docx",
            "--password-file",
            "p",
        ],
        // None is legal: it means "prompt".
        vec!["msoffice-crypto", "decrypt", "f.docx"],
    ] {
        assert!(
            cli().try_get_matches_from(&one).is_ok(),
            "{one:?} must parse"
        );
    }
}

#[test]
fn the_password_argument_is_hidden_but_recognised() {
    // Plan §2. It parses -- so S3's `cmd_crypt` can explain *why* it does not exist,
    // rather than clap reporting a generic unexpected argument -- and it never appears
    // in help, so it is never offered.
    let m = cli()
        .try_get_matches_from([
            "msoffice-crypto",
            "decrypt",
            "f.docx",
            "--password",
            "secret",
        ])
        .expect("the trap argument must parse");
    let sub = m.subcommand_matches("decrypt").expect("decrypt");
    assert_eq!(
        sub.get_one::<String>(PASSWORD_TRAP).map(String::as_str),
        Some("secret")
    );
}

#[test]
fn no_help_output_ever_offers_a_password_value_argument() {
    // The predicate is "no line DEFINES the option", not "the string never appears":
    // the after-help text mentions `--password VALUE` precisely to say it does not
    // exist, and a substring check would fail on that prose.
    let mut cmd = cli();
    let mut texts = vec![cmd.render_long_help().to_string()];
    for name in ["classify", "decrypt", "encrypt"] {
        texts.push(
            cmd.find_subcommand_mut(name)
                .expect("subcommand")
                .render_long_help()
                .to_string(),
        );
    }
    for text in texts {
        for line in text.lines() {
            let t = line.trim_start();
            assert!(
                !(t.starts_with("--password ") || t.starts_with("--password=")),
                "a `--password <VALUE>` option must never be defined in help: {line:?}"
            );
        }
    }
}

#[test]
fn help_advertises_the_three_real_password_sources() {
    // The other half of the test above: that one proves the trap is absent, this one
    // proves the replacements are present. Plan § 2 asks the help text to name the
    // sources that do exist, and `exactly_one_password_source_is_accepted` cannot
    // stand in for it — it parses argv, where a `.hide(true)` on one of the three is
    // invisible. Measured: hiding `--password-stdin` fails this test and only this
    // test (17 passed, 1 failed).
    let mut cmd = cli();
    for name in ["decrypt", "encrypt"] {
        let text = cmd
            .find_subcommand_mut(name)
            .expect("subcommand")
            .render_long_help()
            .to_string();
        for flag in ["--password-env", "--password-file", "--password-stdin"] {
            assert!(text.contains(flag), "{name} help must offer {flag}");
        }
    }
}

#[test]
fn help_names_the_environment_variable_it_will_not_read() {
    // Plan §2 asks the help to name the obvious variable AND to say it is never read
    // implicitly; the behavioural half of that promise is
    // `the_named_environment_variable_is_never_read_unless_it_is_named` in tests/cli.rs,
    // and the pair is what makes it checkable from both sides.
    let mut cmd = cli();
    for name in ["decrypt", "encrypt"] {
        let text = cmd
            .find_subcommand_mut(name)
            .expect("subcommand")
            .render_long_help()
            .to_string();
        assert!(
            text.contains("MSOFFICE_CRYPTO_PASSWORD"),
            "{name} help must name the variable `--password-env` is for"
        );
        assert!(
            text.contains("naming it is the only way this tool reads it"),
            "{name} help must say the variable is not read implicitly"
        );
    }
}

#[test]
fn the_help_denies_the_password_flag_in_the_prose_which_is_why_the_check_is_line_based() {
    // This is the assertion that makes `no_help_output_ever_offers_a_password_value_argument`'s
    // line-definition predicate necessary rather than fussy. The prose names `--password`
    // precisely in order to deny it, and `--password-env` contains it as a substring
    // besides, so a naive `!text.contains("--password")` check would either fail always
    // or force the denial out of the help. "Simplifying" that test this way finds it
    // contradicting this one in the same file.
    let mut cmd = cli();
    for name in ["decrypt", "encrypt"] {
        let text = cmd
            .find_subcommand_mut(name)
            .expect("subcommand")
            .render_long_help()
            .to_string();
        assert!(
            text.contains("`--password"),
            "{name} help must deny `--password` by name in the PASSWORDS prose"
        );
        assert!(text.contains("world-readable"), "{name} help must say why");
    }
}

#[test]
fn the_password_prompt_verb_differs_by_direction() {
    // The prompt is the one password path no subprocess test can drive -- it needs a
    // terminal, and every test in tests/cli.rs redirects stdin -- so the verb is pinned
    // directly here. Without it, `encrypt`'s prompt could say "Password:" for a file that
    // has none and nothing would notice.
    assert_eq!(Direction::Decrypt.prompt_verb(), "Password");
    assert_eq!(Direction::Encrypt.prompt_verb(), "New password");
}

#[test]
fn the_enum_spellings_are_lower_kebab_and_stable() {
    // The greppable contract. A rename here is a breaking change for every script
    // parsing this output, so it is pinned rather than left to the renderer.
    assert_eq!(container_name(Container::Cfb), "cfb");
    assert_eq!(container_name(Container::Zip), "zip");
    assert_eq!(container_name(Container::Unknown), "unknown");

    assert_eq!(document_name(Document::OoxmlPackage), "ooxml-package");
    assert_eq!(document_name(Document::ZipArchive), "zip-archive");
    assert_eq!(document_name(Document::WordBinary), "word-binary");
    assert_eq!(document_name(Document::ExcelBinary), "excel-binary");
    assert_eq!(
        document_name(Document::PowerPointBinary),
        "powerpoint-binary"
    );
    assert_eq!(document_name(Document::Unknown), "unknown");

    assert_eq!(family_name(Family::Unencrypted), "unencrypted");
    assert_eq!(family_name(Family::Agile), "agile");
    assert_eq!(family_name(Family::Standard), "standard");
    assert_eq!(family_name(Family::Rc4CryptoApi), "rc4-cryptoapi");
    assert_eq!(family_name(Family::Rc4), "rc4");
    assert_eq!(family_name(Family::XorObfuscation), "xor-obfuscation");
    assert_eq!(family_name(Family::Unsupported), "unsupported");
    assert_eq!(family_name(Family::Unknown), "unknown");

    assert_eq!(integrity_name(IntegrityDeclaration::Declared), "declared");
    assert_eq!(
        integrity_name(IntegrityDeclaration::Incomplete),
        "incomplete"
    );
    assert_eq!(integrity_name(IntegrityDeclaration::Absent), "absent");
    assert_eq!(
        integrity_name(IntegrityDeclaration::NotApplicable),
        "not-applicable"
    );
    assert_eq!(integrity_name(IntegrityDeclaration::Unknown), "unknown");

    // Not kebab, deliberately: these are the spellings the spec and the file use.
    assert_eq!(cipher_name(CipherAlgorithm::Aes), "AES");
    assert_eq!(cipher_name(CipherAlgorithm::Rc4), "RC4");
    assert_eq!(hash_name(HashAlgorithm::Sha1), "SHA-1");
    assert_eq!(hash_name(HashAlgorithm::Sha256), "SHA-256");
    assert_eq!(hash_name(HashAlgorithm::Sha384), "SHA-384");
    assert_eq!(hash_name(HashAlgorithm::Sha512), "SHA-512");

    for (p, n) in [
        (IntegrityPolicy::Require, "require"),
        (
            IntegrityPolicy::RequireWhereDefined,
            "require-where-defined",
        ),
        (IntegrityPolicy::VerifyIfPresent, "verify-if-present"),
        (IntegrityPolicy::Skip, "skip"),
    ] {
        assert_eq!(policy_name(p), n);
    }
}

#[test]
fn human_output_answers_for_bytes_that_are_not_an_office_file() {
    // No fixture: `classify` answers every input, so sixteen bytes of junk is a legal
    // `Classification` and the only one this module can obtain (the struct is
    // `#[non_exhaustive]`, so it cannot be built by hand from here).
    let text = classification_human(&classify(b"not an office file"));
    assert_eq!(
        text,
        "container:      unknown\n\
         container-read: not-attempted\n\
         document:       unknown\n\
         family:         unknown\n\
         encrypted:      no\n\
         supported:      no\n\
         integrity:      unknown\n"
    );
    // A `None` field is omitted entirely, never printed as an empty value or a dash.
    assert!(!text.contains("version:"));
    assert!(!text.contains("key-"));
    assert!(!text.contains("pw-"));
}

#[test]
fn the_parameter_block_prints_six_fields_and_omits_the_absent_ones() {
    let p = params(|p| {
        p.cipher = Some(CipherAlgorithm::Aes);
        p.hash = Some(HashAlgorithm::Sha512);
        p.key_bits = Some(256);
        // block_size and salt_size left None on purpose: they must print nothing.
        p.spin_count = Some(100_000);
    });
    assert_eq!(
        params_lines("key", &p),
        "key-cipher:     AES\n\
         key-hash:       SHA-512\n\
         key-bits:       256\n\
         key-spin:       100000\n"
    );
    // The same six-field loop under the other prefix -- one function, run twice.
    assert_eq!(
        params_lines("pw", &params(|p| p.block_size = Some(16))),
        "pw-block:       16\n"
    );
    // An all-absent block prints nothing at all rather than six empty lines.
    assert_eq!(params_lines("key", &params(|_| {})), "");
}

#[test]
fn the_json_object_carries_the_full_key_set_for_an_input_that_classifies_as_nothing() {
    // No fixture: sixteen bytes of junk is a legal `Classification` (T4).
    let text = classification_json(&classify(b"not an office file"));
    let v: Value = serde_json::from_str(&text).expect("valid JSON");
    let o = v.as_object().expect("top level is an object");
    let mut keys: Vec<&str> = o.keys().map(String::as_str).collect();
    keys.sort_unstable();
    assert_eq!(
        keys,
        [
            "container",
            "container_read",
            "data_integrity",
            "document",
            "encrypted",
            "family",
            "key_data",
            "password_key",
            "supported",
            "version",
        ]
    );
    assert!(o["version"].is_null());
    assert!(o["key_data"].is_null());
    assert!(o["password_key"].is_null());
}

#[test]
fn a_parameter_block_keeps_all_six_json_keys_when_the_fields_are_absent() {
    // The direct JSON mirror of `the_parameter_block_prints_six_fields_and_omits_the_absent_ones`
    // above: the pair is what pins the human/JSON asymmetry as deliberate rather than an
    // oversight in one renderer.
    let v = params_json(&params(|_| {}));
    let o = v.as_object().expect("object");
    let mut keys: Vec<&str> = o.keys().map(String::as_str).collect();
    keys.sort_unstable();
    assert_eq!(
        keys,
        [
            "block_size",
            "cipher",
            "hash",
            "key_bits",
            "salt_size",
            "spin_count"
        ]
    );
    for k in keys {
        assert!(o[k].is_null(), "{k} must be null, not absent");
    }
}

#[test]
fn a_parameter_block_renders_the_fields_it_has_and_nulls_the_rest() {
    let p = params(|p| {
        p.cipher = Some(CipherAlgorithm::Aes);
        p.hash = Some(HashAlgorithm::Sha512);
        p.key_bits = Some(256);
        // block_size and salt_size left None on purpose: they must render `null`.
    });
    let v = params_json(&p);
    assert_eq!(v["cipher"], "AES");
    assert_eq!(v["hash"], "SHA-512");
    // A JSON number, not a quoted string -- `assert_eq!` against an integer literal
    // fails if `key_bits` were serialised as `"256"`.
    assert_eq!(v["key_bits"], 256);
    assert!(v["block_size"].is_null());
    assert!(v["salt_size"].is_null());
    assert!(v["spin_count"].is_null());
}

#[test]
fn the_json_flag_is_off_by_default_and_exists_only_on_classify() {
    let default = cli()
        .try_get_matches_from(["msoffice-crypto", "classify", "f.docx"])
        .expect("classify must parse with no --json");
    assert!(
        !default
            .subcommand_matches("classify")
            .expect("classify")
            .get_flag("json"),
        "--json must default to off"
    );

    let on = cli()
        .try_get_matches_from(["msoffice-crypto", "classify", "f.docx", "--json"])
        .expect("--json must parse on classify");
    assert!(on
        .subcommand_matches("classify")
        .expect("classify")
        .get_flag("json"));

    // `decrypt` and `encrypt` do not define the flag at all -- clap rejects it as an
    // unknown argument, not merely leaves it false.
    for sub in ["decrypt", "encrypt"] {
        assert!(
            cli()
                .try_get_matches_from(["msoffice-crypto", sub, "f.docx", "--json"])
                .is_err(),
            "{sub} must not accept --json"
        );
    }
}

// --- S4: decrypt dispatch, --integrity, and the atomic write --------------

/// A directory under the system temp dir, created here and removed on drop. Not a
/// fixture read (T4): nothing in it comes from tests/fixtures/.
struct UnitScratch(PathBuf);
impl UnitScratch {
    fn new(tag: &str) -> Self {
        let d =
            std::env::temp_dir().join(format!("msoffice-crypto-unit-{tag}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&d);
        std::fs::create_dir_all(&d).expect("scratch");
        UnitScratch(d)
    }
    fn entries(&self) -> Vec<String> {
        let mut v: Vec<String> = std::fs::read_dir(&self.0)
            .expect("read_dir")
            .map(|e| e.expect("entry").file_name().to_string_lossy().into_owned())
            .collect();
        v.sort();
        v
    }
}
impl Drop for UnitScratch {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

#[test]
fn the_policy_table_round_trips_and_the_default_is_in_it() {
    for p in [
        IntegrityPolicy::Require,
        IntegrityPolicy::RequireWhereDefined,
        IntegrityPolicy::VerifyIfPresent,
        IntegrityPolicy::Skip,
    ] {
        assert_eq!(parse_policy(policy_name(p)), Some(p));
    }
    // The claim S4 exists to pin: the rendered default parses back to the library's.
    assert_eq!(
        parse_policy(policy_name(IntegrityPolicy::default())),
        Some(IntegrityPolicy::default())
    );
    assert_eq!(parse_policy("unrecognised"), None);
}

#[test]
fn the_help_renders_the_librarys_default_not_a_literal() {
    // Computed from the library, never typed. If `.default_value(...)` were the literal
    // that happens to be right today and the library default moved again (it did once,
    // GH #12), the literal would stay and this assertion -- which follows the library --
    // would fail naming both spellings.
    let mut cmd = cli();
    let text = cmd
        .find_subcommand_mut("decrypt")
        .expect("decrypt")
        .render_long_help()
        .to_string();
    let want = format!("[default: {}]", policy_name(IntegrityPolicy::default()));
    assert!(
        text.contains(&want),
        "decrypt --help must show {want:?}, got:\n{text}"
    );
    for n in POLICY_NAMES {
        assert!(text.contains(n), "help must list {n}");
    }
}

#[test]
fn outcome_names_are_lower_kebab() {
    assert_eq!(outcome_name(IntegrityOutcome::Verified), "verified");
    assert_eq!(outcome_name(IntegrityOutcome::NotDeclared), "not-declared");
    assert_eq!(
        outcome_name(IntegrityOutcome::NotApplicable),
        "not-applicable"
    );
    assert_eq!(outcome_name(IntegrityOutcome::Skipped), "skipped");
}

#[test]
fn direction_suffixes_are_the_plans() {
    assert_eq!(Direction::Decrypt.suffix(), "decrypted");
    assert_eq!(Direction::Encrypt.suffix(), "encrypted");
}

#[test]
fn the_format_table_round_trips_over_both_names() {
    // Not tautological: `format_name` maps a variant to a slot of `FORMAT_NAMES` and
    // `parse_format` maps the slots back independently, so swapping either one's arms
    // breaks the round trip here. It says nothing about clap's default -- that claim
    // belongs to `the_encrypt_default_parses_back_to_agile`, which reads `cli()`.
    for f in [Format::Agile, Format::Standard] {
        assert_eq!(parse_format(format_name(f)), Some(f));
    }
    assert_eq!(parse_format("unrecognised"), None);
}

#[test]
fn the_encrypt_default_parses_back_to_agile() {
    let mut cmd = cli();
    let sub = cmd.find_subcommand_mut("encrypt").expect("encrypt");
    // Read what clap actually stored, then push it back through the table. Computing
    // the expectation from the same `format_name(Format::Agile)` call the source passes
    // to `.default_value(..)` would be tautological -- swapped arms in `format_name`
    // would wire `[default: standard]` and the assertion would move with the bug.
    let arg = sub
        .get_arguments()
        .find(|a| a.get_id() == "format")
        .expect("--format");
    let defaults = arg.get_default_values();
    assert_eq!(defaults.len(), 1, "one default, got {defaults:?}");
    let rendered = defaults[0].to_str().expect("utf-8 default").to_string();
    assert_eq!(
        parse_format(&rendered),
        Some(Format::Agile),
        "encrypt --format defaults to {rendered:?}, which is not agile"
    );

    // And it reaches the help, where the user reads it.
    let text = sub.render_long_help().to_string();
    let want = format!("[default: {rendered}]");
    assert!(
        text.contains(&want),
        "encrypt --help must show {want:?}, got:\n{text}"
    );
    for n in FORMAT_NAMES {
        assert!(text.contains(n), "help must list {n}");
    }
}

#[test]
fn the_encrypt_help_says_why_standard_exists_rather_than_leaving_it_to_folklore() {
    let mut cmd = cli();
    let text = cmd
        .find_subcommand_mut("encrypt")
        .expect("encrypt")
        .render_long_help()
        .to_string();
    // Whitespace-flattened because clap rewraps help to its own width, so a phrase can
    // cross a line break in the rendering without crossing one in the source string.
    let flat = text.split_whitespace().collect::<Vec<_>>().join(" ");
    // Phrases, not the bare words "dataIntegrity"/"Office 2007"/"Office 2010": all three
    // of those appear in the one-line help this slice replaced
    // ("agile (Office 2010+, has a dataIntegrity HMAC) or standard (Office 2007)"), so a
    // test built on them passes against the text it is supposed to have superseded and
    // guards nothing (CLAUDE.md, Evidence over intent).
    for token in [
        "refused rather than silently opened",
        "predates agile",
        "defines no integrity element at all",
    ] {
        assert!(
            flat.contains(token),
            "encrypt help must explain the cost of standard ({token:?} missing), got: {flat}"
        );
    }
    // `decrypt` offers `--integrity`; `encrypt` must not -- agile always writes the
    // element and standard cannot, so the flag would be a lie in one direction and a
    // no-op in the other.
    assert!(
        !text.contains("--integrity"),
        "encrypt must not offer --integrity: agile always writes the element and \
         standard cannot, so the flag would be a lie in one direction and a no-op in \
         the other"
    );
}

#[test]
fn the_integrity_word_for_each_format_is_the_librarys_declaration() {
    assert_eq!(
        integrity_name(declared_integrity(Format::Agile)),
        "declared"
    );
    assert_eq!(
        integrity_name(declared_integrity(Format::Standard)),
        "not-applicable"
    );
    // The claim is about the library's own enum, not about two strings.
    assert_eq!(
        declared_integrity(Format::Agile),
        IntegrityDeclaration::Declared
    );
}

#[test]
fn junk_routes_to_not_office_with_the_unknown_container_wording() {
    // Fixture-free: sixteen bytes of junk is a legal Classification (T4).
    let r = route_for(&classify(b"sixteen bytes!!!"), IntegrityPolicy::default());
    let Err(Refusal { code, why }) = r else {
        panic!("junk must be refused, got {r:?}")
    };
    assert_eq!(code, EX_NOT_OFFICE);
    // The wording distinguishes this check from the Document::Unknown one below: with
    // the container check deleted, junk still exits 3 but says "CFB container".
    assert!(why.contains("container: unknown"), "{why}");
}

#[test]
fn a_plain_zip_routes_to_refused_with_the_shared_not_encrypted_sentence() {
    // `is_zip` needs four bytes: `PK\x03\x04` classifies as Zip / ZipArchive /
    // Unencrypted with no fixture at all -- `ZipArchive` rather than `OoxmlPackage`
    // because nothing inside the archive is read.
    assert_eq!(
        route_for(&classify(b"PK\x03\x04"), IntegrityPolicy::default()),
        Err(Refusal {
            code: EX_REFUSED,
            why: NOT_ENCRYPTED.to_string()
        })
    );
}

#[test]
fn a_cfb_that_is_not_office_routes_to_not_office() {
    // The eight magic bytes and nothing else: `is_cfb_office` is true, the container
    // cannot be opened, the binary probe finds nothing -> Cfb / Document::Unknown.
    let magic = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let class = classify(&magic);
    assert_eq!(class.container, Container::Cfb);
    assert_eq!(class.document, Document::Unknown);
    let r = route_for(&class, IntegrityPolicy::default());
    let Err(Refusal { code, why }) = r else {
        panic!("got {r:?}")
    };
    assert_eq!(code, EX_NOT_OFFICE);
    assert!(why.contains("CFB container"), "{why}");
}

#[test]
fn a_cfb_input_is_refused_at_five_without_naming_a_document_kind() {
    // The same eight CFB magic bytes as `a_cfb_that_is_not_office_routes_to_not_office`:
    // Cfb / Document::Unknown / not encrypted.
    // The guard moved into the library, so this is now about how the CLI renders the
    // variant it raises rather than about the classification that raised it. The
    // sentence and the code are unchanged, which is the point.
    let why = describe(&Error::NotAPlainPackage);
    assert_eq!(exit_code(&Error::NotAPlainPackage), EX_REFUSED);
    assert!(why.contains("CFB container"), "{why}");
    assert!(why.contains("OOXML package"), "{why}");
    // The load-bearing negative: `Document::Unknown` for these bytes must not be
    // interpolated into a claim about "a 97-2003 document" or "unknown".
    assert!(
        !why.contains("97-2003 document") && !why.contains("unknown"),
        "the message named a document kind classify refused to name: {why}"
    );
}

#[test]
fn encrypt_and_decrypt_answer_a_bare_cfb_with_different_codes() {
    // decrypt asks "can I decrypt this" (3, not a document it knows); encrypt asks "is
    // this a plain package" (5, nothing to do). Aligning them loses one of the two
    // facts -- CONTRACT §5 / plan §7 call this deliberate.
    let magic = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let class = classify(&magic);
    let decrypt_code = route_for(&class, IntegrityPolicy::default())
        .expect_err("a bare CFB is not decryptable")
        .code;
    let encrypt_code = exit_code(&Error::NotAPlainPackage);
    assert_eq!(decrypt_code, EX_NOT_OFFICE);
    assert_eq!(encrypt_code, EX_REFUSED);
    assert_ne!(
        decrypt_code, encrypt_code,
        "encrypt refuses a CFB because it is not a package (5, nothing to do); decrypt \
         refuses it because it is not a document it knows (3). Aligning them loses one \
         of the two facts"
    );
}

#[test]
fn junk_is_refused_by_encrypt_with_the_same_sentence_decrypt_uses() {
    let why_encrypt = describe(&Error::UnknownContainer);
    let why_decrypt = route_for(&classify(b"sixteen bytes!!!"), IntegrityPolicy::default())
        .expect_err("junk is not decryptable")
        .why;
    assert_eq!(why_encrypt, NOT_OFFICE);
    assert_eq!(why_decrypt, why_encrypt, "one fact must have one sentence");
}

#[cfg(feature = "legacy-binary")]
#[test]
fn the_not_encrypted_sentence_is_shared_with_the_library_variant() {
    assert_eq!(describe(&Error::NotEncrypted), NOT_ENCRYPTED);
}

#[test]
fn integrity_failures_are_worded_by_the_cli_not_forwarded() {
    let missing = describe(&Error::IntegrityElementMissing);
    assert!(
        missing.contains("--integrity verify-if-present"),
        "{missing}"
    );
    assert!(
        !missing.contains("IntegrityPolicy::"),
        "a Rust path reached the user: {missing}"
    );
    let failed = describe(&Error::IntegrityCheckFailed);
    assert!(failed.contains("modified after"), "{failed}");
    assert!(
        !failed.contains("--integrity"),
        "there is no opt-out for a failed MAC: {failed}"
    );
    let unavailable = describe(&Error::IntegrityUnavailable("x"));
    assert!(
        unavailable.contains("--integrity require-where-defined"),
        "{unavailable}"
    );
    // Forwarded variants keep the library's words.
    assert_eq!(
        describe(&Error::WrongPassword),
        Error::WrongPassword.to_string()
    );
}

#[test]
fn temp_path_lives_beside_the_target() {
    let t = Path::new("/a/b/report.docx");
    let tmp = temp_path_for(t);
    assert_eq!(tmp, PathBuf::from("/a/b/.report.docx.msoffice-crypto.tmp"));
    assert_eq!(
        tmp.parent(),
        t.parent(),
        "same directory, or the rename is not atomic"
    );
    assert_eq!(
        temp_path_for(Path::new("out.docx")),
        PathBuf::from(".out.docx.msoffice-crypto.tmp")
    );
}

#[test]
fn write_atomically_leaves_no_temp_and_the_target_holds_the_bytes() {
    let s = UnitScratch::new("atomic-ok");
    let target = s.0.join("out.docx");
    write_atomically(&target, b"NEW BYTES").expect("write");
    assert_eq!(std::fs::read(&target).expect("read"), b"NEW BYTES");
    assert_eq!(
        s.entries(),
        vec!["out.docx".to_string()],
        "a temp file survived a successful write"
    );
}

#[test]
fn write_atomically_never_opens_the_target_itself() {
    // THE INTERRUPTED-RUN PROOF. The target holds ORIGINAL; the temp path is blocked by a
    // directory so `File::create(tmp)` fails. A correct implementation returns Err and
    // ORIGINAL is byte-identical: the target was never opened. An implementation that
    // writes the target directly succeeds here, and that is the half-written-.docx bug --
    // a kill between its first write and its last leaves a file that looks complete.
    let s = UnitScratch::new("atomic-target-untouched");
    let target = s.0.join("out.docx");
    std::fs::write(&target, b"ORIGINAL").expect("seed");
    std::fs::create_dir(temp_path_for(&target)).expect("block the temp path");
    assert!(
        write_atomically(&target, b"NEW").is_err(),
        "the blocked temp must fail the write"
    );
    assert_eq!(
        std::fs::read(&target).expect("read"),
        b"ORIGINAL",
        "the target was opened for writing"
    );
}

#[test]
fn write_atomically_cleans_the_temp_when_the_rename_fails() {
    // The rename fails (the target is a directory) AFTER the temp was fully written.
    let s = UnitScratch::new("atomic-rename-fails");
    let target = s.0.join("taken");
    std::fs::create_dir(&target).expect("dir");
    std::fs::write(target.join("marker"), b"m").expect("marker");
    assert!(write_atomically(&target, b"NEW").is_err());
    assert_eq!(std::fs::read(target.join("marker")).expect("marker"), b"m");
    assert_eq!(
        s.entries(),
        vec!["taken".to_string()],
        "temp left behind after a failed rename"
    );
}
