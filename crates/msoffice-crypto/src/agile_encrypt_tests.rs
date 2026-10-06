//! The generated encryptor, put through the real verification path.
//!
//! The load-bearing test here is [`the_generated_encryptor_verifies_through_the_real_path`]:
//! serialise the material with `encryption_info::write`, parse it back with
//! `agile::parse_encryption_info`, and hand the result to `agile::verify_password` — the
//! same three functions a decrypt runs, in the same order, with nothing re-derived for the
//! test's convenience. A generator that agreed with a bespoke check written beside it would
//! prove nothing; this one has to satisfy the code that opens files.
//!
//! What is *not* proven here, and waits for #8: that another implementation accepts it.
//! Steps 5 and 6 assemble a whole file, and `office-crypto` and `msoffcrypto-tool` are the
//! oracles that make the claim external. Everything below is self-consistency plus a
//! committed golden, which is what step 4 can honestly offer.

use super::*;
use crate::agile;
use crate::encryption_info::{self, EncryptionInfoParams};
use rand::SeedableRng;

const PASSWORD: &str = "testpass";
const SPIN_COUNT: u32 = 100_000;

/// The tuple every golden below was measured under: `EncryptParams::default`, which is
/// the measured Office 16 tuple, with `SPIN_COUNT` stated rather than inherited so that
/// a change to the default's spin count shows up here as a moved golden and not as a
/// silently different test.
///
/// **A golden that moves means the default draw path moved**, and the change is wrong
/// until proven otherwise — re-measuring is the last thing to do, not the first.
fn params() -> EncryptParams {
    EncryptParams {
        spin_count: SPIN_COUNT,
        ..Default::default()
    }
}

/// Every `(hash, keyBits)` pair this crate will write, stated as a table rather than
/// computed.
///
/// Ten, not twelve. `keyBits / 8` must fit inside the named hash's digest — §2.3.4.11
/// truncates and this crate refuses to `0x36`-pad the shortfall — so SHA-1's 20 bytes
/// admit 128 alone, and `(Sha1, 192)` and `(Sha1, 256)` are unwritable.
///
/// **Typed literally, never derived from `can_carry_key_bits`.** A table computed from
/// the predicate it is meant to pin would prove only that the code agrees with itself;
/// this one disagrees with the code if either moves. The length assertion is what stops
/// the sweep silently shrinking — a row lost to an edit would otherwise just make the
/// test faster.
const EVERY_WRITABLE_TUPLE: [(HashAlgorithm, u32); 10] = [
    (HashAlgorithm::Sha1, 128),
    (HashAlgorithm::Sha256, 128),
    (HashAlgorithm::Sha256, 192),
    (HashAlgorithm::Sha256, 256),
    (HashAlgorithm::Sha384, 128),
    (HashAlgorithm::Sha384, 192),
    (HashAlgorithm::Sha384, 256),
    (HashAlgorithm::Sha512, 128),
    (HashAlgorithm::Sha512, 192),
    (HashAlgorithm::Sha512, 256),
];

/// The seed every golden below is measured under. All-zero, so it is obviously arbitrary
/// and obviously not chosen to make an assertion pass.
const SEED: [u8; 32] = [0u8; 32];

fn seeded() -> chacha20::ChaCha12Rng {
    chacha20::ChaCha12Rng::from_seed(SEED)
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

/// Byte-exact output under a seeded RNG — plan D3's payoff.
///
/// Pins two things nothing else does. The **draw order**: password salt, verifier input,
/// session key, keyData salt, which is herumi's (`encode.hpp:167, 174, 180, 187`) and
/// which nothing in the format enforces — reorder two draws and every value below changes
/// while every other test in this file still passes. And the **derivation**: any change to
/// the spin hash, the block keys or the CBC calls moves these bytes.
///
/// Measured, not computed by hand. If a deliberate change moves them, re-measure with the
/// same seed and say in the commit what moved and why.
#[test]
fn the_material_is_byte_exact_under_a_seeded_rng() {
    let material = generate(PASSWORD, params(), &mut seeded()).unwrap();

    assert_eq!(
        hex(&material.password_salt),
        "9bf49a6a0755f953811fce125f2683d5"
    );
    assert_eq!(
        hex(&material.key_data_salt),
        "0bd58841203e74fe86fc71338ce0173d"
    );
    assert_eq!(
        hex(&material.encrypted_verifier_hash_input),
        "0fe715e6275524574c2bc39c92ebec0b"
    );
    assert_eq!(
        hex(&material.encrypted_verifier_hash_value),
        "7984ae30716d67e29d21eaf75dd4e7d980c9565e89ba7f2d11e31cd74fc5c51f\
         53ae214d9799a2557393e8b758e5400b8f08135e84cb24a42917056a874c10f6"
    );
    assert_eq!(
        hex(&material.encrypted_key_value),
        "5c667ede00de124d64f6295629a3b9ec57521a0095fecd3c375b675aa3f7ba9a"
    );
    assert_eq!(
        material.session_key.with_secret(|k| hex(k)),
        "0564f879d27ae3c02ce82834acfa8c793a629f2ca0de6919610be82f411326be"
    );
}

/// Dummy `dataIntegrity` blobs, so step 3's writer can be driven before step 5 exists.
///
/// The writer requires them to be `roundUp(hashSize, blockSize)` = 64 bytes and does not
/// otherwise care; nothing in this file verifies an HMAC. Generating them for real is
/// step 5.
const NO_INTEGRITY_YET: [u8; 64] = [0u8; 64];

/// **Write, parse, verify — the three functions a decrypt actually runs.**
///
/// This is the test the module exists to pass. `generate` produces the encryptor,
/// `encryption_info::write` serialises it, `agile::parse_encryption_info` reads it back
/// from bytes, and `agile::verify_password` — the function `decrypt` calls, unmodified —
/// accepts the password. Nothing is re-derived here for the test's benefit, so a generator
/// that disagreed with the reader in any of the four places they must agree (salt, spin
/// hash, block key, CBC IV) fails.
///
/// The wrong-password case is the control. Without it this would pass on a
/// `verify_password` that accepted everything, which is the shape of the bug that makes a
/// verifier worthless.
#[test]
fn the_generated_encryptor_verifies_through_the_real_path() {
    let material = generate(PASSWORD, params(), &mut seeded()).unwrap();

    let stream = encryption_info::write(&EncryptionInfoParams {
        params: params(),
        key_data_salt: &material.key_data_salt,
        encrypted_hmac_key: &NO_INTEGRITY_YET,
        encrypted_hmac_value: &NO_INTEGRITY_YET,
        password_salt: &material.password_salt,
        encrypted_verifier_hash_input: &material.encrypted_verifier_hash_input,
        encrypted_verifier_hash_value: &material.encrypted_verifier_hash_value,
        encrypted_key_value: &material.encrypted_key_value,
    })
    .expect("generated material must satisfy the writer's own length rules");

    let parsed = agile::parse_encryption_info(&stream[8..])
        .expect("the writer's output must parse -- that is step 3's contract");

    let h_final = agile::spin_hash(
        HashAlgorithm::Sha512,
        PASSWORD,
        &material.password_salt,
        SPIN_COUNT,
    );
    agile::verify_password(&parsed, &h_final)
        .expect("the password used to generate the encryptor must verify against it");

    // The control.
    let wrong = agile::spin_hash(
        HashAlgorithm::Sha512,
        "not the password",
        &material.password_salt,
        SPIN_COUNT,
    );
    assert!(
        matches!(
            agile::verify_password(&parsed, &wrong),
            Err(Error::WrongPassword)
        ),
        "a wrong password must be refused, and by name"
    );
}

/// Re-derive the three block keys from the password and unwrap all three blobs.
///
/// The decrypt-side view of what `generate` produced, using `agile`'s own
/// `derive_block_key` and `aes_cbc_decrypt`: the verifier pair must satisfy
/// `SHA512(input) == value`, which is the relation `verify_password` checks, and
/// `encryptedKeyValue` must unwrap to exactly the session key the material carries.
///
/// The second half is what the test above cannot see. `verify_password` never touches
/// `encryptedKeyValue` — a generator that wrapped the session key under the *wrong* block
/// key would pass every password check and produce a file whose package decrypts to noise,
/// with no `<dataIntegrity>` element obliged to say so.
#[test]
fn every_blob_unwraps_to_what_produced_it() {
    let material = generate(PASSWORD, params(), &mut seeded()).unwrap();
    let h_final = agile::spin_hash(
        HashAlgorithm::Sha512,
        PASSWORD,
        &material.password_salt,
        SPIN_COUNT,
    );

    let unwrap_with = |block_key: &[u8; 8], blob: &[u8]| -> Vec<u8> {
        let key = agile::derive_block_key(
            HashAlgorithm::Sha512,
            &h_final,
            block_key,
            // `p:encryptedKey/@keyBits` -- the password half. `key_data_key_bits` is
            // the same 256 at the default tuple, and using it here would be the silent
            // crossing the two fields exist to make visible.
            params().password_key_bits,
        )
        .expect("SHA-512 yields more than 32 bytes");
        key.with_secret(|k| agile::aes_cbc_decrypt(blob, k, &material.password_salt))
            .expect("the blob was encrypted under this key and IV")
    };

    let verifier_input = unwrap_with(
        &agile::BLOCK_VERIFIER_INPUT,
        &material.encrypted_verifier_hash_input,
    );
    let verifier_value = unwrap_with(
        &agile::BLOCK_VERIFIER_HASH,
        &material.encrypted_verifier_hash_value,
    );
    assert_eq!(
        HashAlgorithm::Sha512.digest(&verifier_input),
        verifier_value,
        "SHA512(verifierHashInput) must equal the decrypted verifierHashValue"
    );

    let key_value = unwrap_with(&agile::BLOCK_KEY_VALUE, &material.encrypted_key_value);
    assert_eq!(
        material.session_key.with_secret(|k| k.to_vec()),
        key_value,
        "encryptedKeyValue must unwrap to the session key it was made from"
    );

    // And the three block keys are genuinely three: wrapping every blob under one key
    // would satisfy neither assertion above by accident, but it is cheap to say outright.
    let keys: Vec<Vec<u8>> = [
        agile::BLOCK_VERIFIER_INPUT,
        agile::BLOCK_VERIFIER_HASH,
        agile::BLOCK_KEY_VALUE,
    ]
    .iter()
    .map(|bk| {
        agile::derive_block_key(
            HashAlgorithm::Sha512,
            &h_final,
            bk,
            params().password_key_bits,
        )
        .unwrap()
        .with_secret(|k| k.to_vec())
    })
    .collect();
    let mut unique = keys.clone();
    unique.sort();
    unique.dedup();
    assert_eq!(unique.len(), 3, "the three block keys must differ");
}

/// **The session key is 32 drawn bytes, never a short draw padded to length.**
///
/// The reference this module was ported from once drew `encryptedKey.saltSize` bytes and
/// `normalizeKey`d up to `keyBits / 8` — `resize(32, 0x36)`, an AES-256 key with 128 bits of
/// entropy and a constant top half. Nothing downstream noticed: the file decrypts perfectly
/// in every reader, because the key is whatever the writer says it is, so no round-trip test
/// in any implementation could find it. It was reported privately and **fixed upstream on
/// 2026-09-10** (herumi/msoffice `b5fed299`); see `agile_encrypt`'s module header.
///
/// The test stays, because the property is worth guarding on its own: a short draw padded to
/// length is a mistake any implementation can make, and this is the assertion that catches
/// it here. The last 16 bytes must not be a constant filler, and — the stronger form — every
/// byte position must vary across seeds, which a constant tail cannot do.
#[test]
fn the_session_key_is_fully_drawn_never_padded_to_length() {
    let material = generate(PASSWORD, params(), &mut seeded()).unwrap();
    material.session_key.with_secret(|k| {
        assert_eq!(k.len(), 32, "AES-256 takes 32 bytes");
        assert_ne!(
            &k[16..],
            &[0x36u8; 16],
            "the top half must be entropy, not a constant filler"
        );
    });

    // Every byte position varies across seeds. A padded tail is constant at 16 positions,
    // so this fails on herumi's construction however the filler byte is chosen.
    let mut varies = [false; 32];
    let first = material.session_key.with_secret(|k| k.to_vec());
    for seed in 1u8..24 {
        let other = generate(
            PASSWORD,
            // A low spin count: this loop is about the RNG, not the KDF.
            EncryptParams {
                spin_count: 1,
                ..params()
            },
            &mut chacha20::ChaCha12Rng::from_seed([seed; 32]),
        )
        .unwrap();
        other.session_key.with_secret(|k| {
            for i in 0..32 {
                varies[i] |= k[i] != first[i];
            }
        });
    }
    assert!(
        varies.iter().all(|v| *v),
        "some byte of the session key never changed across 23 seeds: {varies:?}"
    );
}

/// The same seed reproduces everything; a different seed reproduces nothing.
///
/// The first half is what makes the golden above meaningful, and the second is what says
/// the generator is actually consuming the RNG rather than deriving from the password
/// alone — which would compile, pass the verification test, and hand every document
/// written with one password the same salts and the same session key.
#[test]
fn the_rng_decides_the_output_and_the_seed_decides_the_rng() {
    let a = generate(PASSWORD, params(), &mut seeded()).unwrap();
    let b = generate(PASSWORD, params(), &mut seeded()).unwrap();
    assert_eq!(a.password_salt, b.password_salt);
    assert_eq!(a.key_data_salt, b.key_data_salt);
    assert_eq!(a.encrypted_key_value, b.encrypted_key_value);
    assert_eq!(
        a.session_key.with_secret(|k| k.to_vec()),
        b.session_key.with_secret(|k| k.to_vec())
    );

    let c = generate(
        PASSWORD,
        params(),
        &mut chacha20::ChaCha12Rng::from_seed([1u8; 32]),
    )
    .unwrap();
    assert_ne!(a.password_salt, c.password_salt);
    assert_ne!(a.key_data_salt, c.key_data_salt);
    assert_ne!(
        a.encrypted_verifier_hash_input,
        c.encrypted_verifier_hash_input
    );
    assert_ne!(
        a.session_key.with_secret(|k| k.to_vec()),
        c.session_key.with_secret(|k| k.to_vec())
    );
}

/// The public entry point works and is not accidentally deterministic.
///
/// `crate::encrypt_ooxml` is [`encrypt`] plus the shape guard and the system RNG, which
/// is the point — but "the seeded path and the real path are the same code" is a claim,
/// and an untested wrapper is where that claim would quietly stop being true. Two calls
/// must differ; if they ever match, the system RNG is not being consumed. The four
/// bytes of ZIP magic below are what carries the input past the guard.
#[test]
fn the_public_entry_point_produces_a_fresh_file_each_call() {
    let a = crate::encrypt_ooxml(b"PK\x03\x04 not really a zip", PASSWORD).unwrap();
    let b = crate::encrypt_ooxml(b"PK\x03\x04 not really a zip", PASSWORD).unwrap();
    assert_ne!(
        a, b,
        "two encryptions of the same input must not share salts or keys"
    );
    assert_eq!(a.len(), b.len(), "but they are the same shape");
}

/// The whole write path under a seeded RNG is byte-stable, and this is its golden.
///
/// Plan D3's payoff at the level that matters: not the key schedule alone but the
/// complete container — `EncryptionInfo`, `EncryptedPackage`, the four `DataSpaces`
/// blobs and the CFB directory `cfb` builds around them, which has no timestamps and no
/// random identifiers. A change to any step of the assembly moves this digest.
///
/// Measured, not computed. If a deliberate change moves it, re-measure with the same seed
/// and say in the commit what moved and why. The plaintext is `plain.docx`, the file the
/// standard fixture decrypts to, so the round-trip tests in `lib.rs` and this golden are
/// about the same bytes.
#[test]
fn the_whole_container_is_byte_exact_under_a_seeded_rng() {
    use sha2::{Digest, Sha256};
    let plain = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/plain.docx"
    ))
    .expect("plain.docx is committed");
    let container = encrypt(&plain, PASSWORD, params(), &mut seeded()).unwrap();
    let digest = hex(&Sha256::digest(&container));
    assert_eq!(
        (container.len(), digest.as_str()),
        (GOLDEN_LEN, GOLDEN_SHA256),
        "the container's (length, SHA-256) under the seeded RNG"
    );

    // And the same seed reproduces it exactly; a golden that only held once is a fluke.
    let again = encrypt(&plain, PASSWORD, params(), &mut seeded()).unwrap();
    assert_eq!(container, again);
}

/// Measured on 2026-09-05 under `SEED`, `PASSWORD`, `SPIN_COUNT` and `plain.docx`.
const GOLDEN_LEN: usize = 41984;
const GOLDEN_SHA256: &str = "b4cc009e94bddf00c6610e0b897e5ee0a4485bd5bc9e40a024131d9e789a2a6b";

/// Every generated blob is the length step 3's writer demands.
///
/// Stated directly as well as through the writer, because the writer's refusal names a
/// field and this names the rule: `roundUp(p:encryptedKey saltSize, blockSize)` for the
/// verifier input, `roundUp(hashSize, blockSize)` for its hash, and
/// `roundUp(keyData keyBits / 8, blockSize)` for the wrapped key.
///
/// The default tuple first, then three that pull the seven lengths apart. At
/// AES-256/SHA-512 with 16-byte salts almost every rule coincides -- 16, 64 and 32 are
/// each produced by more than one formula -- so the default alone cannot tell a correct
/// derivation from a lucky one.
#[test]
fn every_generated_length_is_the_one_the_format_fixes() {
    let m = generate(PASSWORD, params(), &mut seeded()).unwrap();
    assert_eq!(m.password_salt.len(), 16);
    assert_eq!(m.key_data_salt.len(), 16);
    assert_eq!(m.encrypted_verifier_hash_input.len(), 16); // roundUp(16, 16)
    assert_eq!(m.encrypted_verifier_hash_value.len(), 64); // roundUp(64, 16)
    assert_eq!(m.encrypted_key_value.len(), 32); // roundUp(256 / 8, 16)
    m.session_key.with_secret(|k| assert_eq!(k.len(), 32));

    // The two salts are drawn independently. Equal salts would make the crossing bug
    // `AgileParams` is shaped to prevent invisible in every test that uses this material.
    assert_ne!(m.password_salt, m.key_data_salt);

    // Now the tuples where the formulas separate. Each row is the tuple and then
    // (password salt, keyData salt, verifier input blob, verifier hash blob, wrapped key
    // blob, session key).
    for (p, want) in [
        // Different salt sizes on the two elements, neither a block multiple: the
        // verifier input blob follows the **password** salt and nothing else, and a
        // generator reading `key_data_salt_size` there would produce 32 rather than 16.
        (
            EncryptParams {
                password_salt_size: 9,
                key_data_salt_size: 24,
                ..params()
            },
            (9, 24, 16, 64, 32, 32),
        ),
        // AES-192 on `<keyData>` and AES-256 on `<p:encryptedKey>`: a 24-byte session key
        // in a 32-byte blob. `keyBits / 8` alone would make that blob 24, and
        // `agile.rs:1160-1170` would refuse the file this crate had just written.
        (
            EncryptParams {
                key_data_key_bits: 192,
                password_key_bits: 256,
                ..params()
            },
            (16, 16, 16, 64, 32, 24),
        ),
        // SHA-1, where `roundUp(hashSize, blockSize)` is 32 rather than the digest's 20 --
        // the case the default tuple cannot see, 64 already being a multiple of 16.
        (
            EncryptParams {
                hash: HashAlgorithm::Sha1,
                key_data_key_bits: 128,
                password_key_bits: 128,
                ..params()
            },
            (16, 16, 16, 32, 16, 16),
        ),
    ] {
        let m = generate(PASSWORD, p, &mut seeded()).unwrap();
        let got = (
            m.password_salt.len(),
            m.key_data_salt.len(),
            m.encrypted_verifier_hash_input.len(),
            m.encrypted_verifier_hash_value.len(),
            m.encrypted_key_value.len(),
            m.session_key.with_secret(|k| k.len()),
        );
        assert_eq!(got, want, "{p:?}");
    }
}

/// **A tuple the caller chose survives the round trip, through the public reader.**
///
/// The point of threading `EncryptParams` this far down: `encrypt` writes it,
/// `decrypt_ooxml_with_policy` reads the file back under the fail-closed default, and the
/// `dataIntegrity` HMAC -- computed over the ciphertext under `<keyData>`'s hash and salt
/// -- has to verify. A writer that sized any of the seven blobs from the wrong element
/// fails here, at the reader, instead of producing a file only this crate could open.
///
/// Three of these are cases the default tuple hides:
///
/// * **`saltSize` 8 and 40** -- both sides of `hash::fit_iv`. The three password blobs
///   take the salt as their CBC IV, and [MS-OFFCRYPTO] §2.3.4.12's third step pads a short
///   one with `0x36` and truncates a long one. Until the `fit_iv` call landed in
///   `generate` this side handed AES the raw salt: 8 bytes was an IV-length refusal out of
///   `check_cbc_lengths`, and 40 bytes was encrypted under an IV the reader -- which has
///   always gone through `AgileParams::password_blob_iv` -- does not reproduce. Neither is
///   reachable at the default 16, where `fit_iv` is the identity, which is why nothing
///   caught it.
/// * **AES-192 on `<keyData>`** -- 24 bytes of key in a 32-byte `encryptedKeyValue`.
/// * **SHA-1** -- a 20-byte digest in a 32-byte verifier blob, and a 32-byte
///   `encryptedHmacKey`, which is §3.11's worked example rather than §2.3.4.14 step 2's
///   `saltSize`.
///
/// The wrong password is the control on every row: a round trip alone would pass on a
/// verifier that accepted anything.
///
/// The hand-picked rows above are joined by [`EVERY_WRITABLE_TUPLE`], so the sweep is
/// exhaustive over the `(hash, keyBits)` space rather than a selection from it. That
/// matters because the interesting cases are the ones nobody would think to pick: the
/// first version of this test covered two of the ten, and neither was SHA-256.
#[test]
fn a_caller_chosen_tuple_round_trips_through_the_public_reader() {
    let plain = plain_docx();
    for p in [
        params(),
        EncryptParams {
            password_salt_size: 8,
            ..params()
        },
        EncryptParams {
            password_salt_size: 40,
            ..params()
        },
        EncryptParams {
            key_data_salt_size: 24,
            password_salt_size: 9,
            ..params()
        },
        EncryptParams {
            key_data_key_bits: 192,
            ..params()
        },
        EncryptParams {
            hash: HashAlgorithm::Sha1,
            key_data_key_bits: 128,
            password_key_bits: 128,
            ..params()
        },
        EncryptParams {
            hash: HashAlgorithm::Sha384,
            key_data_key_bits: 192,
            password_key_bits: 256,
            key_data_salt_size: 20,
            password_salt_size: 24,
            ..params()
        },
    ]
    .into_iter()
    .chain(EVERY_WRITABLE_TUPLE.iter().map(|&(hash, key_bits)| {
        EncryptParams {
            hash,
            key_data_key_bits: key_bits,
            password_key_bits: key_bits,
            // A cheap spin count, and the reason it is safe to make it cheap: what this
            // sweep is for is the hash/keyBits coupling and the seven lengths it moves,
            // none of which `spinCount` participates in -- it is a loop count over
            // `H_final` and the rows would be identical at 100 000. Ten tuples times
            // three key derivations at Office's default is about a minute of SHA for no
            // additional fact. The default spin count is round-tripped by every other
            // test in this file, including both goldens.
            spin_count: 4,
            ..Default::default()
        }
    })) {
        let container = encrypt(&plain, PASSWORD, p, &mut seeded())
            .unwrap_or_else(|e| panic!("{p:?} must be writable: {e}"));
        let crate::Decrypted {
            package: back,
            integrity: outcome,
        } = crate::decrypt_ooxml_with_policy(
            &container,
            PASSWORD,
            crate::IntegrityPolicy::default(),
        )
        .unwrap_or_else(|e| panic!("{p:?} must read back: {e}"));
        assert_eq!(outcome, crate::IntegrityOutcome::Verified, "{p:?}");
        assert_eq!(back, plain, "{p:?}");

        assert!(
            matches!(
                crate::decrypt_ooxml(&container, "not the password"),
                Err(Error::WrongPassword)
            ),
            "{p:?}"
        );
    }
}

/// The plaintext behind `encryptedKeyValue` is the session key **zero-padded** to the
/// blob -- not the key alone, and not a `0x36` pad.
///
/// AES-192 is the only supported tuple where the three are distinguishable: 24 bytes of
/// key in a 32-byte blob leave 8 bytes something has to choose. Zero is what
/// `tools/gen_agile_fixtures.py:63-78` writes and what the thirteen measured variants at
/// `integrity.rs:453-466` carry. `0x36` is the plausible wrong answer -- it is the pad
/// byte two sentences earlier in §2.3.4.12, and herumi's `normalizeKey` uses it -- and no
/// round-trip test could tell the two apart, because this crate's reader ignores the tail.
///
/// Unwrapped through `agile`'s own `derive_block_key` and `aes_cbc_decrypt`, so this is
/// the reader's view of the blob rather than a restatement of what the writer did.
#[test]
fn the_wrapped_key_blob_is_the_session_key_with_a_zero_tail() {
    let p = EncryptParams {
        key_data_key_bits: 192,
        ..params()
    };
    let m = generate(PASSWORD, p, &mut seeded()).unwrap();
    assert_eq!(m.encrypted_key_value.len(), 32, "roundUp(192 / 8, 16)");

    let h_final = agile::spin_hash(p.hash, PASSWORD, &m.password_salt, p.spin_count);
    let key = agile::derive_block_key(
        p.hash,
        &h_final,
        &agile::BLOCK_KEY_VALUE,
        p.password_key_bits,
    )
    .expect("SHA-512 carries 256 bits");
    let plaintext = key
        .with_secret(|k| {
            agile::aes_cbc_decrypt(
                &m.encrypted_key_value,
                k,
                &crate::hash::fit_iv(&m.password_salt, 16),
            )
        })
        .expect("the blob was wrapped under this key and IV");

    assert_eq!(plaintext.len(), 32);
    m.session_key
        .with_secret(|sk| assert_eq!(&plaintext[..24], sk, "the first keyBits / 8 bytes"));
    assert_eq!(&plaintext[24..], &[0u8; 8], "and a zero tail, not 0x36");
}

/// `encryptedVerifierHashValue` is the hash of the **`saltSize` random bytes**, not of the
/// padded blob they travel in.
///
/// [MS-OFFCRYPTO] §2.3.4.13 step 1: "Obtain the hash value of the random array of bytes
/// generated in step 1 of the steps for encryptedVerifierHashInput", and that array is
/// `saltSize` bytes. The `0x00` pad to a block multiple is step 3, applied when the array
/// is encrypted — after the hash.
///
/// **This test exists because no other test in the crate could have caught the bug, and
/// that is the interesting part.** Until 2026-09-21 the writer hashed the padded buffer
/// and `agile::verify_password` digested the whole decrypted blob to match. Two halves
/// agreeing with each other and disagreeing with the format: every round trip here
/// passed, including the ten-tuple sweep at `password_salt_size` 8, 9, 24 and 40. Real
/// Word 16 refused those files with `0x800A1520`, the wrong-password code, on the right
/// password — which is what the artifact set and a human opening it exist to find.
///
/// So this asserts against a digest computed **in the test**, never against the crate's
/// own reader. A test that decrypted the blob and re-digested it the way
/// `verify_password` does would have agreed with the bug.
///
/// `saltSize = 9` because the two lengths must differ: at any multiple of 16 the padded
/// and unpadded arrays are the same bytes and the assertion cannot fail.
#[test]
fn the_verifier_hash_covers_the_salt_sized_array_and_not_its_padding() {
    let p = EncryptParams {
        password_salt_size: 9,
        ..params()
    };
    let m = generate(PASSWORD, p, &mut seeded()).unwrap();
    assert_eq!(m.encrypted_verifier_hash_input.len(), 16, "roundUp(9, 16)");

    let h_final = agile::spin_hash(p.hash, PASSWORD, &m.password_salt, p.spin_count);
    let blob_iv = crate::hash::fit_iv(&m.password_salt, 16);
    let unwrap = |block_key: &[u8; 8], blob: &[u8]| {
        agile::derive_block_key(p.hash, &h_final, block_key, p.password_key_bits)
            .expect("SHA-512 carries 256 bits")
            .with_secret(|k| agile::aes_cbc_decrypt(blob, k, &blob_iv))
            .expect("wrapped under this key and IV")
    };

    let input = unwrap(
        &agile::BLOCK_VERIFIER_INPUT,
        &m.encrypted_verifier_hash_input,
    );
    let value = unwrap(
        &agile::BLOCK_VERIFIER_HASH,
        &m.encrypted_verifier_hash_value,
    );

    assert_eq!(&input[9..], &[0u8; 7], "the pad is 0x00, per step 3");

    // Computed here, from the spec's sentence, against the file's own bytes.
    let want = p.hash.digest(&input[..9]);
    assert_eq!(&value[..want.len()], &want[..], "H(the saltSize bytes)");

    // And the negative that names the old behaviour, so a revert fails by meaning rather
    // than by a number nobody can read.
    let padded = p.hash.digest(&input[..]);
    assert_ne!(
        &value[..padded.len()],
        &padded[..],
        "must NOT be H(the padded blob) -- that is the bug Word refused"
    );
}

// ---- the public entry point, end to end (GH #6 step 6) --------------------------------

fn plain_docx() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/plain.docx"
    ))
    .expect("plain.docx is committed")
}

/// Flip one byte of `/EncryptedPackage` inside a container this crate wrote, deep in the
/// ciphertext body — segment 1, so past the first 4096-byte segment and its IV — leaving
/// the header, the XML and the container untouched, so a refusal is the HMAC's and not a
/// parse failure's.
fn tampered(container: Vec<u8>) -> Vec<u8> {
    use std::io::{Read, Seek, SeekFrom, Write};
    let mut cursor = std::io::Cursor::new(container);
    {
        let mut cfb = cfb::CompoundFile::open(&mut cursor).expect("we wrote a CFB");
        let mut stream = cfb
            .open_stream("/EncryptedPackage")
            .expect("we wrote the stream");
        let at = 8 + 4096 + 100;
        stream.seek(SeekFrom::Start(at)).unwrap();
        let mut b = [0u8; 1];
        stream.read_exact(&mut b).unwrap();
        stream.seek(SeekFrom::Start(at)).unwrap();
        stream.write_all(&[b[0] ^ 0x01]).unwrap();
        stream.flush().unwrap();
        cfb.flush().unwrap();
    }
    cursor.into_inner()
}

/// What this crate writes, this crate reads back — under the **fail-closed default**, so
/// the `dataIntegrity` element it wrote is required and verified, not merely tolerated —
/// and a wrong password is a wrong password.
#[test]
fn encrypt_ooxml_round_trips_under_the_fail_closed_default() {
    let plain = plain_docx();
    let container = crate::encrypt_ooxml(&plain, PASSWORD).expect("encrypt");
    assert!(crate::is_cfb_office(&container));

    let crate::Decrypted {
        package: back,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(&container, PASSWORD, crate::IntegrityPolicy::default())
        .expect("what this crate wrote, it must read");
    assert_eq!(outcome, crate::IntegrityOutcome::Verified);
    assert_eq!(
        back, plain,
        "the package must survive the round trip byte for byte"
    );

    assert!(matches!(
        crate::decrypt_ooxml(&container, "not the password"),
        Err(Error::WrongPassword)
    ));
}

/// The HMAC the writer produces is one the reader enforces: a flipped ciphertext byte is
/// refused by name under the default, and the `Skip` control decrypts the same file to
/// the wrong bytes — so the refusal is the tag's work and not some unrelated check.
#[test]
fn a_tampered_encrypt_ooxml_output_is_refused_by_the_hmac_it_wrote() {
    let plain = plain_docx();
    let bad = tampered(crate::encrypt_ooxml(&plain, PASSWORD).unwrap());
    assert!(matches!(
        crate::decrypt_ooxml(&bad, PASSWORD),
        Err(Error::IntegrityCheckFailed)
    ));
    let crate::Decrypted {
        package: wrong,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(&bad, PASSWORD, crate::IntegrityPolicy::Skip)
        .expect("Skip returns unauthenticated bytes");
    assert_eq!(outcome, crate::IntegrityOutcome::Skipped);
    assert_ne!(wrong, plain, "the flipped byte must change the plaintext");
}

/// `classify` sees exactly the file Office 16 would have written: agile, AES-256/SHA-512,
/// 100 000 rounds, `dataIntegrity` declared. The detection half of the crate and the
/// encrypt half agree about what a file is.
#[test]
fn encrypt_ooxml_output_classifies_as_the_tuple_office_writes() {
    let c = crate::classify(&crate::encrypt_ooxml(&plain_docx(), PASSWORD).unwrap());
    assert_eq!(c.container, crate::Container::Cfb);
    assert_eq!(c.document, crate::Document::OoxmlPackage);
    assert_eq!(c.family, crate::Family::Agile);
    assert_eq!(c.data_integrity, crate::IntegrityDeclaration::Declared);
    assert!(c.is_encrypted() && c.is_supported());
    let key_data = c.key_data.expect("agile declares <keyData>");
    assert_eq!(key_data.hash, Some(crate::HashAlgorithm::Sha512));
    assert_eq!(key_data.key_bits, Some(256));
    let password = c.password_key.expect("and a password key encryptor");
    assert_eq!(password.spin_count, Some(100_000));
}

/// The degenerate package. Zero segments, an 8-byte `EncryptedPackage` of nothing but
/// its prefix, an HMAC over those 8 bytes — and it round-trips to an empty `Vec`.
///
/// Driven through the seeded core rather than [`crate::encrypt_ooxml`], because the
/// public entry point now refuses an input that is not a plain package and `&[]` is not
/// one — that refusal is `bytes_that_are_no_container_are_refused` in `lib.rs`. The two
/// facts are separate: whether the guard turns an empty input away, and whether the
/// writer handles a degenerate payload correctly if it ever reaches it. This is the
/// second, and nothing else covers the container-and-HMAC half of it.
#[test]
fn an_empty_package_round_trips() {
    let container = encrypt(&[], PASSWORD, params(), &mut seeded()).unwrap();
    let crate::Decrypted {
        package: back,
        integrity: outcome,
    } = crate::decrypt_ooxml_with_policy(&container, PASSWORD, crate::IntegrityPolicy::default())
        .unwrap();
    assert_eq!(outcome, crate::IntegrityOutcome::Verified);
    assert!(back.is_empty());
}

/// **An independent implementation reads a file this crate generated** — key material,
/// `EncryptionInfo`, `EncryptedPackage`, `dataIntegrity`, container, all ours.
///
/// Until this test every "another implementation reads it" claim in the crate was about a
/// container filled with streams lifted from a working fixture. `office-crypto` is a
/// separate MIT crate with its own CFB reader and its own parse of [MS-OFFCRYPTO];
/// agreeing with it is evidence about the file. It is one of GH #8's four readers; the
/// other three are run on the artifact the last test below writes, and recorded.
#[test]
fn an_independent_implementation_reads_what_encrypt_ooxml_wrote() {
    let plain = plain_docx();
    let container = crate::encrypt_ooxml(&plain, PASSWORD).unwrap();
    let theirs = office_crypto::decrypt_from_bytes(container, PASSWORD)
        .expect("office-crypto must open a file this crate wrote");
    assert_eq!(theirs, plain);
}

/// A package over the ceiling is refused before any key is derived or byte encrypted:
/// the same 1 GiB the decrypt side refuses, checked on the way in so that a file this
/// crate writes is a file it can read back. The allocation is lazy on every platform
/// this runs on, so the test costs an inequality, not a gigabyte.
///
/// The input must be a **plain ZIP** to reach the ceiling at all: the shape guard runs
/// first, so a gigabyte of zeros is now `UnknownContainer` rather than a size refusal.
/// That ordering has its own test, `the_shape_guard_runs_before_the_payload_ceiling`,
/// and this one is its negative control — a real oversized package still answers by
/// name.
#[test]
fn a_package_over_the_ceiling_is_refused_before_any_work() {
    // `vec![0u8; N]` is `alloc_zeroed`, so these are lazy zero pages. Writing the magic
    // in place touches one of them; `resize` from a 4-byte `Vec` would memset a
    // gigabyte and make the comment above false.
    let mut big = vec![0u8; crate::limits::PAYLOAD_CEILING + 1];
    big[..4].copy_from_slice(b"PK\x03\x04");
    let got = crate::encrypt_ooxml(&big, PASSWORD).map(|c| c.len());
    assert!(
        matches!(&got, Err(Error::BadParameters(msg)) if msg.contains("PAYLOAD_CEILING")),
        "over the ceiling must be refused by name, got: {got:?}"
    );
}

/// The shape guard runs before the payload ceiling, and the order is the claim.
///
/// A gigabyte of zeros is over the ceiling *and* not a package. Both facts are true, and
/// which one the caller is told decides whether the message is useful: "these bytes are
/// not a package" is the primary one, and `BadParameters(… PAYLOAD_CEILING …)` for junk
/// would be the worse answer. Reverse the two checks in `crate::encrypt_ooxml` and this
/// fails with the ceiling message; `a_package_over_the_ceiling_is_refused_before_any_work`
/// is the control that shows the ceiling is still reachable for input that is a package.
#[test]
fn the_shape_guard_runs_before_the_payload_ceiling() {
    let big = vec![0u8; crate::limits::PAYLOAD_CEILING + 1];
    let got = crate::encrypt_ooxml(&big, PASSWORD).map(|c| c.len());
    assert!(
        matches!(&got, Err(Error::UnknownContainer)),
        "oversized junk is not a package first and oversized second, got: {got:?}"
    );
}

/// Writes the artifact GH #8's external readers are run against — real Word over COM,
/// LibreOffice over UNO, `msoffcrypto-tool` and `office-crypto` — and prints where.
/// `cargo test` cannot drive the first two; `tools/acceptance_gate.py` drives all four
/// and their verdicts are recorded in `CHANGELOG.md` against this exact file, the same
/// way the container spike's were.
///
/// The directory is `MSOFFICE_CRYPTO_ARTIFACT_DIR` when set, else the system temp
/// directory. Set it, and to a private directory: the file name is fixed, so every
/// session's `cargo test` on this machine writes the same path, and a gate pointed there
/// can measure another build's artifact — that happened during #8, and is why the
/// evidence in `CHANGELOG.md` is recorded against a SHA-256 rather than a path. CI's gate
/// job sets it for the same reason. The digest printed below is the one the gate prints
/// on its first line: if they differ, the gate did not read this run's file.
#[test]
fn encrypt_ooxml_writes_the_artifact_the_external_readers_are_run_on() {
    let dir = std::env::var_os("MSOFFICE_CRYPTO_ARTIFACT_DIR")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(std::env::temp_dir);
    let path = dir.join("msoffice_crypto_encrypt_ooxml.docx");
    std::fs::write(
        &path,
        crate::encrypt_ooxml(&plain_docx(), PASSWORD).unwrap(),
    )
    .unwrap();
    // Digested from the file rather than from the buffer that was written -- see the note
    // on the standard sibling. Both writers say this the same way on purpose; the last
    // time these two drifted, the one without the mitigation was the one that changed.
    let written = std::fs::read(&path).unwrap();
    let digest = hex(&crate::hash::HashAlgorithm::Sha256.digest(&written));
    println!(
        "encrypt_ooxml artifact written to {} ({} bytes, SHA-256 {digest})",
        path.display(),
        written.len()
    );
}

// ---- plan slice 10: the durable double-click artifact set --------------------------

/// `tests/fixtures/{name}`, read or panic — the same contract `plain_docx` already keeps,
/// generalised so this section can also reach the two real-Office `.xlsx`/`.pptx`
/// fixtures without inventing a second helper for one caller.
fn fixture_bytes(name: &str) -> Vec<u8> {
    std::fs::read(format!(
        "{}/tests/fixtures/{name}",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap_or_else(|e| panic!("fixture {name} is committed: {e}"))
}

/// The plaintext behind `plain_content.txt` — the one prose fact a human checking
/// `agile_sha512_256_default.docx` and its siblings by hand can compare against without
/// opening a second window.
const PLAIN_DOCX_TEXT: &str = "Hello, encrypted Office world!";

/// Plan slice 10, the part this machine can do. Writes the durable artifact set the plan's
/// § *Item 7* describes — one file per writable `(hash, keyBits)` tuple, three salt sizes
/// at the default tuple, the three standard `AlgID`s, and `.xlsx`/`.pptx` at the default
/// agile tuple — into a directory the owner can double-click through, plus a
/// `MANIFEST.md` that carries per-file provenance and the two-case procedure for reading a
/// Word refusal. This is *not* the automated acceptance gate (`tools/acceptance_gate.py`
/// runs over this directory separately) and it is not `MSOFFICE_CRYPTO_ARTIFACT_DIR`
/// (a `mktemp` directory the automation consumes and forgets, per that env var's own
/// doc comment above): the whole point of this directory is that it survives past the
/// test run, at a name a human can find again.
///
/// **Named for the date this slice was executed, not computed from today's clock** — this
/// machine reads local time (UTC-7) and CI reads UTC, so a `chrono::Local::today()` string
/// would print a different directory name depending on which one generated it. Matches the
/// plan's own `artifacts/tuples-2026-09-20/`.
///
/// **The default agile file is asserted byte-identical to the seeded golden**, which is
/// what turns the owner's manual open of this one file into a regression check as well as
/// an interop check: if this assertion ever fails, the manual check below would have been
/// run against a file that no longer represents what the golden pins, silently.
///
/// Every agile file here is written under the fixed `seeded()` RNG rather than the public
/// `encrypt_ooxml_with_params` entry point's `SysRng` — the only way to make the default
/// file's bytes reproducible against the committed golden, and every sibling file inherits
/// the same determinism so a re-run of this test does not silently replace a file the owner
/// already opened with a different (still-conforming) one. The three standard files use
/// the public `encrypt_ooxml_standard_with_key_bits` and real randomness: no golden pins
/// them, and using the production entry point there is one fewer thing to keep in step
/// with `standard_encrypt`'s own seeded core.
#[test]
#[cfg_attr(
    not(fixture_corpus),
    ignore = "needs the fixture corpus, which the published crate does not ship"
)]
fn the_writable_tuple_matrix_is_written_to_the_durable_artifact_directory() {
    let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("artifacts/tuples-2026-09-20");
    std::fs::create_dir_all(&dir)
        .unwrap_or_else(|e| panic!("{} must be creatable: {e}", dir.display()));

    let mut manifest = String::new();
    manifest.push_str(
        "# Tuple artifacts — 2026-09-20\n\n\
         Generated by `the_writable_tuple_matrix_is_written_to_the_durable_artifact_directory` \
         (`src/agile_encrypt_tests.rs`), the deliverable of plan slice 10 in \
         `docs/plans/msoffice-crypto-encrypt-params-2026-09-20.md`. Every file below opens \
         with the password **`testpass`**.\n\n\
         ## Reading a refusal — the two-case procedure (plan Item 7)\n\n\
         The ultimate check is a human opening these files in real Word, Excel and \
         PowerPoint. Word is not one of the four automated readers for this purpose — it \
         is *the reference*, because Microsoft wrote both [MS-OFFCRYPTO] and this \
         implementation of it. A refusal from Word on a file this crate believes conforms \
         is never filed as \"a reader limitation\"; it is exactly one of two things, and \
         both are worked through rather than caveated:\n\n\
         1. **Our output is wrong** (the default assumption, and it blocks). Re-derive the \
         bytes from the cited [MS-OFFCRYPTO] clause for the field in question and check our \
         output against *that* — not against msoffcrypto-tool, LibreOffice, office-crypto \
         or herumi, all of which are interpretations of the same document and never the \
         standard. This is how the zero-pad rule (`0x800A1066`) and the verifier-value tail \
         (`0x800A1520`) were both found, in both cases as our bugs.\n\
         2. **Word diverges from its own published spec.** Possible, and then the finding \
         *is* the deliverable: record the exact build, tuple and `AlgID`/attribute, cite \
         the clause it contradicts, and bring it back for a decision. It does not silently \
         become a documented caveat.\n\n\
         A zero-pass row (no external reader, including Word, has opened a given tuple) is \
         a fact to record, not a reason to stop writing that tuple: the spec defines it, so \
         this crate implements it.\n\n\
         ## Files\n\n",
    );

    let mut record =
        |name: &str, tuple_words: &str, bytes: &[u8], expected: &str, meaning: &str| {
            let path = dir.join(name);
            std::fs::write(&path, bytes).unwrap_or_else(|e| panic!("{name}: {e}"));
            let digest = hex(&crate::hash::HashAlgorithm::Sha256.digest(bytes));
            manifest.push_str(&format!(
                "### `{name}`\n\n\
             - **Tuple:** {tuple_words}\n\
             - **Password:** `testpass`\n\
             - **SHA-256:** `{digest}`\n\
             - **Size:** {} bytes\n\
             - **Expected content:** {expected}\n\
             - **What a refusal would mean:** {meaning}\n\n",
                bytes.len(),
            ));
            digest
        };

    let plain = plain_docx();

    // The ten writable `(hash, keyBits)` combinations -- EVERY_WRITABLE_TUPLE is the
    // authoritative list (see its own doc comment: typed literally, not derived from
    // `can_carry_key_bits`, so this sweep disagrees with the code if either moves).
    for &(hash, key_bits) in EVERY_WRITABLE_TUPLE.iter() {
        let is_default = hash == HashAlgorithm::Sha512 && key_bits == 256;
        let p = EncryptParams {
            hash,
            key_data_key_bits: key_bits,
            password_key_bits: key_bits,
            ..params()
        };
        let container = encrypt(&plain, PASSWORD, p, &mut seeded())
            .unwrap_or_else(|e| panic!("{p:?} must be writable: {e}"));
        let name = if is_default {
            "agile_sha512_256_default.docx".to_string()
        } else {
            format!("agile_{}_{key_bits}.docx", hash.name().to_lowercase())
        };
        let words = format!(
            "agile (ECMA-376), hash {}, keyBits {key_bits}/{key_bits} \
             (keyData/p:encryptedKey), saltSize 16/16, spinCount {SPIN_COUNT}",
            hash.name()
        );
        let meaning = if is_default {
            "This is the exact tuple Office 16 itself writes. There is no plausible \
             excuse for a refusal here: work case 1 first, and only report case 2 if \
             re-deriving §2.3.4.10-.14 by hand still disagrees with what this file holds."
        } else {
            "One of the ten tuples [MS-OFFCRYPTO] §2.3.4.10 permits a writer to choose; \
             keyBits carries no upper bound at all. **But that same section's trailing \
             sentence -- \"values that are not defined MAY be used, and a compliant \
             implementation is not required to support all defined values\" -- is the \
             clause that authorises a READER to refuse.** So \"this reader implements a \
             subset\" is a conforming outcome here, not a defect, and there are three \
             shapes rather than two: the file broke the format, or a reader declines what \
             the format permits it to decline, or our bytes are wrong. Only the third is \
             ours. **Word is the exception** -- Microsoft wrote both the format and that \
             reader -- so a Word refusal still goes through the two-case procedure above."
        };
        let digest = record(&name, &words, &container, PLAIN_DOCX_TEXT, meaning);
        if is_default {
            assert_eq!(
                (container.len(), digest.as_str()),
                (GOLDEN_LEN, GOLDEN_SHA256),
                "the durable default artifact must be byte-identical to the seeded golden \
                 the rest of this file pins -- if this moved, the default write path moved \
                 and the manual check below would be exercising a different file than the \
                 one every other test in this crate measures"
            );
        }
    }

    // Three salt sizes at the default tuple, one a deliberate non-multiple of 16 -- the
    // case `fit_iv` exists for and that no external reader has been measured on (see
    // `EncryptParams::password_salt_size`'s doc comment).
    for salt in [8u32, 17, 32] {
        let p = EncryptParams {
            key_data_salt_size: salt,
            password_salt_size: salt,
            ..params()
        };
        let container = encrypt(&plain, PASSWORD, p, &mut seeded())
            .unwrap_or_else(|e| panic!("{p:?} must be writable: {e}"));
        let name = format!("agile_default_salt{salt}.docx");
        let words = format!(
            "agile (ECMA-376), hash SHA512, keyBits 256/256, saltSize {salt}/{salt} \
             (keyData/p:encryptedKey), spinCount {SPIN_COUNT} -- default tuple apart from \
             the salt"
        );
        let non_multiple = salt % 16 != 0;
        let meaning = if non_multiple {
            "saltSize is not a multiple of 16, the case [MS-OFFCRYPTO] \u{a7}2.3.4.12's \
             `fit_iv` pad/truncate step exists for -- and the case that has already caught \
             one defect. **Word refused both of these files before `ff11e3e`**, with \
             `0x800A1520` on the correct password, because this crate hashed the padded \
             verifier array where §2.3.4.13 hashes the saltSize bytes. Fixed, and both \
             re-opened in real Word afterwards, so a refusal here now is a NEW finding and \
             not the known one. LibreOffice refuses BOTH of these files, and for the \
             very defect `ff11e3e` fixed here: `AgileEngine.cxx:346,354` hashes the \
             verifier array with its padding. That is that reader declining, not our \
             bytes -- real Word opens both. Compare against the \
             multiple-of-16 salt file beside this one to isolate the dimension."
        } else {
            "saltSize is a multiple of 16 but not the Office default (16) -- a control for \
             the non-multiple file beside it. If this one opens and the non-multiple one \
             does not, the finding is specifically about the pad step, not about a \
             non-default saltSize generally."
        };
        record(&name, &words, &container, PLAIN_DOCX_TEXT, meaning);
    }

    // The default agile tuple over a spreadsheet and a deck, not just a document --
    // Excel and PowerPoint have each produced a distinct finding in this crate's history
    // (recorded in docs/design/development-record.md), so Word alone is not the test.
    // The plaintext is what this crate's own reader recovers from the real-Office
    // fixtures already committed, whose SHA-256 is independently pinned against
    // msoffcrypto-tool in tests/real_office_fixtures.rs::AGILE_GOLDENS -- so the content
    // fed in here is not this test's own invention. The expected text below is the marker
    // string each fixture actually carries -- read out of `xl/sharedStrings.xml` and
    // `ppt/slides/slide1.xml` by hand, not guessed -- so `tools/acceptance_gate.py`'s
    // Word/Excel/PowerPoint leg (which asserts the *recovered text*, not only that the
    // file opened) has something real to check on these two as well as on the `.docx`s.
    for (fixture_name, out_ext, app, marker) in [
        (
            "excel16_agile.xlsx",
            "xlsx",
            "Excel",
            "msoffice-crypto fixture: excel16 xlsx agile encryption password testpass",
        ),
        (
            "powerpoint16_agile.pptx",
            "pptx",
            "PowerPoint",
            "msoffice-crypto fixture: powerpoint16 pptx, agile format, password testpass",
        ),
    ] {
        let plain_app = crate::decrypt_ooxml(&fixture_bytes(fixture_name), PASSWORD)
            .unwrap_or_else(|e| panic!("{fixture_name} must decrypt: {e}"));
        let container = encrypt(&plain_app, PASSWORD, params(), &mut seeded())
            .unwrap_or_else(|e| panic!("{fixture_name}: the default tuple must be writable: {e}"));
        let name = format!("agile_sha512_256_default.{out_ext}");
        let words = format!(
            "agile (ECMA-376), hash SHA512, keyBits 256/256, saltSize 16/16, \
             spinCount {SPIN_COUNT} -- default tuple, over a {app} package"
        );
        let expected = format!(
            "the marker string \"{marker}\" (readable in {app} once opened), plus the rest \
             of `tests/fixtures/{fixture_name}`'s content -- that fixture's decrypted \
             SHA-256 is independently pinned against msoffcrypto-tool's own read in \
             `tests/real_office_fixtures.rs::AGILE_GOLDENS`, so identity beyond the marker \
             is already established and is not this test's own invention"
        );
        let meaning = format!(
            "the default agile tuple, applied to a {app} package rather than a Word one. \
             {app} has produced a distinct finding in this crate's history before (see \
             docs/design/development-record.md), so a Word pass above does not predict \
             this file -- apply the two-case procedure above independently. Gate with \
             `--expect-sha256` / `--expect-len` from `AGILE_GOLDENS` and `--expect-text \
             \"{marker}\"`, since no plaintext `.{out_ext}` twin of the decrypted package \
             is committed to diff against."
        );
        record(&name, &words, &container, &expected, &meaning);
    }

    // The three standard AlgIDs, now that the write half supports them. Through the
    // public entry point and real randomness: no golden pins these, so there is nothing
    // to keep reproducible, and using the production entry point is one fewer thing to
    // keep in step with `standard_encrypt`'s own seeded core.
    for bits in [128u32, 192, 256] {
        let container = crate::encrypt_ooxml_standard_with_key_bits(&plain, PASSWORD, bits)
            .unwrap_or_else(|e| panic!("standard AES-{bits} must be writable: {e}"));
        let name = format!("standard_aes{bits}.docx");
        let words = format!(
            "standard (ECMA-376/Office 2007), AES-{bits}, SHA-1 and 50 000 iterations \
             fixed by [MS-OFFCRYPTO] \u{a7}2.3.4.7 (not file fields)"
        );
        let meaning = if bits == 128 {
            "Office 2007 itself only ever writes AES-128 for this format, so this is the \
             one standard-encryption file with a real-Office precedent; a refusal here \
             follows the two-case procedure above with no caveat."
        } else {
            "AES-192/256 standard encryption: no shipping Office reader has been measured \
             at this size before (no fixture exists -- Office cannot produce one), though \
             msoffcrypto-tool reads it back byte-identical and this crate's own reader \
             round-trips it. A refusal is the first real-Office measurement at this size \
             and follows the two-case procedure above; §2.3.2 and §2.3.4.5 are the clauses \
             to re-derive from."
        };
        record(&name, &words, &container, PLAIN_DOCX_TEXT, meaning);
    }

    let manifest_path = dir.join("MANIFEST.md");
    std::fs::write(&manifest_path, &manifest)
        .unwrap_or_else(|e| panic!("{}: {e}", manifest_path.display()));
    println!(
        "wrote the tuple artifact set and MANIFEST.md to {}",
        dir.display()
    );
}
