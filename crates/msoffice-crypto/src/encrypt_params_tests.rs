//! What [`EncryptParams::validate`] accepts, stated as a table rather than as a formula.
//!
//! **Every expectation here is typed out.** None of them calls
//! `HashAlgorithm::can_carry_key_bits`, `digest_len`, or any other part of the code under
//! test to work out what the answer should be. A test that recomputes the implementation
//! proves only that the implementation agrees with itself — it stays green when the
//! predicate is inverted, because the expectation inverts with it. The twelve rows of the
//! sweep below are twelve facts about SHA and AES (FIPS 180-4 digest lengths against the
//! three AES key sizes), written down by hand, and the count of passing rows is asserted
//! separately so that the table cannot lose one quietly.
//!
//! The other tests take one `(param, problem)` cell of [`Error::EncryptParams`] each and
//! assert **both** halves of the payload. Asserting only `is_err()`, or only the `param`,
//! passes when the code refuses for an unrelated reason — CLAUDE.md § *Testing Rules*.
//! Each has a control beside it: a value one step inside the bound that must be `Ok`, so
//! that "refuses at the boundary" is distinguishable from "refuses everything".

use super::*;

/// The whole payload of an [`Error::EncryptParams`] refusal, or a named failure.
///
/// A helper rather than a `matches!` per test because `got`, `min` and `max` are asserted
/// too: the span a consumer renders to its user is part of the contract, and a refusal
/// that reports the wrong `got` sends someone to change a value they did not set.
fn refusal(params: &EncryptParams) -> (EncryptParam, EncryptParamProblem, u32, u32, u32) {
    match params.validate() {
        Err(Error::EncryptParams {
            param,
            problem,
            got,
            min,
            max,
        }) => (param, problem, got, min, max),
        Err(other) => panic!("expected an EncryptParams refusal, got {other:?}"),
        Ok(()) => panic!("expected a refusal, got Ok for {params:?}"),
    }
}

/// Hash × keyBits, all twelve combinations, with the verdict written out.
///
/// Both `keyBits` fields are set to the same value so that each row tests one number; the
/// two are pulled apart in [`the_coupling_binds_the_password_key_bits_only`] below, which
/// is where the per-element split is proved.
///
/// The three refusals are SHA-1's: its digest is 20 bytes, and 192- and 256-bit keys want
/// 24 and 32. [MS-OFFCRYPTO] §2.3.4.11 would have a writer 0x36-pad the difference; this
/// crate refuses, so these rows are `UnusableCombination` and not a spec violation.
const HASH_KEY_BITS_TABLE: [(HashAlgorithm, u32, bool); 12] = [
    (HashAlgorithm::Sha1, 128, true),
    (HashAlgorithm::Sha1, 192, false),
    (HashAlgorithm::Sha1, 256, false),
    (HashAlgorithm::Sha256, 128, true),
    (HashAlgorithm::Sha256, 192, true),
    (HashAlgorithm::Sha256, 256, true),
    (HashAlgorithm::Sha384, 128, true),
    (HashAlgorithm::Sha384, 192, true),
    (HashAlgorithm::Sha384, 256, true),
    (HashAlgorithm::Sha512, 128, true),
    (HashAlgorithm::Sha512, 192, true),
    (HashAlgorithm::Sha512, 256, true),
];

#[test]
fn every_hash_and_key_bits_combination_has_the_typed_verdict() {
    for (hash, key_bits, accepted) in HASH_KEY_BITS_TABLE {
        let params = EncryptParams {
            hash,
            key_data_key_bits: key_bits,
            password_key_bits: key_bits,
            ..Default::default()
        };
        if accepted {
            params
                .validate()
                .unwrap_or_else(|e| panic!("{}/{key_bits} must validate: {e}", hash.name()));
        } else {
            let (param, problem, got, _, _) = refusal(&params);
            assert_eq!(
                (param, problem, got),
                (
                    EncryptParam::KeyBitsWithHash,
                    EncryptParamProblem::UnusableCombination,
                    key_bits
                ),
                "{}/{key_bits} must be refused as the pair",
                hash.name()
            );
        }
    }
}

/// The table's own shape, so it cannot shrink or flip without a red test naming it.
///
/// Counting the literal rows rather than the validated ones on purpose: this asserts what
/// the *expectations* say, which is the thing a well-meaning edit ("SHA-1 with a 256-bit
/// key seems fine now") would change. The test above then asserts the code matches them.
#[test]
fn the_table_still_covers_twelve_combinations_and_admits_ten() {
    assert_eq!(HASH_KEY_BITS_TABLE.len(), 12, "4 hashes × 3 AES key sizes");
    let admitted = HASH_KEY_BITS_TABLE.iter().filter(|row| row.2).count();
    assert_eq!(admitted, 10, "only SHA-1 refuses, and only for 192 and 256");
}

/// Every `keyBits` refusal, with **whose** rule it broke and the exact payload.
///
/// This is the partition the audit found inverted, so it is written out one value at a
/// time rather than derived. `ST_KeyBits` ([MS-OFFCRYPTO] §2.3.4.10) is
/// `minInclusive="8"`, a multiple of 8, and **no maximum**; AES defines 128, 192 and 256.
/// A value can break the first, the second, or neither, and the two refusals are different
/// claims about the world:
///
/// * 0, 7, 12 — `OutsideSpecRange`: no conforming file may carry them. Labelling these
///   `UnsupportedByCipher` tells a caller the format was fine with 7, which it is not.
/// * 8, 64, 160, 512 — `UnsupportedByCipher`: every one is a legal `ST_KeyBits`, including
///   512, because the schema states no ceiling. Labelling these `OutsideSpecRange` tells a
///   caller the format forbids a 512-bit key, which it does not.
///
/// 8 is the floor itself, so it proves `minInclusive` is inclusive: it must reach the
/// cipher check rather than fail the spec one. The third column is the whole reported
/// span, asserted as one number because the payload is degenerate by construction — see
/// [`super::nearest_supported_key_bits`], and [`the_reported_span_never_names_a_refused_size`]
/// for the property that makes that necessary. 160 is the tie: 32 from 128 and 32 from
/// 192, resolving upwards.
const KEY_BITS_REFUSALS: [(u32, EncryptParamProblem, u32); 7] = [
    (0, EncryptParamProblem::OutsideSpecRange, 128),
    (7, EncryptParamProblem::OutsideSpecRange, 128),
    (12, EncryptParamProblem::OutsideSpecRange, 128),
    (8, EncryptParamProblem::UnsupportedByCipher, 128),
    (64, EncryptParamProblem::UnsupportedByCipher, 128),
    (160, EncryptParamProblem::UnsupportedByCipher, 192),
    (512, EncryptParamProblem::UnsupportedByCipher, 256),
];

/// The table above, asserted on **each** element in turn.
///
/// Both elements matter: `ST_KeyBits` types `keyData/@keyBits` and
/// `p:encryptedKey/@keyBits` alike, and a check written on one field only would leave the
/// other able to author a `keyBits="12"` file. The other field is held at 256 so that the
/// refusal under test is the only thing wrong with the struct.
///
/// The control is the default tuple, which must be `Ok`: without it the test cannot tell
/// "refuses 12" from "refuses everything".
#[test]
fn each_key_bits_refusal_names_whose_rule_was_broken() {
    for (offender, problem, span) in KEY_BITS_REFUSALS {
        for (key_data, password) in [(offender, 256), (256, offender)] {
            let params = EncryptParams {
                key_data_key_bits: key_data,
                password_key_bits: password,
                ..Default::default()
            };
            assert_eq!(
                refusal(&params),
                (EncryptParam::KeyBits, problem, offender, span, span),
                "keyBits {offender} on {}",
                if key_data == offender {
                    "keyData"
                } else {
                    "p:encryptedKey"
                }
            );
        }
    }

    EncryptParams::default()
        .validate()
        .expect("control: 256 on both elements is what Office writes");
}

/// The table's own shape: both problems are represented, and 512 is not a spec violation.
///
/// Asserted against the literal rows for the reason
/// [`the_table_still_covers_twelve_combinations_and_admits_ten`] gives — this is what the
/// *expectations* claim, and an edit that quietly relabelled 512 as `OutsideSpecRange`
/// (the exact defect the audit found) would turn this red by name rather than pass along
/// with the implementation.
#[test]
fn the_key_bits_table_partitions_the_two_rules() {
    let spec = KEY_BITS_REFUSALS
        .iter()
        .filter(|row| row.1 == EncryptParamProblem::OutsideSpecRange)
        .map(|row| row.0)
        .collect::<Vec<_>>();
    assert_eq!(spec, vec![0, 7, 12], "below 8, or not a multiple of 8");

    let cipher = KEY_BITS_REFUSALS
        .iter()
        .filter(|row| row.1 == EncryptParamProblem::UnsupportedByCipher)
        .map(|row| row.0)
        .collect::<Vec<_>>();
    assert_eq!(
        cipher,
        vec![8, 64, 160, 512],
        "legal ST_KeyBits values AES has no key for; 512 is legal because the schema \
         states no maximum"
    );
}

/// The span a refusal prints is never a value the next call would refuse.
///
/// The audit's second finding: `{128, 192, 256}` is a set, the message renders
/// `this crate accepts {min}..={max}`, and the ends of the set span 200 — which
/// [`an_invalid_key_size_outranks_the_pair_check_it_would_also_fail`] proves is refused.
/// This sweeps every multiple of 8 from 0 to 1024 and asserts that whatever a refusal
/// reports, both ends of it validate. It fails on the payload the audit flagged: restoring
/// `min: 128, max: 256` leaves `min` accepted and `max` accepted, so the assertion below
/// would still pass — which is why it also asserts the span is a single point, the
/// property that makes "both ends accepted" mean "everything the message names is
/// accepted".
#[test]
fn the_reported_span_never_names_a_refused_size() {
    for requested in (0..=1024).step_by(8) {
        let params = EncryptParams {
            password_key_bits: requested,
            ..Default::default()
        };
        let Err(Error::EncryptParams { min, max, .. }) = params.validate() else {
            continue; // 128, 192 and 256 are accepted; nothing to check.
        };
        assert_eq!(
            min, max,
            "keyBits {requested}: a set must not print as a span"
        );
        EncryptParams {
            password_key_bits: min,
            ..Default::default()
        }
        .validate()
        .unwrap_or_else(|e| {
            panic!("keyBits {requested} was refused with a span naming {min}, which is itself refused: {e}")
        });
    }
}

/// `(KeyBitsWithHash, UnusableCombination)` — both values legal alone, unusable together.
///
/// The span is `128..=128` and not `128..=256`: it names a size that *would* have worked
/// with this hash, so a caller reading it is not sent back to a value SHA-1 also cannot
/// carry — and it is a point rather than a range for the reason
/// [`the_reported_span_never_names_a_refused_size`] asserts globally.
#[test]
fn a_key_longer_than_the_digest_is_an_unusable_combination() {
    let params = EncryptParams {
        hash: HashAlgorithm::Sha1,
        password_key_bits: 256,
        ..Default::default()
    };
    assert_eq!(
        refusal(&params),
        (
            EncryptParam::KeyBitsWithHash,
            EncryptParamProblem::UnusableCombination,
            256,
            128,
            128
        )
    );

    EncryptParams {
        hash: HashAlgorithm::Sha1,
        key_data_key_bits: 128,
        password_key_bits: 128,
        ..Default::default()
    }
    .validate()
    .expect("control: SHA-1's 20 bytes carry a 16-byte key");
}

/// The coupling is `p:encryptedKey`'s, and this is the test that says so.
///
/// [MS-OFFCRYPTO] §2.3.4.11 sizes `Hfinal` against `PasswordKeyEncryptor.keyBits`;
/// §2.3.4.13 step 1 sizes the package key from the RNG against `Encryptor.KeyData.keyBits`
/// and no digest is involved. So SHA-1 with a 256-bit *package* key and a 128-bit KEK is
/// writable, and a validator that checked `key_data_key_bits` against the digest — the
/// obvious mistake, and the one a single merged `key_bits` field would force — refuses it.
#[test]
fn the_coupling_binds_the_password_key_bits_only() {
    EncryptParams {
        hash: HashAlgorithm::Sha1,
        key_data_key_bits: 256,
        password_key_bits: 128,
        ..Default::default()
    }
    .validate()
    .expect("SHA-1 never touches the package key; only the KEK is cut from its digest");
}

/// The two `keyBits` are independent, which is the reason there are two fields.
#[test]
fn an_aes_256_kek_may_wrap_an_aes_128_package_key() {
    EncryptParams {
        key_data_key_bits: 128,
        password_key_bits: 256,
        ..Default::default()
    }
    .validate()
    .expect("[MS-OFFCRYPTO] §2.3.4.10 imposes no equality between the two keyBits");

    EncryptParams {
        key_data_key_bits: 256,
        password_key_bits: 128,
        ..Default::default()
    }
    .validate()
    .expect("and neither direction is privileged");
}

/// `(SaltSize, OutsideSpecRange)` — `ST_SaltSize` is `1..=65536`, on each element alone.
#[test]
fn a_salt_size_outside_the_schema_range_is_refused_on_either_element() {
    for (key_data, password, offender) in [(0, 16, 0), (16, 65_537, 65_537)] {
        let params = EncryptParams {
            key_data_salt_size: key_data,
            password_salt_size: password,
            ..Default::default()
        };
        assert_eq!(
            refusal(&params),
            (
                EncryptParam::SaltSize,
                EncryptParamProblem::OutsideSpecRange,
                offender,
                1,
                65_536
            )
        );
    }

    EncryptParams {
        key_data_salt_size: 1,
        password_salt_size: 65_536,
        ..Default::default()
    }
    .validate()
    .expect("control: both ends of ST_SaltSize are inside the range");
}

/// `(SpinCount, OutsideSpecRange)` — `ST_SpinCount` is `0..=10000000`.
///
/// Written against [`limits::SPIN_COUNT_MAX`] rather than the literal so that the guard
/// moves with the constant, and `min: 0` is asserted because the schema's floor is 0 and
/// the reported span has to say so even though a `u32` cannot break it.
#[test]
fn a_spin_count_above_the_schema_maximum_is_refused() {
    let params = EncryptParams {
        spin_count: limits::SPIN_COUNT_MAX + 1,
        ..Default::default()
    };
    assert_eq!(
        refusal(&params),
        (
            EncryptParam::SpinCount,
            EncryptParamProblem::OutsideSpecRange,
            limits::SPIN_COUNT_MAX + 1,
            0,
            limits::SPIN_COUNT_MAX
        )
    );

    for control in [0, limits::SPIN_COUNT_MAX] {
        EncryptParams {
            spin_count: control,
            ..Default::default()
        }
        .validate()
        .unwrap_or_else(|e| panic!("control: spinCount {control} is inside the range: {e}"));
    }
}

/// A single-field refusal outranks the combination that field takes part in.
///
/// 200 is not an AES key size *and* does not fit SHA-1's digest, so both checks would
/// fire. The answer must be `(KeyBits, UnsupportedByCipher)`: telling the caller the pair
/// is unusable would send it to change the hash, which would not help — no hash makes 200
/// an AES key size. It is a legal `ST_KeyBits` (at least 8, a multiple of 8, and the
/// schema has no maximum), so it is the cipher's refusal and not the format's.
///
/// 200 is also the value that made the old payload a lie: it sits inside the `128..=256`
/// this test used to assert, so the message named it as accepted in the same breath as
/// refusing it. The span is now `192..=192` — the nearest size AES does define, reported
/// as a point.
#[test]
fn an_invalid_key_size_outranks_the_pair_check_it_would_also_fail() {
    let params = EncryptParams {
        hash: HashAlgorithm::Sha1,
        password_key_bits: 200,
        ..Default::default()
    };
    assert_eq!(
        refusal(&params),
        (
            EncryptParam::KeyBits,
            EncryptParamProblem::UnsupportedByCipher,
            200,
            192,
            192
        )
    );

    EncryptParams {
        password_key_bits: 192,
        ..Default::default()
    }
    .validate()
    .expect("the size the refusal named must itself be accepted");
}

/// The default tuple, asserted field by field as literals.
///
/// Not `assert_eq!(EncryptParams::default(), EncryptParams::default())` and not compared
/// against constants pulled from `encryption_info`: these six numbers are what this crate
/// emits when a caller expresses no preference, every acceptance-gate verdict against real
/// Word, real LibreOffice, `msoffcrypto-tool` and `office-crypto` was measured on them,
/// and a change to any one of them must fail here by name so that the evidence is re-run
/// rather than inherited.
#[test]
fn the_default_is_the_tuple_office_16_writes() {
    let d = EncryptParams::default();
    assert_eq!(d.spin_count, 100_000);
    assert_eq!(d.hash, HashAlgorithm::Sha512);
    assert_eq!(d.key_data_key_bits, 256);
    assert_eq!(d.password_key_bits, 256);
    assert_eq!(d.key_data_salt_size, 16);
    assert_eq!(d.password_salt_size, 16);
    d.validate().expect("what Office writes must validate");
}
