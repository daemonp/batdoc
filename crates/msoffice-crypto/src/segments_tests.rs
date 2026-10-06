//! Segmentation and per-segment IV derivation.
//!
//! These are the tests the inline `chunks(4096).enumerate()` in `decrypt_package` could
//! not have: it was reachable only with a session key and a parsed `AgileParams`, so every
//! statement about *segmentation* had to be made through a full decrypt, and a boundary
//! bug was only ever visible as a wrong plaintext. Extracting the loop (plan D4) is what
//! makes the boundaries assertable on their own — and the encrypt path will lean on the
//! same guarantees from the other side.

use super::*;

const SALT: [u8; 16] = [0x5Au8; 16];

/// Segment `data` under SHA-512 with a 16-byte block, the tuple every fixture uses.
fn segments(data: &[u8]) -> Vec<Segment<'_>> {
    Segments::new(data, HashAlgorithm::Sha512, &SALT, 16)
        .expect("SHA-512 with a 16-byte block is always segmentable")
        .map(|s| s.expect("no segment of a validated payload fails"))
        .collect()
}

/// Where the cuts fall, across every shape the final segment can take.
///
/// The three interesting cases are an exact multiple (no short tail), one byte over (a
/// 1-byte tail, the most likely off-by-one), and one byte under (a 4095-byte tail). Empty
/// input yields nothing rather than one empty segment — a writer handed an empty package
/// must emit no ciphertext, not one segment of padding.
#[test]
fn the_cuts_fall_where_the_format_says() {
    for (len, want_lens) in [
        (0usize, vec![]),
        (1, vec![1usize]),
        (SEGMENT_LEN - 1, vec![SEGMENT_LEN - 1]),
        (SEGMENT_LEN, vec![SEGMENT_LEN]),
        (SEGMENT_LEN + 1, vec![SEGMENT_LEN, 1]),
        (SEGMENT_LEN * 2, vec![SEGMENT_LEN, SEGMENT_LEN]),
        (SEGMENT_LEN * 2 + 17, vec![SEGMENT_LEN, SEGMENT_LEN, 17]),
    ] {
        let data = vec![0u8; len];
        let got: Vec<usize> = segments(&data).iter().map(|s| s.bytes.len()).collect();
        assert_eq!(got, want_lens, "wrong cuts for a {len}-byte payload");
    }
}

/// The segments partition the input: concatenating them reproduces it exactly.
///
/// A cut that dropped or duplicated a byte would still produce plausible segment
/// *lengths*, so the lengths above are not sufficient on their own. Distinct bytes rather
/// than zeros, so a duplicated segment is visible.
#[test]
fn the_segments_partition_the_payload_without_loss() {
    let data: Vec<u8> = (0..SEGMENT_LEN * 3 + 123)
        .map(|i| (i % 251) as u8)
        .collect();
    let rejoined: Vec<u8> = segments(&data)
        .iter()
        .flat_map(|s| s.bytes.iter().copied())
        .collect();
    assert_eq!(rejoined, data);
}

/// Each segment's IV is `H(salt || LE32(index))[..block_size]`, derived independently.
///
/// The expectation is recomputed here from the formula rather than taken from the
/// iterator, so this is a statement about the format and not a restatement of the code.
/// The `enumerate()` index is what the segment's IV must agree with — which is also why
/// `Segment` need not carry the index itself.
#[test]
fn each_segment_gets_the_iv_its_index_derives() {
    let data = vec![0u8; SEGMENT_LEN * 3 + 1];
    for (index, segment) in segments(&data).iter().enumerate() {
        let want = derive_iv(
            HashAlgorithm::Sha512,
            &SALT,
            &(index as u32).to_le_bytes(),
            16,
        )
        .unwrap();
        assert_eq!(segment.iv, want, "segment {index} has the wrong IV");
        assert_eq!(segment.iv.len(), 16, "the IV is truncated to blockSize");
    }
}

/// No two segments share an IV.
///
/// The failure this exists for is a counter that never advances — every segment encrypted
/// under IV zero. Under CBC with one key that is a catastrophic reuse, and a round-trip
/// test cannot see it at all: encrypt and decrypt would agree with each other perfectly.
#[test]
fn no_two_segments_share_an_iv() {
    let data = vec![0u8; SEGMENT_LEN * 8];
    let ivs: Vec<Vec<u8>> = segments(&data).into_iter().map(|s| s.iv).collect();
    assert_eq!(ivs.len(), 8);
    let mut unique = ivs.clone();
    unique.sort();
    unique.dedup();
    assert_eq!(unique.len(), ivs.len(), "an IV was reused across segments");
}

/// The IV follows the hash the file names, not a hardcoded SHA-512.
///
/// This is GH #11's bug expressed at the level it actually lived at. Two algorithms over
/// the same salt and index must not produce the same IV; when `decrypt_package` hardcoded
/// SHA-512, a SHA-1 file decrypted under IVs from the wrong hash *after* `integrity::verify`
/// had reported `Verified` under the right ones.
#[test]
fn the_iv_follows_the_named_hash_algorithm() {
    let data = vec![0u8; SEGMENT_LEN + 1];
    let sha512 = segments(&data);
    let sha1: Vec<Segment<'_>> = Segments::new(&data, HashAlgorithm::Sha1, &SALT, 16)
        .unwrap()
        .map(|s| s.unwrap())
        .collect();

    assert_eq!(sha1.len(), sha512.len());
    for (i, (a, b)) in sha1.iter().zip(sha512.iter()).enumerate() {
        assert_ne!(
            a.iv, b.iv,
            "segment {i} derived the same IV under two hashes"
        );
    }
}

/// A block size longer than the digest is an error, not a slice past the end.
///
/// SHA-1 produces 20 bytes, so a 32-byte block size has nothing to truncate from. The
/// constructor is where this is caught, so it costs one check for the whole payload rather
/// than one per segment — and the message names the field the file declared.
#[test]
fn a_block_size_past_the_digest_is_an_error_not_a_panic() {
    let data = vec![0u8; SEGMENT_LEN];
    let got = Segments::new(&data, HashAlgorithm::Sha1, &SALT, 32).map(|s| s.len());
    assert!(
        matches!(&got, Err(Error::BadParameters(msg))
            if msg.contains("blockSize") && msg.contains("SHA1")),
        "a 32-byte block under SHA-1 must be refused, got: {got:?}"
    );

    // The control: 20 is exactly the digest length and is accepted, so the refusal is
    // about the excess and not about SHA-1 being rejected outright.
    assert!(Segments::new(&data, HashAlgorithm::Sha1, &SALT, 20).is_ok());
}

/// Padding rounds the final segment up to an AES block and leaves the rest alone.
///
/// Zeros, matching herumi's `data.resize(RoundUp(size, 16))` (`include/encode.hpp:79`).
/// The pad is invisible to a caller — the reader truncates to the stream's declared
/// plaintext size — but it is part of the ciphertext a writer emits, so it has to match
/// what other implementations produce for the encrypt path's golden output to mean
/// anything.
#[test]
fn the_final_segment_pads_up_to_an_aes_block_with_zeros() {
    let data: Vec<u8> = (0..SEGMENT_LEN + 17).map(|_| 0xAB).collect();
    let segs = segments(&data);
    assert_eq!(segs.len(), 2);

    // A full segment is already a block multiple: padding copies and does not grow.
    assert_eq!(segs[0].padded().len(), SEGMENT_LEN);
    assert_eq!(segs[0].padded(), segs[0].bytes);

    // 17 bytes rounds up to 32, and the twelve added bytes are zero.
    let padded = segs[1].padded();
    assert_eq!(padded.len(), 32);
    assert_eq!(&padded[..17], &[0xABu8; 17]);
    assert_eq!(&padded[17..], &[0u8; 15]);
}

/// `size_hint` and `len` agree with what the iterator actually yields.
///
/// `Vec::with_capacity` in `decrypt_package` and every future `collect()` size off these,
/// so a wrong hint is a silent over-allocation on a file-controlled length.
#[test]
fn the_size_hint_matches_what_is_yielded() {
    for len in [0, 1, SEGMENT_LEN, SEGMENT_LEN * 3 + 5] {
        let data = vec![0u8; len];
        let mut it = Segments::new(&data, HashAlgorithm::Sha512, &SALT, 16).unwrap();
        let expected = it.len();
        assert_eq!(it.size_hint(), (expected, Some(expected)));

        // And it shrinks as the iterator advances, rather than being a constant that
        // happens to be right at the start.
        let mut seen = 0;
        while let Some(segment) = it.next() {
            segment.unwrap();
            seen += 1;
            let left = expected - seen;
            assert_eq!(it.size_hint(), (left, Some(left)));
        }
        assert_eq!(
            seen, expected,
            "a {len}-byte payload yielded {seen} segments"
        );
    }
}
