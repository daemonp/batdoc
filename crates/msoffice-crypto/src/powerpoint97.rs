#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! PowerPoint 97-2003 (`.ppt`) decryption — [MS-PPT] §2.3.7 `CryptSession10Container`,
//! over RC4 CryptoAPI ([MS-OFFCRYPTO] §2.3.5), the only family the format defines.
//!
//! What the format says (§2.3.7): the `Current User` stream is not encrypted; in the
//! `PowerPoint Document` stream the `UserEditAtom`, the `PersistDirectoryAtom` and the
//! `rh` of the `CryptSession10Container` are not encrypted and everything else is; the
//! stream contains exactly one `UserEditAtom`, whose `encryptSessionPersistIdRef` names
//! the persist object holding the container, whose `data` is the encryption header
//! structure. Each persist object is decrypted under the key for the block number equal
//! to its **persist object identifier**, over `8 + recLen` bytes from its offset —
//! except that the `recLen` is itself encrypted, so, like msoffcrypto, this reads each
//! object as extending to the next object's offset (or to the persist directory) and
//! decrypts the gap too, which is a no-op on padding and the same bytes the oracle
//! writes. The `Pictures` stream (§2.1.3, block 0 per picture) is not decrypted; neither
//! the fixture nor msoffcrypto has one.
//!
//! What is written back: every persist object decrypted, the `CryptSession10Container`
//! zeroed (header and data), the `UserEditAtom` shortened to `0x1C` with its
//! `encryptSessionPersistIdRef` zeroed in place, and `CurrentUserAtom.headerToken` set
//! to the unencrypted token `0xE391C05F` ([MS-PPT] §2.3.2; PowerPoint 16 writes that
//! token even when encrypting, so on a real file this last write changes nothing). The
//! persist directory is left exactly as it was: its entry for the container now names a
//! zeroed record, which is an unknown record and "MUST be ignored" (§2.1.2).
//!
//! **That last point is where this output differs from `msoffcrypto-tool -d`'s, and the
//! difference is four bytes.** msoffcrypto decrements the first `PersistDirectoryEntry`'s
//! `cPersist` so the directory no longer counts the container — but leaves the atom's
//! `recLen` and the container's offset in place, so a reader walking the atom finds a
//! second entry made of that stray offset, with `cPersist = 0`, which §2.3.5 forbids
//! ("MUST be greater than or equal to 0x001"). PowerPoint 16 refuses the result
//! outright: `0x80048242`, "Office has detected a problem with this file". Measured on
//! 2026-09-05 by bisecting the rewrite — the oracle's bytes with only that word restored
//! open with the full slide text, and every variant keeping the atom at `0x20` is refused
//! whatever the directory says. So `tests/legacy_binary_fixtures.rs` pins this crate's
//! output to its own digest *and* pins that re-applying msoffcrypto's decrement to it
//! reproduces msoffcrypto's digest exactly: identical everywhere but the one word the
//! real application rejects. Word and Excel outputs match the oracle byte for byte.
//!
//! Behaviour ported from msoffcrypto-tool `msoffcrypto/format/ppt97.py` (MIT); see
//! NOTICE. Two further deliberate departures, both for hostile input: persist object
//! extents come from the offsets sorted, not from the directory's order, so no object
//! can be decrypted twice or overlap the next; and an `offsetLastEdit` other than zero
//! is refused, because the spec makes it a MUST and a second edit's objects would
//! otherwise be left encrypted with no error.

use crate::binary_office::{
    self, UserEditAtom, CURRENT_USER, CURRENT_USER_HEADER_TOKEN_AT, POWERPOINT_DOCUMENT,
    USER_EDIT_ATOM_LEN_ENCRYPTED, USER_EDIT_ATOM_LEN_PLAIN,
};
use crate::error::Error;
use crate::legacy_container::LegacyContainer;
use crate::rc4;
use crate::rc4_cryptoapi::{self, CryptoApiKeySchedule};

/// `CurrentUserAtom.headerToken` for a file that is not encrypted — [MS-PPT] §2.3.2.
const HEADER_TOKEN_PLAIN: u32 = 0xE391_C05F;

fn bad(msg: impl Into<String>) -> Error {
    Error::BadParameters(msg.into())
}

/// Decrypt the container in place. The container is left untouched on any error.
pub(crate) fn decrypt(container: &mut LegacyContainer, password: &str) -> Result<(), Error> {
    let current_user = container.read(CURRENT_USER)?;
    let doc = container.read(POWERPOINT_DOCUMENT)?;
    let mut src: &[u8] = &doc;

    let edit_offset = binary_office::current_edit_offset(&current_user).ok_or(
        Error::MissingStream("Current User stream shorter than its atom"),
    )?;
    let edit_offset = usize::try_from(edit_offset).map_err(|_| bad("offsetToCurrentEdit"))?;
    let atom = binary_office::user_edit_atom(&mut src, edit_offset as u64).ok_or_else(|| {
        bad(format!(
            "no UserEditAtom at offsetToCurrentEdit {edit_offset} (stream of {} bytes)",
            doc.len()
        ))
    })?;
    if atom.rec_len == USER_EDIT_ATOM_LEN_PLAIN {
        return Err(Error::NotEncrypted);
    }
    debug_assert_eq!(atom.rec_len, USER_EDIT_ATOM_LEN_ENCRYPTED);
    if atom.offset_last_edit != 0 {
        return Err(bad(format!(
            "UserEditAtom.offsetLastEdit is {}; an encrypted presentation contains exactly \
             one user edit ([MS-PPT] 2.3.7)",
            atom.offset_last_edit
        )));
    }
    let directory_offset = usize::try_from(atom.offset_persist_directory)
        .map_err(|_| bad("offsetPersistDirectory"))?;
    let directory = binary_office::persist_directory(&mut src, directory_offset as u64)
        .ok_or_else(|| {
            bad(format!(
                "no readable PersistDirectoryAtom at offsetPersistDirectory {directory_offset}"
            ))
        })?;
    let session_id = atom
        .encrypt_session_persist_id_ref
        .ok_or_else(|| bad("UserEditAtom carries no encryptSessionPersistIdRef"))?;
    let session_offset = *directory.get(&session_id).ok_or_else(|| {
        bad(format!(
            "encryptSessionPersistIdRef {session_id} is not in the persist directory"
        ))
    })?;
    let structure = binary_office::crypt_session_container(&mut src, u64::from(session_offset))
        .ok_or_else(|| {
            bad(format!(
                "no CryptSession10Container at persist object {session_id} (offset \
                 {session_offset})"
            ))
        })?;
    let session_len = binary_office::record_header(&mut src, u64::from(session_offset))
        .map(|rh| rh.rec_len)
        .unwrap_or(0);

    let header = rc4_cryptoapi::parse(&structure)?;
    let schedule = CryptoApiKeySchedule::new(password, &header.salt, header.key_bits)?;
    schedule.verify(&header)?;

    // Every persist object must lie below the persist directory ([MS-PPT] §2.3.5), which
    // bounds the total work by the stream's length once the extents are disjoint.
    let mut objects: Vec<(u32, usize)> = Vec::with_capacity(directory.len());
    for (&id, &offset) in &directory {
        let offset = usize::try_from(offset).map_err(|_| bad("persist offset"))?;
        if offset >= directory_offset {
            return Err(bad(format!(
                "persist object {id} at offset {offset} lies at or beyond the persist \
                 directory at {directory_offset}"
            )));
        }
        objects.push((id, offset));
    }
    objects.sort_by_key(|&(_, offset)| offset);

    let mut plain = doc.clone();
    for (i, &(id, offset)) in objects.iter().enumerate() {
        let end = objects
            .get(i + 1)
            .map_or(directory_offset, |&(_, next)| next);
        if offset == usize::try_from(session_offset).unwrap_or(usize::MAX) {
            // The container: header and data zeroed, bounded by its own recLen, which
            // `crypt_session_container` has already read within the stream.
            let len = 8usize.saturating_add(usize::try_from(session_len).unwrap_or(0));
            let zero_end = offset.saturating_add(len).min(plain.len());
            plain[offset..zero_end].fill(0);
            continue;
        }
        if end <= offset {
            // Two objects at one offset: the sort put them adjacent and the second has no
            // extent. Nothing to decrypt twice.
            continue;
        }
        rc4::decrypt_with_block(&schedule, &mut plain[offset..end], id)?;
    }

    // UserEditAtom: recLen 0x20 -> 0x1C, and the reference field zeroed where it stood.
    let rec_len_at = edit_offset + UserEditAtom::REC_LEN_AT;
    plain[rec_len_at..rec_len_at + 4].copy_from_slice(&USER_EDIT_ATOM_LEN_PLAIN.to_le_bytes());
    let ref_at = edit_offset + UserEditAtom::ENCRYPT_SESSION_REF_AT;
    plain[ref_at..ref_at + 4].fill(0);

    // The persist directory is deliberately not touched -- see the module docs for the
    // four bytes msoffcrypto changes here and what PowerPoint makes of them.

    // The container guarantees the same length; the header token is the one field of
    // `Current User` that changes, and only if it was the encrypted token.
    let mut current_user_plain = current_user.clone();
    if let Some(slot) =
        current_user_plain.get_mut(CURRENT_USER_HEADER_TOKEN_AT..CURRENT_USER_HEADER_TOKEN_AT + 4)
    {
        slot.copy_from_slice(&HEADER_TOKEN_PLAIN.to_le_bytes());
    }

    container.overwrite(POWERPOINT_DOCUMENT, &plain)?;
    if current_user_plain != current_user {
        container.overwrite(CURRENT_USER, &current_user_plain)?;
    }
    Ok(())
}
