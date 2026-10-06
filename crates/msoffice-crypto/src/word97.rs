#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! Word 97-2003 (`.doc`) decryption — [MS-DOC] §2.2.6, over the RC4 families of
//! [MS-OFFCRYPTO] §2.3.5 and §2.3.6.
//!
//! What the format says ([MS-DOC] §2.2.6.2 / §2.2.6.3, identical for both families):
//! the `EncryptionHeader` is written in the clear in the first `FibBase.lKey` bytes of
//! the table stream; the rest of the table stream, the `WordDocument` stream beyond its
//! first 68 bytes, and the whole `Data` stream are encrypted in 512-byte blocks numbered
//! from zero at the start of each stream — and "the encryption algorithm MUST be carried
//! out at the beginning of the Table stream and the WordDocument stream even though some
//! of the bytes are written in unencrypted form". So each stream is decrypted whole,
//! from offset 0, and the clear bytes are put back afterwards. The FIB is put back with
//! `fEncrypted`, `fObfuscated` and `lKey` cleared, which is what `msoffcrypto-tool -d`
//! writes and what makes the output a document rather than an encrypted one with its
//! ciphertext replaced; the header at the top of the table stream is left as the
//! keystream made it, as msoffcrypto and LibreOffice (`ww8par.cxx:5795-5808`, behaviour
//! only) both leave it, since nothing in a decrypted document refers to those bytes.
//!
//! XOR obfuscation of a `.doc` ([MS-DOC] §2.2.6.1, Method 2 of [MS-OFFCRYPTO] §2.3.7) is
//! refused by name; see `xor_obfuscation`.
//!
//! Behaviour ported from office-crypto `src/format/doc97.rs` and msoffcrypto-tool
//! `msoffcrypto/format/doc97.py` (both MIT); see NOTICE.

use crate::binary_office::{
    self, FibBase, DATA, FIB_CLEAR_LEN, FIB_FLAGS_OFFSET, FIB_F_ENCRYPTED, FIB_F_OBFUSCATED,
    FIB_L_KEY_OFFSET, WORD_DOCUMENT,
};
use crate::error::Error;
use crate::legacy_container::LegacyContainer;
use crate::limits::ENCRYPTION_HEADER_STRUCTURE_MAX;
use crate::rc4::{self, BlockKeySchedule};
use crate::{rc4_cryptoapi, rc4_office97};

/// [MS-DOC] §2.2.6.2 / §2.2.6.3: 512-byte blocks.
const BLOCK_SIZE: usize = 0x200;

/// Decrypt the container in place. The container is left untouched on any error.
pub(crate) fn decrypt(container: &mut LegacyContainer, password: &str) -> Result<(), Error> {
    let word = container.read(WORD_DOCUMENT)?;
    let fib = FibBase::parse(&word).ok_or(Error::BadParameters(
        "the WordDocument stream does not begin with a FIB (wIdent 0xA5EC)".to_string(),
    ))?;
    if word.len() < FIB_CLEAR_LEN {
        return Err(Error::MissingStream(
            "WordDocument stream shorter than the 68-byte FIB",
        ));
    }
    if !fib.encrypted() {
        return Err(Error::NotEncrypted);
    }
    if fib.obfuscated() {
        return Err(Error::UnsupportedAlgorithm {
            what: "FibBase.fObfuscated",
            name: "XOR obfuscation of a Word document ([MS-OFFCRYPTO] 2.3.7.4)".to_string(),
        });
    }

    // [MS-DOC] §2.5.2: lKey is the size of the EncryptionHeader at the start of the
    // table stream. A file-declared length, bounded before it slices.
    let table_name = fib.table_stream();
    let table = container.read(table_name)?;
    let l_key = usize::try_from(fib.l_key).unwrap_or(usize::MAX);
    if l_key > ENCRYPTION_HEADER_STRUCTURE_MAX {
        return Err(Error::BadParameters(format!(
            "FibBase.lKey is {}, over the {ENCRYPTION_HEADER_STRUCTURE_MAX} an RC4 \
             encryption header can occupy",
            fib.l_key
        )));
    }
    let Some(structure) = table.get(..l_key) else {
        return Err(Error::BadParameters(format!(
            "FibBase.lKey is {} but the table stream holds {} bytes",
            fib.l_key,
            table.len()
        )));
    };

    // [MS-DOC] §2.2.6: EncryptionVersionInfo names the family.
    let schedule: Box<dyn BlockKeySchedule> = match (
        binary_office::le16(structure, 0),
        binary_office::le16(structure, 2),
    ) {
        (Some(1), Some(1)) => {
            let header = rc4_office97::parse(structure)?;
            let schedule = rc4_office97::Office97KeySchedule::new(password, &header.salt);
            schedule.verify(&header)?;
            Box::new(schedule)
        }
        (Some(2..=4), Some(2)) => {
            let header = rc4_cryptoapi::parse(structure)?;
            let schedule =
                rc4_cryptoapi::CryptoApiKeySchedule::new(password, &header.salt, header.key_bits)?;
            schedule.verify(&header)?;
            Box::new(schedule)
        }
        (Some(major), Some(minor)) => {
            return Err(Error::UnsupportedEncryptionVersion(major, minor))
        }
        _ => {
            return Err(Error::BadParameters(
                "FibBase.lKey is too small to hold an EncryptionVersionInfo".to_string(),
            ))
        }
    };

    // The password is verified; decrypt every stream whole, then put the clear FIB back
    // with the encryption flags and lKey cleared.
    let mut word_plain = word.clone();
    rc4::decrypt_in_blocks(schedule.as_ref(), &mut word_plain, BLOCK_SIZE)?;
    word_plain[..FIB_CLEAR_LEN].copy_from_slice(&word[..FIB_CLEAR_LEN]);
    let flags = fib.flags & !(FIB_F_ENCRYPTED | FIB_F_OBFUSCATED);
    word_plain[FIB_FLAGS_OFFSET..FIB_FLAGS_OFFSET + 2].copy_from_slice(&flags.to_le_bytes());
    word_plain[FIB_L_KEY_OFFSET..FIB_L_KEY_OFFSET + 4].copy_from_slice(&0u32.to_le_bytes());

    let mut table_plain = table;
    rc4::decrypt_in_blocks(schedule.as_ref(), &mut table_plain, BLOCK_SIZE)?;

    let data_plain = if container.exists(DATA) {
        let mut data = container.read(DATA)?;
        rc4::decrypt_in_blocks(schedule.as_ref(), &mut data, BLOCK_SIZE)?;
        Some(data)
    } else {
        None
    };

    container.overwrite(WORD_DOCUMENT, &word_plain)?;
    container.overwrite(table_name, &table_plain)?;
    if let Some(data) = data_plain {
        container.overwrite(DATA, &data)?;
    }
    Ok(())
}
