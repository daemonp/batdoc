#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
//! The container a legacy binary document is decrypted *inside*.
//!
//! Unlike an encrypted OOXML file, whose plaintext is a ZIP pulled out of one stream, a
//! `.doc`, `.xls` or `.ppt` is decrypted in place: the same CFB, the same directory, the
//! same summary streams, with the encrypted streams' bytes replaced by their plaintext.
//! Every scheme here is a stream cipher or a byte transformation, so each stream keeps
//! its length, and the rewrite touches nothing but the sectors those streams own — which
//! is what makes the Word and Excel results byte-identical to `msoffcrypto-tool -d`'s,
//! measured on every fixture (its `olefile.write_stream` does the same in-place
//! overwrite). The `.ppt` output is the one exception and differs from that tool's by
//! four bytes on purpose; `powerpoint97` says which and why.
//!
//! `cfb` (MIT) is the container layer, per plan D2; nothing here walks a FAT. An
//! in-place write of an unchanged length updates no directory entry — `cfb` stamps a
//! modified time only through `touch`/`set_modified_time`, never on `write`
//! (`cfb-0.14.0/src/internal/stream.rs:216-233`) — so the file's own timestamps survive,
//! as they do through the oracle.

use crate::binary_office::{self, BinaryFormat};
use crate::cfb_reader::read_capped;
use crate::error::Error;
use crate::limits::LEGACY_STREAM_READ_CAP;
use std::io::{Cursor, Seek, SeekFrom, Write};

pub(crate) struct LegacyContainer {
    cfb: cfb::CompoundFile<Cursor<Vec<u8>>>,
}

impl LegacyContainer {
    /// Open a copy of `data`. The copy is the output: the caller gets it back from
    /// [`Self::into_bytes`] with its encrypted streams rewritten.
    pub(crate) fn open(data: &[u8]) -> Result<Self, Error> {
        let cfb =
            cfb::CompoundFile::open(Cursor::new(data.to_vec())).map_err(|_| Error::NotACfbFile)?;
        Ok(Self { cfb })
    }

    pub(crate) fn format(&self) -> Option<BinaryFormat> {
        binary_office::format_of(&self.cfb)
    }

    pub(crate) fn workbook_stream_name(&self) -> Option<&'static str> {
        binary_office::workbook_stream_name(&self.cfb)
    }

    pub(crate) fn exists(&self, name: &str) -> bool {
        self.cfb.exists(name)
    }

    /// The whole stream, or [`Error::MissingStream`] when the container has
    /// no such stream, [`Error::BadParameters`] when it is longer than
    /// [`LEGACY_STREAM_READ_CAP`], or [`Error::Io`] when the read itself
    /// fails. The third caller of [`crate::cfb_reader::read_capped`], and the `Io` arm is
    /// the one its two others document for the same propagation.
    pub(crate) fn read(&mut self, name: &'static str) -> Result<Vec<u8>, Error> {
        let what = name.trim_start_matches('/');
        let stream = self
            .cfb
            .open_stream(name)
            .map_err(|_| Error::MissingStream(what))?;
        read_capped(stream, LEGACY_STREAM_READ_CAP, what)
    }

    /// Replace a stream's bytes in place. `bytes` must be exactly the stream's current
    /// length — a decrypt never changes one, so a mismatch is a bug in the caller and is
    /// reported rather than resized.
    pub(crate) fn overwrite(&mut self, name: &'static str, bytes: &[u8]) -> Result<(), Error> {
        let what = name.trim_start_matches('/');
        let mut stream = self
            .cfb
            .open_stream(name)
            .map_err(|_| Error::MissingStream(what))?;
        if stream.len() != bytes.len() as u64 {
            return Err(Error::BadParameters(format!(
                "rewriting {what} with {} bytes where the stream holds {}",
                bytes.len(),
                stream.len()
            )));
        }
        stream.seek(SeekFrom::Start(0))?;
        stream.write_all(bytes)?;
        stream.flush()?;
        Ok(())
    }

    /// The container with every rewrite flushed.
    pub(crate) fn into_bytes(mut self) -> Result<Vec<u8>, Error> {
        self.cfb.flush()?;
        Ok(self.cfb.into_inner().into_inner())
    }
}
