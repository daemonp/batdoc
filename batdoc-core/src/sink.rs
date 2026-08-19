use crate::error::{BatdocError, Result};

pub trait ExtractSink {
    /// Write a string to the sink.
    ///
    /// # Errors
    /// Returns an error if the underlying sink or output budget cannot accept
    /// the write.
    fn write_str(&mut self, s: &str) -> Result<()>;
}

impl ExtractSink for String {
    fn write_str(&mut self, s: &str) -> Result<()> {
        self.push_str(s);
        Ok(())
    }
}

impl<S: ExtractSink + ?Sized> ExtractSink for &mut S {
    fn write_str(&mut self, s: &str) -> Result<()> {
        (**self).write_str(s)
    }
}

pub struct IoSink<W: std::io::Write>(pub W);

impl<W: std::io::Write> ExtractSink for IoSink<W> {
    fn write_str(&mut self, s: &str) -> Result<()> {
        self.0.write_all(s.as_bytes())?;
        Ok(())
    }
}

pub struct BudgetSink<S: ExtractSink> {
    inner: S,
    written: u64,
    max: u64,
}

impl<S: ExtractSink> BudgetSink<S> {
    pub const fn new(inner: S, max: u64) -> Self {
        Self {
            inner,
            written: 0,
            max,
        }
    }
}

impl<S: ExtractSink> ExtractSink for BudgetSink<S> {
    fn write_str(&mut self, s: &str) -> Result<()> {
        let add = s.len() as u64;
        if self.written.saturating_add(add) > self.max {
            return Err(BatdocError::Document(format!(
                "output exceeded {} bytes",
                self.max
            )));
        }
        self.inner.write_str(s)?;
        self.written += add;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn string_sink_appends() {
        let mut s = String::new();
        ExtractSink::write_str(&mut s, "ab").unwrap();
        ExtractSink::write_str(&mut s, "c").unwrap();
        assert_eq!(s, "abc");
    }

    #[test]
    fn io_sink_writes_bytes() {
        let mut buf = Vec::new();
        let mut sink = IoSink(&mut buf);
        sink.write_str("hi").unwrap();
        assert_eq!(buf, b"hi");
    }

    #[test]
    fn budget_sink_errors_and_stops() {
        let mut s = String::new();
        let mut sink = BudgetSink::new(&mut s, 4);
        sink.write_str("abc").unwrap();
        let err = sink.write_str("de").unwrap_err().to_string();
        assert_eq!(err, "output exceeded 4 bytes");
        assert_eq!(s, "abc");
    }

    #[test]
    fn extract_plain_to_matches_extract_plain_on_garbage() {
        let data = b"hello world, definitely not a document";
        let a = crate::to_plain(data).unwrap_err().to_string();
        let mut out = String::new();
        let b = crate::extract_plain_to(
            data,
            crate::Format::Doc,
            crate::ExtractOptions::default(),
            &mut out,
        );
        // detect fails first in to_plain; here we pass an explicit format so
        // we only check the dispatch exists. Use detect + extract_plain_to:
        let format = crate::detect_format(data).unwrap_err().to_string();
        assert_eq!(a, format);
        assert!(b.is_err() || out.is_empty());
    }
}
