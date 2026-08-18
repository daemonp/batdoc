pub(crate) struct StringArena {
    buf: Vec<u8>,
    spans: Vec<(u32, u32)>,
}

impl StringArena {
    pub(crate) fn new() -> Self {
        Self {
            buf: Vec::new(),
            spans: Vec::new(),
        }
    }

    pub(crate) fn push(&mut self, s: &str) -> usize {
        let start = self.buf.len() as u32;
        self.buf.extend_from_slice(s.as_bytes());
        self.spans.push((start, s.len() as u32));
        self.spans.len() - 1
    }

    pub(crate) fn get(&self, i: usize) -> Option<&str> {
        let (start, len) = *self.spans.get(i)?;
        let bytes = &self.buf[start as usize..][..len as usize];
        Some(std::str::from_utf8(bytes).unwrap_or(""))
    }

    #[allow(dead_code)]
    pub(crate) fn len(&self) -> usize {
        self.spans.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn push_get_roundtrip() {
        let mut a = StringArena::new();
        let i = a.push("hello");
        let j = a.push("world");
        assert_eq!(a.get(i), Some("hello"));
        assert_eq!(a.get(j), Some("world"));
        assert_eq!(a.len(), 2);
    }

    #[test]
    fn get_oob_is_none() {
        let a = StringArena::new();
        assert_eq!(a.get(0), None);
    }

    #[test]
    fn empty_string_is_distinct_slot() {
        let mut a = StringArena::new();
        let i = a.push("");
        assert_eq!(a.get(i), Some(""));
        assert_eq!(a.len(), 1);
    }
}
