//! OCR text extraction via [ocrs](https://github.com/robertknight/ocrs) (rten backend).
//!
//! Models are downloaded on first use to a cache directory and loaded once
//! per process. Cache resolution: `$BATDOC_MODELS_DIR` → `$XDG_CACHE_HOME/batdoc/models`
//! → `~/.cache/batdoc/models`.

use crate::error::{BatdocError, Result};
use std::io::Read;
use std::path::{Path, PathBuf};

/// Detection model (text regions). 2.4 MB.
const DETECTION_MODEL_URL: &str =
    "https://ocrs-models.s3-accelerate.amazonaws.com/text-detection.rten";
/// Recognition model (text lines). 9.3 MB.
const RECOGNITION_MODEL_URL: &str =
    "https://ocrs-models.s3-accelerate.amazonaws.com/text-recognition.rten";
const DETECTION_MODEL_FILE: &str = "text-detection.rten";
const RECOGNITION_MODEL_FILE: &str = "text-recognition.rten";

/// Resolve the model cache directory, in priority order. Injectable for testing.
fn cache_dir_from(get_env: impl Fn(&str) -> Option<String>) -> PathBuf {
    if let Some(dir) = get_env("BATDOC_MODELS_DIR") {
        return PathBuf::from(dir);
    }
    if let Some(xdg) = get_env("XDG_CACHE_HOME") {
        return PathBuf::from(xdg).join("batdoc").join("models");
    }
    get_env("HOME").map_or_else(
        || PathBuf::from(".cache").join("batdoc").join("models"),
        |home| PathBuf::from(home).join(".cache").join("batdoc").join("models"),
    )
}

fn cache_dir() -> PathBuf {
    cache_dir_from(|key| std::env::var(key).ok())
}

/// Download `url` to `path` unless it already exists.
///
/// Writes to a `.tmp` sibling and renames, so concurrent processes cannot
/// observe a partially written model.
fn ensure_file(path: &Path, url: &str) -> Result<()> {
    if path.exists() {
        return Ok(());
    }
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| {
            BatdocError::Document(format!(
                "failed to create OCR model cache directory {}: {e}",
                parent.display()
            ))
        })?;
    }
    let file_name = path.file_name().and_then(|n| n.to_str()).unwrap_or("model");
    let tmp = path.with_file_name(format!("{file_name}.tmp.{}", std::process::id()));
    let response = ureq::get(url).call().map_err(|e| {
        BatdocError::Document(format!(
            "failed to download OCR model from {url}: {e}\n\
             (set BATDOC_MODELS_DIR to a directory containing the model files)"
        ))
    })?;
    let mut buf = Vec::new();
    response
        .into_reader()
        .read_to_end(&mut buf)
        .map_err(|e| BatdocError::Document(format!("failed to read OCR model download: {e}")))?;
    std::fs::write(&tmp, &buf)
        .map_err(|e| BatdocError::Document(format!("failed to write OCR model to {}: {e}", tmp.display())))?;
    if let Err(e) = std::fs::rename(&tmp, path) {
        let _ = std::fs::remove_file(&tmp);
        return Err(BatdocError::Document(format!(
            "failed to install OCR model to {}: {e}",
            path.display()
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn no_env(_key: &str) -> Option<String> {
        None
    }

    #[test]
    fn cache_dir_uses_env_override() {
        let dir = cache_dir_from(|key| match key {
            "BATDOC_MODELS_DIR" => Some("/opt/models".into()),
            _ => None,
        });
        assert_eq!(dir, PathBuf::from("/opt/models"));
    }

    #[test]
    fn cache_dir_uses_xdg_before_home() {
        let dir = cache_dir_from(|key| match key {
            "XDG_CACHE_HOME" => Some("/xdg".into()),
            "HOME" => Some("/home/u".into()),
            _ => None,
        });
        assert_eq!(dir, PathBuf::from("/xdg/batdoc/models"));
    }

    #[test]
    fn cache_dir_falls_back_to_home() {
        let dir = cache_dir_from(|key| match key {
            "HOME" => Some("/home/u".into()),
            _ => None,
        });
        assert_eq!(dir, PathBuf::from("/home/u/.cache/batdoc/models"));
    }

    #[test]
    fn cache_dir_without_home_is_relative() {
        let dir = cache_dir_from(no_env);
        assert_eq!(dir, PathBuf::from(".cache/batdoc/models"));
    }

    #[test]
    fn ensure_file_skips_existing_without_network() {
        let tmp = std::env::temp_dir().join(format!("batdoc-ocr-t1-{}", std::process::id()));
        std::fs::create_dir_all(&tmp).unwrap();
        let file = tmp.join(DETECTION_MODEL_FILE);
        std::fs::write(&file, b"fake model").unwrap();
        // Unreachable URL: must never be contacted because the file exists.
        let result = ensure_file(&file, "https://127.0.0.1:1/unreachable");
        assert!(result.is_ok());
        std::fs::remove_dir_all(&tmp).ok();
    }

    #[test]
    fn ensure_file_failure_leaves_no_partial_file() {
        let tmp = std::env::temp_dir().join(format!("batdoc-ocr-t1b-{}", std::process::id()));
        let file = tmp.join(DETECTION_MODEL_FILE);
        let result = ensure_file(&file, "https://127.0.0.1:1/unreachable");
        assert!(result.is_err());
        assert!(!file.exists(), "no partial model file may remain");
        std::fs::remove_dir_all(&tmp).ok();
    }
}
