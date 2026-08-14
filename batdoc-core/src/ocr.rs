//! OCR text extraction via [ocrs](https://github.com/robertknight/ocrs) (rten backend).
//!
//! Models are downloaded on first use to a cache directory and loaded once
//! per process. Cache resolution: `$BATDOC_MODELS_DIR` → `$XDG_CACHE_HOME/batdoc/models`
//! → `~/.cache/batdoc/models`.

use crate::error::{BatdocError, Result};
use ocrs::{ImageSource, OcrEngine, OcrEngineParams};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;
use std::time::Duration;

/// Detection model (text regions). 2.4 MB.
const DETECTION_MODEL_URL: &str =
    "https://ocrs-models.s3-accelerate.amazonaws.com/text-detection.rten";
/// Recognition model (text lines). 9.3 MB.
const RECOGNITION_MODEL_URL: &str =
    "https://ocrs-models.s3-accelerate.amazonaws.com/text-recognition.rten";
const DETECTION_MODEL_FILE: &str = "text-detection.rten";
const RECOGNITION_MODEL_FILE: &str = "text-recognition.rten";

/// Network timeout for model downloads.
const DOWNLOAD_TIMEOUT: Duration = Duration::from_mins(2);
/// Maximum accepted model download size (the real models are 2.4 MB and
/// 9.3 MB; a larger response means a compromised/redirected URL).
const MAX_MODEL_DOWNLOAD_BYTES: u64 = 64 * 1024 * 1024;
/// Age after which leftover `*.tmp.<pid>` download files are swept.
const STALE_TMP_AGE: Duration = Duration::from_hours(1);

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
        |home| {
            PathBuf::from(home)
                .join(".cache")
                .join("batdoc")
                .join("models")
        },
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
    let response = ureq::get(url)
        .timeout(DOWNLOAD_TIMEOUT)
        .call()
        .map_err(|e| {
            BatdocError::Document(format!(
                "failed to download OCR model from {url}: {e}\n\
             (set BATDOC_MODELS_DIR to a directory containing the model files)"
            ))
        })?;
    // Read at most 64 MiB so an oversized/redirected response cannot fill
    // memory; a longer body is rejected before any file is written.
    let mut buf = Vec::new();
    response
        .into_reader()
        .take(MAX_MODEL_DOWNLOAD_BYTES + 1)
        .read_to_end(&mut buf)
        .map_err(|e| BatdocError::Document(format!("failed to read OCR model download: {e}")))?;
    if buf.len() as u64 > MAX_MODEL_DOWNLOAD_BYTES {
        return Err(BatdocError::Document(
            "OCR model file exceeds 64 MiB".into(),
        ));
    }
    std::fs::write(&tmp, &buf).map_err(|e| {
        BatdocError::Document(format!(
            "failed to write OCR model to {}: {e}",
            tmp.display()
        ))
    })?;
    if let Err(e) = std::fs::rename(&tmp, path) {
        let _ = std::fs::remove_file(&tmp);
        return Err(BatdocError::Document(format!(
            "failed to install OCR model to {}: {e}",
            path.display()
        )));
    }
    Ok(())
}

/// Remove stale `*.tmp.<pid>` download leftovers in `dir` (from interrupted
/// runs). Only files older than [`STALE_TMP_AGE`] are touched, so an active
/// concurrent download is never disturbed. Best-effort: errors ignored.
fn sweep_stale_tmp_files(dir: &Path) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        let Some(name) = path.file_name().and_then(|n| n.to_str()) else {
            continue;
        };
        let modified = entry.metadata().ok().and_then(|m| m.modified().ok());
        if is_stale_tmp_file(name, modified) {
            std::fs::remove_file(&path).ok();
        }
    }
}

/// `true` for `*.tmp.<pid>` download leftovers older than [`STALE_TMP_AGE`].
fn is_stale_tmp_file(name: &str, modified: Option<std::time::SystemTime>) -> bool {
    name.contains(".tmp.")
        && modified
            .and_then(|m| m.elapsed().ok())
            .is_some_and(|age| age > STALE_TMP_AGE)
}

/// Paths of the two OCR model files once present locally.
struct ModelPaths {
    detection: PathBuf,
    recognition: PathBuf,
}

/// Ensure both model files exist locally, downloading when needed.
fn ensure_models() -> Result<ModelPaths> {
    let dir = cache_dir();
    sweep_stale_tmp_files(&dir);
    let detection = dir.join(DETECTION_MODEL_FILE);
    let recognition = dir.join(RECOGNITION_MODEL_FILE);
    ensure_file(&detection, DETECTION_MODEL_URL)?;
    ensure_file(&recognition, RECOGNITION_MODEL_URL)?;
    Ok(ModelPaths {
        detection,
        recognition,
    })
}

fn build_engine() -> std::result::Result<OcrEngine, String> {
    let paths = ensure_models().map_err(|e| e.to_string())?;
    let detection_model = rten::Model::load_file(&paths.detection)
        .map_err(|e| format!("failed to load OCR detection model: {e}"))?;
    let recognition_model = rten::Model::load_file(&paths.recognition)
        .map_err(|e| format!("failed to load OCR recognition model: {e}"))?;
    OcrEngine::new(OcrEngineParams {
        detection_model: Some(detection_model),
        recognition_model: Some(recognition_model),
        ..Default::default()
    })
    .map_err(|e| format!("failed to initialize OCR engine: {e}"))
}

/// Process-wide OCR engine, built once on first use.
///
/// Failures are NOT cached: a transient error (e.g. a network failure
/// during the first model download) is returned to the caller, and a later
/// call retries. In a concurrent first-build race the loser discards its
/// engine and shares the stored one.
fn engine() -> Result<&'static OcrEngine> {
    static ENGINE: OnceLock<OcrEngine> = OnceLock::new();
    if let Some(engine) = ENGINE.get() {
        return Ok(engine);
    }
    let built = build_engine().map_err(BatdocError::Document)?;
    Ok(ENGINE.get_or_init(|| built))
}

/// `true` when both OCR model files already exist in the cache directory
/// (i.e. OCR will not trigger a download). Used by the CLI to print a
/// first-use download notice.
#[must_use]
pub fn models_present() -> bool {
    models_present_in(&cache_dir())
}

fn models_present_in(dir: &Path) -> bool {
    dir.join(DETECTION_MODEL_FILE).exists() && dir.join(RECOGNITION_MODEL_FILE).exists()
}

/// OCR an already-decoded RGB image. `None` when no text was detected.
pub(crate) fn ocr_rgb_image(img: &image::RgbImage) -> Result<Option<String>> {
    let source = ImageSource::from_bytes(img.as_raw(), img.dimensions())
        .map_err(|e| BatdocError::Document(format!("OCR input preparation failed: {e}")))?;
    let engine = engine()?;
    let input = engine
        .prepare_input(source)
        .map_err(|e| BatdocError::Document(format!("OCR preprocessing failed: {e}")))?;
    let text = engine
        .get_text(&input)
        .map_err(|e| BatdocError::Document(format!("OCR failed: {e}")))?;
    let text = text.trim();
    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(text.to_string()))
    }
}

/// Maximum width/height accepted when decoding images for OCR. Strict
/// decoder-side guard (shared with PDF embedded-image decoding) so a
/// crafted image whose header claims huge dimensions cannot allocate past
/// the OCR pixel budget (`10_000²` ≈ 100 MP ≈ 300 MB RGB).
pub(crate) const MAX_OCR_IMAGE_DIM: u32 = 10_000;

/// Decode limits for OCR image input, matching PDF embedded-image decoding:
/// `Limits::default()` (512 MiB alloc cap) plus width/height ≤
/// [`MAX_OCR_IMAGE_DIM`].
fn ocr_decode_limits() -> image::Limits {
    let mut limits = image::Limits::default();
    limits.max_image_width = Some(MAX_OCR_IMAGE_DIM);
    limits.max_image_height = Some(MAX_OCR_IMAGE_DIM);
    limits
}

/// Decode and OCR an image byte slice (PNG/JPEG/GIF/WebP/BMP).
///
/// `None` when the bytes are not a decodable image (including images beyond
/// the decode limits) or OCR found no text.
pub(crate) fn ocr_image_bytes(data: &[u8]) -> Result<Option<String>> {
    let Ok(mut reader) = image::ImageReader::new(std::io::Cursor::new(data)).with_guessed_format()
    else {
        return Ok(None);
    };
    reader.limits(ocr_decode_limits());
    let img = match reader.decode() {
        Ok(img) => img.into_rgb8(),
        Err(_) => return Ok(None),
    };
    ocr_rgb_image(&img)
}

/// Extract text from an image file (top-level `Format::Image` entry point).
pub(crate) fn extract_image_plain(data: &[u8]) -> Result<String> {
    ocr_image_bytes(data)?.map_or_else(
        || Err(BatdocError::Document("no text found in image".into())),
        Ok,
    )
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

    #[test]
    fn ocr_image_bytes_undecodable_returns_none() {
        // Garbage bytes: decode fails before any model/network access.
        assert_eq!(ocr_image_bytes(b"not an image").unwrap(), None);
    }

    #[test]
    fn extract_image_plain_undecodable_errors() {
        let err = extract_image_plain(b"garbage").unwrap_err().to_string();
        assert!(err.contains("no text found in image"));
    }

    #[test]
    fn stale_tmp_file_detection() {
        use std::time::SystemTime;
        // Fresh download in progress: recent mtime → keep.
        assert!(!is_stale_tmp_file(
            "text-detection.rten.tmp.123",
            Some(SystemTime::now())
        ));
        // Old leftover from a crashed run → sweep.
        assert!(is_stale_tmp_file(
            "text-detection.rten.tmp.123",
            Some(SystemTime::now() - 2 * STALE_TMP_AGE)
        ));
        // Not a tmp file → never sweep, however old.
        assert!(!is_stale_tmp_file(
            "text-detection.rten",
            Some(SystemTime::now() - 2 * STALE_TMP_AGE)
        ));
        // Unknown mtime → leave it alone.
        assert!(!is_stale_tmp_file("x.tmp.1", None));
    }

    #[test]
    fn models_present_in_requires_both_files() {
        let tmp = std::env::temp_dir().join(format!("batdoc-ocr-t2-{}", std::process::id()));
        std::fs::create_dir_all(&tmp).unwrap();
        assert!(!models_present_in(&tmp));
        std::fs::write(tmp.join(DETECTION_MODEL_FILE), b"x").unwrap();
        assert!(!models_present_in(&tmp));
        std::fs::write(tmp.join(RECOGNITION_MODEL_FILE), b"x").unwrap();
        assert!(models_present_in(&tmp));
        std::fs::remove_dir_all(&tmp).ok();
    }
}
