//! Browser-facing entry points for the `wasm32-unknown-unknown` build.
//!
//! Bytes in, text out — no file I/O, no terminal. The OCR *model download*
//! path is feature-gated off (`--no-default-features`); OCR still works when
//! model files have been seeded via `BATDOC_MODELS_DIR` (or embedded by the
//! caller before invoking these functions).
//!
//! This module is only compiled for `wasm32` targets with the `wasm-bindgen`
//! cargo feature enabled; native builds never pull in `wasm-bindgen`.

#![cfg(target_arch = "wasm32")]

use crate::{detect_format, extract_markdown_with, extract_plain_with, ExtractOptions};
use wasm_bindgen::prelude::*;

/// Detect the document format from magic bytes and return its name
/// (`DOC`, `XLS`, `DOCX`, `XLSX`, `PPTX`, `PDF`, `IMAGE`) or an error string.
#[wasm_bindgen]
pub fn detect(data: &[u8]) -> String {
    detect_format(data)
        .map(|f| f.to_string())
        .unwrap_or_else(|e| e.to_string())
}

/// Detect + extract plain text. `data` is the raw file bytes (a `Uint8Array`
/// from JS). Returns the extracted text, or throws with a descriptive message.
#[wasm_bindgen]
pub fn to_plain(data: &[u8]) -> Result<String, String> {
    let format = detect_format(data).map_err(|e| e.to_string())?;
    extract_plain_with(data, format, ExtractOptions::default()).map_err(|e| e.to_string())
}

/// Detect + extract Markdown. `images` embeds DOCX/XLSX/PPTX images as
/// base64 data URIs; `ocr` OCRs DOCX/PPTX embedded images (needs seeded
/// models). Returns the Markdown, or throws with a descriptive message.
#[wasm_bindgen]
pub fn to_markdown(data: &[u8], images: bool, ocr: bool) -> Result<String, String> {
    let format = detect_format(data).map_err(|e| e.to_string())?;
    extract_markdown_with(
        data,
        format,
        ExtractOptions {
            images,
            ocr,
            ..Default::default()
        },
    )
    .map_err(|e| e.to_string())
}
