//! Browser-facing entry points for the `wasm32-unknown-unknown` build.
//!
//! Bytes in, text out — no file I/O, no terminal. This module is only
//! meaningful when built with `--no-default-features --features wasm-bindgen`,
//! which also disables `net` (model download) and `ocr` (inference) — so no
//! image OCR or textless-PDF fallback is available in the browser build. The
//! `ocr` parameter of `to_markdown` is accepted but is a no-op here.
//!
//! This module is only compiled for `wasm32` targets with the `wasm-bindgen`
//! cargo feature enabled; native builds never pull in `wasm-bindgen`.

#![cfg(target_arch = "wasm32")]

use crate::{
    detect_format, extract_markdown_with, extract_plain_with, extract_sheets_to, ExtractOptions,
    Sheet, SheetSink,
};
use js_sys::{Array, Function, Object, Reflect};
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
/// base64 data URIs; `ocr` is accepted for API compatibility but is a no-op
/// in this build (the `ocr` feature is off). Returns the Markdown, or
/// throws with a descriptive message.
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

/// Detect + extract tabular data (XLS / XLSX) as an array of sheet objects,
/// each `{ name, rows: [[cell, …], …] }`. `max_output_bytes` bounds the
/// payload estimate (sheet name bytes + per-cell `len+1`); pass `null` from
/// JS for unlimited.
///
/// This collecting path is O(total cells) in both the Rust and JS heaps and
/// is intended for small files only — prefer [`to_sheets_stream`] (or the
/// Rust rlib [`crate::SheetSink`] in a Worker) for large workbooks. Returns
/// the sheet array, or throws with a descriptive message.
#[wasm_bindgen]
pub fn to_sheets(data: &[u8], max_output_bytes: Option<u64>) -> Result<Array, String> {
    let format = detect_format(data).map_err(|e| e.to_string())?;
    let sheets = crate::extract_sheets_with(
        data,
        format,
        ExtractOptions {
            max_output_bytes,
            ..Default::default()
        },
    )
    .map_err(|e| e.to_string())?;
    Ok(sheets_to_js_array(&sheets))
}

fn sheets_to_js_array(sheets: &[Sheet]) -> Array {
    let out = Array::new();
    for sheet in sheets {
        let obj = Object::new();
        Reflect::set(
            &obj,
            &JsValue::from_str("name"),
            &JsValue::from_str(&sheet.name),
        )
        .ok();
        let rows = Array::new();
        for row in &sheet.rows {
            let cells = Array::new();
            for c in row {
                cells.push(&JsValue::from_str(c));
            }
            rows.push(&cells);
        }
        Reflect::set(&obj, &JsValue::from_str("rows"), &rows).ok();
        out.push(&obj);
    }
    out
}

struct JsSheetSink<'a> {
    on_begin: &'a Function,
    on_row: &'a Function,
    on_end: &'a Function,
}

impl SheetSink for JsSheetSink<'_> {
    fn begin_sheet(&mut self, name: &str) -> crate::Result<()> {
        self.on_begin
            .call1(&JsValue::NULL, &JsValue::from_str(name))
            .map_err(|e| crate::BatdocError::Document(format!("{e:?}")))?;
        Ok(())
    }
    fn row(&mut self, cells: Vec<String>) -> crate::Result<()> {
        let arr = Array::new();
        for c in cells {
            arr.push(&JsValue::from_str(&c));
        }
        self.on_row
            .call1(&JsValue::NULL, &arr)
            .map_err(|e| crate::BatdocError::Document(format!("{e:?}")))?;
        Ok(())
    }
    fn end_sheet(&mut self) -> crate::Result<()> {
        self.on_end
            .call0(&JsValue::NULL)
            .map_err(|e| crate::BatdocError::Document(format!("{e:?}")))?;
        Ok(())
    }
}

/// Detect + stream tabular data (XLS / XLSX) into JS callbacks, one sheet at
/// a time. `on_begin_sheet(name)`, `on_row(cells)`, and `on_end_sheet()` are
/// invoked synchronously on the CPU; Promises returned by the callbacks are
/// NOT awaited. Prefer the Rust rlib [`crate::SheetSink`] in a Worker for
/// large, long-running workbooks. `max_output_bytes` bounds the payload
/// estimate (see [`to_sheets`]); pass `null` from JS for unlimited. Returns
/// `undefined`, or throws with a descriptive message.
#[wasm_bindgen]
pub fn to_sheets_stream(
    data: &[u8],
    max_output_bytes: Option<u64>,
    on_begin_sheet: &Function,
    on_row: &Function,
    on_end_sheet: &Function,
) -> Result<(), String> {
    let format = detect_format(data).map_err(|e| e.to_string())?;
    let mut sink = JsSheetSink {
        on_begin: on_begin_sheet,
        on_row,
        on_end: on_end_sheet,
    };
    extract_sheets_to(
        data,
        format,
        ExtractOptions {
            max_output_bytes,
            ..Default::default()
        },
        &mut sink,
    )
    .map_err(|e| e.to_string())
}
