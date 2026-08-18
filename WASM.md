# Compiling batdoc to WebAssembly

Status: **implemented.** `batdoc-core` builds a browser-loadable
`cdylib` for `wasm32-unknown-unknown` with `--no-default-features`, exposes
`wasm-bindgen` entry points (`detect`, `to_plain`, `to_markdown`), and ships a
working browser demo in `web/` (see [Demo](#demo)). Native builds are
unchanged.

## Why you'd care

`batdoc-core` is a pure-Rust document-to-text/OCR library (docx/xlsx/pptx/doc/xls/
pdf → plain text or markdown). The entire *document* pipeline (via `lopdf`,
`pdf-extract`, `zip`, `cfb`, `image`) is wasm-clean. Only the OCR model
**download** path and the terminal-oriented **CLI binary** are host-only. That
means a wasm build of the library — e.g. an in-browser doc viewer/extractor — is
very achievable with modest, well-scoped changes.

## Findings (empirically verified)

All of the following came from actually running
`cargo check --target wasm32-unknown-unknown` against the workspace, not from
reading dependency READMEs.

### State (2026-08-17)

1. **Done — model download is feature-gated behind `net`.** `ureq` is now an
   `optional` dependency behind a `net` feature (default-on). Building with
   `--no-default-features` drops `ureq`/`rustls`/`ring`, and the download
   branch of `ocr.rs::ensure_file` is compiled out — it only checks for
   pre-seeded model files and returns a descriptive error otherwise. Model
   loading/inference (`ocrs`/`rten`) are untouched.
2. **Done — `getrandom` wasm backend configured.** `.cargo/config.toml` sets
   `getrandom_backend="wasm_js"` for `wasm32-unknown-unknown`, and the
   `batdoc-core` manifest enables `getrandom/0.3`'s `wasm_js` feature on the
   wasm target only. `getrandom 0.2` disappears with `ring` (state 1).
3. **Done — wasm-facing entry points + browser demo.** `batdoc-core` is now a
   `rlib`+`cdylib`; `src/wasm.rs` (wasm32-only) exports `detect`, `to_plain`,
   and `to_markdown` via `wasm-bindgen`. The `web/` demo builds the `.wasm`,
   generates the JS glue, and runs entirely in the browser (see [Demo](#demo)).

### Good news

- **`rayon` compiles for `wasm32-unknown-unknown`.** Tested standalone: a tiny
  crate using `rayon::prelude::ParallelIterator` (`.par_iter().sum()`) builds and
  type-checks for wasm. rayon-core now ships a single-threaded wasm path.
- **`ocrs` is wasm-aware by design.** Its `Cargo.toml` declares a dedicated
  `[target.'cfg(target_arch = "wasm32")'.dependencies.wasm-bindgen]` and the
  author ships in-browser builds, so `ocrs` + `rten` (the OCR engine) are expected
  to build on wasm32 once the broader graph is unblocked.
- **`getrandom` is a config issue, not a rewrite** (see blockers).

### Blockers to a plain `cargo check --target wasm32-unknown-unknown`

1. **`ureq`/`rustls`/`ring` → startup/clock failure (hard blocker) — RESOLVED.**
   `rustls` fails to compile for wasm because it needs `UnixTime::now()` /
   `std::time::SystemTime` (needs a clock). This whole stack exists **only** for
   `batdoc-core/src/ocr.rs::ensure_file`, which downloads the two OCR model
   files (`.rten`) over HTTP on first use. On wasm there is no socket/clock, so
   this path is gated out by the `net` feature (see State 1).

2. **`getrandom` (two versions) needs the wasm backend config — RESOLVED.**
   - `getrandom 0.2` — pulled via `ring → rustls → ureq` (gone with `net` off).
   - `getrandom 0.3` — pulled via `rand → lopdf`.
   Both reject wasm unless you pass a config flag *and* enable a feature:
   ```toml
   [target.wasm32-unknown-unknown]
   rustflags = "--cfg getrandom_backend=\"wasm_js\""
   ```
   plus the `wasm_js` feature (0.3) / `js` feature (0.2) enabled on the
   dependency. Now wired up in `.cargo/config.toml` and the `batdoc-core`
   manifest (see State 2).

3. **The CLI crate is host-shaped.** `src/main.rs` uses `std::fs`, stdin/stdout,
   `is-terminal`, and the `bat` pager; `build.rs` synthesizes a man page. A
   `batdoc` **binary** on wasm is meaningless. Only `batdoc-core` the **library**
   is a realistic target, and it currently exposes no wasm-friendly entry point.

4. **The OCR model download is network-at-runtime.** On wasm the model files are
   not download-and-cache; they must be embedded at build time or supplied
   out-of-band (mirror the existing `BATDOC_MODELS_DIR` seeding, but baked in).

## Reproduction

```sh
# add the wasm target once
rustup target add wasm32-unknown-unknown

# compile the library for wasm (model download is feature-gated off)
cargo check --target wasm32-unknown-unknown -p batdoc-core --no-default-features

# produce the rlib (links; verifies codegen end-to-end)
cargo build --target wasm32-unknown-unknown --release -p batdoc-core --no-default-features

# build the browser demo (wasm-bindgen glue + .wasm into web/pkg/)
./web/build.sh
```

## Porting plan (whenever someone picks this up)

Ordered so each step is independently shippable and keeps the native CLI intact.

1. **Make networking optional — DONE.** Added the `net` feature (default-on)
   in `batdoc-core`:
   - `ureq` is `optional` behind `net = ["dep:ureq"]`, so
     `--no-default-features` removes `ureq` (and `rustls`/`ring`), and
   - `#[cfg(feature = "net")]` gates the download branch of `ocr.rs::ensure_file`;
     without it the path only looks up pre-seeded model files (or returns a
     "models must be provided/embedded" error). Model *loading* and inference
     (`ocrs`/`rten`) remain wasm-safe.

2. **Fix `getrandom` for wasm — DONE.** Added `.cargo/config.toml` with the
   `getrandom_backend="wasm_js"` rustflag and enabled `getrandom/0.3`'s
   `wasm_js` feature on the wasm target only. The 0.2 instance disappeared with
   `ring`; only the 0.3 one (via `lopdf`) remains.

3. **Expose a library entry point — DONE.** `batdoc-core` is now
   `crate-type = ["rlib", "cdylib"]`, and `src/wasm.rs` (compiled only for
   `wasm32`) exports `detect`, `to_plain`, and `to_markdown` via `wasm-bindgen`,
   each delegating to the existing `detect_format` / `extract_plain_with` /
   `extract_markdown_with` API. Native builds never compile `wasm-bindgen`.

4. **Verify + demo — DONE.** `web/build.sh` builds the wasm cdylib and runs
   `wasm-bindgen --target web`; `web/index.html` + `web/demo.js` load it and
   convert a chosen file (or the generated `samples/*.docx`, `*.pdf`) to
   markdown/plain text entirely in the browser. See [Demo](#demo).

## Scope notes

- The **document/OCR accuracy** code paths are cross-platform and need no changes.
- Only the **model acquisition** (`ensure_file`, `ureq`) and the **CLI shell**
  (`src/main.rs`, `build.rs`) are wasm-hostile.
- Native builds (Linux/macOS/Windows) are unchanged: `net` is a **default-on**
  feature, so `ureq` and the download path are present exactly as before. Only
  `--no-default-features` (or wasm, which needs it) opts out. The `getrandom`
  `wasm_js` feature is enabled on the wasm target only, so native is untouched.

### Memory (streaming extract)

Non-OCR extract now streams through `ExtractSink` (`extract_*_to`); the old
`extract_*` functions are collecting wrappers. Measured on a synthetic
13.45 MB XLSX (300,000 rows × 8 cols = 2,400,000 cells, 20,000 shared strings,
deflate-9) with `resource.getrusage(RUSAGE_CHILDREN).ru_maxrss`:

| file size | cells | plain RSS (old → new) | markdown RSS (old → new) |
|-----------|-------|------------------------|--------------------------|
| 13.45 MB  | 2.4 M | 326.5 → 109.8 MiB      | 694.1 → 114.9 MiB        |

The 64 MiB goal is **not yet met**: peak RSS is still ~110 MiB because the CLI
writes output via the buffered `String` API, so the ~70 MB of extracted text
is materialized in full before hitting stdout. A follow-up is to wire
`src/main.rs` to `extract_*_to(..., &mut IoSink(stdout))`, which drops the
output copy and should bring peak RSS under 64 MiB.

For Workers (128 MiB default memory limit), a 15 MB XLSX is currently
borderline through the `String` API. Prefer `extract_*_to` with an
incremental sink, and set `ExtractOptions.max_output_bytes` to bound output
(the budget error is `"output exceeded {n} bytes"`).

## Demo

The browser demo lives in `web/`. It builds `batdoc-core` as a wasm cdylib,
emits the `wasm-bindgen` glue, and serves a single static page that converts a
file to markdown/plain text client-side.

Prereqs: `rustup target add wasm32-unknown-unknown` and
`cargo install wasm-bindgen-cli --version 0.2.108`.

```sh
./web/make_samples.py   # once: write web/samples/{sample.docx,sample.pdf}
./web/build.sh          # build wasm + JS glue into web/pkg/

# serve (wasm must be fetched over http, not file://) and open the browser
cd web && python3 -m http.server 8000    # then open http://localhost:8000/
```

Notes:

- `web/pkg/` and `web/samples/` are generated and git-ignored.
- The exported functions are `detect(data)`, `to_plain(data)`, and
  `to_markdown(data, images, ocr)` — `data` is a `Uint8Array`; errors throw a
  JS `Error` with the `BatdocError` message.
- This build disables `net`, so OCR model files are **not** downloaded — they
  must be seeded (e.g. preloaded into the cache dir the library resolves). The
  document/text pipeline (DOCX/XLSX/PPTX/DOC/XLS/PDF-text) needs no models and
  works out of the box.

## Related

- Source of the full dependency graph: `cargo tree -p batdoc-core`,
  `cargo tree -i getrandom`, `cargo tree -i rayon`.
- OCR engine: [ocrs](https://github.com/robertknight/ocrs), [rten](https://github.com/robertknight/rten).
