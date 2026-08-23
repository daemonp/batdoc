#!/usr/bin/env bash
# Build the wasm module and generate the browser JS glue into web/pkg/.
set -euo pipefail
cd "$(dirname "$0")"
ROOT="$(cd .. && pwd)"

cargo build \
  --target wasm32-unknown-unknown \
  --release \
  -p batdoc-core \
  --no-default-features \
  --manifest-path "$ROOT/Cargo.toml"

wasm-bindgen \
  --target web \
  --out-dir pkg \
  "$ROOT/target/wasm32-unknown-unknown/release/batdoc_core.wasm"

echo "Built web/pkg/ — serve this directory over HTTP to demo (e.g. python3 -m http.server)."
