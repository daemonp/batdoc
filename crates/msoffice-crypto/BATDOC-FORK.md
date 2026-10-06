# batdoc fork of msoffice-crypto 0.1.0-rc.5

Upstream: <https://github.com/Slurp9187/msoffice-crypto> (`0.1.0-rc.5`, MIT OR Apache-2.0).
Not affiliated with Microsoft. License texts are `LICENSE-MIT` and `LICENSE-APACHE`;
attribution for the ported encrypt path is `NOTICE`.

## Why this copy exists

`cfb` 0.14 timestamps its directory entries with `web_time::SystemTime`. Off wasm that
type is `std::time::SystemTime`. On `wasm32-unknown-unknown` it is a different type, so
upstream `cfb_zero_time() -> std::time::SystemTime` does not compile: `set_created_time`
/ `set_modified_time` reject it.

`src/dataspaces.rs` now returns `web_time::SystemTime`. Native ciphertext is unchanged,
because `web_time` re-exports `std` there and the CFB zero (1601-01-01) still subtracts
from the Unix epoch. This repo patches the crates.io crate via `[patch.crates-io]` so
`batdoc-core` links the same Office decrypt path on wasm as on native.

## What wasm still cannot do

`web_time`'s wasm `SystemTime` cannot represent a time before the Unix epoch
(`checked_sub` returns `None`). `cfb_zero_time` therefore panics if encrypt runs on
wasm. Decrypt does not call it. Encrypted legacy `.doc` / `.xls` stay out of scope
(`legacy-binary` is off). A proper encrypt fix belongs in `cfb`, which would need to
write a zero FILETIME without a pre-epoch `SystemTime`.
