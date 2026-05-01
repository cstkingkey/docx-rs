![GitHub Workflow Status](https://img.shields.io/github/actions/workflow/status/cstkingkey/docx-rs/test.yml?branch=master)
[![Crates.io](https://img.shields.io/crates/v/docx-rust.svg)](https://crates.io/crates/docx-rust)
[![Document](https://docs.rs/docx/badge.svg)](https://docs.rs/docx-rust)

# docx

A Rust library for parsing and generating docx files.

fork of https://github.com/PoiScript/docx-rs

[Document](https://docs.rs/docx-rust)

## Embedding inline images

```rust
use docx_rust::{Docx, document::{Paragraph, Run}, media::{MediaType, Pic}};
use std::fs;

let png = fs::read("logo.png").unwrap();

let mut docx = Docx::default();
let rid = docx.add_image("logo.png", MediaType::Image, &png);
let drawing = Pic::new(rid).size_px(120, 60).into_drawing();

let para = Paragraph::default()
    .push(Run::default().push_image(drawing));
docx.document.push(para);

docx.write_file("with-logo.docx").unwrap();
```

`Docx::add_image` registers the bytes, allocates a relationship id,
and auto-adds the matching `<Default>` Content Types entry (PNG, JPEG,
BMP, GIF, TIFF). `Pic` produces the `wp:inline` drawing chain. Use
`size_px` for 96-DPI pixel sizing or `size_emu` for direct EMU.

See [`examples/image.rs`](./examples/image.rs) for an end-to-end run.

## License

MIT
