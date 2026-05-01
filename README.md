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

## Headers, footers, and page-number fields

```rust
use docx_rust::{
    document::{num_pages_field, page_field, Footer, Header, Paragraph, Run},
    Docx,
};

let mut docx = Docx::default();

let mut header = Header::default();
header.push(Paragraph::default().push(Run::default().push_text("Confidential")));
docx.add_header(header);

let mut footer = Footer::default();
footer.push(
    Paragraph::default()
        .push(Run::default().push_text("Workbench Export"))
        .push(Run::default().push_tab())
        .push(Run::default().push_text("Page "))
        .push(page_field())
        .push(Run::default().push_text(" of "))
        .push(num_pages_field()),
);
docx.add_footer(footer);

docx.write_file("with-header-and-footer.docx").unwrap();
```

`Docx::add_header` / `Docx::add_footer` allocate the next `headerN.xml` /
`footerN.xml` slot, register the relationship, add the matching
`<Override>` Content Types entry, and attach a `<w:headerReference>` /
`<w:footerReference>` to the trailing `<w:sectPr>` (creating one if
absent). The `_with_type` variants take a
`HeaderFooterReferenceType::{Default, Even, First}` for first-page-only
or even-page variants.

`page_field()` and `num_pages_field()` return runs containing the
`{ PAGE }` and `{ NUMPAGES }` field codes respectively. Word
substitutes the live values at render time. `field_run(instr)` is the
escape hatch for arbitrary field instructions.

See [`examples/header_footer.rs`](./examples/header_footer.rs) for an
end-to-end run.

## SVG with raster fallback

Word 2016+ renders embedded SVG vectorially via the `asvg:svgBlip`
extension on the standard `<a:blip>`. Older Word ignores the extension
and falls back to the raster image. Both companion files live in
`word/media/`, both get relationships, both extensions are registered
in `[Content_Types].xml`.

```rust
use docx_rust::{Docx, document::{Paragraph, Run}, media::Pic};
use std::fs;

let svg = fs::read("logo.svg").unwrap();
let png = fs::read("logo.png").unwrap();   // pre-rasterised fallback

let mut docx = Docx::default();
let ids = docx.add_svg("logo", &svg, &png);
let drawing = Pic::with_svg(ids).size_px(120, 60).into_drawing();

let para = Paragraph::default()
    .push(Run::default().push_image(drawing));
docx.document.push(para);
docx.write_file("with-svg.docx").unwrap();
```

The default build does not bundle an SVG rasteriser — callers supply
the PNG fallback themselves. For one-call convenience, enable the
opt-in `svg-rasterize` feature, which pulls in `resvg` + `usvg` +
`tiny-skia`:

```toml
[dependencies]
docx-rust = { version = "0.2", features = ["svg-rasterize"] }
```

```rust
use docx_rust::Docx;
use docx_rust::media::{rasterize_svg, Pic};

let svg = std::fs::read("logo.svg")?;
let png = rasterize_svg(&svg, 240, 120)?;   // 2x display size for crispness

let mut docx = Docx::default();
let ids = docx.add_svg("logo", &svg, &png);
let drawing = Pic::with_svg(ids).size_px(120, 60).into_drawing();
```

See [`examples/svg.rs`](./examples/svg.rs) (caller-supplied) and
[`examples/svg_auto.rs`](./examples/svg_auto.rs) (feature-gated).

## License

MIT
