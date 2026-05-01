//! Embed an SVG in a docx with an auto-rasterised PNG fallback,
//! using the opt-in `svg-rasterize` cargo feature.
//!
//! Run with:
//!
//! ```sh
//! cargo run --example svg_auto --features svg-rasterize
//! ```
//!
//! Without the feature this example is a no-op stub that prints a
//! helpful message — keeping the default-feature build honest about
//! what's available.

#[cfg(feature = "svg-rasterize")]
fn main() -> docx_rust::DocxResult<()> {
    use docx_rust::{
        document::{Paragraph, Run},
        media::{rasterize_svg, Pic},
        Docx,
    };

    const SVG: &[u8] = br#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" fill="steelblue"/><circle cx="50" cy="50" r="30" fill="white"/></svg>"#;

    let svg = SVG.to_vec();
    let png = rasterize_svg(&svg, 192, 192).expect("rasterise SVG");

    let mut docx = Docx::default();
    let ids = docx.add_svg("logo", &svg, &png);
    let drawing = Pic::with_svg(ids)
        .name("logo")
        .size_px(96, 96)
        .into_drawing();

    let para = Paragraph::default()
        .push(Run::default().push_text("Auto-rasterised SVG: "))
        .push(Run::default().push_image(drawing));
    docx.document.push(para);

    docx.write_file("svg_auto.docx")?;
    Ok(())
}

#[cfg(not(feature = "svg-rasterize"))]
fn main() {
    eprintln!(
        "This example requires the `svg-rasterize` cargo feature. \
         Run again with: `cargo run --example svg_auto --features svg-rasterize`."
    );
}
