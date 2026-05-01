//! Embed an SVG in a docx with a caller-supplied PNG raster fallback.
//!
//! Run with:
//!
//! ```sh
//! cargo run --example svg
//! ```
//!
//! Output: `svg.docx` in the working directory. Open in Word 2019+ /
//! Microsoft 365 — the red circle should render as crisp vector
//! graphics. Open in Word 2013 — the same circle renders from the
//! PNG fallback (slightly blurry at very large zoom).

use docx_rust::{
    document::{Paragraph, Run},
    media::Pic,
    Docx, DocxResult,
};

const SVG: &[u8] =
    br#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="50" cy="50" r="45" fill="crimson"/><text x="50" y="58" font-size="32" text-anchor="middle" fill="white" font-family="sans-serif">SVG</text></svg>"#;

/// Encode a `side x side` solid-color RGB PNG to use as a raster
/// fallback. In a real app the fallback would typically be a
/// rasterised version of the SVG itself.
fn solid_png(side: u32, rgb: [u8; 3]) -> Vec<u8> {
    let mut buf = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut buf, side, side);
        encoder.set_color(png::ColorType::Rgb);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().expect("png header");
        let pixels: Vec<u8> = (0..side * side).flat_map(|_| rgb).collect();
        writer.write_image_data(&pixels).expect("png pixels");
    }
    buf
}

fn main() -> DocxResult<()> {
    let svg = SVG.to_vec();
    let png = solid_png(96, [0xDC, 0x14, 0x3C]); // crimson placeholder

    let mut docx = Docx::default();
    let ids = docx.add_svg("logo", &svg, &png);
    let drawing = Pic::with_svg(ids)
        .name("logo")
        .size_px(96, 96)
        .into_drawing();

    let para = Paragraph::default()
        .push(Run::default().push_text("SVG with raster fallback: "))
        .push(Run::default().push_image(drawing));
    docx.document.push(para);

    docx.write_file("svg.docx")?;
    Ok(())
}
