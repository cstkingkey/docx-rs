//! Embed a visible inline PNG into a docx.
//!
//! Generates a 32x32 solid red square at runtime using the `png`
//! dev-dependency, embeds it via the new ergonomic API, and writes
//! `image.docx` in the working directory. Open in Word: a red square
//! should appear inline with the surrounding text.
//!
//! Run with:
//!
//! ```sh
//! cargo run --example image
//! ```

use docx_rust::{
    document::{Paragraph, Run},
    media::{MediaType, Pic},
    Docx, DocxResult,
};

fn red_square_png(side: u32) -> Vec<u8> {
    let mut buf = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut buf, side, side);
        encoder.set_color(png::ColorType::Rgb);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().expect("png header");
        let pixels: Vec<u8> = (0..side * side)
            .flat_map(|_| [0xCC_u8, 0x22, 0x22])
            .collect();
        writer.write_image_data(&pixels).expect("png pixels");
    }
    buf
}

fn main() -> DocxResult<()> {
    let png = red_square_png(32);

    let mut docx = Docx::default();
    let rid = docx.add_image("square.png", MediaType::Image, &png);
    let drawing = Pic::new(rid).name("square").size_px(32, 32).into_drawing();

    let para = Paragraph::default()
        .push(Run::default().push_text("Behold a red square: "))
        .push(Run::default().push_image(drawing));
    docx.document.push(para);

    docx.write_file("image.docx")?;
    Ok(())
}
