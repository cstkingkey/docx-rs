//! End-to-end tests for SVG-with-fallback embedding via
//! `Docx::add_svg` + `Pic::with_svg`. The XML output must carry both
//! a standard `<a:blip r:embed="...">` (PNG fallback for legacy Word)
//! and an `<asvg:svgBlip r:embed="...">` extension (SVG for modern
//! Word), with distinct relationship ids.

use std::io::{Cursor, Read};

use docx_rust::{
    document::{Paragraph, Run},
    media::{MediaType, Pic, SvgImageIds},
    Docx,
};
use zip::ZipArchive;

/// Smallest valid PNG (1x1 transparent). Structural-only; rendering
/// is irrelevant to these tests.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0xF8, 0x0F, 0x04, 0x00,
    0x00, 0x09, 0xFB, 0x03, 0xFD, 0x8E, 0xB8, 0x8C, 0x7E, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

const SVG: &[u8] =
    br#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="red"/></svg>"#;

fn write_to_zip<'a>(docx: &'a mut Docx<'a>) -> Vec<u8> {
    let buf = Cursor::new(Vec::new());
    let result = docx.write(buf).expect("write docx");
    result.into_inner()
}

fn read_part(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> String {
    let mut entry = zip.by_name(name).expect(name);
    let mut s = String::new();
    entry.read_to_string(&mut s).unwrap();
    s
}

#[test]
fn add_svg_writes_both_parts_and_both_rels_and_both_content_types() {
    let svg = SVG.to_vec();
    let png = PNG.to_vec();
    let mut docx = Docx::default();

    let ids: SvgImageIds = docx.add_svg("logo", &svg, &png);
    assert_ne!(ids.svg_rid, ids.png_rid, "rids must differ");

    let drawing = Pic::with_svg(ids.clone()).size_px(64, 64).into_drawing();
    let para = Paragraph::default().push(Run::default().push_image(drawing));
    docx.document.push(para);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    assert!(
        zip.by_name("word/media/logo.svg").is_ok(),
        "svg part missing"
    );
    assert!(
        zip.by_name("word/media/logo.png").is_ok(),
        "png part missing"
    );

    let rels = read_part(&mut zip, "word/_rels/document.xml.rels");
    assert!(rels.contains(r#"Target="media/logo.svg""#));
    assert!(rels.contains(r#"Target="media/logo.png""#));
    assert!(rels.contains(&format!(r#"Id="{}""#, ids.svg_rid)));
    assert!(rels.contains(&format!(r#"Id="{}""#, ids.png_rid)));

    let cts = read_part(&mut zip, "[Content_Types].xml");
    assert!(
        cts.contains(r#"Extension="svg""#) && cts.contains(r#"ContentType="image/svg+xml""#),
        "svg Default missing: {}",
        cts
    );
    assert!(cts.contains(r#"Extension="png""#));
}

#[test]
fn drawing_xml_carries_both_blip_and_svg_blip_extension() {
    let svg = SVG.to_vec();
    let png = PNG.to_vec();
    let mut docx = Docx::default();

    let ids = docx.add_svg("logo", &svg, &png);
    let png_rid = ids.png_rid.clone();
    let svg_rid = ids.svg_rid.clone();
    let drawing = Pic::with_svg(ids).size_px(64, 64).into_drawing();
    let para = Paragraph::default().push(Run::default().push_image(drawing));
    docx.document.push(para);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    let doc = read_part(&mut zip, "word/document.xml");

    // Standard blip points at the PNG fallback.
    assert!(
        doc.contains(&format!(r#"<a:blip r:embed="{}""#, png_rid)),
        "standard a:blip should embed the PNG rid: {}",
        doc
    );

    // Extension list with SVG-blip URI.
    assert!(doc.contains("<a:extLst>"), "a:extLst missing");
    assert!(
        doc.contains("{96DAC541-7B7A-43D3-8B79-37D633B846F1}"),
        "SVG ext URI missing: {}",
        doc
    );

    // SVG blip points at the SVG rid.
    assert!(
        doc.contains(&format!(r#"r:embed="{}""#, svg_rid)),
        "asvg:svgBlip should embed the SVG rid: {}",
        doc
    );
    assert!(doc.contains("<asvg:svgBlip"), "<asvg:svgBlip> tag missing");
    assert!(
        doc.contains("xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\""),
        "asvg namespace declaration missing: {}",
        doc
    );
}

#[test]
fn raster_only_image_does_not_emit_svg_extension() {
    let png = PNG.to_vec();
    let mut docx = Docx::default();
    let rid = docx.add_image("plain.png", MediaType::Image, &png);
    let drawing = Pic::new(rid).size_px(32, 32).into_drawing();
    docx.document
        .push(Paragraph::default().push(Run::default().push_image(drawing)));

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();
    let doc = read_part(&mut zip, "word/document.xml");
    assert!(
        !doc.contains("<a:extLst>"),
        "plain raster path should not emit a:extLst"
    );
    assert!(!doc.contains("<asvg:svgBlip"));
}

#[cfg(feature = "svg-rasterize")]
#[test]
fn rasterize_svg_produces_valid_png_bytes() {
    let bytes = docx_rust::media::rasterize_svg(SVG, 64, 64).expect("rasterise");
    assert!(bytes.starts_with(&[0x89, 0x50, 0x4E, 0x47]), "not a PNG");
    // Sanity-check the rasterised PNG round-trips through add_svg.
    let svg = SVG.to_vec();
    let mut docx = Docx::default();
    let ids = docx.add_svg("logo", &svg, &bytes);
    assert_ne!(ids.svg_rid, ids.png_rid);
}
