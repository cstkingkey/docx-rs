//! End-to-end tests for the image-embedding ergonomics
//! (`Docx::add_image`, `Pic`, `Run::push_image`).
//!
//! Tests serialise the document into an in-memory zip, then parse the
//! resulting parts as XML to assert presence of relationships, content
//! types, and the drawing chain.

use std::io::Cursor;

use docx_rust::{
    document::{Paragraph, Run},
    media::{MediaType, Pic},
    Docx,
};
use zip::ZipArchive;

/// Smallest valid PNG (1x1 transparent).
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0xF8, 0x0F, 0x04, 0x00,
    0x00, 0x09, 0xFB, 0x03, 0xFD, 0x8E, 0xB8, 0x8C, 0x7E, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn write_to_zip<'a>(docx: &'a mut Docx<'a>) -> Vec<u8> {
    let buf = Cursor::new(Vec::new());
    let result = docx.write(buf).expect("write docx");
    result.into_inner()
}

fn read_part(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> String {
    use std::io::Read;
    let mut entry = zip.by_name(name).expect(name);
    let mut s = String::new();
    entry.read_to_string(&mut s).unwrap();
    s
}

#[test]
fn single_image_registers_media_rel_and_content_type() {
    let png = PNG.to_vec();
    let mut docx = Docx::default();

    let rid = docx.add_image("dot.png", MediaType::Image, &png);
    let drawing = Pic::new(rid.clone()).size_px(16, 16).into_drawing();
    let para = Paragraph::default().push(Run::default().push_image(drawing));
    docx.document.push(para);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    assert!(
        zip.by_name("word/media/dot.png").is_ok(),
        "image bytes missing from zip"
    );

    let rels = read_part(&mut zip, "word/_rels/document.xml.rels");
    assert!(
        rels.contains(&format!(r#"Id="{}""#, rid)),
        "rel id {} missing in {}",
        rid,
        rels
    );
    assert!(rels.contains(r#"Target="media/dot.png""#));
    assert!(rels.contains("/relationships/image"));

    let cts = read_part(&mut zip, "[Content_Types].xml");
    assert!(
        cts.contains(r#"Extension="png""#) && cts.contains(r#"ContentType="image/png""#),
        "png Default not registered: {}",
        cts
    );

    let doc = read_part(&mut zip, "word/document.xml");
    assert!(doc.contains(r#"r:embed="rId1""#) || doc.contains(&format!(r#"r:embed="{}""#, rid)));
    assert!(doc.contains("<w:drawing>"));
    assert!(doc.contains("<pic:pic"));
}

#[test]
fn ten_images_get_distinct_rids_and_one_content_type() {
    let png = PNG.to_vec();
    let mut docx = Docx::default();

    let mut rids = Vec::new();
    for i in 0..10 {
        let name = format!("img{}.png", i);
        let rid = docx.add_image(name, MediaType::Image, &png);
        rids.push(rid.clone());
        let drawing = Pic::new(rid).doc_id(i + 1).into_drawing();
        let para = Paragraph::default().push(Run::default().push_image(drawing));
        docx.document.push(para);
    }

    let unique: std::collections::HashSet<_> = rids.iter().collect();
    assert_eq!(unique.len(), 10, "rids must be distinct: {:?}", rids);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    let cts = read_part(&mut zip, "[Content_Types].xml");
    let png_default_count = cts.matches(r#"Extension="png""#).count();
    assert_eq!(png_default_count, 1, "png Default duplicated: {}", cts);
}

#[test]
fn jpeg_extension_registers_image_jpeg_content_type() {
    let bytes = vec![0xFF, 0xD8, 0xFF, 0xE0]; // not a valid JPEG but irrelevant for the test
    let mut docx = Docx::default();
    let _rid = docx.add_image("photo.jpg", MediaType::Image, &bytes);

    let zip_bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(zip_bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    let cts = read_part(&mut zip, "[Content_Types].xml");
    assert!(cts.contains(r#"Extension="jpg""#));
    assert!(cts.contains(r#"ContentType="image/jpeg""#));
}

#[test]
fn pic_size_px_converts_to_emu() {
    let drawing = Pic::new("rId99").size_px(100, 50).into_drawing();
    let inline = drawing.inline.expect("inline");
    let extent = inline.extent.expect("extent");
    assert_eq!(extent.cx, 100 * 9_525);
    assert_eq!(extent.cy, 50 * 9_525);
}
