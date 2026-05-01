//! End-to-end tests for header / footer ergonomics
//! (`Docx::add_header`, `Docx::add_footer`, `page_field`,
//! `num_pages_field`, `Run::push_tab`, and the
//! `SCHEMA_HEADER` -> `SCHEMA_FOOTER` bug fix in the write loop).

use std::io::{Cursor, Read};

use docx_rust::{
    document::{
        num_pages_field, page_field, Footer, Header, HeaderFooterReferenceType, Paragraph, Run,
    },
    Docx,
};
use zip::ZipArchive;

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

fn build_minimal_with_footer() -> Vec<u8> {
    let mut docx = Docx::default();
    let para = Paragraph::default().push(Run::default().push_text("body"));
    docx.document.push(para);

    let mut footer = Footer::default();
    footer.push(
        Paragraph::default()
            .push(Run::default().push_text("Page "))
            .push(page_field())
            .push(Run::default().push_text(" of "))
            .push(num_pages_field()),
    );
    let _rid = docx.add_footer(footer);

    write_to_zip(&mut docx)
}

#[test]
fn add_footer_writes_part_rel_override_and_section_reference() {
    let bytes = build_minimal_with_footer();
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    assert!(
        zip.by_name("word/footer1.xml").is_ok(),
        "footer1.xml missing"
    );

    let rels = read_part(&mut zip, "word/_rels/document.xml.rels");
    assert!(rels.contains(r#"Target="footer1.xml""#));
    assert!(
        rels.contains("/relationships/footer"),
        "footer rel schema missing: {}",
        rels
    );

    let cts = read_part(&mut zip, "[Content_Types].xml");
    assert!(
        cts.contains(r#"PartName="/word/footer1.xml""#)
            && cts.contains("wordprocessingml.footer+xml"),
        "footer Override missing: {}",
        cts
    );

    let doc = read_part(&mut zip, "word/document.xml");
    assert!(
        doc.contains("<w:footerReference") && doc.contains(r#"w:type="default""#),
        "footerReference missing in sectPr: {}",
        doc
    );
}

#[test]
fn footer_contains_page_and_numpages_field_codes() {
    let bytes = build_minimal_with_footer();
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    let footer = read_part(&mut zip, "word/footer1.xml");
    assert!(footer.contains(r#"<w:fldChar w:fldCharType="begin"/>"#));
    assert!(footer.contains(r#"<w:fldChar w:fldCharType="end"/>"#));
    assert!(
        footer.contains("PAGE") && footer.contains("NUMPAGES"),
        "field instructions missing: {}",
        footer
    );
    assert!(
        footer.contains(r#"xml:space="preserve""#),
        "instr text missing space=preserve (Word may strip leading/trailing spaces): {}",
        footer
    );
}

#[test]
fn add_header_writes_part_rel_override_and_section_reference() {
    let mut docx = Docx::default();
    let para = Paragraph::default().push(Run::default().push_text("body"));
    docx.document.push(para);

    let mut header = Header::default();
    header.push(Paragraph::default().push(Run::default().push_text("HEADER TEXT")));
    let _rid = docx.add_header(header);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    assert!(
        zip.by_name("word/header1.xml").is_ok(),
        "header1.xml missing"
    );

    let rels = read_part(&mut zip, "word/_rels/document.xml.rels");
    assert!(rels.contains(r#"Target="header1.xml""#));
    assert!(rels.contains("/relationships/header"));

    let cts = read_part(&mut zip, "[Content_Types].xml");
    assert!(cts.contains(r#"PartName="/word/header1.xml""#));
    assert!(cts.contains("wordprocessingml.header+xml"));

    let doc = read_part(&mut zip, "word/document.xml");
    assert!(doc.contains("<w:headerReference"));
}

#[test]
fn header_and_footer_can_coexist_in_one_section() {
    let mut docx = Docx::default();
    let para = Paragraph::default().push(Run::default().push_text("body"));
    docx.document.push(para);

    let mut header = Header::default();
    header.push(Paragraph::default().push(Run::default().push_text("H")));
    let h_rid = docx.add_header(header);

    let mut footer = Footer::default();
    footer.push(Paragraph::default().push(Run::default().push_text("F")));
    let f_rid = docx.add_footer(footer);

    assert_ne!(h_rid, f_rid, "header and footer rids must differ");

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    let doc = read_part(&mut zip, "word/document.xml");
    assert!(doc.contains("<w:headerReference"));
    assert!(doc.contains("<w:footerReference"));
}

#[test]
fn first_page_footer_uses_first_reference_type() {
    let mut docx = Docx::default();
    let para = Paragraph::default().push(Run::default().push_text("body"));
    docx.document.push(para);

    let mut footer = Footer::default();
    footer.push(Paragraph::default().push(Run::default().push_text("first-page-only")));
    docx.add_footer_with_type(footer, HeaderFooterReferenceType::First);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();
    let doc = read_part(&mut zip, "word/document.xml");
    assert!(
        doc.contains(r#"<w:footerReference w:type="first""#),
        "first-page footer ref type missing: {}",
        doc
    );
    // Without `<w:titlePg/>` Word silently ignores a `w:type="first"`
    // reference. The helper must auto-add it.
    assert!(
        doc.contains("<w:titlePg"),
        "first-page reference type without <w:titlePg/>: Word will ignore it. doc.xml: {}",
        doc
    );
}

#[test]
fn even_page_header_writes_even_and_odd_headers_setting() {
    let mut docx = Docx::default();
    docx.document
        .push(Paragraph::default().push(Run::default().push_text("body")));

    let mut header = Header::default();
    header.push(Paragraph::default().push(Run::default().push_text("even")));
    docx.add_header_with_type(header, HeaderFooterReferenceType::Even);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    // Without `<w:evenAndOddHeaders/>` in settings.xml Word silently
    // ignores `w:type="even"` references.
    let settings = zip
        .by_name("word/settings.xml")
        .map(|mut e| {
            let mut s = String::new();
            std::io::Read::read_to_string(&mut e, &mut s).unwrap();
            s
        })
        .expect("settings.xml must be written when an even-page reference is added");
    assert!(
        settings.contains("<w:evenAndOddHeaders"),
        "evenAndOddHeaders missing from settings.xml: {}",
        settings
    );
}

#[test]
fn registering_two_footers_of_the_same_type_replaces_the_old_reference() {
    let mut docx = Docx::default();
    docx.document
        .push(Paragraph::default().push(Run::default().push_text("body")));

    // First registration of type Default.
    let mut a = Footer::default();
    a.push(Paragraph::default().push(Run::default().push_text("v1")));
    docx.add_footer(a);

    // Second registration of the same type. Should replace the
    // section reference, not stack two of them.
    let mut b = Footer::default();
    b.push(Paragraph::default().push(Run::default().push_text("v2")));
    docx.add_footer(b);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();
    let doc = read_part(&mut zip, "word/document.xml");

    let count = doc
        .matches(r#"<w:footerReference w:type="default""#)
        .count();
    assert_eq!(
        count, 1,
        "expected exactly one default footerReference; got {} in doc.xml: {}",
        count, doc
    );
}

#[test]
fn two_footers_get_distinct_part_names() {
    let mut docx = Docx::default();
    let para = Paragraph::default().push(Run::default().push_text("body"));
    docx.document.push(para);

    let mut a = Footer::default();
    a.push(Paragraph::default().push(Run::default().push_text("default")));
    docx.add_footer(a);

    let mut b = Footer::default();
    b.push(Paragraph::default().push(Run::default().push_text("first")));
    docx.add_footer_with_type(b, HeaderFooterReferenceType::First);

    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();

    assert!(zip.by_name("word/footer1.xml").is_ok());
    assert!(zip.by_name("word/footer2.xml").is_ok());
}

#[test]
fn run_push_tab_emits_tab_element() {
    let mut docx = Docx::default();
    let para = Paragraph::default().push(
        Run::default()
            .push_text("left")
            .push_tab()
            .push_text("right"),
    );
    docx.document.push(para);
    let bytes = write_to_zip(&mut docx);
    let cursor = Cursor::new(bytes.as_slice());
    let mut zip = ZipArchive::new(cursor).unwrap();
    let doc = read_part(&mut zip, "word/document.xml");
    assert!(doc.contains("<w:tab/>"), "expected <w:tab/>: {}", doc);
}
