extern crate docx_rust;
use docx_rust::DocxFile;

#[test]
fn read_and_replace() {
    // reader
    let path = std::path::Path::new("./tests/aaa/aa.docx");
    let book = DocxFile::from_file(path).unwrap();
    let mut docx = book.parse().unwrap();
    docx.document.body.replace_text_simple("好日子", "好天气");

    let path2 = std::path::Path::new("./tests/bbb/aa.docx");
    let _d = docx.write_file(path2).unwrap();
}

#[test]
fn read_list() {
    let path = std::path::Path::new("./tests/aaa/aa_list.docx");
    let book = DocxFile::from_file(path).unwrap();
    let docx = book.parse().unwrap();
    let text = docx.document.body.text();
    assert_eq!("Dock。 List \r\nTest list\r\nNano editor\r\nTest\r\nNano", text);
}