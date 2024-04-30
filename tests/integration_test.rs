extern crate docx_rust;
use std::fs::read_dir;

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
    assert_eq!(
        "Dock。 List \r\nTest list\r\nNano editor\r\nTest\r\nNano",
        text
    );
}

#[test]
fn read_pandocs() {
    if let Ok(dir) = read_dir("./tests/pandoc/") {
        for entry in dir {
            if let Ok(entry) = entry {
                let path = entry.path();
                // Check if the entry is a file
                if path.is_file() {
                    match DocxFile::from_file(path) {
                        Ok(docx_file) => {
                            // Process the DocxFile as needed
                            match docx_file.parse() {
                                Ok(_) => assert!(true),
                                Err(err) => assert!(false, "Error processing file: {:?}", err),
                            }
                        }
                        Err(err) => {
                            // Handle the error if DocxFile::from_file() fails
                            assert!(false, "Error processing file: {:?}", err);
                        }
                    }
                }
            }
        }
    }
}
