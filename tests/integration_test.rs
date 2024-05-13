extern crate docx_rust;
use std::fs::read_dir;

use docx_rust::{
    document::{BodyContent, ParagraphContent, RunContent},
    DocxFile,
};

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

#[test]
fn read_image() {
    let path = std::path::Path::new("./tests/pandoc/image.docx");
    let book = DocxFile::from_file(path).unwrap();
    let docx = book.parse().unwrap();
    let mut is_first = true;
    for content in docx.document.body.content {
        if !is_first {
            return ();
        }
        match content {
            BodyContent::Paragraph(paragraph) => {
                for para_content in paragraph.content {
                    match para_content {
                        ParagraphContent::Run(run) => {
                            for run_content in run.content {
                                match run_content {
                                    RunContent::Drawing(drawing) => {
                                        is_first = false;
                                        if let Some(inline) = drawing.inline {
                                            if let Some(extent) = inline.extent {
                                                assert_eq!(1905000, extent.cx);
                                                assert_eq!(1905000, extent.cy);
                                            }
                                            if let Some(graphic) = inline.graphic {
                                                if let Some(cnvpr) =
                                                    graphic.data.pic.nv_pic_pr.c_nv_pr
                                                {
                                                    assert_eq!("lalune.jpg", cnvpr.descr.unwrap());
                                                    assert_eq!(22, cnvpr.id.unwrap());
                                                }
                                                assert_eq!(
                                                    "rId20",
                                                    graphic.data.pic.fill.blip.embed
                                                );
                                                if let Some(relationships) = &docx.document_rels {
                                                    if let Some(target) =
                                                        relationships.get_target("rId20")
                                                    {
                                                        assert_eq!("media/rId20.jpg", target);
                                                    } else {
                                                        assert!(false)
                                                    }
                                                }
                                            } else {
                                                assert!(false)
                                            }
                                        }
                                        ()
                                    }
                                    _ => (),
                                }
                            }
                            ()
                        }
                        _ => (),
                    }
                }
                ()
            }
            _ => (),
        }
    }
}
