//! A Rust library for parsing and generating docx files.
//!
//! # Create a new document
//!
//! Use `Docx::default` to create a new empty `Docx`, then use
//! [`Docx::write_file`] for saving it to a file.
//!
//! [`Docx::write_file`]: struct.Docx.html#method.write_file
//!
//! ```no_run
//! use docx_rust::document::Paragraph;
//! use docx_rust::Docx;
//!
//! let mut docx = Docx::default();
//!
//! // create a new paragraph and insert it
//! let para = Paragraph::default().push_text("Lorem Ipsum");
//! docx.document.push(para);
//!
//! docx.write_file("demo.docx").unwrap();
//! ```
//!
//! Also see: [`Docx::write`].
//!
//! [`Docx::write`]: struct.Docx.html#method.write
//!
//! # Reading from files
//!
//! Use [`DocxFile::from_file`] to extract content from docx files, then use
//! [`DocxFile::parse`] to generate a `Docx` struct.
//!
//! [`DocxFile::from_file`]: struct.DocxFile.html#method.from_file
//! [`DocxFile::parse`]: struct.DocxFile.html#method.parse
//!
//! ```no_run
//! use docx_rust::document::Paragraph;
//! use docx_rust::DocxFile;
//!
//! let docx = DocxFile::from_file("origin.docx").unwrap();
//! let mut docx = docx.parse().unwrap();
//!
//! let para = Paragraph::default().push_text("Lorem Ipsum");
//! docx.document.push(para);
//!
//! docx.write_file("origin_appended.docx").unwrap();
//! ```
//!
//! To reduce allocations, `DocxFile::parse` returns a `Docx` struct contains
//! references to `DocxFile` itself. It means you have to make sure that
//! `DocxFile` lives as long as its returned `Docx`:
//!
//! ```compile_fail
//! use docx_rust::DocxFile;
//!
//! let mut docx_option = None;
//! {
//!     let docx_file = DocxFile::from_file("foo.docx").unwrap();
//!     let mut docx = docx_file.parse().unwrap();
//!     docx_option = Some(docx);
//!     // `docx_file` gets dropped here and code fails to compile
//! }
//! docx_option.unwrap().write_file("foo.docx").unwrap();
//! ```
//!
//! Also see: [`DocxFile::from_reader`].
//!
//! [`DocxFile::from_reader`]: struct.DocxFile.html#method.from_reader
//!
//! # Similar Projects
//!
//! [`bokuweb/docx-rs`]: A .docx file writer with Rust/WebAssembly.
//!
//! [`bokuweb/docx-rs`]: https://github.com/bokuweb/docx-rs
//!
//! # License
//!
//! MIT
//!

mod macros;

pub mod app;
pub mod content_type;
pub mod core;
pub mod document;
mod docx;
mod error;
pub mod font_table;
pub mod formatting;
pub mod media;
pub mod rels;
mod schema;
pub mod settings;
pub mod styles;
pub mod web_settings;

use std::io::Write;

use hard_xml::{XmlWrite, XmlWriter};

pub use crate::docx::{Docx, DocxFile};
pub use crate::error::{DocxError, DocxResult};

pub mod rounded_float {
    use std::num::ParseFloatError;
    pub fn from_xml(mode: &str) -> hard_xml::XmlResult<isize> {
        let f: f64 = mode
            .parse()
            .map_err(|e: ParseFloatError| hard_xml::XmlError::FromStr(e.into()))?;

        let r = f
            .is_finite()
            .then_some(f.round())
            .ok_or_else(|| hard_xml::XmlError::FromStr("f64 must be finite".into()))?;

        hard_xml::XmlResult::Ok(r as isize)
    }
    pub fn to_xml(mode: &isize) -> hard_xml::XmlResult<String> {
        Ok(mode.to_string())
    }
}

pub fn write_attr<W: Write, T: XmlWrite>(
    element: &Option<T>,
    writer: &mut XmlWriter<W>,
) -> Result<(), hard_xml::XmlError> {
    if let Some(e) = element {
        e.to_writer(writer)?;
    };
    Ok(())
}
