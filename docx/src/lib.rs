//! A Rust library for parsing and generating docx files.
//!
//! ## Example
//!
//! Using methods [`from_file`] and [`write_file`] for reading from and writing to file directly.
//!
//! [`from_file`]: struct.Docx.html#method.from_file
//! [`write_file`]: struct.Docx.html#method.write_file
//!
//! ```no_run
//! use docx::Docx;
//! use docx::document::Para;
//!
//! // reading docx from file
//! let mut docx = Docx::from_file("demo.docx").unwrap();
//!
//! // do what you want to do…
//! // for example, appending something
//! let mut para = Para::default();
//! para.text("Lorem Ipsum");
//! docx.insert_para(para);
//!
//! // writing back to the original file
//! docx.write_file("demo.docx").unwrap();
//! ```
//!
//! Alternatively, you can use [`parse`] (accepts [`Read`] + [`Seek`]) and [`generate`] (accepts [`Write`] + [`Seek`]).
//!
//! [`parse`]: struct.Docx.html#method.parse
//! [`generate`]: struct.Docx.html#method.generate
//! [`Read`]: std::io::Read
//! [`Write`]: std::io::Write
//! [`Seek`]: std::io::Seek
//!
//! ## Glossary
//!
//! Some terms used in this crate.
//!
//! * Body: main surface for editing
//! * Paragraph: block-level container of content, begins with a new line
//! * Run(Character): non-block region of text
//! * Style: a set of paragraph and character properties which can be applied to text
//!
//! ## Note
//!
//! ### Toggle Properties
//!
//! Some fields in this crate (e.g. [`bold`] and [`italics`]) are declared as `Option<bool>` and this is not redundant at all.
//!
//! This indicates that they are **toggle properties** which can be **inherited** from style (`None`) or **disabled/enabled explicitly** (`Some`).
//!
//! [`bold`]: style/struct.CharStyle.html#structfield.bold
//! [`italics`]: style/struct.CharStyle.html#structfield.italics
//!
//! For example, you can disable bold of a run within a paragraph specifies bold by setting it to `Some(false)`:
//!
//! ```rust
//! use docx::Docx;
//! use docx::document::{Para, Run};
//!
//! let mut docx = Docx::default();
//!
//! docx
//!   .create_style()
//!   .name("Normal")
//!   .char()
//!   .bold(true)
//!   .italics(true);
//!
//! let mut para = Para::default();
//! para.prop().name("Normal");
//!
//! para.text("I'm bold and italics.").text_break();
//!
//! let mut run = Run::text("poi");
//! run.prop().bold(false).italics(false);
//! para.run(run);
//!
//! docx.insert_para(para);
//! ```

#[macro_use]
extern crate docx_codegen;
#[macro_use]
extern crate log;
extern crate quick_xml;
extern crate zip;

#[macro_use]
mod macros;

pub mod app;
pub mod content_type;
pub mod core;
pub mod document;
mod docx;
pub mod errors;
pub mod font_table;
pub mod rels;
mod schema;
pub mod style;

pub mod prelude {
    //! Prelude module

    pub use crate::document::{Para, Run};
    pub use crate::docx::Docx;
    pub use crate::style::Style;
}

pub use crate::docx::Docx;
pub use crate::errors::{Error, Result};