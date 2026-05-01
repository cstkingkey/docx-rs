//! Field-code helpers for page numbering in headers and footers.
//!
//! Word represents a simple field like `{ PAGE }` as a sequence of
//! three run-content elements:
//!
//! 1. `<w:fldChar w:fldCharType="begin"/>`
//! 2. `<w:instrText> PAGE </w:instrText>`
//! 3. `<w:fldChar w:fldCharType="end"/>`
//!
//! These helpers build a single [`Run`] containing that sequence,
//! ready to push into a [`Paragraph`](crate::document::Paragraph). Word
//! recomputes the displayed value on open or when the user presses F9.
//!
//! ```rust
//! use docx_rust::document::{Paragraph, Run, page_field, num_pages_field};
//!
//! let footer_para = Paragraph::default()
//!     .push(Run::default().push_text("Page "))
//!     .push(page_field())
//!     .push(Run::default().push_text(" of "))
//!     .push(num_pages_field());
//! ```

use std::borrow::Cow;

use crate::document::{
    field_char::{CharType, FieldChar},
    instrtext::{InstrText, TextSpace},
    Run, RunContent,
};

/// A run containing a `{ PAGE }` field. Word substitutes the current
/// page number at render time.
pub fn page_field<'a>() -> Run<'a> {
    field_run(" PAGE ")
}

/// A run containing a `{ NUMPAGES }` field. Word substitutes the
/// total page count at render time.
pub fn num_pages_field<'a>() -> Run<'a> {
    field_run(" NUMPAGES ")
}

/// Build a run holding an arbitrary simple Word field. The provided
/// instruction string is wrapped in `begin`/`end` field-char markers
/// with `xml:space="preserve"` so leading and trailing spaces survive.
///
/// Use [`page_field`] / [`num_pages_field`] for the common cases.
pub fn field_run<'a>(instr: impl Into<Cow<'a, str>>) -> Run<'a> {
    let mut run = Run::default();
    run.content.push(RunContent::FieldChar(FieldChar {
        ty: Some(CharType::Begin),
    }));
    run.content.push(RunContent::InstrText(InstrText {
        space: Some(TextSpace::Preserve),
        text: instr.into(),
    }));
    run.content.push(RunContent::FieldChar(FieldChar {
        ty: Some(CharType::End),
    }));
    run
}
