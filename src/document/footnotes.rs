//! FootNotes part
//!
//! The corresponding ZIP item is `/word/footnotes.xml`.
//!

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::borrow::Cow;
use std::io::Write;

use crate::schema::{SCHEMA_MAIN, SCHEMA_WORDML_14};
use crate::{__string_enum, __xml_test_suites};

use crate::document::BodyContent;

#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnotes")]
pub struct FootNotes<'a> {
    #[xml(child = "w:footnote")]
    pub content: Vec<FootNote<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnote")]
pub struct FootNote<'a> {
    #[xml(attr = "w:type")]
    pub ty: Option<NoteSeparator>,
    #[xml(attr = "w:id")]
    pub id: Option<isize>,
    #[xml(child = "w:sdt", child = "w:p", child = "w:tbl", child = "w:sectPr")]
    pub content: Vec<BodyContent<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnoteRef")]
pub struct FootnoteRef;

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnoteReference")]
pub struct FootnoteReference<'a> {
    /// Specifies the HeaderReference type of this HeaderReference.
    #[xml(attr = "w:customMarkFollows")]
    pub supress_reference_mark: Option<bool>,
    #[xml(attr = "w:id")]
    pub id: Option<Cow<'a, str>>,
}

impl<'a> FootNote<'a> {
    pub fn push<T: Into<BodyContent<'a>>>(&mut self, content: T) -> &mut Self {
        self.content.push(content.into());
        self
    }
}

impl<'a> XmlWrite for FootNotes<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let FootNotes { content } = self;

        log::debug!("[FootNotes] Started writing.");
        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:footnotes")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        for c in content {
            c.to_writer(writer)?;
        }

        writer.write_element_end_close("w:footnotes")?;

        log::debug!("[Document] Finished writing.");

        Ok(())
    }
}

#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum NoteSeparator {
    Separator,
    ContinuationSeparator,
    ContinuationNotice,
}

__string_enum! {
    NoteSeparator {
        Separator = "separator",
        ContinuationSeparator = "continuationSeparator",
        ContinuationNotice = "continuationNotice",
    }
}

__xml_test_suites!(
    FootNotes,
    FootNotes::default(),
    format!(
        r#"{}<w:footnotes xmlns:w="{}" xmlns:w14="{}"></w:footnotes>"#,
        crate::schema::SCHEMA_XML,
        SCHEMA_MAIN,
        SCHEMA_WORDML_14
    )
    .as_str(),
);
