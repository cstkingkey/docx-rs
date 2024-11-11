//! EndNotes part
//!
//! The corresponding ZIP item is `/word/endnotes.xml`.
//!

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::borrow::Cow;
use std::io::Write;

use crate::__xml_test_suites;
use crate::schema::{SCHEMA_MAIN, SCHEMA_WORDML_14};

use crate::document::BodyContent;

use super::NoteSeparator;

#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnotes")]
pub struct EndNotes<'a> {
    #[xml(child = "w:endnote")]
    pub content: Vec<EndNote<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnote")]
pub struct EndNote<'a> {
    #[xml(attr = "w:type")]
    pub ty: Option<NoteSeparator>,
    #[xml(attr = "w:id")]
    pub id: Option<isize>,
    #[xml(child = "w:sdt", child = "w:p", child = "w:tbl", child = "w:sectPr")]
    pub content: Vec<BodyContent<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnoteRef")]
pub struct EndnoteRef;

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnoteReference")]
pub struct EndnoteReference<'a> {
    /// Specifies the HeaderReference type of this HeaderReference.
    #[xml(attr = "w:customMarkFollows")]
    pub supress_reference_mark: Option<bool>,
    #[xml(attr = "w:id")]
    pub id: Option<Cow<'a, str>>,
}

impl<'a> EndNote<'a> {
    pub fn push<T: Into<BodyContent<'a>>>(&mut self, content: T) -> &mut Self {
        self.content.push(content.into());
        self
    }
}

impl<'a> XmlWrite for EndNotes<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let EndNotes { content } = self;

        log::debug!("[EndNotes] Started writing.");
        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:endnotes")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        for c in content {
            c.to_writer(writer)?;
        }

        writer.write_element_end_close("w:endnotes")?;

        log::debug!("[Document] Finished writing.");

        Ok(())
    }
}

__xml_test_suites!(
    EndNotes,
    EndNotes::default(),
    format!(
        r#"{}<w:endnotes xmlns:w="{}" xmlns:w14="{}"></w:endnotes>"#,
        crate::schema::SCHEMA_XML,
        SCHEMA_MAIN,
        SCHEMA_WORDML_14
    )
    .as_str(),
);
