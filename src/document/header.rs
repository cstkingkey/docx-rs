//! Header part
//!
//! The corresponding ZIP item is `/word/header{n}.xml`.
//!

use std::io::Write;
use strong_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};

use crate::__xml_test_suites;
use crate::schema::{SCHEMA_MAIN, SCHEMA_WORDML_14};

use crate::document::BodyContent;

/// The root element of the main document part.
#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:hdr")]
pub struct Header<'a> {
    #[xml(child = "w:p", child = "w:tbl", child = "w:sectPr", child = "w:sdt")]
    pub content: Vec<BodyContent<'a>>,
}

impl<'a> Header<'a> {
    pub fn push<T: Into<BodyContent<'a>>>(&mut self, content: T) -> &mut Self {
        self.content.push(content.into());
        self
    }
    pub fn replace_text_simple<S>(&mut self, old: S, new: S)
    where
        S: AsRef<str>,
    {
        let dic = (old, new);
        let dic = vec![dic];
        let _d = self.replace_text(&dic);
    }

    pub fn replace_text<'b, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = &'b (S, S)> + std::marker::Copy,
    {
        for content in self.content.iter_mut() {
            match content {
                BodyContent::Paragraph(p) => {
                    p.replace_text(dic)?;
                }
                BodyContent::Table(_) => {}
                BodyContent::SectionProperty(_) => {}
                BodyContent::Sdt(_) => {}
                BodyContent::TableCell(_) => {}
            }
        }
        Ok(())
    }
}

impl<'a> XmlWrite for Header<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let Header { content } = self;

        log::debug!("[Header] Started writing.");
        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:hdr")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        for c in content {
            c.to_writer(writer)?;
        }

        writer.write_element_end_close("w:hdr")?;

        log::debug!("[Document] Finished writing.");

        Ok(())
    }
}

__xml_test_suites!(
    Header,
    Header::default(),
    format!(
        r#"{}<w:hdr xmlns:w="{}" xmlns:w14="{}"></w:hdr>"#,
        crate::schema::SCHEMA_XML,
        SCHEMA_MAIN,
        SCHEMA_WORDML_14
    )
    .as_str(),
);
