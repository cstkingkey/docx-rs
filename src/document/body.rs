use derive_more::From;
use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Borrow;

use crate::__xml_test_suites;
use crate::document::{Paragraph, Run, Table, TableCell};
use crate::formatting::SectionProperty;

use super::SDT;

/// Document Body
///
/// This is the main document editing surface.
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:body")]
pub struct Body<'a> {
    /// Specifies the contents of the body of the document.
    #[xml(child = "w:p", child = "w:tbl", child = "w:sectPr", child = "w:sdt")]
    pub content: Vec<BodyContent<'a>>,
}

impl<'a> Body<'a> {
    pub fn push<T: Into<BodyContent<'a>>>(&mut self, content: T) -> &mut Self {
        self.content.push(content.into());
        self
    }

    pub fn text(&self) -> String {
        let v: Vec<_> = self
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(para) => Some(para.text()),
                BodyContent::Table(_) => None,
                BodyContent::SectionProperty(_) => None,
                BodyContent::Sdt(sdt) => Some(sdt.text()),
                BodyContent::TableCell(_) => None,
                BodyContent::Run(_) => None,
            })
            .collect();
        v.join("\r\n")
    }

    pub fn replace_text_simple<S>(&mut self, old: S, new: S)
    where
        S: AsRef<str>,
    {
        let _d = self.replace_text(&[(old, new)]);
    }

    pub fn replace_text<'b, I, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = I> + Copy,
        I: Borrow<(S, S)>,
    {
        for content in self.content.iter_mut() {
            match content {
                BodyContent::Paragraph(p) => {
                    p.replace_text(dic)?;
                }
                BodyContent::Table(t) => {
                    t.replace_text(dic)?;
                }
                BodyContent::SectionProperty(_) => {}
                BodyContent::Sdt(_) => {}
                BodyContent::TableCell(_) => {}
                BodyContent::Run(_) => {}
            }
        }
        Ok(())
    }

    // pub fn iter_text(&self) -> impl Iterator<Item = &Cow<'a, str>> {
    //     self.content
    //         .iter()
    //         .filter_map(|content| match content {
    //             BodyContent::Paragraph(para) => Some(para.iter_text()),
    //         })
    //         .flatten()
    // }

    // pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
    //     self.content
    //         .iter_mut()
    //         .filter_map(|content| match content {
    //             BodyContent::Paragraph(para) => Some(para.iter_text_mut()),
    //         })
    //         .flatten()
    // }
}

/// A set of elements that can be contained in the body
#[derive(Debug, From, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum BodyContent<'a> {
    #[xml(tag = "w:p")]
    Paragraph(Paragraph<'a>),
    #[xml(tag = "w:tbl")]
    Table(Table<'a>),
    #[xml(tag = "w:sdt")]
    Sdt(SDT<'a>),
    #[xml(tag = "w:sectPr")]
    SectionProperty(SectionProperty<'a>),
    #[xml(tag = "w:tc")]
    TableCell(TableCell<'a>),
    #[xml(tag = "w:r")]
    Run(Run<'a>),
}

__xml_test_suites!(
    Body,
    Body::default(),
    r#"<w:body/>"#,
    Body {
        content: vec![Paragraph::default().into()]
    },
    r#"<w:body><w:p/></w:body>"#,
    Body {
        content: vec![Table::default().into()]
    },
    r#"<w:body><w:tbl><w:tblPr/><w:tblGrid/></w:tbl></w:body>"#,
);
