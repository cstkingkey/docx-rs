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

    /// Insert content immediately before the trailing `<w:sectPr>` if
    /// one exists; otherwise append. Use this instead of [`push`] when
    /// adding paragraphs/tables to a document that already has headers
    /// or footers wired in — appending after the trailing `<w:sectPr>`
    /// creates a second section, which Word treats as a section break
    /// and which detaches the new content from the configured
    /// headers/footers.
    ///
    /// [`push`]: Body::push
    pub fn push_before_section_property<T: Into<BodyContent<'a>>>(
        &mut self,
        content: T,
    ) -> &mut Self {
        if matches!(self.content.last(), Some(BodyContent::SectionProperty(_))) {
            let pos = self.content.len() - 1;
            self.content.insert(pos, content.into());
        } else {
            self.content.push(content.into());
        }
        self
    }

    /// Return the trailing `<w:sectPr>` of this body, creating an
    /// empty one if absent.
    ///
    /// Headers and footers are wired into a section via
    /// `<w:headerReference>` / `<w:footerReference>` children of a
    /// `<w:sectPr>`. This helper guarantees a trailing section
    /// property *at the moment it is called* so callers (typically
    /// [`Docx::add_header`] / [`Docx::add_footer`]) can attach a
    /// reference without walking the content vector.
    ///
    /// **Caveat:** [`Body::push`] always appends, so any body content
    /// added *after* this helper runs lands after the trailing
    /// `<w:sectPr>`, producing an unintended second section. Order of
    /// operations matters:
    ///
    /// * Add all body content (paragraphs, tables, SDTs) **first**.
    /// * Then call `add_header` / `add_footer` once everything is in.
    /// * If you must add body content after wiring headers/footers,
    ///   use [`Body::push_before_section_property`] instead of `push`.
    ///
    /// [`Docx::add_header`]: crate::Docx::add_header
    /// [`Docx::add_footer`]: crate::Docx::add_footer
    pub fn last_section_property_mut_or_create(&mut self) -> &mut SectionProperty<'a> {
        let last_is_sect_pr = matches!(self.content.last(), Some(BodyContent::SectionProperty(_)));
        if !last_is_sect_pr {
            self.content
                .push(BodyContent::SectionProperty(SectionProperty::default()));
        }
        match self.content.last_mut() {
            Some(BodyContent::SectionProperty(sp)) => sp,
            _ => unreachable!("just pushed or matched a SectionProperty"),
        }
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
