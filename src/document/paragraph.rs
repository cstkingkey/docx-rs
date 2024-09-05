#![allow(unused_must_use)]
use derive_more::From;
use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{
    __setter, __xml_test_suites,
    document::{
        BookmarkEnd, BookmarkStart, CommentRangeEnd, CommentRangeStart, Hyperlink, Run, RunContent,
        Text, SDT,
    },
    formatting::ParagraphProperty,
};

/// Paragraph
///
/// Paragraph is the main block-level container for content.
///
/// ```rust
/// use docx_rust::document::*;
/// use docx_rust::formatting::*;
///
/// let par = Paragraph::default()
///     .property(ParagraphProperty::default())
///     .push_text("hello,")
///     .push_text((" world.", TextSpace::Preserve))
///     .push(Run::default())
///     .push(BookmarkStart::default())
///     .push(BookmarkEnd::default());
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:p")]
pub struct Paragraph<'a> {
    //#[xml(attr = "w14:paraId")]
    //pub id: Option<Cow<'a, str>>,
    //#[xml(attr = "w14:textId")]
    //pub text_id: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidR")]
    pub rsid_r: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidRDefault")]
    pub rsid_r_default: Option<Cow<'a, str>>,
    /// Specifies the properties of a paragraph
    ///
    /// This information is applied to all the contents of the paragraph.
    #[xml(child = "w:pPr")]
    pub property: Option<ParagraphProperty<'a>>,
    /// Specifes the run contents of a paragraph
    ///
    /// Run is a region of text with properties. Each paragraph containes one or more runs.
    #[xml(
        child = "w:commentRangeStart",
        child = "w:commentRangeEnd",
        child = "w:r",
        child = "w:hyperlink",
        child = "w:bookmarkStart",
        child = "w:bookmarkEnd",
        child = "w:sdt"
    )]
    pub content: Vec<ParagraphContent<'a>>,
}

impl<'a> Paragraph<'a> {
    __setter!(property: Option<ParagraphProperty<'a>>);

    #[inline(always)]
    pub fn push<T: Into<ParagraphContent<'a>>>(mut self, content: T) -> Self {
        self.content.push(content.into());
        self
    }

    #[inline(always)]
    pub fn push_text<T: Into<Text<'a>>>(mut self, content: T) -> Self {
        self.content.push(ParagraphContent::Run(Run {
            content: vec![RunContent::Text(content.into())],
            ..Default::default()
        }));
        self
    }

    pub fn text(&self) -> String {
        self.iter_text()
            .map(|c| c.to_string())
            .collect::<Vec<_>>()
            .join("")
    }

    pub fn iter_text(&self) -> Box<dyn Iterator<Item = &Cow<'a, str>> + '_> {
        Box::new(
            self.content
                .iter()
                .filter_map(|content| match content {
                    ParagraphContent::Run(run) => Some(run.iter_text()),
                    ParagraphContent::Link(link) => Some(link.content.iter_text()),
                    ParagraphContent::SDT(sdt) => Some(sdt.iter_text()),
                    _ => None,
                })
                .flatten()
        )
    }

    pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
        self.content
            .iter_mut()
            .filter_map(|content| match content {
                ParagraphContent::Run(run) => Some(run.iter_text_mut()),
                ParagraphContent::Link(link) => Some(link.content.iter_text_mut()),
                _ => None,
            })
            .flatten()
    }

    pub fn replace_text<'b, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = &'b (S, S)> + std::marker::Copy,
    {
        for content in self.content.iter_mut() {
            match content {
                ParagraphContent::Run(r) => {
                    r.replace_text(dic)?;
                }
                _ => {}
            }
        }

        Ok(())
    }
}

/// A set of elements that can be contained as the content of a paragraph.
#[derive(Debug, From, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ParagraphContent<'a> {
    #[xml(tag = "w:commentRangeStart")]
    CommentRangeStart(CommentRangeStart<'a>),
    #[xml(tag = "w:commentRangeEnd")]
    CommentRangeEnd(CommentRangeEnd<'a>),
    #[xml(tag = "w:r")]
    Run(Run<'a>),
    #[xml(tag = "w:hyperlink")]
    Link(Hyperlink<'a>),
    #[xml(tag = "w:bookmarkStart")]
    BookmarkStart(BookmarkStart<'a>),
    #[xml(tag = "w:bookmarkEnd")]
    BookmarkEnd(BookmarkEnd<'a>),
    #[xml(tag = "w:sdt")]
    SDT(SDT<'a>),
}

__xml_test_suites!(
    Paragraph,
    Paragraph::default(),
    r#"<w:p/>"#,
    Paragraph::default().push(Run::default()),
    r#"<w:p><w:r/></w:p>"#,
    Paragraph::default().push(Hyperlink::default()),
    r#"<w:p><w:hyperlink><w:r/></w:hyperlink></w:p>"#,
    Paragraph::default().push(BookmarkStart::default()),
    r#"<w:p><w:bookmarkStart/></w:p>"#,
    Paragraph::default().push(BookmarkEnd::default()),
    r#"<w:p><w:bookmarkEnd/></w:p>"#,
);
