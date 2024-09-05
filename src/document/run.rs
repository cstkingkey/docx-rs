#![allow(unused_must_use)]
use derive_more::From;
use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{
    __setter, __xml_test_suites,
    document::{
        drawing::Drawing, field_char::FieldChar, instrtext::InstrText, r#break::Break,
        r#break::LastRenderedPageBreak, tab::Tab, text::Text,
    },
    formatting::CharacterProperty,
    DocxResult, __define_enum, __define_struct,
};

use super::{
    date::{DayLong, DayShort, MonthLong, MonthShort, YearLong, YearShort},
    instrtext::DelInstrText,
    sym::Sym,
    AnnotationRef, CarriageReturn, CommentReference, DelText, EndnoteRef, EndnoteReference,
    FootnoteRef, FootnoteReference,
};

/// Run
///
/// Run is a non-block region of text with properties.
///
/// ```rust
/// use docx_rust::document::*;
/// use docx_rust::formatting::*;
///
/// let run = Run::default()
///     .property(CharacterProperty::default())
///     .push_text("text")
///     .push_break(None)
///     .push_text((" text ", TextSpace::Preserve))
///     .push_break(BreakType::Column);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:r")]
pub struct Run<'a> {
    /// Specifies the properties of a run
    ///
    #[xml(attr = "w:rsidR")]
    pub rsid_r: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidRDefault")]
    pub rsid_r_default: Option<Cow<'a, str>>,
    /// Just as paragraph, a run's properties is applied to all the contents of the run.
    #[xml(child = "w:rPr")]
    pub property: Option<CharacterProperty<'a>>,
    #[xml(
        child = "w:br", //Break
        child = "w:t", //Text
        child = "w:delText", //Deleted Text
        child = "w:instrText", //Field Code
        child = "w:delInstrText", //Deleted Field Code
        child = "w:noBreakHyphen", //Non Breaking Hyphen Character
        child = "w:softHyphen", //Optional Hyphen Character
        child = "w:dayShort", //Date Block - Short Day Format
        child = "w:monthShort", //Date Block - Short Month Format
        child = "w:yearShort", //Date Block - Short Year Format
        child = "w:dayLong", //Date Block - Long Day Format
        child = "w:monthLong", //Date Block - Long Month Format
        child = "w:yearLong", //Date Block - Long Year Format
        child = "w:annotationRef", //Comment Information Block
        child = "w:footnoteRef", //Footnote Reference Mark
        child = "w:endnoteRef", //Endnote Reference Mark
        child = "w:separator", //Footnote/Endnote Separator Mark
        child = "w:continuationSeparator", //Continuation Separator Mark
        child = "w:sym", //Symbol Character
        child = "w:pgNum", //Page Number Block
        child = "w:cr", //Carriage Return
        child = "w:tab", //Tab Character
        //child = "w:object", //Inline Embedded Object
        //child = "w:pict", //VML Object
        child = "w:fldChar", //Complex Field Character
        //child = "w:ruby", //Phonetic Guide
        child = "w:footnoteReference", //Footnote Reference
        child = "w:endnoteReference", //Endnote Reference
        child = "w:commentReference", //Comment Content Reference Mark
        child = "w:drawing", //DrawingML Object
        child = "w:ptab", //Absolute Position Tab Character
        child = "w:lastRenderedPageBreak", //Position of Last Calculated Page Break
    )]
    /// Specifies the content of a run
    pub content: Vec<RunContent<'a>>,
}

impl<'a> Run<'a> {
    __setter!(property: Option<CharacterProperty<'a>>);

    #[inline(always)]
    pub fn push<T: Into<RunContent<'a>>>(mut self, content: T) -> Self {
        self.content.push(content.into());
        self
    }

    #[inline(always)]
    pub fn push_text<T: Into<Text<'a>>>(mut self, content: T) -> Self {
        self.content.push(RunContent::Text(content.into()));
        self
    }

    #[inline(always)]
    pub fn push_break<T: Into<Break>>(mut self, br: T) -> Self {
        self.content.push(RunContent::Break(br.into()));
        self
    }

    pub fn text(&self) -> String {
        self.iter_text()
            .map(|c| c.to_string())
            .collect::<Vec<_>>()
            .join("")
    }

    pub fn iter_text(&self) -> Box<dyn Iterator<Item = &Cow<'a, str>> + '_> {
        Box::new(self.content.iter().filter_map(|content| match content {
            RunContent::Text(Text { text, .. }) => Some(text),
            RunContent::InstrText(InstrText { text, .. }) => Some(text),
            RunContent::Break(_) => None,
            RunContent::LastRenderedPageBreak(_) => None,
            RunContent::FieldChar(_) => None,
            RunContent::Separator(_) => None,
            RunContent::ContinuationSeparator(_) => None,
            RunContent::Tab(_) => None,
            RunContent::CarriageReturn(_) => None,
            RunContent::Drawing(_) => None,
            _ => None,
        }))
    }

    pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
        self.content.iter_mut().filter_map(|content| match content {
            RunContent::Text(Text { text, .. }) => Some(text),
            RunContent::InstrText(InstrText { text, .. }) => Some(text),
            RunContent::Break(_) => None,
            RunContent::LastRenderedPageBreak(_) => None,
            RunContent::FieldChar(_) => None,
            RunContent::Separator(_) => None,
            RunContent::ContinuationSeparator(_) => None,
            RunContent::Tab(_) => None,
            RunContent::CarriageReturn(_) => None,
            RunContent::Drawing(_) => None,
            _ => None,
        })
    }

    pub fn replace_text_simple<S>(&mut self, old: S, new: S)
    where
        S: AsRef<str>,
    {
        let dic = (old, new);
        let dic = vec![dic];
        self.replace_text(&dic);
    }

    pub fn replace_text<'b, T, S>(&mut self, dic: T) -> DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = &'b (S, S)> + std::marker::Copy,
    {
        for c in self.content.iter_mut() {
            match c {
                RunContent::Text(t) => {
                    let mut tc = t.text.to_string();
                    for p in dic {
                        tc = tc.replace(p.0.as_ref(), p.1.as_ref());
                    }
                    t.text = tc.into();
                }
                _ => {}
            }
        }

        Ok(())
    }
}

/// A set of elements that can be contained as the content of a run.
#[derive(Debug, From, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum RunContent<'a> {
    #[xml(tag = "w:br")]
    Break(Break),
    #[xml(tag = "w:t")]
    Text(Text<'a>),
    #[xml(tag = "w:delText")]
    DelText(DelText<'a>),
    #[xml(tag = "w:instrText")]
    InstrText(InstrText<'a>),
    #[xml(tag = "w:delInstrText")]
    DelInstrText(DelInstrText<'a>),
    #[xml(tag = "w:noBreakHyphen")]
    NoBreakHyphen(NoBreakHyphen),
    #[xml(tag = "w:softHyphen")]
    SoftHyphen(SoftHyphen),
    #[xml(tag = "w:dayShort")]
    DayShort(DayShort),
    #[xml(tag = "w:monthShort")]
    MonthShort(MonthShort),
    #[xml(tag = "w:yearShort")]
    YearShort(YearShort),
    #[xml(tag = "w:dayLong")]
    DayLong(DayLong),
    #[xml(tag = "w:monthLong")]
    MonthLong(MonthLong),
    #[xml(tag = "w:yearLong")]
    YearLong(YearLong),
    #[xml(tag = "w:annotationRef")]
    AnnotationRef(AnnotationRef),
    #[xml(tag = "w:footnoteRef")]
    FootnoteRef(FootnoteRef),
    #[xml(tag = "w:endnoteRef")]
    EndnoteRef(EndnoteRef),
    #[xml(tag = "w:separator")]
    Separator(Separator),
    #[xml(tag = "w:continuationSeparator")]
    ContinuationSeparator(ContinuationSeparator),
    #[xml(tag = "w:sym")]
    Sym(Sym<'a>),
    #[xml(tag = "w:pgNum")]
    PgNum(PgNum),
    #[xml(tag = "w:cr")]
    CarriageReturn(CarriageReturn),
    #[xml(tag = "w:tab")]
    Tab(Tab),
    //#[xml(tag = "w:object")]
    //Object(Object<'a>),
    //#[xml(tag = "w:pict")]
    //Pict(Pict<'a>),
    #[xml(tag = "w:fldChar")]
    FieldChar(FieldChar),
    //#[xml(tag = "w:ruby")]
    //Ruby(Ruby<'a>),
    #[xml(tag = "w:footnoteReference")]
    FootnoteReference(FootnoteReference<'a>),
    #[xml(tag = "w:endnoteReference")]
    EndnoteReference(EndnoteReference<'a>),
    #[xml(tag = "w:commentReference")]
    CommentReference(CommentReference<'a>),
    #[xml(tag = "w:drawing")]
    Drawing(Drawing<'a>),
    #[xml(tag = "w:ptab")]
    PTab(PTab),
    #[xml(tag = "w:lastRenderedPageBreak")]
    LastRenderedPageBreak(LastRenderedPageBreak),
}

__define_struct! {
    ("w:ptab", PTab) {
        "w:alignment", alignment,	PTabAlignment	//Positional Tab Stop Alignment
        "w:relativeTo",	relative_to,	PTabRelativeTo	//Positional Tab Base
        "w:leader",	leader,	PTabLeader	//Tab Leader Character
    }
}

__define_enum! {
    PTabAlignment {
        Left = "left", // Left
        Center = "center", // Center
        Right = "right", // Right
    }
}

__define_enum! {
    PTabRelativeTo {
        Margin = "margin", // Relative To Text Margins
        Indent = "indent", // Relative To Indents
    }
}

__define_enum! {
    PTabLeader {
        None = "none", // No Leader Character
        Dot = "dot", // Dot Leader Character
        Hyphen = "hyphen", // Hyphen Leader Character
        Underscore = "underscore", // Underscore Leader Character
        MiddleDot = "middleDot", // Centered Dot Leader Character
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:noBreakHyphen")]
pub struct NoBreakHyphen;

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:softHyphen")]
pub struct SoftHyphen {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:separator")]
pub struct Separator {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:continuationSeparator")]
pub struct ContinuationSeparator {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pgNum")]
pub struct PgNum {}

__xml_test_suites!(
    Run,
    Run::default(),
    r#"<w:r/>"#,
    Run::default().push_break(None),
    r#"<w:r><w:br/></w:r>"#,
    Run::default().push_text("text"),
    r#"<w:r><w:t>text</w:t></w:r>"#,
);
