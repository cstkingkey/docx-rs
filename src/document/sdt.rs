#![allow(unused_must_use)]
use std::borrow::Cow;

use strong_xml::{XmlRead, XmlWrite};

use crate::{__setter, __xml_test_suites};

use super::BodyContent;

/// SDT
///
// /// ```rust
// /// use docx_rust::document::*;
// /// use docx_rust::formatting::*;
// ///
// /// let tbl = SDT::default()
// ///     .property(SDTProperty::default())
// ///     .push_grid(vec![1, 2, 3])
// ///     .push_grid(SDTGrid::default())
// ///     .push_row(SDTRow::default());
// /// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sdt")]
pub struct SDT<'a> {
    #[xml(child = "w:sdtPr")]
    pub property: Option<SDTProperty<'a>>,
    #[xml(child = "w:sdtEndPr")]
    pub end_property: Option<SDTEndProperty>,
    #[xml(child = "w:sdtContent")]
    pub content: Option<SDTContent<'a>>,
}

impl<'a> SDT<'a> {
    __setter!(property: Option<SDTProperty<'a>>);
    __setter!(end_property: Option<SDTEndProperty>);
    __setter!(content: Option<SDTContent<'a>>);
}

/// Section Property
///
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sdtPr")]
pub struct SDTProperty<'a> {
    #[xml(child = "w:id")]
    pub id: Option<STDId>,
    #[xml(child = "w:docPartObj")]
    pub doc_part_obj: Option<DocPartObj<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:id")]
pub struct STDId {
    #[xml(attr = "w:val")]
    pub id: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:docPartObj")]
pub struct DocPartObj<'a> {
    #[xml(child = "w:docPartGallery")]
    pub doc_part_gallery: Option<DocPartGallery<'a>>,
    #[xml(child = "w:docPartUnique")]
    pub doc_part_unique: Option<DocPartUnique>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:docPartGallery")]
pub struct DocPartGallery<'a> {
    #[xml(attr = "w:val")]
    pub name: Option<Cow<'a, str>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:docPartUnique")]
pub struct DocPartUnique {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sdtEndPr")]
pub struct SDTEndProperty {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sdtContent")]
pub struct SDTContent<'a> {
    #[xml(child = "w:p", child = "w:tc", child = "w:tbl", child = "w:sectPr", child = "w:sdt")]
    pub content: Vec<BodyContent<'a>>,
}

__xml_test_suites!(SDT, SDT::default(), "<w:sdt/>",);
