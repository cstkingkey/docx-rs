//! Comments part
//!
//! The corresponding ZIP item is `/word/numbering.xml`.
#![allow(unused_must_use)]

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::{borrow::Cow, io::Write};

use crate::{
    formatting::{CharacterProperty, Indent},
    schema::{SCHEMA_MAIN, SCHEMA_WORDML_14},
};

#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numbering")]
pub struct Numbering<'a> {
    #[xml(child = "w:abstractNum")]
    pub abstract_nums: Vec<AbstractNum<'a>>,
    #[xml(child = "w:num")]
    pub nums: Vec<Num<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:abstractNum")]
pub struct AbstractNum<'a> {
    #[xml(attr = "w:abstractNumId")]
    pub abstract_num_id: Cow<'a, str>,
    #[xml(child = "w:nsid")]
    pub nsid: Nsid<'a>,
    #[xml(child = "w:multiLevelType")]
    pub multi_level_type: MultiLevelType<'a>,
    #[xml(child = "w:lvl")]
    pub levels: Vec<Level<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:nsid")]
pub struct Nsid<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:multiLevelType")]
pub struct MultiLevelType<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvl")]
pub struct Level<'a> {
    #[xml(attr = "w:ilvl")]
    pub ilvl: Option<isize>,
    #[xml(child = "w:numFmt")]
    pub number_format: Option<NumFmt<'a>>,
    #[xml(child = "w:lvlText")]
    pub level_text: Option<LevelText<'a>>,
    #[xml(child = "w:lvlJc")]
    pub level_jc: Option<LevelJc<'a>>,
    #[xml(child = "w:pPr")]
    pub w_pr: Option<PPr>,
    #[xml(child = "w:rPr")]
    pub character: Option<CharacterProperty<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pPr")]
pub struct PPr {
    #[xml(child = "w:ind")]
    pub indent: Option<Indent>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numFmt")]
/// TODO Replace by enum NumberFormat
pub struct NumFmt<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvlText")]
pub struct LevelText<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvlJc")]
pub struct LevelJc<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:num")]
pub struct Num<'a> {
    #[xml(attr = "w:numId")]
    pub num_id: Cow<'a, str>,
    #[xml(child = "w:abstractNumId")]
    pub abstract_num_id: Option<AbstractNumId<'a>>,
    #[xml(child = "w:lvlOverride")]
    pub lvl_overrides: Vec<LvlOverride<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvlOverride")]
pub struct LvlOverride<'a> {
    #[xml(attr = "w:ilvl")]
    pub ilvl: Cow<'a, str>,
    #[xml(child = "w:startOverride")]
    pub start_override: StartOverride<'a>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:startOverride")]
pub struct StartOverride<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:abstractNumId")]
pub struct AbstractNumId<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

impl<'a> XmlWrite for Numbering<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        log::debug!("[Numbering] Started writing.");

        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:numbering")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        writer.write_element_end_close("w:comments")?;

        log::debug!("[Comments] Finished writing.");

        Ok(())
    }
}
