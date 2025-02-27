//! Comments part
//!
//! The corresponding ZIP item is `/word/numbering.xml`.
#![allow(unused_must_use)]

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::{borrow::Cow, io::Write};

use crate::{
    formatting::{CharacterProperty, Indent, JustificationVal},
    schema::{SCHEMA_MAIN, SCHEMA_WORDML_14},
};

#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numbering")]
/// Numbering defines ordered and unordered lists.
pub struct Numbering<'a> {
    #[xml(child = "w:abstractNum")]
    /// Abstract numberings are not used directly when laying out your document.
    /// Instead, they are referred to by the numberings.
    pub abstract_numberings: Vec<AbstractNum<'a>>,
    #[xml(child = "w:num")]
    /// Numberings are used by your document and refer to abstract numberings for layout.
    pub numberings: Vec<Num>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:abstractNum")]
pub struct AbstractNum<'a> {
    #[xml(attr = "w:abstractNumId")]
    pub abstract_num_id: Option<isize>,
    #[xml(child = "w:nsid")]
    pub nsid: Option<Nsid<'a>>,
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
    pub i_level: Option<isize>,
    #[xml(child = "w:start")]
    pub start: Option<LevelStart>,
    #[xml(child = "w:numFmt")]
    pub number_format: Option<NumFmt<'a>>,
    #[xml(child = "w:lvlText")]
    pub level_text: Option<LevelText<'a>>,
    #[xml(child = "w:lvlJc")]
    pub justification: Option<LevelJustification>,
    #[xml(child = "w:pPr")]
    pub p_pr: Option<PPr>,
    #[xml(child = "w:rPr")]
    pub r_pr: Vec<CharacterProperty<'a>>,
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
#[xml(tag = "w:start")]
/// TODO Replace by enum NumberFormat
pub struct LevelStart {
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvlText")]
pub struct LevelText<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvlJc")]
pub struct LevelJustification {
    #[xml(attr = "w:val")]
    pub value: JustificationVal,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:num")]
pub struct Num {
    #[xml(attr = "w:numId")]
    pub num_id: Option<isize>,
    #[xml(child = "w:abstractNumId")]
    pub abstract_num_id: Option<AbstractNumId>,
    #[xml(child = "w:lvlOverride")]
    pub level_overrides: Vec<LevelOverride>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:lvlOverride")]
pub struct LevelOverride {
    #[xml(attr = "w:ilvl")]
    pub i_level: Option<isize>,
    #[xml(child = "w:startOverride")]
    pub start_override: Option<StartOverride>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:startOverride")]
pub struct StartOverride {
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:abstractNumId")]
pub struct AbstractNumId {
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
}

impl<'a> Numbering<'a> {
    /// Actual numberings refer to abstract numberings, and may overrule some settings.
    /// This helper function takes an numbering id that is provided in a paragraph, looks up
    /// the details in the numbering section and merges it with the abstract numbering to get
    /// a complete AbstractNum object.
    pub fn numbering_details(&self, id: isize) -> Option<AbstractNum> {
        self.numberings.iter().find_map(|n| {
            if n.num_id != Some(id) || n.abstract_num_id.is_none() {
                None
            } else {
                if let Some(abstract_num_id) = &n.abstract_num_id {
                    if let Some(abstract_numbering) = self
                        .abstract_numberings
                        .iter()
                        .find(|an| an.abstract_num_id == abstract_num_id.value)
                    {
                        let mut an = abstract_numbering.clone();
                        n.level_overrides.iter().for_each(|o| {
                            let LevelOverride {
                                i_level,
                                start_override,
                            } = o;
                            if i_level.is_some() && start_override.is_some() {
                                if let Some(level) =
                                    an.levels.iter_mut().find(|level| level.i_level == *i_level)
                                {
                                    level.start = Some(LevelStart {
                                        value: start_override.as_ref().unwrap().value.clone(),
                                    });
                                }
                            }
                        });
                        return Some(an);
                    }
                }

                None
            }
        })
    }
}

impl<'a> XmlWrite for Numbering<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let Numbering {
            abstract_numberings: abstract_nums,
            numberings: nums,
        } = self;

        log::debug!("[Numbering] Started writing.");

        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:numbering")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        for an in abstract_nums {
            an.to_writer(writer)?;
        }

        for num in nums {
            num.to_writer(writer)?;
        }

        writer.write_element_end_close("w:numbering")?;

        log::debug!("[Numbering] Finished writing.");

        Ok(())
    }
}

#[cfg(test)]

const NUMBERING_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="990">
            <w:nsid w:val="0000A990" />
            <w:multiLevelType w:val="multilevel" />
            <w:lvl w:ilvl="0">
                <w:numFmt w:val="bullet" />
                <w:lvlText w:val=" " />
                <w:lvlJc w:val="left" />
                <w:pPr>
                    <w:ind w:left="720" w:hanging="360" />
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="1">
                <w:numFmt w:val="bullet" />
                <w:lvlText w:val=" " />
                <w:lvlJc w:val="left" />
                <w:pPr>
                    <w:ind w:left="1440" w:hanging="360" />
                </w:pPr>
            </w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="99411">
            <w:nsid w:val="00A99411" />
            <w:multiLevelType w:val="multilevel" />
            <w:lvl w:ilvl="0">
                <w:start w:val="1" />
                <w:numFmt w:val="decimal" />
                <w:lvlText w:val="%1." />
                <w:lvlJc w:val="left" />
                <w:pPr>
                    <w:ind w:left="720" w:hanging="360" />
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="1">
                <w:start w:val="1" />
                <w:numFmt w:val="decimal" />
                <w:lvlText w:val="%2." />
                <w:lvlJc w:val="left" />
                <w:pPr>
                    <w:ind w:left="1440" w:hanging="360" />
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="2">
                <w:start w:val="1" />
                <w:numFmt w:val="decimal" />
                <w:lvlText w:val="%3." />
                <w:lvlJc w:val="left" />
                <w:pPr>
                    <w:ind w:left="2160" w:hanging="360" />
                </w:pPr>
            </w:lvl>
        </w:abstractNum>
        <w:num w:numId="1000">
            <w:abstractNumId w:val="990" />
        </w:num>
        <w:num w:numId="1001">
            <w:abstractNumId w:val="99411" />
            <w:lvlOverride w:ilvl="0">
                <w:startOverride w:val="1" />
            </w:lvlOverride>
            <w:lvlOverride w:ilvl="1">
                <w:startOverride w:val="1" />
            </w:lvlOverride>
        </w:num>
    </w:numbering>
"#;

#[test]
fn xml_parsing() {
    let numbering = Numbering::from_str(NUMBERING_XML).unwrap();
    assert_eq!(numbering.abstract_numberings.len(), 2);
    assert_eq!(numbering.numberings.len(), 2);
    assert_eq!(numbering.abstract_numberings[0].nsid.value, "0000A990");
    assert_eq!(
        numbering.abstract_numberings[0].levels[0]
            .number_format
            .as_ref()
            .unwrap()
            .value,
        "bullet"
    );
    assert_eq!(
        numbering.numberings[0]
            .abstract_num_id
            .as_ref()
            .unwrap()
            .value,
        Some(990_isize)
    );
}

#[test]
fn find_numbering_details() {
    let numbering = Numbering::from_str(NUMBERING_XML).unwrap();
    if let Some(num) = numbering.numbering_details(
        numbering.numberings[0]
            .abstract_num_id
            .as_ref()
            .unwrap()
            .value
            .unwrap(),
    ) {
        assert_eq!(
            num.levels[0].number_format,
            Some(NumFmt {
                value: Cow::Borrowed("bullet")
            })
        );
    }
    if let Some(num) = numbering.numbering_details(1001) {
        assert_eq!(
            num.levels[0].number_format,
            Some(NumFmt {
                value: Cow::Borrowed("decimal")
            })
        );
        assert_eq!(
            num.levels[1].level_text,
            Some(LevelText {
                value: Cow::Borrowed("%2.")
            })
        );
    }
}

#[test]
fn xml_writing() {
    fn replace_whitespace(input: &str, replacement: &str) -> String {
        let mut result = String::new();
        let mut last_was_whitespace_or_bracket = false;

        for c in input.chars() {
            if c.is_whitespace() {
                if !last_was_whitespace_or_bracket {
                    result.push_str(replacement);
                    last_was_whitespace_or_bracket = true;
                }
            } else {
                result.push(c);
                last_was_whitespace_or_bracket = if c == '>' || c == '"' { true } else { false };
            }
        }

        result
    }

    let numbering = Numbering::from_str(NUMBERING_XML).unwrap();
    let result = numbering.to_string().unwrap();
    assert_eq!(
        replace_whitespace(NUMBERING_XML, " "),
        replace_whitespace(
            &result.replace(&format!(" xmlns:w14=\"{SCHEMA_WORDML_14}\""), ""),
            " "
        )
    );
}
