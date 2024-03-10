use hard_xml::{XmlRead, XmlWrite};

use crate::{__setter, __xml_test_suites};

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tcPr")]
pub struct TableCellProperty {
    #[xml(child = "w:tcW")]
    pub wide: Option<super::TableCellWidth>,
    #[xml(default, child = "w:vAlign")]
    pub v_align: super::VAlign
}

impl TableCellProperty {
    __setter!(v_align: super::VAlign);
    __setter!(wide: Option<super::TableCellWidth>);
}

__xml_test_suites!(
    TableCellProperty,
    TableCellProperty::default(),
    r#"<w:tcPr><w:vAlign w:val="top"/></w:tcPr>"#,
    TableCellProperty::default().v_align(super::VAlignType::Bottom),
    r#"<w:tcPr><w:vAlign w:val="bottom"/></w:tcPr>"#,
);