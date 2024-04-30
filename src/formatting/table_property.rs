use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{
    __setter, __string_enum, __xml_test_suites,
    formatting::{TableBorders, TableIndent, TableJustification, TableWidth},
};

use super::table_margin::TableMargins;

/// Table Property
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let prop = TableProperty::default()
///     .style_id("foo")
///     .justification(TableJustificationVal::Start)
///     .indent((50, TableIndentUnit::Pct))
///     .width((50, TableWidthUnit::Pct));
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tblPr")]
pub struct TableProperty<'a> {
    #[xml(child = "w:tblStyle")]
    pub style_id: Option<TableStyleId<'a>>,
    #[xml(child = "w:tblW")]
    pub width: Option<TableWidth>,
    #[xml(child = "w:jc")]
    pub justification: Option<TableJustification>,
    #[xml(child = "w:tblInd")]
    pub indent: Option<TableIndent>,
    #[xml(child = "w:tblBorders")]
    pub borders: Option<TableBorders<'a>>,
    #[xml(child = "w:tblCellMar")]
    pub margins: Option<TableMargins<'a>>,
}

impl<'a> TableProperty<'a> {
    __setter!(style_id: Option<TableStyleId<'a>>);
    __setter!(justification: Option<TableJustification>);
    __setter!(borders: Option<TableBorders<'a>>);
    __setter!(indent: Option<TableIndent>);
    __setter!(width: Option<TableWidth>);
}

#[derive(Debug, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tblStyle")]
pub struct TableStyleId<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

impl<'a, T: Into<Cow<'a, str>>> From<T> for TableStyleId<'a> {
    fn from(val: T) -> Self {
        TableStyleId { value: val.into() }
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tblStylePr")]
pub struct ConditionalTableProperty<'a> {
    #[xml(attr = "type")]
    pub condition: Option<ConditionType>,
    /// Specifies a set of paragraph properties
    #[xml(default, child = "w:pPr")]
    pub paragraph: Option<super::ParagraphProperty<'a>>,
    /// Specifies a set of character properties
    #[xml(default, child = "w:rPr")]
    pub character: Option<super::CharacterProperty<'a>>,
    #[xml(default, child = "w:tblPr")]
    pub table: Option<TableProperty<'a>>,
    #[xml(child = "w:trPr")]
    pub table_row: Option<crate::formatting::TableRowProperty>,
    #[xml(child = "w:tcPr")]
    pub table_cell: Option<crate::formatting::TableCellProperty>,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ConditionType {
    #[default]
    WholeTable, //Whole table formatting.
    FirstRow,  //First Row Conditional Formatting.
    LastRow,   //Last table row formatting.
    FirstCol,  //First Column Conditional Formatting.
    LastCol,   //Last table column formatting.
    Band1Vert, //Banded Column Conditional Formatting.
    Band2Vert, //Even Column Stripe Conditional Formatting.
    Band1Horz, //Banded Row Conditional Formatting.
    Band2Horz, //Even Row Stripe Conditional Formatting.
    NeCell,    //Top right table cell formatting.
    NwCell,    //Top left table cell formatting.
    SeCell,    //Bottom right table cell formatting.
    SwCell,    //Bottom left table cell formatting.
}

__string_enum! {
    ConditionType {
        WholeTable = "wholeTable",
        FirstRow = "firstRow",
        LastRow = "lastRow",
        FirstCol = "firstCol",
        LastCol = "lastCol",
        Band1Vert = "band1Vert",
        Band2Vert = "band2Vert",
        Band1Horz = "band1Horz",
        Band2Horz = "band2Horz",
        NeCell = "neCell",
        NwCell = "nwCell",
        SeCell = "seCell",
        SwCell = "swCell",
    }
}

__xml_test_suites!(
    TableProperty,
    TableProperty::default(),
    r#"<w:tblPr/>"#,
    TableProperty::default().style_id("id"),
    r#"<w:tblPr><w:tblStyle w:val="id"/></w:tblPr>"#,
    TableProperty::default().justification(crate::formatting::TableJustificationVal::Start),
    r#"<w:tblPr><w:jc w:val="start"/></w:tblPr>"#,
    TableProperty::default().borders(TableBorders::default()),
    r#"<w:tblPr><w:tblBorders/></w:tblPr>"#,
    TableProperty::default().indent(TableIndent::default()),
    r#"<w:tblPr><w:tblInd/></w:tblPr>"#,
    TableProperty::default().width(TableWidth::default()),
    r#"<w:tblPr><w:tblW/></w:tblPr>"#,
);
