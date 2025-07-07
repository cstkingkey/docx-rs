use hard_xml::{XmlRead, XmlWrite};

use crate::{__string_enum, __xml_test_suites};

/// Table Width
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let width = TableWidth::from(42isize);
/// let width = TableWidth::from(TableWidthUnit::Pct);
/// let width = TableWidth::from((42, TableWidthUnit::Dxa));
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tblW")]
pub struct TableWidth {
    #[xml(attr = "w:w", with = "crate::rounded_float")]
    pub value: Option<isize>,
    #[xml(attr = "w:type")]
    pub unit: Option<TableWidthUnit>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tcW")]
pub struct TableCellWidth {
    #[xml(attr = "w:w", with = "crate::rounded_float")]
    pub value: Option<isize>,
    #[xml(attr = "w:type")]
    pub unit: Option<TableWidthUnit>,
}

impl From<isize> for TableWidth {
    fn from(val: isize) -> Self {
        TableWidth {
            value: Some(val),
            unit: None,
        }
    }
}

impl From<TableWidthUnit> for TableWidth {
    fn from(val: TableWidthUnit) -> Self {
        TableWidth {
            value: None,
            unit: Some(val),
        }
    }
}

impl From<(isize, TableWidthUnit)> for TableWidth {
    fn from(val: (isize, TableWidthUnit)) -> Self {
        TableWidth {
            value: Some(val.0),
            unit: Some(val.1),
        }
    }
}

#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum TableWidthUnit {
    Auto,
    Dxa,
    Nil,
    Pct,
}

__string_enum! {
    TableWidthUnit {
        Auto = "auto",
        Dxa = "dxa",
        Nil = "nil",
        Pct = "pct",
    }
}

__xml_test_suites!(
    TableWidth,
    TableWidth::default(),
    "<w:tblW/>",
    TableWidth::from(42),
    r#"<w:tblW w:w="42"/>"#,
    TableWidth::from(TableWidthUnit::Pct),
    r#"<w:tblW w:type="pct"/>"#,
    TableWidth::from((42, TableWidthUnit::Dxa)),
    r#"<w:tblW w:w="42" w:type="dxa"/>"#,
);
