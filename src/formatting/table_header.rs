use hard_xml::{XmlRead, XmlWrite};

use crate::{__string_enum, __xml_test_suites};

/// Table Justification
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let th = TableHeader::from(OnOffOnlyType::On);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tblHeader")]
pub struct TableHeader {
    #[xml(attr = "w:val")]
    pub value: Option<OnOffOnlyType>,
}

impl From<OnOffOnlyType> for TableHeader {
    fn from(val: OnOffOnlyType) -> Self {
        TableHeader { value: Some(val) }
    }
}

#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum OnOffOnlyType {
    On,
    Off,
}

__string_enum! {
    OnOffOnlyType {
        On = "on",
        Off = "off",
    }
}

__xml_test_suites!(
    TableHeader,
    TableHeader::default(),
    "<w:tblHeader/>",
    TableHeader::from(OnOffOnlyType::On),
    r#"<w:tblHeader w:val="on"/>"#,
    TableHeader::from(OnOffOnlyType::Off),
    r#"<w:tblHeader w:val="off"/>"#,
);
