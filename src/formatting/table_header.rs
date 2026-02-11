use hard_xml::{XmlRead, XmlWrite};

use crate::__xml_test_suites;

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

#[derive(Debug, Clone, Copy)]
#[cfg_attr(test, derive(PartialEq))]
#[repr(u8)]
pub enum OnOffOnlyType {
    Off = 0,
    On = 1,
}

impl OnOffOnlyType {
    const STR_REPR: [&str; 2] = ["off", "on"];
}

impl std::fmt::Display for OnOffOnlyType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", Self::STR_REPR[*self as usize])
    }
}

impl std::str::FromStr for OnOffOnlyType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            s if s == Self::STR_REPR[0] => Ok(Self::Off),
            s if s == Self::STR_REPR[1] => Ok(Self::On),
            "0" => Ok(Self::Off),
            "1" => Ok(Self::On),
            s => Err(format!(
                "Unkown Value. Found `{}`, Expected `{:?}`",
                s,
                Self::STR_REPR
            )),
        }
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

#[test]
fn parse_bool_from_int() {
    use std::str::FromStr;

    for (i, v) in OnOffOnlyType::STR_REPR.iter().enumerate() {
        assert_eq!(
            OnOffOnlyType::from_str(&format!("{}", i)).unwrap(),
            OnOffOnlyType::from_str(v).unwrap(),
        );
    }
}
