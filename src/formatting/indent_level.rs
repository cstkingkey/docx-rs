use hard_xml::{XmlRead, XmlWrite};

use crate::__xml_test_suites;

/// Indent Level
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let lvl = IndentLevel::from(42isize);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:ilvl")]
pub struct IndentLevel {
    #[xml(attr = "w:val", with = "crate::rounded_float")]
    pub value: isize,
}

impl<T: Into<isize>> From<T> for IndentLevel {
    fn from(val: T) -> Self {
        IndentLevel { value: val.into() }
    }
}

__xml_test_suites!(
    IndentLevel,
    IndentLevel::from(40isize),
    r#"<w:ilvl w:val="40"/>"#,
);
