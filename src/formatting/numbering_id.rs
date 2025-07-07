use hard_xml::{XmlRead, XmlWrite};

use crate::__xml_test_suites;

/// Numbering Id
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let id = NumberingId::from(42isize);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numId")]
pub struct NumberingId {
    #[xml(attr = "w:val", with = "crate::rounded_float")]
    pub value: isize,
}

impl<T: Into<isize>> From<T> for NumberingId {
    fn from(val: T) -> Self {
        NumberingId { value: val.into() }
    }
}

__xml_test_suites!(
    NumberingId,
    NumberingId::from(40isize),
    r#"<w:numId w:val="40"/>"#,
);
