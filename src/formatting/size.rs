use hard_xml::{XmlRead, XmlWrite};

use crate::__xml_test_suites;

/// Size
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let sz = Size::from(42isize);
/// ```
#[derive(Debug, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sz")]
pub struct Size {
    #[xml(attr = "w:val", with = "crate::rounded_float")]
    pub value: isize,
}

impl<T: Into<isize>> From<T> for Size {
    fn from(val: T) -> Self {
        Size { value: val.into() }
    }
}

__xml_test_suites!(Size, Size::from(42isize), r#"<w:sz w:val="42"/>"#,);
