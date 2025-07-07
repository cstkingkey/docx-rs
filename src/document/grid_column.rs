use hard_xml::{XmlRead, XmlWrite};

use crate::__xml_test_suites;

/// Grid Column
///
/// ```rust
/// use docx_rust::document::*;
///
/// let col = GridColumn::from(42);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:gridCol")]
pub struct GridColumn {
    #[xml(attr = "w:w", with = "crate::rounded_float")]
    pub width: isize,
}

impl From<isize> for GridColumn {
    fn from(width: isize) -> GridColumn {
        GridColumn { width }
    }
}

__xml_test_suites!(
    GridColumn,
    GridColumn::from(42isize),
    r#"<w:gridCol w:w="42"/>"#,
);
