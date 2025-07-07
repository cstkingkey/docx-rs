use hard_xml::{XmlRead, XmlWrite};

/// WidowControl
///
/// ```rust
/// use docx_rust::formatting::*;
///
/// let WidowControl = WidowControl::from(false);
/// let WidowControl = WidowControl::from(true);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:uiPriority")]
pub struct Priority {
    #[xml(attr = "w:val", with = "crate::rounded_float")]
    pub value: Option<isize>,
}
