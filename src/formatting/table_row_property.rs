use hard_xml::{XmlRead, XmlWrite};

use crate::{__setter, __xml_test_suites, formatting::TableHeader, formatting::TableJustification};

/// Table Row Property
///
/// ```rust
/// use docx_rust::formatting::{TableRowProperty, TableJustificationVal};
///
/// let prop = TableRowProperty::default()
///     .justification(TableJustificationVal::Start);
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:trPr")]
pub struct TableRowProperty {
    /// Specifies the alignment of the row with respect to the text margins in the section.
    #[xml(child = "w:jc")]
    pub justification: Option<TableJustification>,
    /// Repeat Table Row on Every New Page
    #[xml(child = "w:tblHeader")]
    pub table_header: Option<TableHeader>,
}

impl TableRowProperty {
    __setter!(justification: Option<TableJustification>);
    __setter!(table_header: Option<TableHeader>);
}

__xml_test_suites!(
    TableRowProperty,
    TableRowProperty::default(),
    r#"<w:trPr/>"#,
    TableRowProperty::default()
        .justification(crate::formatting::TableJustificationVal::Start)
        .table_header(crate::formatting::OnOffOnlyType::On),
    r#"<w:trPr><w:jc w:val="start"/><w:tblHeader w:val="on"/></w:trPr>"#,
);
