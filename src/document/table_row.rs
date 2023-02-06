#![allow(unused_must_use)]
use std::borrow::Cow;

use super::SDT;

use strong_xml::{XmlRead, XmlWrite};

use crate::{__setter, __xml_test_suites, document::TableCell, formatting::TableRowProperty};

/// Table Row
///
/// ```rust
/// use docx_rust::document::*;
/// use docx_rust::formatting::*;
///
/// let row = TableRow::default()
///     .property(TableRowProperty::default())
///     .push_cell(Paragraph::default())
///     .push_cell(
///         TableCell::pargraph(Paragraph::default())
///             .property(TableCellProperty::default())
///     );
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tr")]
pub struct TableRow<'a> {
    #[xml(default, child = "w:trPr")]
    pub property: TableRowProperty,
    #[xml(child = "w:tc", child = "w:sdt")]
    pub cells: Vec<TableRowContent<'a>>,
}

#[derive(Debug, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum TableRowContent<'a> {
    #[xml(tag = "w:tc")]
    TableCell(TableCell<'a>),
    #[xml(tag = "w:sdt")]
    SDT(SDT<'a>)
}

impl<'a> TableRow<'a> {
    __setter!(property: TableRowProperty);

    pub fn push_cell<T: Into<TableRowContent<'a>>>(mut self, cell: T) -> Self {
        self.cells.push(cell.into());
        self
    }

    
    pub fn iter_text(&self) -> impl Iterator<Item = &Cow<'a, str>> {
        self.cells
            .iter()
            .filter_map(|content| match content {
                //Some(content.iter_text())
                TableRowContent::TableCell(tc) => Some(tc.iter_text()),
                TableRowContent::SDT(_) => None
            })
            .flatten()
    }
    
    
    pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
        self.cells
            .iter_mut()
            .filter_map(|content| match content {
                //Some(content.iter_text_mut())
                TableRowContent::TableCell(tc) => Some(tc.iter_text_mut()),
                TableRowContent::SDT(_) => None
            })
            .flatten()
    }
    
}

#[cfg(test)]
use crate::document::Paragraph;

__xml_test_suites!(
    TableRow,
    TableRow::default(),
    "<w:tr><w:trPr/></w:tr>",
    TableRow::default().push_cell(Paragraph::default()),
    "<w:tr><w:trPr/><w:tc><w:tcPr/><w:p/></w:tc></w:tr>",
);
