#![allow(unused_must_use)]
use std::borrow::Cow;

use derive_more::From;
use strong_xml::{XmlRead, XmlWrite};

use crate::{__setter, __xml_test_suites, document::Paragraph, formatting::TableCellProperty};

/// Table Cell
///
/// ```rust
/// use docx_rust::document::*;
/// use docx_rust::formatting::*;
///
/// let cell = TableCell::from(Paragraph::default());
///
/// let cell = TableCell::paragraph(Paragraph::default())
///     .property(TableCellProperty::default());
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tc")]
pub struct TableCell<'a> {
    #[xml(default, child = "w:tcPr")]
    pub property: TableCellProperty,
    #[xml(child = "w:p")]
    pub content: Vec<TableCellContent<'a>>,
}

impl<'a> TableCell<'a> {
    __setter!(property: TableCellProperty);

    pub fn paragraph<T: Into<Paragraph<'a>>>(par: T) -> Self {
        TableCell {
            property: TableCellProperty::default(),
            content: vec![TableCellContent::Paragraph(par.into())],
        }
    }

    pub fn iter_text(&self) -> impl Iterator<Item = &Cow<'a, str>> {
        self.content
            .iter()
            .filter_map(|content| match content {
                TableCellContent::Paragraph(p) => Some(p.iter_text()),
            })
            .flatten()
    }

    pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
        self.content
            .iter_mut()
            .filter_map(|content| match content {
                TableCellContent::Paragraph(p) => Some(p.iter_text_mut()),
            })
            .flatten()
    }

    pub fn replace_text<'b, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = &'b (S, S)> + std::marker::Copy,
    {
        for content in self.content.iter_mut() {
            if let TableCellContent::Paragraph(p) = content {
                p.replace_text(dic)?
            }
        }
        Ok(())
    }
}

impl<'a, T: Into<TableCellContent<'a>>> From<T> for TableCell<'a> {
    fn from(content: T) -> Self {
        TableCell {
            property: TableCellProperty::default(),
            content: vec![content.into()],
        }
    }
}

#[derive(Debug, From, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum TableCellContent<'a> {
    #[xml(tag = "w:p")]
    Paragraph(Paragraph<'a>),
    // #[xml(tag = "w:tbl")]
    // Table(Table<'a>),
}

__xml_test_suites!(
    TableCell,
    TableCell::paragraph(Paragraph::default()),
    "<w:tc><w:tcPr/><w:p/></w:tc>",
);
