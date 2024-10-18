#![allow(unused_must_use)]
use hard_xml::{XmlRead, XmlWrite};
use std::borrow::{Borrow, Cow};

use crate::{
    __setter, __xml_test_suites,
    document::{TableGrid, TableRow},
    formatting::TableProperty,
};

/// Table
///
/// ```rust
/// use docx_rust::document::*;
/// use docx_rust::formatting::*;
///
/// let tbl = Table::default()
///     .property(TableProperty::default())
///     .push_row(TableRow::default());
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tbl")]
pub struct Table<'a> {
    #[xml(default, child = "w:tblPr")]
    pub property: TableProperty<'a>,
    #[xml(child = "w:tblGrid")]
    pub grids: TableGrid,
    #[xml(child = "w:tr")]
    pub rows: Vec<TableRow<'a>>,
}

impl<'a> Table<'a> {
    __setter!(property: TableProperty<'a>);

    pub fn push_row<T: Into<TableRow<'a>>>(mut self, row: T) -> Self {
        self.rows.push(row.into());
        self
    }

    pub fn iter_text(&self) -> impl Iterator<Item = &Cow<'a, str>> {
        self.rows.iter().flat_map(|content| content.iter_text())
    }

    pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
        self.rows
            .iter_mut()
            .flat_map(|content| content.iter_text_mut())
    }

    pub fn replace_text<'b, I, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = I> + Copy,
        I: Borrow<(S, S)>,
    {
        for row in self.rows.iter_mut() {
            row.replace_text(dic)?;
        }
        Ok(())
    }
}

__xml_test_suites!(
    Table,
    Table::default(),
    "<w:tbl><w:tblPr/><w:tblGrid/></w:tbl>",
    Table::default().push_row(TableRow::default()),
    "<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:trPr/></w:tr></w:tbl>",
);
