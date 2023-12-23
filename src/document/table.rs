#![allow(unused_must_use)]
use std::borrow::Cow;
use strong_xml::{XmlRead, XmlWrite};

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
///     .push_grid(vec![1, 2, 3])
///     .push_grid(TableGrid::default())
///     .push_row(TableRow::default());
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:tbl")]
pub struct Table<'a> {
    #[xml(default, child = "w:tblPr")]
    pub property: TableProperty<'a>,
    #[xml(child = "w:tblGrid")]
    pub grids: Vec<TableGrid>,
    #[xml(child = "w:tr")]
    pub rows: Vec<TableRow<'a>>,
}

impl<'a> Table<'a> {
    __setter!(property: TableProperty<'a>);

    pub fn push_grid<T: Into<TableGrid>>(mut self, grid: T) -> Self {
        self.grids.push(grid.into());
        self
    }

    pub fn push_row<T: Into<TableRow<'a>>>(mut self, row: T) -> Self {
        self.rows.push(row.into());
        self
    }

    pub fn iter_text(&self) -> impl Iterator<Item = &Cow<'a, str>> {
        self.rows
            .iter()
            .map(|content| content.iter_text())
            .flatten()
    }

    pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
        self.rows
            .iter_mut()
            .map(|content| content.iter_text_mut())
            .flatten()
    }

    pub fn replace_text<'b, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = &'b (S, S)> + std::marker::Copy,
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
    "<w:tbl><w:tblPr/></w:tbl>",
    Table::default().push_grid(TableGrid::default()),
    "<w:tbl><w:tblPr/><w:tblGrid/></w:tbl>",
    Table::default().push_row(TableRow::default()),
    "<w:tbl><w:tblPr/><w:tr><w:trPr/></w:tr></w:tbl>",
);
