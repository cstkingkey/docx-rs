//! Formatting
//!
//! Formatting can be used to declare a style,
//! or used within a document directly.

mod bold;
mod border;
mod borders;
mod character_property;
mod color;
mod dstrike;
mod fonts;
mod indent;
mod indent_level;
mod italics;
mod justification;
mod lang;
mod line_rule;
mod margin;
mod numbering_id;
mod numbering_property;
mod outline;
mod page_cols;
mod page_grid;
mod page_margin;
mod page_size;
mod paragraph_property;
mod section_property;
mod size;
mod spacing;
mod strike;
mod table_borders;
mod table_cell_property;
mod table_header;
mod table_indent;
mod table_justification;
mod table_margin;
mod table_property;
mod table_row_property;
mod table_width;
mod underline;
mod widow_control;

// re-export
pub use self::{
    bold::*, border::*, borders::*, character_property::*, color::*, dstrike::*, fonts::*,
    indent::*, indent_level::*, italics::*, justification::*, lang::*, numbering_id::*,
    numbering_property::*, outline::*, page_cols::*, page_grid::*, page_margin::*, page_size::*,
    paragraph_property::*, section_property::*, size::*, spacing::*, strike::*, table_borders::*,
    table_cell_property::*, table_header::*, table_indent::*, table_justification::*,
    table_property::*, table_row_property::*, table_width::*, underline::*, widow_control::*,
};
