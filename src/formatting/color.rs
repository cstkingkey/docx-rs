use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{__string_enum, __xml_test_suites};

/// Text Color
///
/// Specifies the color to be used to display text.
///
/// ```rust
/// use docx_rust::formatting::Color;
///
/// let color = Color::from("000000");
/// let color = Color::from(String::from("000000"));
/// let color = Color::from(0u32); // "000000"
/// let color = Color::from((0u8, 0u8, 0u8)); // "000000"
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:color")]
pub struct Color<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

impl<'a> From<&'a str> for Color<'a> {
    fn from(val: &'a str) -> Self {
        Color {
            value: Cow::Borrowed(val),
        }
    }
}

impl From<String> for Color<'_> {
    fn from(val: String) -> Self {
        Color {
            value: Cow::Owned(val),
        }
    }
}

impl From<u32> for Color<'_> {
    fn from(val: u32) -> Self {
        Color {
            value: Cow::Owned(format!("{:06x}", val)),
        }
    }
}

impl From<(u8, u8, u8)> for Color<'_> {
    fn from(val: (u8, u8, u8)) -> Self {
        Color {
            value: Cow::Owned(format!("{:02x}{:02x}{:02x}", val.0, val.1, val.2)),
        }
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:highlight")]
pub struct Highlight {
    #[xml(attr = "w:val")]
    pub value: Option<HighlightType>,
}

impl From<HighlightType> for Highlight {
    fn from(val: HighlightType) -> Self {
        Highlight { value: Some(val) }
    }
}

#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum HighlightType {
    Black,       //Black Highlighting Color
    Blue,        //Blue Highlighting Color
    Cyan,        //Cyan Highlighting Color
    Green,       //Green Highlighting Color
    Magenta,     //Magenta Highlighting Color
    Red,         //Red Highlighting Color
    Yellow,      //Yellow Highlighting Color
    White,       //White Highlighting Color
    DarkBlue,    //Dark Blue Highlighting Color
    DarkCyan,    //Dark Cyan Highlighting Color
    DarkGreen,   //Dark Green Highlighting Color
    DarkMagenta, //Dark Magenta Highlighting Color
    DarkRed,     //Dark Red Highlighting Color
    DarkYellow,  //Dark Yellow Highlighting Color
    DarkGray,    //Dark Gray Highlighting Color
    LightGray,   //Light Gray Highlighting Color
    None,        //No Text Highlighting
}

__string_enum! {
    HighlightType {
        Black = "black",
        Blue = "blue",
        Cyan = "cyan",
        Green = "green",
        Magenta = "magenta",
        Red = "red",
        Yellow = "yellow",
        White = "white",
        DarkBlue = "darkBlue",
        DarkCyan = "darkCyan",
        DarkGreen = "darkGreen",
        DarkMagenta = "darkMagenta",
        DarkRed = "darkRed",
        DarkYellow = "darkYellow",
        DarkGray = "darkGray",
        LightGray = "lightGray",
        None = "none",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:vertAlign")]
pub struct VertAlign {
    #[xml(attr = "w:val")]
    pub value: Option<VertAlignType>,
}

#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum VertAlignType {
    Baseline,    //Regular Vertical Positioning
    Superscript, //	Superscript
    Subscript,   //	Subscript
}

__string_enum! {
    VertAlignType {
        Baseline = "baseline",
        Subscript = "subscript",
        Superscript = "superscript",
    }
}

__xml_test_suites!(
    Color,
    Color::from("000000"),
    r#"<w:color w:val="000000"/>"#,
    Color::from(0u32),
    r#"<w:color w:val="000000"/>"#,
    Color::from((0u8, 0u8, 0u8)),
    r#"<w:color w:val="000000"/>"#,
);
