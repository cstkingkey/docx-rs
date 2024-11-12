//! Theme part
//!
//! The corresponding ZIP item is `/word/theme/theme{n}.xml`.
//!
#![allow(unused_must_use)]
#![allow(unused_imports)]
use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::borrow::Cow;
use std::io::Write;

use crate::schema::{SCHEMA_DRAWINGML, SCHEMA_MAIN, SCHEMA_WORDML_14};
use crate::{__define_struct, __define_struct_vec, __string_enum, __xml_test_suites, write_attr};

/// The root element of the main document part.
#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:theme")]
pub struct Theme<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,
    #[xml(child = "a:themeElements")]
    pub elements: ThemeElements<'a>,
    #[xml(child = "a:objectDefaults")]
    pub defaults: Option<ObjectDefaults>,
    #[xml(child = "a:extraClrSchemeLst")]
    pub extra_clr_scheme_lst: Option<ExtraClrSchemeLst>,
    #[xml(child = "a:custClrLst")]
    pub cust_clr_lst: Option<CustClrLst<'a>>,
    #[xml(child = "a:extLst")]
    pub ext_lst: Option<ExtLst>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:custClrLst")]
pub struct CustClrLst<'a> {
    #[xml(child = "a:custClr")]
    pub contents: Vec<CustClr<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:dk1")]
pub struct Dk1<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:lt1")]
pub struct Lt1<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:dk2")]
pub struct Dk2<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:lt2")]
pub struct Lt2<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:accent1")]
pub struct Accent1<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:accent2")]
pub struct Accent2<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:accent3")]
pub struct Accent3<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:accent4")]
pub struct Accent4<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:accent5")]
pub struct Accent5<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:accent6")]
pub struct Accent6<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:hlink")]
pub struct HLink<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:folHlink")]
pub struct FolHlink<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:custClr")]
pub struct CustClr<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,

    #[xml(
        child = "a:scrgbClr",
        child = "a:srgbClr",
        child = "a:hslClr",
        child = "a:sysClr",
        child = "a:schemeClr",
        child = "a:prstClr"
    )]
    pub custom_color: Vec<CustClrChoice>,
}

#[derive(Debug, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum CustClrChoice {
    ///  RGB Color Model - Percentage Variant
    #[xml(tag = "a:scrgbClr")]
    ScrgbClr(ScrgbClr),
    ///  RGB Color Model - Hex Variant
    #[xml(tag = "a:srgbClr")]
    SrgbClr(SrgbClr),
    ///  Hue, Saturation, Luminance Color Model
    #[xml(tag = "a:hslClr")]
    HslClr(HslClr),
    ///  System Color
    #[xml(tag = "a:sysClr")]
    SysClr(SysClr),
    ///  Scheme Color
    #[xml(tag = "a:schemeClr")]
    SchemeClr(SchemeClr),
    ///  Preset Color
    #[xml(tag = "a:prstClr")]
    PrstClr(PrstClr),
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:scrgbClr")]
pub struct ScrgbClr {
    #[xml(attr = "r")]
    pub r: u8,
    #[xml(attr = "g")]
    pub g: u8,
    #[xml(attr = "b")]
    pub b: u8,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:srgbClr")]
pub struct SrgbClr {
    #[xml(attr = "val")]
    pub value: Option<String>,
    #[xml(child = "a:alpha")]
    pub alpha: Option<Alpha>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:alpha")]
pub struct Alpha {
    #[xml(attr = "val")]
    pub value: usize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:hslClr")]
pub struct HslClr {
    #[xml(attr = "hue")]
    pub hue: Option<isize>,
    #[xml(attr = "sat")]
    pub sat: Option<isize>,
    #[xml(attr = "lum")]
    pub lum: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:sysClr")]
pub struct SysClr {
    #[xml(attr = "val")]
    pub val: SysClrType,
    #[xml(attr = "lastClr")]
    pub last_color: Option<String>,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum SysClrType {
    #[default]
    ScrollBar, //Scroll Bar System Color.
    Background,              //Background System Color.
    ActiveCaption,           //Active Caption System Color.
    InactiveCaption,         //Inactive Caption System Color.
    Menu,                    //Menu System Color.
    Window,                  //Window System Color.
    WindowFrame,             //Window Frame System Color.
    MenuText,                //Menu Text System Color.
    WindowText,              //Window Text System Color.
    CaptionText,             //Caption Text System Color.
    ActiveBorder,            //Active Border System Color.
    InactiveBorder,          //Inactive Border System Color.
    AppWorkspace,            //Application Workspace System Color.
    Highlight,               //Highlight System Color.
    HighlightText,           //Highlight Text System Color.
    BtnFace,                 //Button Face System Color.
    BtnShadow,               //Button Shadow System Color.
    GrayText,                //Gray Text System Color.
    BtnText,                 //Button Text System Color.
    InactiveCaptionText,     //Inactive Caption Text System Color.
    BtnHighlight,            //Button Highlight System Color.
    TdDkShadow,              //3D Dark System Color.
    TdLight,                 //3D Light System Color.
    InfoText,                //Info Text System Color.
    InfoBk,                  //Info Back System Color.
    HotLight,                //Hot Light System Color.
    GradientActiveCaption,   //Gradient Active Caption System Color.
    GradientInactiveCaption, //Gradient Inactive Caption System Color.
    MenuHighlight,           //Menu Highlight System Color.
    MenuBar,                 //Menu Bar System Color.
}

__string_enum! {
    SysClrType {
        ScrollBar = "scrollBar",
        Background = "background",
        ActiveCaption = "activeCaption",
        InactiveCaption = "inactiveCaption",
        Menu = "menu",
        Window = "window",
        WindowFrame = "windowFrame",
        MenuText = "menuText",
        WindowText = "windowText",
        CaptionText = "captionText",
        ActiveBorder = "activeBorder",
        InactiveBorder = "inactiveBorder",
        AppWorkspace = "appWorkspace",
        Highlight = "highlight",
        HighlightText = "highlightText",
        BtnFace = "btnFace",
        BtnShadow = "btnShadow",
        GrayText = "grayText",
        BtnText = "btnText",
        InactiveCaptionText = "inactiveCaptionText",
        BtnHighlight = "btnHighlight",
        TdDkShadow = "3dDkShadow",
        TdLight = "3dLight",
        InfoText = "infoText",
        InfoBk = "infoBk",
        HotLight = "hotLight",
        GradientActiveCaption = "gradientActiveCaption",
        GradientInactiveCaption = "gradientInactiveCaption",
        MenuHighlight = "menuHighlight",
        MenuBar = "menuBar",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:schemeClr")]
pub struct SchemeClr {
    #[xml(attr = "val")]
    pub val: SchemeClrType,
    #[xml(child = "a:lumMod")]
    pub lum_mod: Option<LuminanceModulation>,
    #[xml(child = "a:satMod")]
    pub sat_mod: Option<SaturationModulation>,
    #[xml(child = "a:tint")]
    pub tint: Option<Tint>,
    #[xml(child = "a:shade")]
    pub shade: Option<Shade>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:lumMod")]
pub struct LuminanceModulation {
    #[xml(attr = "val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:satMod")]
pub struct SaturationModulation {
    #[xml(attr = "val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:tint")]
pub struct Tint {
    #[xml(attr = "val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:shade")]
pub struct Shade {
    #[xml(attr = "val")]
    pub val: isize,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum SchemeClrType {
    #[default]
    Bg1, //Background Color 1.
    Tx1,      //Text Color 1.
    Bg2,      //Background Color 2.
    Tx2,      //Text Color 2.
    Accent1,  //Accent Color 1.
    Accent2,  //Accent Color 2.
    Accent3,  //Accent Color 3.
    Accent4,  //Accent Color 4.
    Accent5,  //Accent Color 5.
    Accent6,  //Accent Color 6.
    Hlink,    //Hyperlink Color.
    FolHlink, //Followed Hyperlink Color.
    PhClr,    //Style Color.
    Dk1,      //Dark Color 1.
    Lt1,      //Light Color 1.
    Dk2,      //Dark Color 2.
    Lt2,      //Light Color 2.
}

__string_enum! {
    SchemeClrType {
        Bg1 = "bg1",
        Tx1 = "tx1",
        Bg2 = "bg2",
        Tx2 = "tx2",
        Accent1 = "accent1",
        Accent2 = "accent2",
        Accent3 = "accent3",
        Accent4 = "accent4",
        Accent5 = "accent5",
        Accent6 = "accent6",
        Hlink = "hlink",
        FolHlink = "folHlink",
        PhClr = "phClr",
        Dk1 = "dk1",
        Lt1 = "lt1",
        Dk2 = "dk2",
        Lt2 = "lt2",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:prstClr")]
pub struct PrstClr {
    #[xml(attr = "val")]
    pub val: Option<PrstClrType>,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum PrstClrType {
    #[default]
    AliceBlue, //Alice Blue Preset Color.
    AntiqueWhite,      //Antique White Preset Color.
    Aqua,              //Aqua Preset Color.
    Aquamarine,        //Aquamarine Preset Color.
    Azure,             //Azure Preset Color.
    Beige,             //Beige Preset Color.
    Bisque,            //Bisque Preset Color.
    Black,             //Black Preset Color.
    BlanchedAlmond,    //Blanched Almond Preset Color.
    Blue,              //Blue Preset Color.
    BlueViolet,        //Blue Violet Preset Color.
    Brown,             //Brown Preset Color.
    BurlyWood,         //Burly Wood Preset Color.
    CadetBlue,         //Cadet Blue Preset Color.
    Chartreuse,        //Chartreuse Preset Color.
    Chocolate,         //Chocolate Preset Color.
    Coral,             //Coral Preset Color.
    CornflowerBlue,    //Cornflower Blue Preset Color.
    Cornsilk,          //Cornsilk Preset Color.
    Crimson,           //Crimson Preset Color.
    Cyan,              //Cyan Preset Color.
    DkBlue,            //Dark Blue Preset Color.
    DkCyan,            //Dark Cyan Preset Color.
    DkGoldenrod,       //Dark Goldenrod Preset Color.
    DkGray,            //Dark Gray Preset Color.
    DkGreen,           //Dark Green Preset Color.
    DkKhaki,           //Dark Khaki Preset Color.
    DkMagenta,         //Dark Magenta Preset Color.
    DkOliveGreen,      //Dark Olive Green Preset Color.
    DkOrange,          //Dark Orange Preset Color.
    DkOrchid,          //Dark Orchid Preset Color.
    DkRed,             //Dark Red Preset Color.
    DkSalmon,          //Dark Salmon Preset Color.
    DkSeaGreen,        //Dark Sea Green Preset Color.
    DkSlateBlue,       //Dark Slate Blue Preset Color.
    DkSlateGray,       //Dark Slate Gray Preset Color.
    DkTurquoise,       //Dark Turquoise Preset Color.
    DkViolet,          //Dark Violet Preset Color.
    DeepPink,          //Deep Pink Preset Color.
    DeepSkyBlue,       //Deep Sky Blue Preset Color.
    DimGray,           //Dim Gray Preset Color.
    DodgerBlue,        //Dodger Blue Preset Color.
    Firebrick,         //Firebrick Preset Color.
    FloralWhite,       //Floral White Preset Color.
    ForestGreen,       //Forest Green Preset Color.
    Fuchsia,           //Fuchsia Preset Color.
    Gainsboro,         //Gainsboro Preset Color.
    GhostWhite,        //Ghost White Preset Color.
    Gold,              //Gold Preset Color.
    Goldenrod,         //Goldenrod Preset Color.
    Gray,              //Gray Preset Color.
    Green,             //Green Preset Color.
    GreenYellow,       //Green Yellow Preset Color.
    Honeydew,          //Honeydew Preset Color.
    HotPink,           //Hot Pink Preset Color.
    IndianRed,         //Indian Red Preset Color.
    Indigo,            //Indigo Preset Color.
    Ivory,             //Ivory Preset Color.
    Khaki,             //Khaki Preset Color.
    Lavender,          //Lavender Preset Color.
    LavenderBlush,     //Lavender Blush Preset Color.
    LawnGreen,         //Lawn Green Preset Color.
    LemonChiffon,      //Lemon Chiffon Preset Color.
    LtBlue,            //Light Blue Preset Color.
    LtCoral,           //Light Coral Preset Color.
    LtCyan,            //Light Cyan Preset Color.
    LtGoldenrodYellow, //Light Goldenrod Yellow Preset Color.
    LtGray,            //Light Gray Preset Color.
    LtGreen,           //Light Green Preset Color.
    LtPink,            //Light Pink Preset Color.
    LtSalmon,          //Light Salmon Preset Color.
    LtSeaGreen,        //Light Sea Green Preset Color.
    LtSkyBlue,         //Light Sky Blue Preset Color.
    LtSlateGray,       //Light Slate Gray Preset Color.
    LtSteelBlue,       //Light Steel Blue Preset Color.
    LtYellow,          //Light Yellow Preset Color.
    Lime,              //Lime Preset Color.
    LimeGreen,         //Lime Green Preset Color.
    Linen,             //Linen Preset Color.
    Magenta,           //Magenta Preset Color.
    Maroon,            //Maroon Preset Color.
    MedAquamarine,     //Medium Aquamarine Preset Color.
    MedBlue,           //Medium Blue Preset Color.
    MedOrchid,         //Medium Orchid Preset Color.
    MedPurple,         //Medium Purple Preset Color.
    MedSeaGreen,       //Medium Sea Green Preset Color.
    MedSlateBlue,      //Medium Slate Blue Preset Color.
    MedSpringGreen,    //Medium Spring Green Preset Color.
    MedTurquoise,      //Medium Turquoise Preset Color.
    MedVioletRed,      //Medium Violet Red Preset Color.
    MidnightBlue,      //Midnight Blue Preset Color.
    MintCream,         //Mint Cream Preset Color.
    MistyRose,         //Misty Rose Preset Color.
    Moccasin,          //Moccasin Preset Color.
    NavajoWhite,       //Navajo White Preset Color.
    Navy,              //Navy Preset Color.
    OldLace,           //Old Lace Preset Color.
    Olive,             //Olive Preset Color.
    OliveDrab,         //Olive Drab Preset Color.
    Orange,            //Orange Preset Color.
    OrangeRed,         //Orange Red Preset Color.
    Orchid,            //Orchid Preset Color.
    PaleGoldenrod,     //Pale Goldenrod Preset Color.
    PaleGreen,         //Pale Green Preset Color.
    PaleTurquoise,     //Pale Turquoise Preset Color.
    PaleVioletRed,     //Pale Violet Red Preset Color.
    PapayaWhip,        //Papaya Whip Preset Color.
    PeachPuff,         //Peach Puff Preset Color.
    Peru,              //Peru Preset Color.
    Pink,              //Pink Preset Color.
    Plum,              //Plum Preset Color.
    PowderBlue,        //Powder Blue Preset Color.
    Purple,            //Purple Preset Color.
    Red,               //Red Preset Color.
    RosyBrown,         //Rosy Brown Preset Color.
    RoyalBlue,         //Royal Blue Preset Color.
    SaddleBrown,       //Saddle Brown Preset Color.
    Salmon,            //Salmon Preset Color.
    SandyBrown,        //Sandy Brown Preset Color.
    SeaGreen,          //Sea Green Preset Color.
    SeaShell,          //Sea Shell Preset Color.
    Sienna,            //Sienna Preset Color.
    Silver,            //Silver Preset Color.
    SkyBlue,           //Sky Blue Preset Color.
    SlateBlue,         //Slate Blue Preset Color.
    SlateGray,         //Slate Gray Preset Color.
    Snow,              //Snow Preset Color.
    SpringGreen,       //Spring Green Preset Color.
    SteelBlue,         //Steel Blue Preset Color.
    Tan,               //Tan Preset Color.
    Teal,              //Teal Preset Color.
    Thistle,           //Thistle Preset Color.
    Tomato,            //Tomato Preset Color.
    Turquoise,         //Turquoise Preset Color.
    Violet,            //Violet Preset Color.
    Wheat,             //Wheat Preset Color.
    White,             //White Preset Color.
    WhiteSmoke,        //White Smoke Preset Color.
    Yellow,            //Yellow Preset Color.
    YellowGreen,       //Yellow Green Preset Color.
}

__string_enum! {
    PrstClrType {
        AliceBlue = "aliceBlue",
        AntiqueWhite = "antiqueWhite",
        Aqua = "aqua",
        Aquamarine = "aquamarine",
        Azure = "azure",
        Beige = "beige",
        Bisque = "bisque",
        Black = "black",
        BlanchedAlmond = "blanchedAlmond",
        Blue = "blue",
        BlueViolet = "blueViolet",
        Brown = "brown",
        BurlyWood = "burlyWood",
        CadetBlue = "cadetBlue",
        Chartreuse = "chartreuse",
        Chocolate = "chocolate",
        Coral = "coral",
        CornflowerBlue = "cornflowerBlue",
        Cornsilk = "cornsilk",
        Crimson = "crimson",
        Cyan = "cyan",
        DkBlue = "dkBlue",
        DkCyan = "dkCyan",
        DkGoldenrod = "dkGoldenrod",
        DkGray = "dkGray",
        DkGreen = "dkGreen",
        DkKhaki = "dkKhaki",
        DkMagenta = "dkMagenta",
        DkOliveGreen = "dkOliveGreen",
        DkOrange = "dkOrange",
        DkOrchid = "dkOrchid",
        DkRed = "dkRed",
        DkSalmon = "dkSalmon",
        DkSeaGreen = "dkSeaGreen",
        DkSlateBlue = "dkSlateBlue",
        DkSlateGray = "dkSlateGray",
        DkTurquoise = "dkTurquoise",
        DkViolet = "dkViolet",
        DeepPink = "deepPink",
        DeepSkyBlue = "deepSkyBlue",
        DimGray = "dimGray",
        DodgerBlue = "dodgerBlue",
        Firebrick = "firebrick",
        FloralWhite = "floralWhite",
        ForestGreen = "forestGreen",
        Fuchsia = "fuchsia",
        Gainsboro = "gainsboro",
        GhostWhite = "ghostWhite",
        Gold = "gold",
        Goldenrod = "goldenrod",
        Gray = "gray",
        Green = "green",
        GreenYellow = "greenYellow",
        Honeydew = "honeydew",
        HotPink = "hotPink",
        IndianRed = "indianRed",
        Indigo = "indigo",
        Ivory = "ivory",
        Khaki = "khaki",
        Lavender = "lavender",
        LavenderBlush = "lavenderBlush",
        LawnGreen = "lawnGreen",
        LemonChiffon = "lemonChiffon",
        LtBlue = "ltBlue",
        LtCoral = "ltCoral",
        LtCyan = "ltCyan",
        LtGoldenrodYellow = "ltGoldenrodYellow",
        LtGray = "ltGray",
        LtGreen = "ltGreen",
        LtPink = "ltPink",
        LtSalmon = "ltSalmon",
        LtSeaGreen = "ltSeaGreen",
        LtSkyBlue = "ltSkyBlue",
        LtSlateGray = "ltSlateGray",
        LtSteelBlue = "ltSteelBlue",
        LtYellow = "ltYellow",
        Lime = "lime",
        LimeGreen = "limeGreen",
        Linen = "linen",
        Magenta = "magenta",
        Maroon = "maroon",
        MedAquamarine = "medAquamarine",
        MedBlue = "medBlue",
        MedOrchid = "medOrchid",
        MedPurple = "medPurple",
        MedSeaGreen = "medSeaGreen",
        MedSlateBlue = "medSlateBlue",
        MedSpringGreen = "medSpringGreen",
        MedTurquoise = "medTurquoise",
        MedVioletRed = "medVioletRed",
        MidnightBlue = "midnightBlue",
        MintCream = "mintCream",
        MistyRose = "mistyRose",
        Moccasin = "moccasin",
        NavajoWhite = "navajoWhite",
        Navy = "navy",
        OldLace = "oldLace",
        Olive = "olive",
        OliveDrab = "oliveDrab",
        Orange = "orange",
        OrangeRed = "orangeRed",
        Orchid = "orchid",
        PaleGoldenrod = "paleGoldenrod",
        PaleGreen = "paleGreen",
        PaleTurquoise = "paleTurquoise",
        PaleVioletRed = "paleVioletRed",
        PapayaWhip = "papayaWhip",
        PeachPuff = "peachPuff",
        Peru = "peru",
        Pink = "pink",
        Plum = "plum",
        PowderBlue = "powderBlue",
        Purple = "purple",
        Red = "red",
        RosyBrown = "rosyBrown",
        RoyalBlue = "royalBlue",
        SaddleBrown = "saddleBrown",
        Salmon = "salmon",
        SandyBrown = "sandyBrown",
        SeaGreen = "seaGreen",
        SeaShell = "seaShell",
        Sienna = "sienna",
        Silver = "silver",
        SkyBlue = "skyBlue",
        SlateBlue = "slateBlue",
        SlateGray = "slateGray",
        Snow = "snow",
        SpringGreen = "springGreen",
        SteelBlue = "steelBlue",
        Tan = "tan",
        Teal = "teal",
        Thistle = "thistle",
        Tomato = "tomato",
        Turquoise = "turquoise",
        Violet = "violet",
        Wheat = "wheat",
        White = "white",
        WhiteSmoke = "whiteSmoke",
        Yellow = "yellow",
        YellowGreen = "yellowGreen",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:themeElements")]
pub struct ThemeElements<'a> {
    #[xml(child = "a:clrScheme")]
    pub clr_scheme: ClrScheme<'a>,
    #[xml(child = "a:fontScheme")]
    pub font_scheme: FontScheme<'a>,
    #[xml(child = "a:fmtScheme")]
    pub fmt_scheme: FmtScheme<'a>,
    #[xml(child = "a:extLst")]
    pub ext_lst: Option<ExtLst>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:clrScheme")]
pub struct ClrScheme<'a> {
    #[xml(attr = "name")]
    pub name: Cow<'a, str>,

    ///    Dark 1
    #[xml(child = "a:dk1")]
    pub dk1: Dk1<'a>,
    ///    Light 1
    #[xml(child = "a:lt1")]
    pub lt1: Lt1<'a>,
    ///    Dark 2
    #[xml(child = "a:dk2")]
    pub dk2: Dk2<'a>,
    ///    Light 2
    #[xml(child = "a:lt2")]
    pub lt2: Lt2<'a>,
    ///    Accent
    #[xml(child = "a:accent1")]
    pub accent1: Accent1<'a>,
    ///    Accent 2
    #[xml(child = "a:accent2")]
    pub accent2: Accent2<'a>,
    ///    Accent 3
    #[xml(child = "a:accent3")]
    pub accent3: Accent3<'a>,
    ///    Accent 4
    #[xml(child = "a:accent4")]
    pub accent4: Accent4<'a>,
    ///    Accent 5
    #[xml(child = "a:accent5")]
    pub accent5: Accent5<'a>,
    ///    Accent 6
    #[xml(child = "a:accent6")]
    pub accent6: Accent6<'a>,
    ///    Hyperlink
    #[xml(child = "a:hlink")]
    pub hlink: HLink<'a>,
    ///    Followed Hyperlink
    #[xml(child = "a:folHlink")]
    pub fol_hlink: FolHlink<'a>,
    ///    Extension List
    #[xml(child = "a:extLst")]
    pub ext_lst: Option<ExtLst>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:fontScheme")]
pub struct FontScheme<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,
    #[xml(child = "a:majorFont")]
    pub major_font: MajorFont<'a>,
    #[xml(child = "a:minorFont")]
    pub minor_font: MinorFont<'a>,
    #[xml(child = "a:extLst")]
    pub ext_lst: Option<ExtLst>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:majorFont")]
pub struct MajorFont<'a> {
    #[xml(child = "a:latin")]
    pub latin: Latin<'a>,
    #[xml(child = "a:ea")]
    pub ea: EA<'a>,
    #[xml(child = "a:cs")]
    pub cs: CS<'a>,
    #[xml(child = "a:font")]
    pub fonts: Vec<Font<'a>>,
    #[xml(child = "a:extLst")]
    pub ext_lst: Option<ExtLst>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:minorFont")]
pub struct MinorFont<'a> {
    #[xml(child = "a:latin")]
    pub latin: Latin<'a>,
    #[xml(child = "a:ea")]
    pub ea: EA<'a>,
    #[xml(child = "a:cs")]
    pub cs: CS<'a>,
    #[xml(child = "a:font")]
    pub fonts: Vec<Font<'a>>,
    #[xml(child = "a:extLst")]
    pub ext_lst: Option<ExtLst>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:ea")]
pub struct EA<'a> {
    #[xml(attr = "typeface")]
    pub typeface: Option<Cow<'a, str>>,
    #[xml(attr = "panose")]
    pub panose: Option<Cow<'a, str>>,
    #[xml(attr = "pitchFamily")]
    pub pitch_family: Option<i8>,
    #[xml(attr = "charset")]
    pub charset: Option<i8>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:cs")]
pub struct CS<'a> {
    #[xml(attr = "typeface")]
    pub typeface: Option<Cow<'a, str>>,
    #[xml(attr = "panose")]
    pub panose: Option<Cow<'a, str>>,
    #[xml(attr = "pitchFamily")]
    pub pitch_family: Option<i8>,
    #[xml(attr = "charset")]
    pub charset: Option<i8>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:latin")]
pub struct Latin<'a> {
    #[xml(attr = "typeface")]
    pub typeface: Option<Cow<'a, str>>,
    #[xml(attr = "panose")]
    pub panose: Option<Cow<'a, str>>,
    #[xml(attr = "pitchFamily")]
    pub pitch_family: Option<i8>,
    #[xml(attr = "charset")]
    pub charset: Option<i8>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:font")]
pub struct Font<'a> {
    #[xml(attr = "script")]
    pub script: Cow<'a, str>,
    #[xml(attr = "typeface")]
    pub typeface: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:fmtScheme")]
pub struct FmtScheme<'a> {
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,
    #[xml(child = "a:fillStyleLst")]
    pub fill_style_lst: FillStyleLst,
    #[xml(child = "a:lnStyleLst")]
    pub in_style_lst: LineStyleList,
    #[xml(child = "a:effectStyleLst")]
    pub effect_style_lst: EffectStyleLst,
    #[xml(child = "a:bgFillStyleLst")]
    pub bg_fill_style_lst: BgFillStyleLst,
}

__define_struct_vec! {
    ("a:fillStyleLst", FillStyleLst, FillStyleLstChoice) {} {
        "a:noFill", NoFill    //No Fill
        "a:solidFill", SolidFill    //Solid Fill
        "a:gradFill", GradFill    //Gradient Fill
        //"a:blipFill", blipFill    //Picture Fill
        //"a:pattFill", pattFill    //Pattern Fill
        "a:grpFill", GrpFill    //Group Fill
    }
}

__define_struct! {
    ("a:noFill", NoFill) {

    }
}

__define_struct! {
    ("a:grpFill", GrpFill) {

    }
}

__define_struct_vec! {
    ("a:solidFill", SolidFill, SolidFillChoice) {} {
        "a:scrgbClr", ScrgbClr
        "a:srgbClr", SrgbClr
        "a:hslClr", HslClr
        "a:sysClr", SysClr
        "a:schemeClr", SchemeClr
        "a:prstClr", PrstClr
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:gradFill")]
pub struct GradFill {
    #[xml(attr = "rotWithShape")]
    pub rotate_with_shape: bool,
    #[xml(child = "a:lin")]
    pub linear_gradient_fill: Option<LinearGradientFill>,
    #[xml(child = "a:gsLst")]
    pub gradient_stop_list: Option<GradientStopList>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:gsLst")]
pub struct GradientStopList {
    #[xml(child = "a:gs")]
    pub gradient_stop: Vec<GradientStop>,
}

__define_struct_vec! {
    ("a:gs", GradientStop, GradientStopChoice) {
        "pos", position, isize
    } {
        "a:scrgbClr", ScrgbClr
        "a:srgbClr", SrgbClr
        "a:hslClr", HslClr
        "a:sysClr", SysClr
        "a:schemeClr", SchemeClr
        "a:prstClr", PrstClr
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:lin")]
pub struct LinearGradientFill {
    #[xml(attr = "ang")]
    pub angle: isize,
    #[xml(attr = "scaled")]
    pub scaled: bool,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:lnStyleLst")]
pub struct LineStyleList {
    #[xml(child = "a:ln")]
    pub outline: Vec<Outline>,
}

__define_struct_vec! {
    ("a:ln", Outline, OutlineChoice) {
        "w", line_with, isize
        "cap", cap, CapType
        "cmpd", cmpd, CompoundLineType
        "algn", algn, PenAlignment
    } {
        "a:solidFill", SolidFill
        "a:prstDash", PresetDash
        "a:miter", MiterLineJoin
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:prstDash")]
pub struct PresetDash {
    #[xml(attr = "val")]
    pub val: PresetDashType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum PresetDashType {
    #[default]
    Solid,
    Dot,
    Dash,
    LargeDash,
    LargeDashDot,
    LargeDashDotDot,
    SystemDot,
    SystemDash,
    SystemDashDot,
    SystemDashDotDot,
}

__string_enum! {
    PresetDashType {
        Solid = "solid",
        Dot = "dot",
        Dash = "dash",
        LargeDash = "lgDash",
        LargeDashDot = "lgDashDot",
        LargeDashDotDot = "lgDashDotDot",
        SystemDot = "sysDot",
        SystemDash = "sysDash",
        SystemDashDot = "sysDashDot",
        SystemDashDotDot = "sysDashDotDot",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:miter")]
pub struct MiterLineJoin {
    #[xml(attr = "lim")]
    pub limit: isize,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum PenAlignment {
    #[default]
    Center,
    Insert,
}

__string_enum! {
    PenAlignment {
        Center = "ctr",
        Insert = "in",
    }
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum CapType {
    #[default]
    Square,
    Round,
    Flat,
}

__string_enum! {
    CapType {
        Square = "sq",
        Round = "rnd",
        Flat = "flat",
    }
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum CompoundLineType {
    #[default]
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple,
}

__string_enum! {
    CompoundLineType {
        Single = "sng",
        Double = "dbl",
        ThickThin = "thickThin",
        ThinThick = "thinThick",
        Triple = "tri",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:effectStyleLst")]
pub struct EffectStyleLst {
    #[xml(child = "a:effectStyle")]
    pub effect_styles: Vec<EffectStyle>,
}

__define_struct_vec! {
    ("a:effectStyle", EffectStyle, EffectStyleChoice) {} {
        "a:effectLst", EffectList
    }
}

__define_struct_vec! {
    ("a:effectLst", EffectList, EffectListChoice) {} {
        "a:outerShdw", OuterShadow
        // "a:prstShdw", PresetShadow
        // "a:reflection", Reflection
        // "a:softEdge", SoftEdge
        // "a:glow", Glow
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:outerShdw")]
pub struct OuterShadow {
    #[xml(attr = "blurRad")]
    pub blur_radius: Option<isize>,
    #[xml(attr = "dist")]
    pub distance: Option<isize>,
    #[xml(attr = "dir")]
    pub direction: Option<isize>,
    #[xml(attr = "algn")]
    pub alignment: Option<PenAlignment>,
    #[xml(attr = "kx")]
    pub kx: Option<isize>,
    #[xml(attr = "ky")]
    pub ky: Option<isize>,
    #[xml(attr = "rotWithShape")]
    pub rotate_with_shape: Option<bool>,
    #[xml(child = "a:hslClr")]
    pub hsl_color: Option<HslClr>,
    #[xml(child = "a:prstClr")]
    pub prst_color: Option<PrstClr>,
    #[xml(child = "a:schemeClr")]
    pub scheme_color: Option<SchemeClr>,
    #[xml(child = "a:scrgbClr")]
    pub scrgb_color: Option<ScrgbClr>,
    #[xml(child = "a:srgbClr")]
    pub srgb_color: Option<SrgbClr>,
    #[xml(child = "a:sysClr")]
    pub sys_color: Option<SysClr>,
}

__define_struct_vec! {
    ("a:bgFillStyleLst", BgFillStyleLst, BgFillStyleLstChoice) {} {
        "a:noFill", NoFill    //No Fill
        "a:solidFill", SolidFill    //Solid Fill
        "a:gradFill", GradFill    //Gradient Fill
        //"a:blipFill", blipFill    //Picture Fill
        //"a:pattFill", pattFill    //Pattern Fill
        "a:grpFill", GrpFill    //Group Fill
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:objectDefaults")]
pub struct ObjectDefaults {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:extraClrSchemeLst")]
pub struct ExtraClrSchemeLst {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:extLst")]
pub struct ExtLst {
    #[xml(child = "a:ext")]
    pub ext: Vec<ExtAny>,
}

__define_struct! {
    ("a:ext", ExtAny) {
        "uri", uri, String
    }
}

impl<'a> XmlWrite for Theme<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let Theme {
            name,
            elements,
            defaults,
            extra_clr_scheme_lst,
            cust_clr_lst,
            ext_lst,
        } = self;

        log::debug!("[Theme] Started writing.");
        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("a:theme")?;

        writer.write_attribute("xmlns:a", SCHEMA_DRAWINGML)?;

        if let Some(n) = name {
            writer.write_attribute("name", n)?;
        }

        writer.write_element_end_open()?;

        elements.to_writer(writer)?;
        write_attr(defaults, writer)?;
        write_attr(extra_clr_scheme_lst, writer)?;
        write_attr(cust_clr_lst, writer)?;
        write_attr(ext_lst, writer)?;

        writer.write_element_end_close("a:theme")?;

        log::debug!("[Theme] Finished writing.");

        Ok(())
    }
}

__xml_test_suites!(
    Theme,
    Theme::default(),
    format!(
        r#"{}<a:theme xmlns:a="{}"><a:themeElements><a:clrScheme name=""><a:dk1/><a:lt1/><a:dk2/><a:lt2/><a:accent1/><a:accent2/><a:accent3/><a:accent4/><a:accent5/><a:accent6/><a:hlink/><a:folHlink/></a:clrScheme><a:fontScheme><a:majorFont><a:latin/><a:ea/><a:cs/></a:majorFont><a:minorFont><a:latin/><a:ea/><a:cs/></a:minorFont></a:fontScheme><a:fmtScheme><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#,
        crate::schema::SCHEMA_XML,
        SCHEMA_DRAWINGML,
    )
    .as_str(),
);
