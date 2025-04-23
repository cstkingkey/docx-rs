use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{
    __setter, __string_enum, __xml_test_suites,
    formatting::{Bold, Color, Dstrike, Fonts, Italics, Lang, Outline, Size, Strike, Underline},
};

use super::{BoldComplex, Caps, Highlight, ItalicsComplex, Position, SmallCaps, VertAlign};

/// Character Property
///
/// ```rust
/// use docx_rust::formatting::{CharacterProperty, UnderlineStyle, HighlightType};
///
/// let prop = CharacterProperty::default()
///     .style_id("foo")
///     .color("00ff00")
///     .color(0xff0000)
///     .color((0x00, 0x00, 0xff))
///     .size(42isize)
///     .bold(true)
///     .italics(false)
///     .strike(true)
///     .dstrike(false)
///     .outline(true)
///     .highlight(HighlightType::Yellow)
///     .underline("00ff00")
///     .underline(("ff0000", UnderlineStyle::Dash));
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rPr")]
pub struct CharacterProperty<'a> {
    /// Specifies the style ID of the character style.
    #[xml(child = "w:rStyle")]
    pub style_id: Option<CharacterStyleId<'a>>,
    /// Specifies the font.
    #[xml(child = "w:rFonts")]
    pub fonts: Option<Fonts>,
    /// Specifies that the text of the text run is to be bold.
    #[xml(child = "w:b")]
    pub bold: Option<Bold>,
    #[xml(child = "w:bCs")]
    pub bold_complex: Option<BoldComplex>,
    /// Specifies that the text of the text run is to be italics.
    #[xml(child = "w:i")]
    pub italics: Option<Italics>,
    ///  Complex Script Italics
    #[xml(child = "w:iCs")]
    pub italics_complex: Option<ItalicsComplex>,
    ///  Display All Characters As Capital Letters
    #[xml(child = "w:caps")]
    pub caps: Option<Caps>,
    ///  Small Caps
    #[xml(child = "w:smallCaps")]
    pub small_caps: Option<SmallCaps>,
    /// Specifies that the contents are to be displayed with a horizontal line through the center of the line.
    #[xml(child = "w:strike")]
    pub strike: Option<Strike>,
    /// Specifies that the contents are to be displayed with two horizontal lines through each character.
    #[xml(child = "w:dstrike")]
    pub dstrike: Option<Dstrike>,
    /// Specifies that the content should be displayed as if it had an outline.
    #[xml(child = "w:outline")]
    pub outline: Option<Outline>,
    ///  Shadow
    #[xml(child = "w:shadow")]
    pub shadow: Option<Shadow>,
    ///  Embossing
    #[xml(child = "w:emboss")]
    pub emboss: Option<Emboss>,
    ///  Imprinting
    #[xml(child = "w:imprint")]
    pub imprint: Option<Imprint>,
    ///  Do Not Check Spelling or Grammar
    #[xml(child = "w:noProof")]
    pub no_proof: Option<NoProof>,
    /// Use Document Grid Settings For Inter-Character Spacing
    #[xml(child = "w:snapToGrid")]
    pub snap_to_grid: Option<super::SnapToGrid>,
    ///  Hidden Text
    #[xml(child = "w:vanish")]
    pub vanish: Option<Vanish>,
    ///  Web Hidden Text
    #[xml(child = "w:webHidden")]
    pub web_hidden: Option<WebHidden>,
    /// Specifies the color to be used to display text.
    #[xml(child = "w:color")]
    pub color: Option<Color<'a>>,
    ///  Character Spacing Adjustment
    #[xml(child = "w:spacing")]
    pub spacing: Option<TextSpacing>,
    ///Expanded/Compressed Text
    #[xml(child = "w:w")]
    pub scale: Option<Scale>,
    ///  Font Kerning
    #[xml(child = "w:kern")]
    pub kern: Option<Kern>,
    /// Vertically Raised or Lowered Text
    #[xml(child = "w:position")]
    pub position: Option<Position>,
    /// Specifies the font size in half points.
    #[xml(child = "w:sz")]
    pub size: Option<Size>,
    ///  Complex Script Font Size
    #[xml(child = "w:szCs")]
    pub size_complex: Option<SizeComplex>,
    ///  Text Highlighting
    #[xml(child = "w:highlight")]
    pub highlight: Option<Highlight>,
    /// Specifies that the content should be displayed with an underline
    #[xml(child = "w:u")]
    pub underline: Option<Underline<'a>>,
    ///  Animated Text Effect
    #[xml(child = "w:effect")]
    pub effect: Option<Effect>,
    ///  Text Border
    #[xml(child = "w:bdr")]
    pub border: Option<TextBorder<'a>>,
    ///  Run Shading
    #[xml(child = "w:shd")]
    pub shading: Option<Shading<'a>>,
    ///  Manual Run Width
    #[xml(child = "w:fitText")]
    pub fit_text: Option<FitText>,
    /// Subscript/Superscript Text
    #[xml(child = "w:vertAlign")]
    pub vertical_align: Option<VertAlign>,
    ///  Right To Left Text
    #[xml(child = "w:rtl")]
    pub rtl: Option<RightToLeftText>,
    ///  Use Complex Script Formatting on Run
    #[xml(child = "w:cs")]
    pub complex_script: Option<ComplexScript>,
    ///  Emphasis Mark
    #[xml(child = "w:em")]
    pub emphasis: Option<Emphasis>,
    /// Specifies the language to be used.
    #[xml(child = "w:lang")]
    pub lang: Option<Lang<'a>>,
    ///  East Asian Typography Settings
    #[xml(child = "w:eastAsianLayout")]
    pub east_asian_layout: Option<EastAsianLayout>,
    ///  Paragraph Mark Is Always Hidden
    #[xml(child = "w:specVanish")]
    pub spec_vanish: Option<SpecVanish>,
    ///  Office Open XML Math
    #[xml(child = "w:oMath")]
    pub o_math: Option<OMath>,
}

impl<'a> CharacterProperty<'a> {
    __setter!(style_id: Option<CharacterStyleId<'a>>);
    __setter!(color: Option<Color<'a>>);
    __setter!(highlight: Option<Highlight>);
    __setter!(bold: Option<Bold>);
    __setter!(dstrike: Option<Dstrike>);
    __setter!(italics: Option<Italics>);
    __setter!(outline: Option<Outline>);
    __setter!(strike: Option<Strike>);
    __setter!(size: Option<Size>);
    __setter!(underline: Option<Underline<'a>>);
    __setter!(fonts: Option<Fonts>);
}

#[derive(Debug, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rStyle")]
pub struct CharacterStyleId<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

impl<'a, T: Into<Cow<'a, str>>> From<T> for CharacterStyleId<'a> {
    fn from(val: T) -> Self {
        CharacterStyleId { value: val.into() }
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:shadow")]
pub struct Shadow {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:emboss")]
pub struct Emboss {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:imprint")]
pub struct Imprint {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:noProof")]
pub struct NoProof {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

// #[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
// #[cfg_attr(test, derive(PartialEq))]
// #[xml(tag = "w:snapToGrid")]
// pub struct SnapToGrid {
//     #[xml(attr = "w:val")]
//     pub value: Option<bool>,
// }

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:vanish")]
pub struct Vanish {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:webHidden")]
pub struct WebHidden {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rtl")]
pub struct RightToLeftText {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:cs")]
pub struct ComplexScript {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:specVanish")]
pub struct SpecVanish {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:oMath")]
pub struct OMath {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

/// Positive or Negative Value in Twentieths of a Point
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:spacing")]
pub struct TextSpacing {
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
}

/// Text Expansion/Compression Percentage, 0..=600.
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:w")]
pub struct Scale {
    #[xml(attr = "w:val")]
    pub value: Option<u16>,
}

/// Measurement in Half-Points
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:szCs")]
pub struct SizeComplex {
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
}

/// Measurement in Half-Points
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:kern")]
pub struct Kern {
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:effect")]
pub struct Effect {
    #[xml(attr = "w:val")]
    pub value: Option<EffectType>,
}

#[derive(Debug, Clone, Default)]
#[cfg_attr(test, derive(PartialEq))]
pub enum EffectType {
    BlinkBackground, //Blinking Background Animation
    Lights,          //Colored Lights Animation
    AntsBlack,       //Black Dashed Line Animation
    AntsRed,         //Marching Red Ants
    Shimmer,         //Shimmer Animation
    Sparkle,         //Sparkling Lights Animation
    #[default]
    None, //No Animation
}

__string_enum! {
    EffectType {
        BlinkBackground = "blinkBackground",
        Lights = "lights",
        AntsBlack = "antsBlack",
        AntsRed = "antsRed",
        Shimmer = "shimmer",
        Sparkle = "sparkle",
        None = "none",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:eastAsianLayout")]
pub struct EastAsianLayout {
    #[xml(attr = "w:id")]
    pub id: Option<isize>,
    #[xml(attr = "w:combine")]
    pub combine: Option<bool>,
    #[xml(attr = "w:combineBrackets")]
    pub combine_brackets: Option<CombineBracketsType>,
    #[xml(attr = "w:vert")]
    pub vert: Option<bool>,
    #[xml(attr = "w:vertCompress")]
    pub vert_compress: Option<bool>,
}

#[derive(Debug, Clone, Default)]
#[cfg_attr(test, derive(PartialEq))]
pub enum CombineBracketsType {
    #[default]
    None, //	No Enclosing Brackets
    Round,  //	Round Brackets
    Square, //	Square Brackets
    Angle,  //	Angle Brackets
    Curly,  //	Curly Brackets
}

__string_enum! {
    CombineBracketsType {
        None = "none",
        Round = "round",
        Square = "square",
        Angle = "angle",
        Curly = "curly",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:fitText")]
pub struct FitText {
    // Measurement in Twentieths of a Point
    #[xml(attr = "w:val")]
    pub value: Option<isize>,
    #[xml(attr = "w:id")]
    pub id: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bdr")]
pub struct TextBorder<'a> {
    #[xml(attr = "w:val")]
    pub style: super::BorderStyle,
    #[xml(attr = "w:color")]
    pub color: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeColor")]
    pub theme_color: Option<ThemeColor>,
    #[xml(attr = "w:themeTint")]
    pub theme_tint: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeShade")]
    pub theme_shade: Option<Cow<'a, str>>,
    #[xml(attr = "w:sz")]
    pub size: Option<isize>, // Measurement in Eighths of a Point
    #[xml(attr = "w:space")]
    pub space: Option<isize>,
    #[xml(attr = "w:shadow")]
    pub shadow: Option<bool>,
    #[xml(attr = "w:frame")]
    pub frame: Option<bool>,
}

#[derive(Debug, Clone, Default)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ThemeColor {
    #[default]
    Dark1, //Dark 1 Theme Color.
    Light1,            //Light 1 Theme Color.
    Dark2,             //Dark 2 Theme Color.
    Light2,            //Light 2 Theme Color.
    Accent1,           //Accent 1 Theme Color.
    Accent2,           //Accent 2 Theme Color.
    Accent3,           //Accent 3 Theme Color.
    Accent4,           //Accent 4 Theme Color.
    Accent5,           //Accent 5 Theme Color.
    Accent6,           //Accent 6 Theme Color.
    Hyperlink,         //Hyperlink Theme Color.
    FollowedHyperlink, //Followed Hyperlink Theme Color.
    None,              //No Theme Color.
    Background1,       //Background 1 Theme Color.
    Text1,             //Text 1 Theme Color.
    Background2,       //Background 2 Theme Color.
    Text2,             //Text 2 Theme Color.
}

__string_enum! {
    ThemeColor {
        Dark1 = "dark1",
        Light1 = "light1",
        Dark2 = "dark2",
        Light2 = "light2",
        Accent1 = "accent1",
        Accent2 = "accent2",
        Accent3 = "accent3",
        Accent4 = "accent4",
        Accent5 = "accent5",
        Accent6 = "accent6",
        Hyperlink = "hyperlink",
        FollowedHyperlink = "followedHyperlink",
        None = "none",
        Background1 = "background1",
        Text1 = "text1",
        Background2 = "background2",
        Text2 = "text2",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:shd")]
pub struct Shading<'a> {
    #[xml(attr = "w:val")]
    pub style: Option<ShadingStyle>,
    #[xml(attr = "w:color")]
    pub color: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeColor")]
    pub theme_color: Option<ThemeColor>,
    #[xml(attr = "w:themeTint")]
    pub theme_tint: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeShade")]
    pub theme_shade: Option<Cow<'a, str>>,
    #[xml(attr = "w:fill")]
    pub fill: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeFill")]
    pub theme_fill: Option<ThemeColor>,
    #[xml(attr = "w:themeFillTint")]
    pub theme_fill_tint: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeFillShade")]
    pub theme_fill_shade: Option<Cow<'a, str>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:em")]
pub struct Emphasis {
    #[xml(attr = "w:val")]
    pub value: Option<EmphasisType>,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ShadingStyle {
    #[default]
    Nil, //No Pattern.
    Clear,                 //No Pattern.
    Solid,                 //100% Fill Pattern.
    HorzStripe,            //Horizontal Stripe Pattern.
    VertStripe,            //Vertical Stripe Pattern.
    ReverseDiagStripe,     //Reverse Diagonal Stripe Pattern.
    DiagStripe,            //Diagonal Stripe Pattern.
    HorzCross,             //Horizontal Cross Pattern.
    DiagCross,             //Diagonal Cross Pattern.
    ThinHorzStripe,        //Thin Horizontal Stripe Pattern.
    ThinVertStripe,        //Thin Vertical Stripe Pattern.
    ThinReverseDiagStripe, //Thin Reverse Diagonal Stripe Pattern.
    ThinDiagStripe,        //Thin Diagonal Stripe Pattern.
    ThinHorzCross,         //Thin Horizontal Cross Pattern.
    ThinDiagCross,         //Thin Diagonal Cross Pattern.
    Pct5,                  //5% Fill Pattern.
    Pct10,                 //10% Fill Pattern.
    Pct12,                 //12.5% Fill Pattern.
    Pct15,                 //15% Fill Pattern.
    Pct20,                 //20% Fill Pattern.
    Pct25,                 //25% Fill Pattern.
    Pct30,                 //30% Fill Pattern.
    Pct35,                 //35% Fill Pattern.
    Pct37,                 //37.5% Fill Pattern.
    Pct40,                 //40% Fill Pattern.
    Pct45,                 //45% Fill Pattern.
    Pct50,                 //50% Fill Pattern.
    Pct55,                 //55% Fill Pattern.
    Pct60,                 //60% Fill Pattern.
    Pct62,                 //62.5% Fill Pattern.
    Pct65,                 //65% Fill Pattern.
    Pct70,                 //70% Fill Pattern.
    Pct75,                 //75% Fill Pattern.
    Pct80,                 //80% Fill Pattern.
    Pct85,                 //85% Fill Pattern.
    Pct87,                 //87.5% Fill Pattern.
    Pct90,                 //90% Fill Pattern.
    Pct95,                 //95% Fill Pattern.
}

__string_enum! {
    ShadingStyle {
        Nil = "nil",
        Clear = "clear",
        Solid = "solid",
        HorzStripe = "horzStripe",
        VertStripe = "vertStripe",
        ReverseDiagStripe = "reverseDiagStripe",
        DiagStripe = "diagStripe",
        HorzCross = "horzCross",
        DiagCross = "diagCross",
        ThinHorzStripe = "thinHorzStripe",
        ThinVertStripe = "thinVertStripe",
        ThinReverseDiagStripe = "thinReverseDiagStripe",
        ThinDiagStripe = "thinDiagStripe",
        ThinHorzCross = "thinHorzCross",
        ThinDiagCross = "thinDiagCross",
        Pct5 = "pct5",
        Pct10 = "pct10",
        Pct12 = "pct12",
        Pct15 = "pct15",
        Pct20 = "pct20",
        Pct25 = "pct25",
        Pct30 = "pct30",
        Pct35 = "pct35",
        Pct37 = "pct37",
        Pct40 = "pct40",
        Pct45 = "pct45",
        Pct50 = "pct50",
        Pct55 = "pct55",
        Pct60 = "pct60",
        Pct62 = "pct62",
        Pct65 = "pct65",
        Pct70 = "pct70",
        Pct75 = "pct75",
        Pct80 = "pct80",
        Pct85 = "pct85",
        Pct87 = "pct87",
        Pct90 = "pct90",
        Pct95 = "pct95",
    }
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum EmphasisType {
    #[default]
    None, //	No Emphasis Mark
    Dot,      //	Dot Emphasis Mark Above Characters
    Comma,    //	Comma Emphasis Mark Above Characters
    Circle,   //	Circle Emphasis Mark Above Characters
    UnderDot, //	Dot Emphasis Mark Below Characters
}

__string_enum! {
    EmphasisType {
        None = "none",
        Dot = "dot",
        Comma = "comma",
        Circle = "circle",
        UnderDot = "underdot",
    }
}

__xml_test_suites!(
    CharacterProperty,
    CharacterProperty::default(),
    r#"<w:rPr/>"#,
    CharacterProperty::default().style_id("id"),
    r#"<w:rPr><w:rStyle w:val="id"/></w:rPr>"#,
    CharacterProperty::default().color("00ff00"),
    r#"<w:rPr><w:color w:val="00ff00"/></w:rPr>"#,
    CharacterProperty::default().size(42isize),
    r#"<w:rPr><w:sz w:val="42"/></w:rPr>"#,
    CharacterProperty::default().bold(true),
    r#"<w:rPr><w:b w:val="true"/></w:rPr>"#,
    CharacterProperty::default().italics(false),
    r#"<w:rPr><w:i w:val="false"/></w:rPr>"#,
    CharacterProperty::default().outline(true),
    r#"<w:rPr><w:outline w:val="true"/></w:rPr>"#,
    CharacterProperty::default().strike(false),
    r#"<w:rPr><w:strike w:val="false"/></w:rPr>"#,
    CharacterProperty::default().dstrike(true),
    r#"<w:rPr><w:dstrike w:val="true"/></w:rPr>"#,
    CharacterProperty::default().underline(Underline::default()),
    r#"<w:rPr><w:u/></w:rPr>"#,
    CharacterProperty::default().fonts(Fonts::default().east_asia("宋体")),
    r#"<w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr>"#,
);
