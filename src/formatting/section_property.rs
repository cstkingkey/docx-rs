#![allow(dead_code)]
use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{
    __define_enum, __define_struct, __string_enum,
    document::HeaderFooterReference,
    formatting::{PageCols, PageGrid, PageMargin, PageSize},
};

use super::Bidi;

/// Section Properties
///
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sectPr")]
pub struct SectionProperty<'a> {
    //#[xml(attr = "w14:paraId")]
    //pub id: Option<Cow<'a, str>>,
    //#[xml(attr = "w14:textId")]
    //pub text_id: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidR")]
    pub rsid_r: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidRDefault")]
    pub rsid_r_default: Option<Cow<'a, str>>,

    #[xml(child = "w:headerReference", child = "w:footerReference")]
    /// Specifies the content of a run
    pub header_footer_references: Vec<HeaderFooterReference<'a>>,
    ///  Section-Wide Footnote Properties
    #[xml(child = "w:footnotePr")]
    pub footnote_property: Option<FootnoteProperty>,
    ///  Section-Wide Endnote Properties
    #[xml(child = "w:endnotePr")]
    pub endnote_property: Option<EndnoteProperty>,
    ///  Section Type
    #[xml(child = "w:type")]
    pub ty: Option<SectionTypeP>,
    #[xml(child = "w:pgSz")]
    pub page_size: Option<PageSize>,
    #[xml(child = "w:pgMar")]
    pub page_margin: Option<PageMargin>,
    ///  Paper Source Information
    #[xml(child = "w:paperSrc")]
    pub paper_source: Option<PaperSource>,
    ///  Page Borders
    #[xml(child = "w:pgBorders")]
    pub page_borders: Option<PgBorders>,
    ///  Line Numbering Settings
    #[xml(child = "w:lnNumType")]
    pub line_numbering: Option<PgLnNumType>,
    ///  Page Numbering Settings
    #[xml(child = "w:pgNumType")]
    pub page_numbering: Option<PgNumType>,
    ///  Column Definitions
    #[xml(child = "w:cols")]
    pub cols: Option<PageCols>,
    ///  Only Allow Editing of Form Fields
    #[xml(child = "w:formProt")]
    pub form_prot: Option<FormProt>,
    /////  Vertical Text Alignment on Page
    #[xml(child = "w:vAlign")]
    pub v_align: Option<VAlign>,
    ///  Suppress Endnotes In Document
    #[xml(child = "w:noEndnote")]
    pub no_endnote: Option<NoEndnote>,
    ///  Different First Page Headers and Footers
    #[xml(child = "w:titlePg")]
    pub title_page: Option<TitlePage>,
    ///  Text Flow Direction
    #[xml(child = "w:textDirection")]
    pub text_direction: Option<TextDirection>,
    ///  Right to Left Section Layout
    #[xml(child = "w:bidi")]
    pub bidi: Option<Bidi>,
    ///  Gutter on Right Side of Page
    #[xml(child = "w:rtlGutter")]
    pub rtl_gutter: Option<RtlGutter>,
    ///  Document Grid
    #[xml(child = "w:docGrid")]
    pub grid: Option<PageGrid>,
    /////  Reference to Printer Settings Data
    //#[xml(child = "w:printerSettings")]
    //pub printer_settings: Option<PrinterSettings>,
    /// Revision Information for Section Properties
    #[xml(child = "w:sectPrChange")]
    pub revision: Option<Revision<'a>>,
}

/// Previous Section Properties
///
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sectPr")]
pub struct PreviousSectionProperty<'a> {
    //#[xml(attr = "w14:paraId")]
    //pub id: Option<Cow<'a, str>>,
    //#[xml(attr = "w14:textId")]
    //pub text_id: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidR")]
    pub rsid_r: Option<Cow<'a, str>>,
    #[xml(attr = "w:rsidRDefault")]
    pub rsid_r_default: Option<Cow<'a, str>>,

    ///  Section-Wide Footnote Properties
    #[xml(child = "w:footnotePr")]
    pub footnote_property: Option<FootnoteProperty>,
    ///  Section-Wide Endnote Properties
    #[xml(child = "w:endnotePr")]
    pub endnote_property: Option<EndnoteProperty>,
    ///  Section Type
    #[xml(child = "w:type")]
    pub ty: Option<SectionTypeP>,
    #[xml(child = "w:pgSz")]
    pub page_size: Option<PageSize>,
    #[xml(child = "w:pgMar")]
    pub page_margin: Option<PageMargin>,
    ///  Paper Source Information
    #[xml(child = "w:paperSrc")]
    pub paper_source: Option<PaperSource>,
    ///  Page Borders
    #[xml(child = "w:pgBorders")]
    pub page_borders: Option<PgBorders>,
    ///  Line Numbering Settings
    #[xml(child = "w:lnNumType")]
    pub line_numbering: Option<PgLnNumType>,
    ///  Page Numbering Settings
    #[xml(child = "w:pgNumType")]
    pub page_numbering: Option<PgNumType>,
    ///  Column Definitions
    #[xml(child = "w:cols")]
    pub cols: Option<PageCols>,
    ///  Only Allow Editing of Form Fields
    #[xml(child = "w:formProt")]
    pub form_prot: Option<FormProt>,
    /////  Vertical Text Alignment on Page
    #[xml(child = "w:vAlign")]
    pub v_align: Option<VAlign>,
    ///  Suppress Endnotes In Document
    #[xml(child = "w:noEndnote")]
    pub no_endnote: Option<NoEndnote>,
    ///  Different First Page Headers and Footers
    #[xml(child = "w:titlePg")]
    pub title_page: Option<TitlePage>,
    ///  Text Flow Direction
    #[xml(child = "w:textDirection")]
    pub text_direction: Option<TextDirection>,
    ///  Right to Left Section Layout
    #[xml(child = "w:bidi")]
    pub bidi: Option<Bidi>,
    ///  Gutter on Right Side of Page
    #[xml(child = "w:rtlGutter")]
    pub rtl_gutter: Option<RtlGutter>,
    ///  Document Grid
    #[xml(child = "w:docGrid")]
    pub grid: Option<PageGrid>,
    /////  Reference to Printer Settings Data
    //#[xml(child = "w:printerSettings")]
    //pub printer_settings: Option<PrinterSettings>,
}

__define_struct! {
    ("w:pgBorders", PgBorders) {
        "w:zOrder",	z_order, PageBorderZOrder	//Z-Ordering of Page Border
        "w:display", display, PageBorderDisplay //Pages to Display Page Borders
        "w:offsetFrom",	offset_from, PageBorderOffset 	//Page Border Positioning
    } {
        "w:top", top, PgTopBorder
        "w:left", left, PgLeftBorder
        "w:bottom", bottom, PgBottomBorder
        "w:right", right, PgRightBorder
    }
}

__define_enum! {
    PageBorderZOrder {
        Front = "front", // Page Border Ahead of Text
        Back = "back", // Page Border Behind Text
    }
}

__define_enum! {
    PageBorderDisplay {
        AllPages = "allPages", // Display Page Border on All Pages
        FirstPage = "firstPage", // Display Page Border on First Page
        NotFirstPage = "notFirstPage", // Display Page Border on All Pages Except First
    }
}

__define_enum! {
    PageBorderOffset {
        Page = "page", // Page Border Is Positioned Relative to Page Edges
        Text = "text", // Page Border Is Positioned Relative to Text Extents
    }
}

__define_struct! {
    ("w:top", PgTopBorder) {
        "w:val", style, super::BorderStyle
        "w:color", color, String
        "w:themeColor", theme_color, crate::formatting::ThemeColor
        "w:themeTint", theme_tint, String
        "w:themeShade", theme_shade, String
        "w:sz", size, isize // Measurement in Eighths of a Point
        "w:space", space, isize
        "w:shadow", shadow, bool
        "w:frame", frame, bool
    }
}

__define_struct! {
    ("w:bottom", PgBottomBorder) {
        "w:val", style, super::BorderStyle
        "w:color", color, String
        "w:themeColor", theme_color, crate::formatting::ThemeColor
        "w:themeTint", theme_tint, String
        "w:themeShade", theme_shade, String
        "w:sz", size, isize // Measurement in Eighths of a Point
        "w:space", space, isize
        "w:shadow", shadow, bool
        "w:frame", frame, bool
    }
}

__define_struct! {
    ("w:left", PgLeftBorder) {
        "w:val", style, super::BorderStyle
        "w:color", color, String
        "w:themeColor", theme_color, crate::formatting::ThemeColor
        "w:themeTint", theme_tint, String
        "w:themeShade", theme_shade, String
        "w:sz", size, isize // Measurement in Eighths of a Point
        "w:space", space, isize
        "w:shadow", shadow, bool
        "w:frame", frame, bool
    }
}

__define_struct! {
    ("w:right", PgRightBorder) {
        "w:val", style, super::BorderStyle
        "w:color", color, String
        "w:themeColor", theme_color, crate::formatting::ThemeColor
        "w:themeTint", theme_tint, String
        "w:themeShade", theme_shade, String
        "w:sz", size, isize // Measurement in Eighths of a Point
        "w:space", space, isize
        "w:shadow", shadow, bool
        "w:frame", frame, bool
    }
}

__define_struct! {
    ("w:lnNumType", PgLnNumType) {
        "w:countBy", count_by, isize //	[0..1]	w:ST_DecimalNumber	Line Number Increments to Display
        "w:start", start, isize //	[0..1]	w:ST_DecimalNumber	Line Numbering Starting Value
        "w:distance", distance, isize //	[0..1]	w:ST_TwipsMeasure	Distance Between Text and Line Numbering
        "w:restart", restart, LineNumberRestart //	[0..1]	w:ST_LineNumberRestart	Line Numbering Restart Setting
    }
}

__define_enum! {
    LineNumberRestart
    {
        NewPage = "newPage", // Restart Line Numbering on Each Page
        NewSection = "newSection", // Restart Line Numbering for Each Section
        Continuous = "continuous", // Continue Line Numbering From Previous Section
    }
}

__define_struct! {
    ("w:pgNumType", PgNumType) {
        "w:fmt", fmt, NumberFormat	//Page Number Format
        "w:start", start, isize //	[0..1]	w:ST_DecimalNumber	Starting Page Number
        "w:chapStyle", chap_style, isize//	[0..1]	w:ST_DecimalNumber	Chapter Heading Style
        "w:chapSep", chap_sep, ChapterSep//	Chapter Separator Character
    }
}

__define_enum! {
    NumberFormat {
        Decimal = "decimal", // Decimal Numbers
        UpperRoman = "upperRoman", // Uppercase Roman Numerals
        LowerRoman = "lowerRoman", // Lowercase Roman Numerals
        UpperLetter = "upperLetter", // Uppercase Latin Alphabet
        LowerLetter = "lowerLetter", // Lowercase Latin Alphabet
        Ordinal = "ordinal", // Ordinal
        CardinalText = "cardinalText", // Cardinal Text
        OrdinalText = "ordinalText", // Ordinal Text
        Hex = "hex", // Hexadecimal Numbering
        Chicago = "chicago", // Chicago Manual of Style
        IdeographDigital = "ideographDigital", // Ideographs
        JapaneseCounting = "japaneseCounting", // Japanese Counting System
        Aiueo = "aiueo", // AIUEO Order Hiragana
        Iroha = "iroha", // Iroha Ordered Katakana
        DecimalFullWidth = "decimalFullWidth", // Double Byte Arabic Numerals
        DecimalHalfWidth = "decimalHalfWidth", // Single Byte Arabic Numerals
        JapaneseLegal = "japaneseLegal", // Japanese Legal Numbering
        JapaneseDigitalTenThousand = "japaneseDigitalTenThousand", // Japanese Digital Ten Thousand Counting System
        DecimalEnclosedCircle = "decimalEnclosedCircle", // Decimal Numbers Enclosed in a Circle
        DecimalFullWidth2 = "decimalFullWidth2", // Double Byte Arabic Numerals Alternate
        AiueoFullWidth = "aiueoFullWidth", // Full-Width AIUEO Order Hiragana
        IrohaFullWidth = "irohaFullWidth", // Full-Width Iroha Ordered Katakana
        DecimalZero = "decimalZero", // Initial Zero Arabic Numerals
        Bullet = "bullet", // Bullet
        Ganada = "ganada", // Korean Ganada Numbering
        Chosung = "chosung", // Korean Chosung Numbering
        DecimalEnclosedFullstop = "decimalEnclosedFullstop", // Decimal Numbers Followed by a Period
        DecimalEnclosedParen = "decimalEnclosedParen", // Decimal Numbers Enclosed in Parenthesis
        DecimalEnclosedCircleChinese = "decimalEnclosedCircleChinese", // Decimal Numbers Enclosed in a Circle
        IdeographEnclosedCircle = "ideographEnclosedCircle", // Ideographs Enclosed in a Circle
        IdeographTraditional = "ideographTraditional", // Traditional Ideograph Format
        IdeographZodiac = "ideographZodiac", // Zodiac Ideograph Format
        IdeographZodiacTraditional = "ideographZodiacTraditional", // Traditional Zodiac Ideograph Format
        TaiwaneseCounting = "taiwaneseCounting", // Taiwanese Counting System
        IdeographLegalTraditional = "ideographLegalTraditional", // Traditional Legal Ideograph Format
        TaiwaneseCountingThousand = "taiwaneseCountingThousand", // Taiwanese Counting Thousand System
        TaiwaneseDigital = "taiwaneseDigital", // Taiwanese Digital Counting System
        ChineseCounting = "chineseCounting", // Chinese Counting System
        ChineseLegalSimplified = "chineseLegalSimplified", // Chinese Legal Simplified Format
        ChineseCountingThousand = "chineseCountingThousand", // Chinese Counting Thousand System
        KoreanDigital = "koreanDigital", // Korean Digital Counting System
        KoreanCounting = "koreanCounting", // Korean Counting System
        KoreanLegal = "koreanLegal", // Korean Legal Numbering
        KoreanDigital2 = "koreanDigital2", // Korean Digital Counting System Alternate
        VietnameseCounting = "vietnameseCounting", // Vietnamese Numerals
        RussianLower = "russianLower", // Lowercase Russian Alphabet
        RussianUpper = "russianUpper", // Uppercase Russian Alphabet
        None = "none", // No Numbering
        NumberInDash = "numberInDash", // Number With Dashes
        Hebrew1 = "hebrew1", // Hebrew Numerals
        Hebrew2 = "hebrew2", // Hebrew Alphabet
        ArabicAlpha = "arabicAlpha", // Arabic Alphabet
        ArabicAbjad = "arabicAbjad", // Arabic Abjad Numerals
        HindiVowels = "hindiVowels", // Hindi Vowels
        HindiConsonants = "hindiConsonants", // Hindi Consonants
        HindiNumbers = "hindiNumbers", // Hindi Numbers
        HindiCounting = "hindiCounting", // Hindi Counting System
        ThaiLetters = "thaiLetters", // Thai Letters
        ThaiNumbers = "thaiNumbers", // Thai Numerals
        ThaiCounting = "thaiCounting", // Thai Counting System
    }
}

__define_enum! {
    ChapterSep {
        Hyphen = "hyphen", // Hyphen Chapter Separator
        Period = "period", // Period Chapter Separator
        Colon = "colon", // Colon Chapter Separator
        EmDash = "emDash", // Em Dash Chapter Separator
        EnDash = "enDash", // En Dash Chapter Separator
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:sectPrChange")]
pub struct Revision<'a> {
    #[xml(attr = "w:id")]
    pub id: isize,
    #[xml(attr = "w:author")]
    pub author: Cow<'a, str>,
    #[xml(attr = "w:date")]
    pub date: Option<Cow<'a, str>>,

    #[xml(child = "w:secPr")]
    pub section_property: Option<PreviousSectionProperty<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:type")]
pub struct SectionTypeP {
    #[xml(attr = "w:val")]
    pub ty: Option<SectionType>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:formProt")]
pub struct FormProt {
    #[xml(attr = "w:val")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:noEndnote")]
pub struct NoEndnote {
    #[xml(attr = "w:val")]
    pub val: Option<bool>,
}

// #[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
// #[cfg_attr(test, derive(PartialEq))]
// #[xml(tag = "w:bidi")]
// pub struct Bidi {
//     #[xml(attr = "w:val")]
//     pub val: Option<bool>,
// }

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:paperSrc")]
pub struct PaperSource {
    #[xml(attr = "w:first")]
    pub first: Option<isize>,
    #[xml(attr = "w:other")]
    pub other: Option<isize>,
}

__define_struct! {
    ("w:textAlignment", TextAlignment) {
        "w:val", val, TextAlignmentType
    }
}

__define_enum! {
    TextAlignmentType {
        Top = "top", // Align Text at Top
        Center = "center", // Align Text at Center
        Baseline = "baseline", // Align Text at Baseline
        Bottom = "bottom", // Align Text at Bottom
        Auto = "auto", // Automatically Determine Alignment
    }
}

__define_struct! {
    ("w:textboxTightWrap", TextboxTightWrap) {
        "w:val", val, TextboxTightWrapType
    }
}

__define_enum! {
    TextboxTightWrapType {
        None = "none", // Do Not Tight Wrap
        AllLines = "allLines", // Tight Wrap All Lines
        FirstAndLastLine = "firstAndLastLine", // Tight Wrap First and Last Lines
        FirstLineOnly = "firstLineOnly", // Tight Wrap First Line
        LastLineOnly = "lastLineOnly", // Tight Wrap Last Line
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:textDirection")]
pub struct TextDirection {
    #[xml(attr = "w:val")]
    pub val: TextDirectionType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum TextDirectionType {
    #[default]
    LrTb, //	Left to Right, Top to Bottom
    TbRl,  //		Top to Bottom, Right to Left
    BtLr,  //		Bottom to Top, Left to Right
    LrTbV, //		Left to Right, Top to Bottom Rotated
    TbRlV, //		Top to Bottom, Right to Left Rotated
    TbLrV, //		Top to Bottom, Left to Right Rotated
}

__string_enum! {
    TextDirectionType {
        LrTb = "lrTb",
        TbRl = "tbRl",
        BtLr = "btLr",
        LrTbV = "lrTbV",
        TbRlV = "tbRlV",
        TbLrV = "tbLrV",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:vAlign")]
pub struct VAlign {
    #[xml(attr = "w:val")]
    pub val: VAlignType,
}

impl From<VAlignType> for VAlign {
    fn from(al_type: VAlignType) -> Self {
        Self { val: al_type }
    }
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum VAlignType {
    #[default]
    Top, //	Align Top
    Center, //	Align Center
    Both,   //	Vertical Justification
    Bottom, //	Align Bottom
}

__string_enum! {
    VAlignType {
        Top = "top",
        Center = "center",
        Both = "both",
        Bottom = "bottom",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rtlGutter")]
pub struct RtlGutter {
    #[xml(attr = "w:val")]
    pub val: Option<bool>,
}

#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum SectionType {
    NextPage,
    NextColumn,
    Continuous,
    EvenPage,
    OddPage,
}

__string_enum! {
    SectionType {
        NextPage = "nextPage",
        NextColumn = "nextColumn",
        Continuous = "continuous",
        EvenPage = "evenPage",
        OddPage = "oddPage",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:titlePg")]
pub struct TitlePage {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnotePr")]
pub struct FootnoteProperty {
    #[xml(child = "w:pos")]
    pub position: Option<FootnotePosition>,
    #[xml(child = "w:numFmt")]
    pub num_fmt: Option<NumFmt>,
    /// Footnote and Endnote Numbering Starting Value
    #[xml(child = "w:numStart")]
    pub num_start: Option<NumStart>,
    /// Footnote and Endnote Numbering Restart Location
    #[xml(child = "w:numRestart")]
    pub num_restart: Option<NumRestart>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnotePr")]
pub struct EndnoteProperty {
    #[xml(child = "w:pos")]
    pub position: Option<EndnotePosition>,
    #[xml(child = "w:numFmt")]
    pub num_fmt: Option<NumFmt>,
    /// Footnote and Endnote Numbering Starting Value
    #[xml(child = "w:numStart")]
    pub num_start: Option<NumStart>,
    /// Footnote and Endnote Numbering Restart Location
    #[xml(child = "w:numRestart")]
    pub num_restart: Option<NumRestart>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnotePr")]
pub struct FootnoteProperty2 {
    #[xml(child = "w:pos")]
    pub position: Option<FootnotePosition>,
    #[xml(child = "w:numFmt")]
    pub num_fmt: Option<NumFmt>,
    /// Footnote and Endnote Numbering Starting Value
    #[xml(child = "w:numStart")]
    pub num_start: Option<NumStart>,
    /// Footnote and Endnote Numbering Restart Location
    #[xml(child = "w:numRestart")]
    pub num_restart: Option<NumRestart>,
    #[xml(child = "w:footnote")]
    pub footnote: Vec<Footnote>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:footnote")]
pub struct Footnote {
    #[xml(attr = "w:id")]
    pub id: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnotePr")]
pub struct EndnoteProperty2 {
    #[xml(child = "w:pos")]
    pub position: Option<EndnotePosition>,
    #[xml(child = "w:numFmt")]
    pub num_fmt: Option<NumFmt>,
    /// Footnote and Endnote Numbering Starting Value
    #[xml(child = "w:numStart")]
    pub num_start: Option<NumStart>,
    /// Footnote and Endnote Numbering Restart Location
    #[xml(child = "w:numRestart")]
    pub num_restart: Option<NumRestart>,
    #[xml(child = "w:endnote")]
    pub endnote: Vec<Endnote>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:endnote")]
pub struct Endnote {
    #[xml(attr = "w:id")]
    pub id: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numRestart")]
pub struct NumRestart {
    #[xml(attr = "w:val")]
    pub val: NumRestartType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum NumRestartType {
    #[default]
    Continuous, //	Continue Numbering From Previous Section
    EachSect, //	Restart Numbering For Each Section
    EachPage, //	Restart Numbering On Each Page
}

__string_enum! {
    NumRestartType {
        Continuous = "continuous",
        EachSect = "eachSect",
        EachPage = "eachPage",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numStart")]
pub struct NumStart {
    #[xml(attr = "w:val")]
    pub val: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:numFmt")]
pub struct NumFmt {
    #[xml(attr = "w:val")]
    pub ty: NumFmtType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum NumFmtType {
    #[default]
    Decimal, //Decimal Numbers.
    UpperRoman,                   //Uppercase Roman Numerals.
    LowerRoman,                   //Lowercase Roman Numerals.
    UpperLetter,                  //Uppercase Latin Alphabet.
    LowerLetter,                  //Lowercase Latin Alphabet.
    Ordinal,                      //Ordinal.
    CardinalText,                 //Cardinal Text.
    OrdinalText,                  //Ordinal Text.
    Hex,                          //Hexadecimal Numbering.
    Chicago,                      //Chicago Manual of Style.
    IdeographDigital,             //Ideographs.
    JapaneseCounting,             //Japanese Counting System.
    Aiueo,                        //AIUEO Order Hiragana.
    Iroha,                        //Iroha Ordered Katakana.
    DecimalFullWidth,             //Double Byte Arabic Numerals.
    DecimalHalfWidth,             //Single Byte Arabic Numerals.
    JapaneseLegal,                //Japanese Legal Numbering.
    JapaneseDigitalTenThousand,   //Japanese Digital Ten Thousand Counting System.
    DecimalEnclosedCircle,        //Decimal Numbers Enclosed in a Circle.
    DecimalFullWidth2,            //Double Byte Arabic Numerals Alternate.
    AiueoFullWidth,               //Full-Width AIUEO Order Hiragana.
    IrohaFullWidth,               //Full-Width Iroha Ordered Katakana.
    DecimalZero,                  //Initial Zero Arabic Numerals.
    Bullet,                       //Bullet.
    Ganada,                       //Korean Ganada Numbering.
    Chosung,                      //Korean Chosung Numbering.
    DecimalEnclosedFullstop,      //Decimal Numbers Followed by a Period.
    DecimalEnclosedParen,         //Decimal Numbers Enclosed in Parenthesis.
    DecimalEnclosedCircleChinese, //Decimal Numbers Enclosed in a Circle.
    IdeographEnclosedCircle,      //Ideographs Enclosed in a Circle.
    IdeographTraditional,         //Traditional Ideograph Format.
    IdeographZodiac,              //Zodiac Ideograph Format.
    IdeographZodiacTraditional,   //Traditional Zodiac Ideograph Format.
    TaiwaneseCounting,            //Taiwanese Counting System.
    IdeographLegalTraditional,    //Traditional Legal Ideograph Format.
    TaiwaneseCountingThousand,    //Taiwanese Counting Thousand System.
    TaiwaneseDigital,             //Taiwanese Digital Counting System.
    ChineseCounting,              //Chinese Counting System.
    ChineseLegalSimplified,       //Chinese Legal Simplified Format.
    ChineseCountingThousand,      //Chinese Counting Thousand System.
    KoreanDigital,                //Korean Digital Counting System.
    KoreanCounting,               //Korean Counting System.
    KoreanLegal,                  //Korean Legal Numbering.
    KoreanDigital2,               //Korean Digital Counting System Alternate.
    VietnameseCounting,           //Vietnamese Numerals.
    RussianLower,                 //Lowercase Russian Alphabet.
    RussianUpper,                 //Uppercase Russian Alphabet.
    None,                         //No Numbering.
    NumberInDash,                 //Number With Dashes.
    Hebrew1,                      //Hebrew Numerals.
    Hebrew2,                      //Hebrew Alphabet.
    ArabicAlpha,                  //Arabic Alphabet.
    ArabicAbjad,                  //Arabic Abjad Numerals.
    HindiVowels,                  //Hindi Vowels.
    HindiConsonants,              //Hindi Consonants.
    HindiNumbers,                 //Hindi Numbers.
    HindiCounting,                //Hindi Counting System.
    ThaiLetters,                  //Thai Letters.
    ThaiNumbers,                  //Thai Numerals.
    ThaiCounting,                 //Thai Counting System.
}

__string_enum! {
    NumFmtType {
        Decimal = "decimal",
        UpperRoman = "upperRoman",
        LowerRoman = "lowerRoman",
        UpperLetter = "upperLetter",
        LowerLetter = "lowerLetter",
        Ordinal = "ordinal",
        CardinalText = "cardinalText",
        OrdinalText = "ordinalText",
        Hex = "hex",
        Chicago = "chicago",
        IdeographDigital = "ideographDigital",
        JapaneseCounting = "japaneseCounting",
        Aiueo = "aiueo",
        Iroha = "iroha",
        DecimalFullWidth = "decimalFullWidth",
        DecimalHalfWidth = "decimalHalfWidth",
        JapaneseLegal = "japaneseLegal",
        JapaneseDigitalTenThousand = "japaneseDigitalTenThousand",
        DecimalEnclosedCircle = "decimalEnclosedCircle",
        DecimalFullWidth2 = "decimalFullWidth2",
        AiueoFullWidth = "aiueoFullWidth",
        IrohaFullWidth = "irohaFullWidth",
        DecimalZero = "decimalZero",
        Bullet = "bullet",
        Ganada = "ganada",
        Chosung = "chosung",
        DecimalEnclosedFullstop = "decimalEnclosedFullstop",
        DecimalEnclosedParen = "decimalEnclosedParen",
        DecimalEnclosedCircleChinese = "decimalEnclosedCircleChinese",
        IdeographEnclosedCircle = "ideographEnclosedCircle",
        IdeographTraditional = "ideographTraditional",
        IdeographZodiac = "ideographZodiac",
        IdeographZodiacTraditional = "ideographZodiacTraditional",
        TaiwaneseCounting = "taiwaneseCounting",
        IdeographLegalTraditional = "ideographLegalTraditional",
        TaiwaneseCountingThousand = "taiwaneseCountingThousand",
        TaiwaneseDigital = "taiwaneseDigital",
        ChineseCounting = "chineseCounting",
        ChineseLegalSimplified = "chineseLegalSimplified",
        ChineseCountingThousand = "chineseCountingThousand",
        KoreanDigital = "koreanDigital",
        KoreanCounting = "koreanCounting",
        KoreanLegal = "koreanLegal",
        KoreanDigital2 = "koreanDigital2",
        VietnameseCounting = "vietnameseCounting",
        RussianLower = "russianLower",
        RussianUpper = "russianUpper",
        None = "none",
        NumberInDash = "numberInDash",
        Hebrew1 = "hebrew1",
        Hebrew2 = "hebrew2",
        ArabicAlpha = "arabicAlpha",
        ArabicAbjad = "arabicAbjad",
        HindiVowels = "hindiVowels",
        HindiConsonants = "hindiConsonants",
        HindiNumbers = "hindiNumbers",
        HindiCounting = "hindiCounting",
        ThaiLetters = "thaiLetters",
        ThaiNumbers = "thaiNumbers",
        ThaiCounting = "thaiCounting",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pos")]
pub struct FootnotePosition {
    #[xml(attr = "w:val")]
    pub val: PositionType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum PositionType {
    #[default]
    PageBottom, //	Footnotes Positioned at Page Bottom
    BeneathText, //	Footnotes Positioned Beneath Text
    SectEnd,     //	Footnotes Positioned At End of Section
    DocEnd,      //	Footnotes Positioned At End of Document
}

__string_enum! {
    PositionType {
        PageBottom = "pageBottom",
        BeneathText = "beneathText",
        SectEnd = "sectEnd",
        DocEnd = "docEnd",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pos")]
pub struct EndnotePosition {
    #[xml(attr = "w:val")]
    pub val: EndnotePositionType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum EndnotePositionType {
    #[default]
    SectEnd, //	Footnotes Positioned At End of Section
    DocEnd, //	Footnotes Positioned At End of Document
}

__string_enum! {
    EndnotePositionType {
        SectEnd = "sectEnd",
        DocEnd = "docEnd",
    }
}

impl<'a> SectionProperty<'a> {
    //     __setter!(style_id: Option<SectionStyleId<'a>>);
    //     __setter!(justification: Option<Justification>);
    //     __setter!(border: Option<Borders<'a>>);
    //     __setter!(numbering: Option<NumberingProperty>);
    //     __setter!(spacing: Option<Spacing>);
    //     __setter!(indent: Option<Indent>);
    fn first_page_has_diffrent_header_and_footer(&mut self, val: bool) -> &mut Self {
        if val {
            self.title_page = Some(TitlePage::default());
        } else {
            self.title_page = None;
        }
        self
    }
}

// #[derive(Debug, XmlRead, XmlWrite, Clone)]
// #[cfg_attr(test, derive(PartialEq))]
// #[xml(tag = "w:pStyle")]
// pub struct SectionStyleId<'a> {
//     #[xml(attr = "w:val")]
//     pub value: Cow<'a, str>,
// }

// impl<'a, T: Into<Cow<'a, str>>> From<T> for SectionStyleId<'a> {
//     fn from(val: T) -> Self {
//         SectionStyleId { value: val.into() }
//     }
// }

// #[cfg(test)]
// use crate::formatting::JustificationVal;

// __xml_test_suites!(
//     SectionProperty,
//     SectionProperty::default(),
//     r#"<w:pPr/>"#,
//     SectionProperty::default().style_id("id"),
//     r#"<w:pPr><w:pStyle w:val="id"/></w:pPr>"#,
//     SectionProperty::default().justification(JustificationVal::Start),
//     r#"<w:pPr><w:jc w:val="start"/></w:pPr>"#,
//     SectionProperty::default().border(Borders::default()),
//     r#"<w:pPr><w:pBdr/></w:pPr>"#,
//     SectionProperty::default().numbering(NumberingProperty::default()),
//     r#"<w:pPr><w:numPr><w:numId w:val="0"/><w:ilvl w:val="0"/></w:numPr></w:pPr>"#,
// );
