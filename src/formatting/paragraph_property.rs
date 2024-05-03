use hard_xml::{XmlRead, XmlWrite};
use std::borrow::Cow;

use crate::{
    __define_enum, __define_struct, __define_struct_vec, __setter, __xml_test_suites,
    formatting::{Borders, Indent, Justification, NumberingProperty, Spacing, WidowControl},
};

/// Paragraph Property
///
/// ```rust
/// use docx_rust::formatting::{ParagraphProperty, JustificationVal};
///
/// let prop = ParagraphProperty::default()
///     .style_id("foo")
///     .justification(JustificationVal::Start)
///     .numbering((10isize, 20isize));
/// ```
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pPr")]
pub struct ParagraphProperty<'a> {
    /// Specifies the style ID of the paragraph style.
    #[xml(child = "w:pStyle")]
    pub style_id: Option<ParagraphStyleId<'a>>,
    ///  Keep Paragraph With Next Paragraph
    #[xml(child = "w:keepNext")]
    pub keep_next: Option<KeepNext>,
    ///  Keep All Lines On One Page
    #[xml(child = "w:keepLines")]
    pub keep_lines: Option<KeepLines>,
    ///  Start Paragraph on Next Page
    #[xml(child = "w:pageBreakBefore")]
    pub page_break_before: Option<PageBreakBefore>,
    ///  Text Frame Properties
    //#[xml(child = "w:framePr")]
    //pub frame_pr: Option<FramePr>,
    ///  Allow First/Last Line to Display on a Separate Page
    /// Specifies whether enable widow control
    #[xml(child = "w:widowControl")]
    pub widow_control: Option<WidowControl>,
    /// Specifies that the paragraph should be numbered.
    #[xml(child = "w:numPr")]
    pub numbering: Option<NumberingProperty<'a>>,
    ///  Suppress Line Numbers for Paragraph
    #[xml(child = "w:suppressLineNumbers")]
    pub suppress_line_numbers: Option<SuppressLineNumbers>,
    /// Specifies borders for the paragraph.
    #[xml(child = "w:pBdr")]
    pub border: Option<Borders<'a>>,
    ///  Paragraph Shading
    #[xml(child = "w:shd")]
    pub shading: Option<super::Shading<'a>>,
    ///  Set of Custom Tab Stops
    #[xml(child = "w:tabs")]
    pub tabs: Option<CustomTabStopSet>,
    ///  Suppress Hyphenation for Paragraph
    #[xml(child = "w:suppressAutoHyphens")]
    pub suppress_auto_hyphens: Option<SuppressAutoHyphens>,
    ///  Use East Asian Typography Rules for First and Last Character per Line
    #[xml(child = "w:kinsoku")]
    pub kinsoku: Option<Kinsoku>,
    ///  Allow Line Breaking At Character Level
    #[xml(child = "w:wordWrap")]
    pub word_wrap: Option<WordWrap>,
    ///  Allow Punctuation to Extent Past Text Extents
    #[xml(child = "w:overflowPunct")]
    pub overflow_punct: Option<OverflowPunct>,
    ///  Compress Punctuation at Start of a Line
    #[xml(child = "w:topLinePunct")]
    pub top_line_punct: Option<TopLinePunct>,
    ///  Automatically Adjust Spacing of Latin and East Asian Text
    #[xml(child = "w:autoSpaceDE")]
    pub auto_space_de: Option<AutoSpaceDE>,
    ///  Automatically Adjust Spacing of East Asian Text and Numbers
    #[xml(child = "w:autoSpaceDN")]
    pub auto_space_dn: Option<AutoSpaceDN>,
    ///  Right to Left Paragraph Layout
    #[xml(child = "w:bidi")]
    pub bidi: Option<Bidi>,
    ///  Automatically Adjust Right Indent When Using Document Grid
    #[xml(child = "w:adjustRightInd")]
    pub adjust_right_ind: Option<AdjustRightInd>,
    ///  Use Document Grid Settings for Inter-Line Paragraph Spacing
    #[xml(child = "w:snapToGrid")]
    pub snap_to_grid: Option<SnapToGrid>,
    ///  Spacing Between Lines and Above/Below Paragraph
    #[xml(child = "w:spacing")]
    pub spacing: Option<Spacing>,
    ///  Paragraph Indentation
    #[xml(child = "w:ind")]
    pub indent: Option<Indent>,
    ///  Ignore Spacing Above and Below When Using Identical Styles
    #[xml(child = "w:contextualSpacing")]
    pub contextual_spacing: Option<ContextualSpacing>,
    ///  Use Left/Right Indents as Inside/Outside Indents
    #[xml(child = "w:mirrorIndents")]
    pub mirror_indents: Option<MirrorIndents>,
    ///  Prevent Text Frames From Overlapping
    #[xml(child = "w:suppressOverlap")]
    pub suppress_overlap: Option<SuppressOverlap>,
    ///  Paragraph Alignment
    #[xml(child = "w:jc")]
    pub justification: Option<Justification>,
    ///  Paragraph Text Flow Direction
    #[xml(child = "w:textDirection")]
    pub text_direction: Option<super::TextDirection>,
    ///  Vertical Character Alignment on Line
    #[xml(child = "w:textAlignment")]
    pub text_alignment: Option<super::TextAlignment>,
    ///  Allow Surrounding Paragraphs to Tight Wrap to Text Box Contents
    #[xml(child = "w:textboxTightWrap")]
    pub textbox_tight_wrap: Option<super::TextboxTightWrap>,
    ///  Associated Outline Level
    #[xml(child = "w:outlineLvl")]
    pub outline_lvl: Option<OutlineLvl>,
    ///  Associated HTML div ID
    #[xml(child = "w:divId")]
    pub div_id: Option<DivId>,
    ///  Paragraph Conditional Formatting
    #[xml(child = "w:cnfStyle")]
    pub cnf_style: Option<CnfStyle<'a>>,
    #[xml(child = "w:rPr")]
    pub r_pr: Vec<super::CharacterProperty<'a>>,
    #[xml(child = "w:sectPr")]
    pub section_property: Option<SectionProperty<'a>>,
    #[xml(child = "w:pPrChange")]
    pub p_pr_change: Option<RevisionParagraphProperty<'a>>,
}

impl<'a> ParagraphProperty<'a> {
    __setter!(style_id: Option<ParagraphStyleId<'a>>);
    __setter!(justification: Option<Justification>);
    __setter!(border: Option<Borders<'a>>);
    __setter!(numbering: Option<NumberingProperty<'a>>);
    __setter!(spacing: Option<Spacing>);
    __setter!(indent: Option<Indent>);
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pStyle")]
pub struct ParagraphStyleId<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

impl<'a, T: Into<Cow<'a, str>>> From<T> for ParagraphStyleId<'a> {
    fn from(val: T) -> Self {
        ParagraphStyleId { value: val.into() }
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:keepNext")]
pub struct KeepNext {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:keepLines")]
pub struct KeepLines {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pageBreakBefore")]
pub struct PageBreakBefore {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:suppressLineNumbers")]
pub struct SuppressLineNumbers {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:suppressAutoHyphens")]
pub struct SuppressAutoHyphens {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:kinsoku")]
pub struct Kinsoku {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:wordWrap")]
pub struct WordWrap {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:overflowPunct")]
pub struct OverflowPunct {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:topLinePunct")]
pub struct TopLinePunct {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:autoSpaceDE")]
pub struct AutoSpaceDE {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:autoSpaceDN")]
pub struct AutoSpaceDN {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bidi")]
pub struct Bidi {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:adjustRightInd")]
pub struct AdjustRightInd {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:snapToGrid")]
pub struct SnapToGrid {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:contextualSpacing")]
pub struct ContextualSpacing {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:mirrorIndents")]
pub struct MirrorIndents {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:suppressOverlap")]
pub struct SuppressOverlap {
    #[xml(attr = "w:val")]
    pub value: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:outlineLvl")]
pub struct OutlineLvl {
    #[xml(attr = "w:val")]
    pub value: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:divId")]
pub struct DivId {
    #[xml(attr = "w:val")]
    pub value: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:cnfStyle")]
pub struct CnfStyle<'a> {
    #[xml(attr = "w:val")]
    pub value: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pPrChange")]
pub struct RevisionParagraphProperty<'a> {
    #[xml(attr = "w:id")]
    pub id: isize,
    #[xml(attr = "w:author")]
    pub author: Cow<'a, str>,
    #[xml(attr = "w:date")]
    pub date: Option<Cow<'a, str>>,

    #[xml(child = "w:pPr")]
    pub previous_property: Option<PreviousParagraphProperty<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:pPr")]
pub struct PreviousParagraphProperty<'a> {
    /// Specifies the style ID of the paragraph style.
    #[xml(child = "w:pStyle")]
    pub style_id: Option<ParagraphStyleId<'a>>,
    ///  Keep Paragraph With Next Paragraph
    #[xml(child = "w:keepNext")]
    pub keep_next: Option<KeepNext>,
    ///  Keep All Lines On One Page
    #[xml(child = "w:keepLines")]
    pub keep_lines: Option<KeepLines>,
    ///  Start Paragraph on Next Page
    #[xml(child = "w:pageBreakBefore")]
    pub page_break_before: Option<PageBreakBefore>,
    ///  Text Frame Properties
    //#[xml(child = "w:framePr")]
    //pub frame_pr: Option<FramePr>,
    ///  Allow First/Last Line to Display on a Separate Page
    /// Specifies whether enable widow control
    #[xml(child = "w:widowControl")]
    pub widow_control: Option<WidowControl>,
    /// Specifies that the paragraph should be numbered.
    #[xml(child = "w:numPr")]
    pub numbering: Option<NumberingProperty<'a>>,
    ///  Suppress Line Numbers for Paragraph
    #[xml(child = "w:suppressLineNumbers")]
    pub suppress_line_numbers: Option<SuppressLineNumbers>,
    /// Specifies borders for the paragraph.
    #[xml(child = "w:pBdr")]
    pub border: Option<Borders<'a>>,
    ///  Paragraph Shading
    #[xml(child = "w:shd")]
    pub shading: Option<super::Shading<'a>>,
    ///  Set of Custom Tab Stops
    #[xml(child = "w:tabs")]
    pub tabs: Option<CustomTabStopSet>,
    ///  Suppress Hyphenation for Paragraph
    #[xml(child = "w:suppressAutoHyphens")]
    pub suppress_auto_hyphens: Option<SuppressAutoHyphens>,
    ///  Use East Asian Typography Rules for First and Last Character per Line
    #[xml(child = "w:kinsoku")]
    pub kinsoku: Option<Kinsoku>,
    ///  Allow Line Breaking At Character Level
    #[xml(child = "w:wordWrap")]
    pub word_wrap: Option<WordWrap>,
    ///  Allow Punctuation to Extent Past Text Extents
    #[xml(child = "w:overflowPunct")]
    pub overflow_punct: Option<OverflowPunct>,
    ///  Compress Punctuation at Start of a Line
    #[xml(child = "w:topLinePunct")]
    pub top_line_punct: Option<TopLinePunct>,
    ///  Automatically Adjust Spacing of Latin and East Asian Text
    #[xml(child = "w:autoSpaceDE")]
    pub auto_space_de: Option<AutoSpaceDE>,
    ///  Automatically Adjust Spacing of East Asian Text and Numbers
    #[xml(child = "w:autoSpaceDN")]
    pub auto_space_dn: Option<AutoSpaceDN>,
    ///  Right to Left Paragraph Layout
    #[xml(child = "w:bidi")]
    pub bidi: Option<Bidi>,
    ///  Automatically Adjust Right Indent When Using Document Grid
    #[xml(child = "w:adjustRightInd")]
    pub adjust_right_ind: Option<AdjustRightInd>,
    ///  Use Document Grid Settings for Inter-Line Paragraph Spacing
    #[xml(child = "w:snapToGrid")]
    pub snap_to_grid: Option<SnapToGrid>,
    ///  Spacing Between Lines and Above/Below Paragraph
    #[xml(child = "w:spacing")]
    pub spacing: Option<Spacing>,
    ///  Paragraph Indentation
    #[xml(child = "w:ind")]
    pub indent: Option<Indent>,
    ///  Ignore Spacing Above and Below When Using Identical Styles
    #[xml(child = "w:contextualSpacing")]
    pub contextual_spacing: Option<ContextualSpacing>,
    ///  Use Left/Right Indents as Inside/Outside Indents
    #[xml(child = "w:mirrorIndents")]
    pub mirror_indents: Option<MirrorIndents>,
    ///  Prevent Text Frames From Overlapping
    #[xml(child = "w:suppressOverlap")]
    pub suppress_overlap: Option<SuppressOverlap>,
    ///  Paragraph Alignment
    #[xml(child = "w:jc")]
    pub justification: Option<Justification>,
    ///  Paragraph Text Flow Direction
    #[xml(child = "w:textDirection")]
    pub text_direction: Option<super::TextDirection>,
    ///  Vertical Character Alignment on Line
    #[xml(child = "w:textAlignment")]
    pub text_alignment: Option<super::TextAlignment>,
    ///  Allow Surrounding Paragraphs to Tight Wrap to Text Box Contents
    #[xml(child = "w:textboxTightWrap")]
    pub textbox_tight_wrap: Option<super::TextboxTightWrap>,
    ///  Associated Outline Level
    #[xml(child = "w:outlineLvl")]
    pub outline_lvl: Option<OutlineLvl>,
    ///  Associated HTML div ID
    #[xml(child = "w:divId")]
    pub div_id: Option<DivId>,
    ///  Paragraph Conditional Formatting
    #[xml(child = "w:cnfStyle")]
    pub cnf_style: Option<CnfStyle<'a>>,
}

__define_enum! {
    TabStopType  {
        Clear = "clear", // No Tab Stop
        Left = "left", // Left Tab
        Center = "center", // Centered Tab
        Right = "right", // Right Tab
        Decimal = "decimal", // Decimal Tab
        Bar = "bar", // Bar Tab
        Num = "num", // List Tab
    }
}

__define_enum! {
    TabLeaderCharacter {
        None = "none", // No tab stop leader
        Dot = "dot", // Dotted leader line
        Hyphen = "hyphen", // Dashed tab stop leader line
        Underscore = "underscore", // Solid leader line
        Heavy = "heavy", // Heavy solid leader line
        MiddleDot = "middleDot", // Middle dot leader line
    }
}

__define_struct! {
    ("w:tab", CustomTabStop) {
        "w:val", tab_stop_type, TabStopType
        "w:leader", leader, TabLeaderCharacter
        "w:pos", pos, isize
    }
}

__define_struct_vec! {
    ("w:tabs", CustomTabStopSet, CustomTabStopSetChoice) {} {
        "w:tab", CustomTabStop
    }
}

#[cfg(test)]
use crate::formatting::JustificationVal;

use super::SectionProperty;

__xml_test_suites!(
    ParagraphProperty,
    ParagraphProperty::default(),
    r#"<w:pPr/>"#,
    ParagraphProperty::default().style_id("id"),
    r#"<w:pPr><w:pStyle w:val="id"/></w:pPr>"#,
    ParagraphProperty::default().justification(JustificationVal::Start),
    r#"<w:pPr><w:jc w:val="start"/></w:pPr>"#,
    ParagraphProperty::default().border(Borders::default()),
    r#"<w:pPr><w:pBdr/></w:pPr>"#,
    ParagraphProperty::default().numbering(NumberingProperty::default()),
    r#"<w:pPr><w:numPr/></w:pPr>"#,
);
