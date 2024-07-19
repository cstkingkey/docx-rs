//! Settings part
//!
//! The corresponding ZIP item is `/word/settings.xml`.
//!

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::borrow::Cow;
use std::io::Write;

use crate::schema::{SCHEMA_MAIN, SCHEMA_WORDML_14};
use crate::{__string_enum, __xml_test_suites, write_attr};

/// The root element of the main document part.
#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:settings")]
pub struct Settings<'a> {
    ///  Write Protection
    #[xml(child = "w:writeProtection")]
    pub write_protection: Option<WriteProtection>,
    ///  Document View Setting
    #[xml(child = "w:view")]
    pub view: Option<View>,
    ///  Magnification Setting
    #[xml(child = "w:zoom")]
    pub zoom: Option<Zoom>,
    ///  Remove Personal Information from Document Properties
    #[xml(child = "w:removePersonalInformation")]
    pub remove_personal_information: Option<RemovePersonalInformation>,
    ///  Remove Date and Time from Annotations
    #[xml(child = "w:removeDateAndTime")]
    pub remove_date_and_time: Option<RemoveDateAndTime>,
    ///  Do Not Display Visual Boundary For Header/Footer or Between Pages
    #[xml(child = "w:doNotDisplayPageBoundaries")]
    pub do_not_display_page_boundaries: Option<DoNotDisplayPageBoundaries>,
    ///  Display Background Objects When Displaying Document
    #[xml(child = "w:displayBackgroundShape")]
    pub display_background_shape: Option<DisplayBackgroundShape>,
    ///  Print PostScript Codes With Document Text
    #[xml(child = "w:printPostScriptOverText")]
    pub print_post_script_over_text: Option<PrintPostScriptOverText>,
    ///  Print Fractional Character Widths
    #[xml(child = "w:printFractionalCharacterWidth")]
    pub print_fractional_character_width: Option<PrintFractionalCharacterWidth>,
    ///  Only Print Form Field Content
    #[xml(child = "w:printFormsData")]
    pub print_forms_data: Option<PrintFormsData>,
    ///  Embed TrueType Fonts
    #[xml(child = "w:embedTrueTypeFonts")]
    pub embed_true_type_fonts: Option<EmbedTrueTypeFonts>,
    ///  Embed Common System Fonts
    #[xml(child = "w:embedSystemFonts")]
    pub embed_system_fonts: Option<EmbedSystemFonts>,
    ///  Subset Fonts When Embedding
    #[xml(child = "w:saveSubsetFonts")]
    pub save_subset_fonts: Option<SaveSubsetFonts>,
    ///  Only Save Form Field Content
    #[xml(child = "w:saveFormsData")]
    pub save_forms_data: Option<SaveFormsData>,
    ///  Mirror Page Margins
    #[xml(child = "w:mirrorMargins")]
    pub mirror_margins: Option<MirrorMargins>,
    ///  Align Paragraph and Table Borders with Page Border
    #[xml(child = "w:alignBordersAndEdges")]
    pub align_borders_and_edges: Option<AlignBordersAndEdges>,
    ///  Page Border Excludes Header
    #[xml(child = "w:bordersDoNotSurroundHeader")]
    pub borders_do_not_surround_header: Option<BordersDoNotSurroundHeader>,
    ///  Page Border Excludes Footer
    #[xml(child = "w:bordersDoNotSurroundFooter")]
    pub borders_do_not_surround_footer: Option<BordersDoNotSurroundFooter>,
    ///  Position Gutter At Top of Page
    #[xml(child = "w:gutterAtTop")]
    pub gutter_at_top: Option<GutterAtTop>,
    ///  Do Not Display Visual Indication of Spelling Errors
    #[xml(child = "w:hideSpellingErrors")]
    pub hide_spelling_errors: Option<HideSpellingErrors>,
    ///  Do Not Display Visual Indication of Grammatical Errors
    #[xml(child = "w:hideGrammaticalErrors")]
    pub hide_grammatical_errors: Option<HideGrammaticalErrors>,
    ///  Grammar Checking Settings
    #[xml(child = "w:activeWritingStyle")]
    pub active_writing_style: Option<ActiveWritingStyle>,
    ///  Spelling and Grammatical Checking State
    #[xml(child = "w:proofState")]
    pub proof_state: Option<ProofState>,
    ///  Structured Document Tag Placeholder Text Should be Resaved
    #[xml(child = "w:formsDesign")]
    pub forms_design: Option<FormsDesign>,
    ///  Attached Document Template
    #[xml(child = "w:attachedTemplate")]
    pub attached_template: Option<AttachedTemplate<'a>>,
    ///  Automatically Update Styles From Document Template
    #[xml(child = "w:linkStyles")]
    pub link_styles: Option<LinkStyles>,
    ///  Suggested Filtering for List of Document Styles
    #[xml(child = "w:stylePaneFormatFilter")]
    pub style_pane_format_filter: Option<StylePaneFormatFilter>,
    ///  Suggested Sorting for List of Document Styles
    #[xml(child = "w:stylePaneSortMethod")]
    pub style_pane_sort_method: Option<StylePaneSortMethod>,
    ///  Document Classification
    #[xml(child = "w:documentType")]
    pub document_type: Option<DocumentType>,
    ///  Mail Merge Settings
    #[xml(child = "w:mailMerge")]
    pub mail_merge: Option<MailMerge>,
    ///  Visibility of Annotation Types
    #[xml(child = "w:revisionView")]
    pub revision_view: Option<RevisionView>,
    ///  Track Revisions to Document
    #[xml(child = "w:trackRevisions")]
    pub track_revisions: Option<TrackRevisions>,
    ///  Do Not Use Move Syntax When Tracking Revisions
    #[xml(child = "w:doNotTrackMoves")]
    pub do_not_track_moves: Option<DoNotTrackMoves>,
    ///  Do Not Track Formatting Revisions When Tracking Revisions
    #[xml(child = "w:doNotTrackFormatting")]
    pub do_not_track_formatting: Option<DoNotTrackFormatting>,
    ///  Document Editing Restrictions
    #[xml(child = "w:documentProtection")]
    pub document_protection: Option<DocumentProtection>,
    ///  Allow Automatic Formatting to Override Formatting Protection Settings
    #[xml(child = "w:autoFormatOverride")]
    pub auto_format_override: Option<AutoFormatOverride>,
    ///  Prevent Modification of Themes Part
    #[xml(child = "w:styleLockTheme")]
    pub style_lock_theme: Option<StyleLockTheme>,
    ///  Prevent Replacement of Styles Part
    #[xml(child = "w:styleLockQFSet")]
    pub style_lock_qfset: Option<StyleLockQfset>,
    ///  Distance Between Automatic Tab Stops
    #[xml(child = "w:defaultTabStop")]
    pub default_tab_stop: Option<DefaultTabStop>,
    ///  Automatically Hyphenate Document Contents When Displayed
    #[xml(child = "w:autoHyphenation")]
    pub auto_hyphenation: Option<AutoHyphenation>,
    ///  Maximum Number of Consecutively Hyphenated Lines
    #[xml(child = "w:consecutiveHyphenLimit")]
    pub consecutive_hyphen_limit: Option<ConsecutiveHyphenLimit>,
    ///  Hyphenation Zone
    #[xml(child = "w:hyphenationZone")]
    pub hyphenation_zone: Option<HyphenationZone>,
    ///  Do Not Hyphenate Words in ALL CAPITAL LETTERS
    #[xml(child = "w:doNotHyphenateCaps")]
    pub do_not_hyphenate_caps: Option<DoNotHyphenateCaps>,
    ///  Show E-Mail Message Header
    #[xml(child = "w:showEnvelope")]
    pub show_envelope: Option<ShowEnvelope>,
    ///  Percentage of Document to Use When Generating Summary
    #[xml(child = "w:summaryLength")]
    pub summary_length: Option<SummaryLength>,
    ///  Paragraph Style Applied to Automatically Generated Paragraphs
    #[xml(child = "w:clickAndTypeStyle")]
    pub click_and_type_style: Option<ClickAndTypeStyle>,
    ///  Default Table Style for Newly Inserted Tables
    #[xml(child = "w:defaultTableStyle")]
    pub default_table_style: Option<DefaultTableStyle>,
    ///  Different Even/Odd Page Headers and Footers
    #[xml(child = "w:evenAndOddHeaders")]
    pub even_and_odd_headers: Option<EvenAndOddHeaders>,
    ///  Reverse Book Fold Printing
    #[xml(child = "w:bookFoldRevPrinting")]
    pub book_fold_rev_printing: Option<BookFoldRevPrinting>,
    ///  Book Fold Printing
    #[xml(child = "w:bookFoldPrinting")]
    pub book_fold_printing: Option<BookFoldPrinting>,
    ///  Number of Pages Per Booklet
    #[xml(child = "w:bookFoldPrintingSheets")]
    pub book_fold_printing_sheets: Option<BookFoldPrintingSheets>,
    ///  Drawing Grid Horizontal Grid Unit Size
    #[xml(child = "w:drawingGridHorizontalSpacing")]
    pub drawing_grid_horizontal_spacing: Option<DrawingGridHorizontalSpacing>,
    ///  Drawing Grid Vertical Grid Unit Size
    #[xml(child = "w:drawingGridVerticalSpacing")]
    pub drawing_grid_vertical_spacing: Option<DrawingGridVerticalSpacing>,
    ///  Distance between Horizontal Gridlines
    #[xml(child = "w:displayHorizontalDrawingGridEvery")]
    pub display_horizontal_drawing_grid_every: Option<DisplayHorizontalDrawingGridEvery>,
    ///  Distance between Vertical Gridlines
    #[xml(child = "w:displayVerticalDrawingGridEvery")]
    pub display_vertical_drawing_grid_every: Option<DisplayVerticalDrawingGridEvery>,
    ///  Do Not Use Margins for Drawing Grid Origin
    #[xml(child = "w:doNotUseMarginsForDrawingGridOrigin")]
    pub do_not_use_margins_for_drawing_grid_origin: Option<DoNotUseMarginsForDrawingGridOrigin>,
    ///  Drawing Grid Horizontal Origin Point
    #[xml(child = "w:drawingGridHorizontalOrigin")]
    pub drawing_grid_horizontal_origin: Option<DrawingGridHorizontalOrigin>,
    ///  Drawing Grid Vertical Origin Point
    #[xml(child = "w:drawingGridVerticalOrigin")]
    pub drawing_grid_vertical_origin: Option<DrawingGridVerticalOrigin>,
    ///  Do Not Show Visual Indicator For Form Fields
    #[xml(child = "w:doNotShadeFormData")]
    pub do_not_shade_form_data: Option<DoNotShadeFormData>,
    ///  Never Kern Punctuation Characters
    #[xml(child = "w:noPunctuationKerning")]
    pub no_punctuation_kerning: Option<NoPunctuationKerning>,
    ///  Character-Level Whitespace Compression
    #[xml(child = "w:characterSpacingControl")]
    pub character_spacing_control: Option<CharacterSpacingControl>,
    ///  Print Two Pages Per Sheet
    #[xml(child = "w:printTwoOnOne")]
    pub print_two_on_one: Option<PrintTwoOnOne>,
    ///  Use Strict Kinsoku Rules for Japanese Text
    #[xml(child = "w:strictFirstAndLastChars")]
    pub strict_first_and_last_chars: Option<StrictFirstAndLastChars>,
    ///  Custom Set of Characters Which Cannot End a Line
    #[xml(child = "w:noLineBreaksAfter")]
    pub no_line_breaks_after: Option<NoLineBreaksAfter>,
    ///  Custom Set Of Characters Which Cannot Begin A Line
    #[xml(child = "w:noLineBreaksBefore")]
    pub no_line_breaks_before: Option<NoLineBreaksBefore>,
    ///  Generate Thumbnail For Document On Save
    #[xml(child = "w:savePreviewPicture")]
    pub save_preview_picture: Option<SavePreviewPicture>,
    ///  Do Not Validate Custom XML Markup Against Schemas
    #[xml(child = "w:doNotValidateAgainstSchema")]
    pub do_not_validate_against_schema: Option<DoNotValidateAgainstSchema>,
    ///  Allow Saving Document As XML File When Custom XML Markup Is Invalid
    #[xml(child = "w:saveInvalidXml")]
    pub save_invalid_xml: Option<SaveInvalidXml>,
    ///  Ignore Mixed Content When Validating Custom XML Markup
    #[xml(child = "w:ignoreMixedContent")]
    pub ignore_mixed_content: Option<IgnoreMixedContent>,
    ///  Use Custom XML Element Names as Default Placeholder Text
    #[xml(child = "w:alwaysShowPlaceholderText")]
    pub always_show_placeholder_text: Option<AlwaysShowPlaceholderText>,
    ///  Do Not Show Visual Indicator For Invalid Custom XML Markup
    #[xml(child = "w:doNotDemarcateInvalidXml")]
    pub do_not_demarcate_invalid_xml: Option<DoNotDemarcateInvalidXml>,
    ///  Only Save Custom XML Markup
    #[xml(child = "w:saveXmlDataOnly")]
    pub save_xml_data_only: Option<SaveXmlDataOnly>,
    ///  Save Document as XML File through Custom XSL Transform
    #[xml(child = "w:useXSLTWhenSaving")]
    pub use_xsltwhen_saving: Option<UseXsltwhenSaving>,
    ///  Custom XSL Transform To Use When Saving As XML File
    #[xml(child = "w:saveThroughXslt")]
    pub save_through_xslt: Option<SaveThroughXslt>,
    ///  Show Visual Indicators for Custom XML Markup Start/End Locations
    #[xml(child = "w:showXMLTags")]
    pub show_xmltags: Option<ShowXmltags>,
    ///  Do Not Mark Custom XML Elements With No Namespace As Invalid
    #[xml(child = "w:alwaysMergeEmptyNamespace")]
    pub always_merge_empty_namespace: Option<AlwaysMergeEmptyNamespace>,
    ///  Automatically Recalculate Fields on Open
    #[xml(child = "w:updateFields")]
    pub update_fields: Option<UpdateFields>,
    ///  Default Properties for VML Objects in Header and Footer
    #[xml(child = "w:hdrShapeDefaults")]
    pub hdr_shape_defaults: Option<HdrShapeDefaults>,
    ///  Document-Wide Footnote Properties
    #[xml(child = "w:footnotePr")]
    pub footnote_pr: Option<crate::formatting::FootnoteProperty2>,
    ///  Document-Wide Endnote Properties
    #[xml(child = "w:endnotePr")]
    pub endnote_pr: Option<crate::formatting::EndnoteProperty2>,
    ///  Compatibility Settings
    #[xml(child = "w:compat")]
    pub compat: Option<Compat>,
    ///  Document Variables
    #[xml(child = "w:docVars")]
    pub doc_vars: Option<DocVars<'a>>,
    ///  Listing of All Revision Save ID Values
    #[xml(child = "w:rsids")]
    pub rsids: Option<Rsids<'a>>,
    ///// properties of math in the document
    //#[xml(child = "m:mathPr")]
    //pub mathPr: Option<mathPr>,
    ///  Disable Features Incompatible With Earlier Word Processing Formats
    #[xml(child = "w:uiCompat97To2003")]
    pub ui_compat97_to2003: Option<UiCompat97to2003>,
    // /// Attached Custom XML Schema
    // #[xml(child = "w:attachedSchema")]
    // pub attachedSchema: Vec<attachedSchema>,
    ///  Theme Font Languages
    #[xml(child = "w:themeFontLang")]
    pub theme_font_lang: Option<ThemeFontLang<'a>>,
    ///  Theme Color Mappings
    #[xml(child = "w:clrSchemeMapping")]
    pub clr_scheme_mapping: Option<ClrSchemeMapping>,
    ///  Do Not Include Content in Text Boxes, Footnotes, and Endnotes in Document Statistics
    #[xml(child = "w:doNotIncludeSubdocsInStats")]
    pub do_not_include_subdocs_in_stats: Option<DoNotIncludeSubdocsInStats>,
    ///  Do Not Automatically Compress Images
    #[xml(child = "w:doNotAutoCompressPictures")]
    pub do_not_auto_compress_pictures: Option<DoNotAutoCompressPictures>,
    ///  Upgrade Document on Open
    #[xml(child = "w:forceUpgrade")]
    pub force_upgrade: Option<ForceUpgrade>,
    ///  Caption Settings
    #[xml(child = "w:captions")]
    pub captions: Option<Captions>,
    ///  Freeze Document Layout
    #[xml(child = "w:readModeInkLockDown")]
    pub read_mode_ink_lock_down: Option<ReadModeInkLockDown>,
    // /// Supplementary Smart Tag Information
    // #[xml(child = "w:smartTagType")]
    // pub smartTagType: Vec<smartTagType>,
    //     /// Custom XML Schema List
    // #[xml(child = "sl:schemaLibrary")]
    // pub schemaLibrary: Option<schemaLibrary>,
    ///  Default Properties for VML Objects in Main Document
    #[xml(child = "w:shapeDefaults")]
    pub shape_defaults: Option<ShapeDefaults>,
    ///  Remove Smart Tags When Saving
    #[xml(child = "w:doNotEmbedSmartTags")]
    pub do_not_embed_smart_tags: Option<DoNotEmbedSmartTags>,
    ///  Radix Point for Field Code Evaluation
    #[xml(child = "w:decimalSymbol")]
    pub decimal_symbol: Option<DecimalSymbol<'a>>,
    ///  List Separator for Field Code Evaluation
    #[xml(child = "w:listSeparator")]
    pub list_separator: Option<ListSeparator<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:writeProtection")]
pub struct WriteProtection {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:smartTagType")]
pub struct SmartTagType {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:view")]
pub struct View {
    #[xml(attr = "w:val")]
    pub val: ViewType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ViewType {
    #[default]
    None, //	Default View
    Print,       //	Print Layout View
    Outline,     //	Outline View
    MasterPages, //	Master Document View
    Normal,      //	Draft View
    Web,         //	Web Page View
}

__string_enum! {
    ViewType {
        None = "none",
        Print = "print",
        Outline = "outline",
        MasterPages = "masterPages",
        Normal = "normal",
        Web = "web",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:zoom")]
pub struct Zoom {
    #[xml(attr = "w:val")]
    pub val: Option<ZoomType>,
    #[xml(attr = "w:percent")]
    pub percent: Option<u8>,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ZoomType {
    #[default]
    None, //	No Preset Magnification
    FullPage, //	Display One Full Page
    BestFit,  //	Display Page Width
    TextFit,  //	Display Text Width
}

__string_enum! {
    ZoomType {
        None = "none",
        FullPage = "fullPage",
        BestFit = "bestFit",
        TextFit = "textFit",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:removePersonalInformation")]
pub struct RemovePersonalInformation {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:removeDateAndTime")]
pub struct RemoveDateAndTime {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotDisplayPageBoundaries")]
pub struct DoNotDisplayPageBoundaries {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:displayBackgroundShape")]
pub struct DisplayBackgroundShape {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:printPostScriptOverText")]
pub struct PrintPostScriptOverText {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:printFractionalCharacterWidth")]
pub struct PrintFractionalCharacterWidth {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:printFormsData")]
pub struct PrintFormsData {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:embedTrueTypeFonts")]
pub struct EmbedTrueTypeFonts {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:embedSystemFonts")]
pub struct EmbedSystemFonts {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:saveSubsetFonts")]
pub struct SaveSubsetFonts {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:saveFormsData")]
pub struct SaveFormsData {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:mirrorMargins")]
pub struct MirrorMargins {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:alignBordersAndEdges")]
pub struct AlignBordersAndEdges {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bordersDoNotSurroundHeader")]
pub struct BordersDoNotSurroundHeader {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bordersDoNotSurroundFooter")]
pub struct BordersDoNotSurroundFooter {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:gutterAtTop")]
pub struct GutterAtTop {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:hideSpellingErrors")]
pub struct HideSpellingErrors {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:hideGrammaticalErrors")]
pub struct HideGrammaticalErrors {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:activeWritingStyle")]
pub struct ActiveWritingStyle {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:proofState")]
pub struct ProofState {
    #[xml(attr = "w:spelling")]
    pub spelling: Option<ProofStateType>,
    #[xml(attr = "w:grammar")]
    pub grammar: Option<ProofStateType>,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum ProofStateType {
    #[default]
    Clean, //	Check Completed
    Dirty, //	Check Not Completed
}

__string_enum! {
    ProofStateType {
        Clean ="clean",
        Dirty = "dirty",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:formsDesign")]
pub struct FormsDesign {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:attachedTemplate")]
pub struct AttachedTemplate<'a> {
    #[xml(attr = "r:id")]
    pub val: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:linkStyles")]
pub struct LinkStyles {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:stylePaneFormatFilter")]
pub struct StylePaneFormatFilter {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:stylePaneSortMethod")]
pub struct StylePaneSortMethod {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:documentType")]
pub struct DocumentType {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:mailMerge")]
pub struct MailMerge {
    // ///required    Source Document Type
    // #[xml(child = "w:mainDocumentType")]
    // pub main_document_type: Option<mainDocumentType>,
    // ///required    Query Contains Link to External Query File
    // #[xml(child = "w:linkToQuery")]
    // pub link_to_query: Option<linkToQuery>,
    // ///required    Data Source Type
    // #[xml(child = "w:dataType")]
    // pub data_type: Option<dataType>,
    // ///required    Data Source Connection String
    // #[xml(child = "w:connectString")]
    // pub connect_string: Option<connectString>,
    // ///required    Query For Data Source Records To Merge
    // #[xml(child = "w:query")]
    // pub query: Option<query>,
    // ///required    Data Source File Path
    // #[xml(child = "w:dataSource")]
    // pub data_source: Option<dataSource>,
    // ///required    Header Definition File Path
    // #[xml(child = "w:headerSource")]
    // pub header_source: Option<headerSource>,
    // ///required    Remove Blank Lines from Merged Documents
    // #[xml(child = "w:doNotSuppressBlankLines")]
    // pub do_not_suppress_blank_lines: Option<doNotSuppressBlankLines>,
    // ///required    Merged Document Destination
    // #[xml(child = "w:destination")]
    // pub destination: Option<destination>,
    // ///required    Column Containing E-mail Address
    // #[xml(child = "w:addressFieldName")]
    // pub address_field_name: Option<addressFieldName>,
    // ///required    Merged E-mail or Fax Subject Line
    // #[xml(child = "w:mailSubject")]
    // pub mail_subject: Option<mailSubject>,
    // ///required    Merged Document To E-Mail Attachment
    // #[xml(child = "w:mailAsAttachment")]
    // pub mail_as_attachment: Option<mailAsAttachment>,
    // ///required    View Merged Data Within Document
    // #[xml(child = "w:viewMergedData")]
    // pub view_merged_data: Option<viewMergedData>,
    // ///required    Record Currently Displayed In Merged Document
    // #[xml(child = "w:activeRecord")]
    // pub active_record: Option<activeRecord>,
    // ///required    Mail Merge Error Reporting Setting
    // #[xml(child = "w:checkErrors")]
    // pub check_errors: Option<checkErrors>,
    // ///required    Office Data Source Object Settings
    // #[xml(child = "w:odso")]
    // pub odso: Option<odso>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:revisionView")]
pub struct RevisionView {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:trackRevisions")]
pub struct TrackRevisions {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotTrackMoves")]
pub struct DoNotTrackMoves {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotTrackFormatting")]
pub struct DoNotTrackFormatting {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:documentProtection")]
pub struct DocumentProtection {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:autoFormatOverride")]
pub struct AutoFormatOverride {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:styleLockTheme")]
pub struct StyleLockTheme {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:styleLockQFSet")]
pub struct StyleLockQfset {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:defaultTabStop")]
pub struct DefaultTabStop {
    #[xml(attr = "w:val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:autoHyphenation")]
pub struct AutoHyphenation {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:consecutiveHyphenLimit")]
pub struct ConsecutiveHyphenLimit {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:hyphenationZone")]
pub struct HyphenationZone {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotHyphenateCaps")]
pub struct DoNotHyphenateCaps {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:showEnvelope")]
pub struct ShowEnvelope {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:summaryLength")]
pub struct SummaryLength {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:clickAndTypeStyle")]
pub struct ClickAndTypeStyle {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:defaultTableStyle")]
pub struct DefaultTableStyle {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bookFoldRevPrinting")]
pub struct BookFoldRevPrinting {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bookFoldPrinting")]
pub struct BookFoldPrinting {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bookFoldPrintingSheets")]
pub struct BookFoldPrintingSheets {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:drawingGridHorizontalSpacing")]
pub struct DrawingGridHorizontalSpacing {
    #[xml(attr = "w:val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:drawingGridVerticalSpacing")]
pub struct DrawingGridVerticalSpacing {
    #[xml(attr = "w:val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:displayHorizontalDrawingGridEvery")]
pub struct DisplayHorizontalDrawingGridEvery {
    #[xml(attr = "w:val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:displayVerticalDrawingGridEvery")]
pub struct DisplayVerticalDrawingGridEvery {
    #[xml(attr = "w:val")]
    pub val: isize,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotUseMarginsForDrawingGridOrigin")]
pub struct DoNotUseMarginsForDrawingGridOrigin {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:drawingGridHorizontalOrigin")]
pub struct DrawingGridHorizontalOrigin {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:drawingGridVerticalOrigin")]
pub struct DrawingGridVerticalOrigin {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotShadeFormData")]
pub struct DoNotShadeFormData {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:noPunctuationKerning")]
pub struct NoPunctuationKerning {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:characterSpacingControl")]
pub struct CharacterSpacingControl {
    #[xml(attr = "w:val")]
    pub val: CharacterSpacingControlType,
}

#[derive(Debug, Default, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum CharacterSpacingControlType {
    #[default]
    DoNotCompress, //	Do Not Compress Whitespace
    CompressPunctuation, //	Compress Whitespace From Punctuation Characters
    CompressPunctuationAndJapaneseKana, //	Compress Whitespace From Both Japanese Kana And Punctuation Characters
}

__string_enum! {
    CharacterSpacingControlType {
        DoNotCompress = "doNotCompress",
        CompressPunctuation = "compressPunctuation",
        CompressPunctuationAndJapaneseKana = "compressPunctuationAndJapaneseKana",
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:printTwoOnOne")]
pub struct PrintTwoOnOne {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:strictFirstAndLastChars")]
pub struct StrictFirstAndLastChars {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:noLineBreaksAfter")]
pub struct NoLineBreaksAfter {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:noLineBreaksBefore")]
pub struct NoLineBreaksBefore {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:savePreviewPicture")]
pub struct SavePreviewPicture {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotValidateAgainstSchema")]
pub struct DoNotValidateAgainstSchema {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:saveInvalidXml")]
pub struct SaveInvalidXml {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:ignoreMixedContent")]
pub struct IgnoreMixedContent {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:alwaysShowPlaceholderText")]
pub struct AlwaysShowPlaceholderText {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotDemarcateInvalidXml")]
pub struct DoNotDemarcateInvalidXml {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:saveXmlDataOnly")]
pub struct SaveXmlDataOnly {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:useXSLTWhenSaving")]
pub struct UseXsltwhenSaving {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:saveThroughXslt")]
pub struct SaveThroughXslt {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:showXMLTags")]
pub struct ShowXmltags {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:alwaysMergeEmptyNamespace")]
pub struct AlwaysMergeEmptyNamespace {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:updateFields")]
pub struct UpdateFields {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:hdrShapeDefaults")]
pub struct HdrShapeDefaults {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:compat")]
pub struct Compat {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:docVars")]
pub struct DocVars<'a> {
    #[xml(child = "w:docVar")]
    pub vars: Vec<DocVar<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:docVar")]
pub struct DocVar<'a> {
    #[xml(attr = "w:name")]
    pub name: Cow<'a, str>,
    #[xml(attr = "w:val")]
    pub val: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rsids")]
pub struct Rsids<'a> {
    #[xml(child = "w:rsidRoot")]
    pub ro: Option<RsidRoot<'a>>,
    #[xml(child = "w:rsid")]
    pub rsids: Vec<Rsid<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rsid")]
pub struct Rsid<'a> {
    #[xml(attr = "w:val")]
    pub val: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:rsidRoot")]
pub struct RsidRoot<'a> {
    #[xml(attr = "w:val")]
    pub val: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:uiCompat97To2003")]
pub struct UiCompat97to2003 {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:clrSchemeMapping")]
pub struct ClrSchemeMapping {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotIncludeSubdocsInStats")]
pub struct DoNotIncludeSubdocsInStats {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotAutoCompressPictures")]
pub struct DoNotAutoCompressPictures {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:forceUpgrade")]
pub struct ForceUpgrade {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:captions")]
pub struct Captions {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:readModeInkLockDown")]
pub struct ReadModeInkLockDown {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:shapeDefaults")]
pub struct ShapeDefaults {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotEmbedSmartTags")]
pub struct DoNotEmbedSmartTags {
    #[xml(attr = "w:value")]
    pub val: Option<bool>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:decimalSymbol")]
pub struct DecimalSymbol<'a> {
    #[xml(attr = "w:val")]
    pub val: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:listSeparator")]
pub struct ListSeparator<'a> {
    #[xml(attr = "w:val")]
    pub val: Cow<'a, str>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:evenAndOddHeaders")]
pub struct EvenAndOddHeaders {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:themeFontLang")]
pub struct ThemeFontLang<'a> {
    #[xml(attr = "w:val")]
    pub val: Option<Cow<'a, str>>,
    #[xml(attr = "w:eastAsia")]
    pub east_asia: Option<Cow<'a, str>>,
}

impl<'a> XmlWrite for Settings<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let Settings {
            write_protection,
            view,
            zoom,
            remove_personal_information,
            remove_date_and_time,
            do_not_display_page_boundaries,
            display_background_shape,
            print_post_script_over_text,
            print_fractional_character_width,
            print_forms_data,
            embed_true_type_fonts,
            embed_system_fonts,
            save_subset_fonts,
            save_forms_data,
            mirror_margins,
            align_borders_and_edges,
            borders_do_not_surround_header,
            borders_do_not_surround_footer,
            gutter_at_top,
            hide_spelling_errors,
            hide_grammatical_errors,
            active_writing_style,
            proof_state,
            forms_design,
            attached_template,
            link_styles,
            style_pane_format_filter,
            style_pane_sort_method,
            document_type,
            mail_merge,
            revision_view,
            track_revisions,
            do_not_track_moves,
            do_not_track_formatting,
            document_protection,
            auto_format_override,
            style_lock_theme,
            style_lock_qfset,
            default_tab_stop,
            auto_hyphenation,
            consecutive_hyphen_limit,
            hyphenation_zone,
            do_not_hyphenate_caps,
            show_envelope,
            summary_length,
            click_and_type_style,
            default_table_style,
            even_and_odd_headers,
            book_fold_rev_printing,
            book_fold_printing,
            book_fold_printing_sheets,
            drawing_grid_horizontal_spacing,
            drawing_grid_vertical_spacing,
            display_horizontal_drawing_grid_every,
            display_vertical_drawing_grid_every,
            do_not_use_margins_for_drawing_grid_origin,
            drawing_grid_horizontal_origin,
            drawing_grid_vertical_origin,
            do_not_shade_form_data,
            no_punctuation_kerning,
            character_spacing_control,
            print_two_on_one,
            strict_first_and_last_chars,
            no_line_breaks_after,
            no_line_breaks_before,
            save_preview_picture,
            do_not_validate_against_schema,
            save_invalid_xml,
            ignore_mixed_content,
            always_show_placeholder_text,
            do_not_demarcate_invalid_xml,
            save_xml_data_only,
            use_xsltwhen_saving,
            save_through_xslt,
            show_xmltags,
            always_merge_empty_namespace,
            update_fields,
            hdr_shape_defaults,
            footnote_pr,
            endnote_pr,
            compat,
            doc_vars,
            rsids,
            ui_compat97_to2003,
            theme_font_lang,
            clr_scheme_mapping,
            do_not_include_subdocs_in_stats,
            do_not_auto_compress_pictures,
            force_upgrade,
            captions,
            read_mode_ink_lock_down,
            shape_defaults,
            do_not_embed_smart_tags,
            decimal_symbol,
            list_separator,
        } = self;

        log::debug!("[Settings] Started writing.");
        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:settings")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        write_attr(write_protection, writer)?;
        write_attr(view, writer)?;
        write_attr(zoom, writer)?;
        write_attr(remove_personal_information, writer)?;
        write_attr(remove_date_and_time, writer)?;
        write_attr(do_not_display_page_boundaries, writer)?;
        write_attr(display_background_shape, writer)?;
        write_attr(print_post_script_over_text, writer)?;
        write_attr(print_fractional_character_width, writer)?;
        write_attr(print_forms_data, writer)?;
        write_attr(embed_true_type_fonts, writer)?;
        write_attr(embed_system_fonts, writer)?;
        write_attr(save_subset_fonts, writer)?;
        write_attr(save_forms_data, writer)?;
        write_attr(mirror_margins, writer)?;
        write_attr(align_borders_and_edges, writer)?;
        write_attr(borders_do_not_surround_header, writer)?;
        write_attr(borders_do_not_surround_footer, writer)?;
        write_attr(gutter_at_top, writer)?;
        write_attr(hide_spelling_errors, writer)?;
        write_attr(hide_grammatical_errors, writer)?;
        write_attr(active_writing_style, writer)?;
        write_attr(proof_state, writer)?;
        write_attr(forms_design, writer)?;
        write_attr(attached_template, writer)?;
        write_attr(link_styles, writer)?;
        write_attr(style_pane_format_filter, writer)?;
        write_attr(style_pane_sort_method, writer)?;
        write_attr(document_type, writer)?;
        write_attr(mail_merge, writer)?;
        write_attr(revision_view, writer)?;
        write_attr(track_revisions, writer)?;
        write_attr(do_not_track_moves, writer)?;
        write_attr(do_not_track_formatting, writer)?;
        write_attr(document_protection, writer)?;
        write_attr(auto_format_override, writer)?;
        write_attr(style_lock_theme, writer)?;
        write_attr(style_lock_qfset, writer)?;
        write_attr(default_tab_stop, writer)?;
        write_attr(auto_hyphenation, writer)?;
        write_attr(consecutive_hyphen_limit, writer)?;
        write_attr(hyphenation_zone, writer)?;
        write_attr(do_not_hyphenate_caps, writer)?;
        write_attr(show_envelope, writer)?;
        write_attr(summary_length, writer)?;
        write_attr(click_and_type_style, writer)?;
        write_attr(default_table_style, writer)?;
        write_attr(even_and_odd_headers, writer)?;
        write_attr(book_fold_rev_printing, writer)?;
        write_attr(book_fold_printing, writer)?;
        write_attr(book_fold_printing_sheets, writer)?;
        write_attr(drawing_grid_horizontal_spacing, writer)?;
        write_attr(drawing_grid_vertical_spacing, writer)?;
        write_attr(display_horizontal_drawing_grid_every, writer)?;
        write_attr(display_vertical_drawing_grid_every, writer)?;
        write_attr(do_not_use_margins_for_drawing_grid_origin, writer)?;
        write_attr(drawing_grid_horizontal_origin, writer)?;
        write_attr(drawing_grid_vertical_origin, writer)?;
        write_attr(do_not_shade_form_data, writer)?;
        write_attr(no_punctuation_kerning, writer)?;
        write_attr(character_spacing_control, writer)?;
        write_attr(print_two_on_one, writer)?;
        write_attr(strict_first_and_last_chars, writer)?;
        write_attr(no_line_breaks_after, writer)?;
        write_attr(no_line_breaks_before, writer)?;
        write_attr(save_preview_picture, writer)?;
        write_attr(do_not_validate_against_schema, writer)?;
        write_attr(save_invalid_xml, writer)?;
        write_attr(ignore_mixed_content, writer)?;
        write_attr(always_show_placeholder_text, writer)?;
        write_attr(do_not_demarcate_invalid_xml, writer)?;
        write_attr(save_xml_data_only, writer)?;
        write_attr(use_xsltwhen_saving, writer)?;
        write_attr(save_through_xslt, writer)?;
        write_attr(show_xmltags, writer)?;
        write_attr(always_merge_empty_namespace, writer)?;
        write_attr(update_fields, writer)?;
        write_attr(hdr_shape_defaults, writer)?;
        write_attr(footnote_pr, writer)?;
        write_attr(endnote_pr, writer)?;
        write_attr(compat, writer)?;
        write_attr(doc_vars, writer)?;
        write_attr(rsids, writer)?;
        write_attr(ui_compat97_to2003, writer)?;
        write_attr(theme_font_lang, writer)?;
        write_attr(clr_scheme_mapping, writer)?;
        write_attr(do_not_include_subdocs_in_stats, writer)?;
        write_attr(do_not_auto_compress_pictures, writer)?;
        write_attr(force_upgrade, writer)?;
        write_attr(captions, writer)?;
        write_attr(read_mode_ink_lock_down, writer)?;
        write_attr(shape_defaults, writer)?;
        write_attr(do_not_embed_smart_tags, writer)?;
        write_attr(decimal_symbol, writer)?;
        write_attr(list_separator, writer)?;

        writer.write_element_end_close("w:settings")?;

        log::debug!("[Settings] Finished writing.");

        Ok(())
    }
}

__xml_test_suites!(
    Settings,
    Settings::default(),
    format!(
        r#"{}<w:settings xmlns:w="{}" xmlns:w14="{}"></w:settings>"#,
        crate::schema::SCHEMA_XML,
        SCHEMA_MAIN,
        SCHEMA_WORDML_14
    )
    .as_str(),
);
