#![allow(unused_must_use)]

use std::borrow::Cow;

use derive_more::From;
use hard_xml::{XmlRead, XmlWrite};

use crate::{__define_enum, __string_enum};

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:drawing")]
pub struct Drawing<'a> {
    /// comment
    #[xml(child = "wp:anchor")]
    pub anchor: Option<Anchor<'a>>,
    #[xml(child = "wp:inline")]
    pub inline: Option<Inline<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:anchor")]
pub struct Anchor<'a> {
    #[xml(attr = "distT")]
    pub dist_t: Option<isize>,
    #[xml(attr = "distB")]
    pub dist_b: Option<isize>,
    #[xml(attr = "distL")]
    pub dist_l: Option<isize>,
    #[xml(attr = "distR")]
    pub dist_r: Option<isize>,
    #[xml(attr = "simplePos")]
    pub simple_pos_attr: Option<isize>,
    #[xml(attr = "relativeHeight")]
    pub relative_height: Option<isize>,
    #[xml(attr = "behindDoc")]
    pub behind_doc: Option<bool>,
    #[xml(attr = "locked")]
    pub locked: Option<bool>,
    #[xml(attr = "layoutInCell")]
    pub layout_in_cell: Option<bool>,
    #[xml(attr = "allowOverlap")]
    pub allow_overlap: Option<bool>,

    #[xml(child = "wp:simplePos")]
    pub simple_pos: Option<SimplePos>,
    #[xml(child = "wp:positionH")]
    pub position_horizontal: Option<PositionHorizontal>,
    #[xml(child = "wp:positionV")]
    pub position_vertical: Option<PositionVertical>,
    #[xml(child = "wp:extent")]
    pub extent: Option<Extent>,
    #[xml(
        child = "wp:wrapNone",
        child = "wp:wrapSquare",
        child = "wp:wrapTight",
        child = "wp:wrapThrough",
        child = "wp:wrapTopAndBottom"
    )]
    pub wrap: Option<Wrap>,
    #[xml(child = "wp:docPr")]
    pub doc_property: DocPr<'a>,
    #[xml(child = "a:graphic")]
    pub graphic: Option<Graphic<'a>>,
}

#[derive(Debug, From, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum Wrap {
    #[xml(tag = "wp:wrapNone")]
    None(WrapNone),
    #[xml(tag = "wp:wrapSquare")]
    Square(WrapSquare),
    #[xml(tag = "wp:wrapTight")]
    Tight(WrapTight),
    #[xml(tag = "wp:wrapThrough")]
    Through(WrapThrough),
    #[xml(tag = "wp:wrapTopAndBottom")]
    TopAndBottom(WrapTopAndBottom),
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:wrapNone")]
pub struct WrapNone {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:wrapSquare")]
pub struct WrapSquare {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:wrapTight")]
pub struct WrapTight {
    #[xml(attr = "wrapText")]
    pub wrap_text: WrapTextType,
    #[xml(child = "wp:wrapPolygon")]
    pub wrap_polygon: WrapPolygon,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:wrapPolygon")]
pub struct WrapPolygon {
    #[xml(attr = "edited")]
    pub edited: Option<bool>,
    #[xml(child = "wp:start")]
    pub start: WrapPolygonStart,
    #[xml(child = "wp:lineTo")]
    pub lineto: Vec<WrapPolygonLineTo>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:start")]
pub struct WrapPolygonStart {
    #[xml(attr = "x")]
    pub x: Option<isize>,
    #[xml(attr = "y")]
    pub y: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:lineTo")]
pub struct WrapPolygonLineTo {
    #[xml(attr = "x")]
    pub x: Option<isize>,
    #[xml(attr = "y")]
    pub y: Option<isize>,
}

#[derive(Debug, Clone, Default)]
#[cfg_attr(test, derive(PartialEq))]
pub enum WrapTextType {
    #[default]
    Both,
}

__string_enum! {
    WrapTextType {
    Both = "bothSides",
}}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:wrapThrough")]
pub struct WrapThrough {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:wrapTopAndBottom")]
pub struct WrapTopAndBottom {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:positionH")]
pub struct PositionHorizontal {
    #[xml(attr = "relativeFrom")]
    pub relative_from: Option<RelativeFromH>,
    #[xml(flatten_text = "wp:posOffset")]
    pub pos_offset: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:positionV")]
pub struct PositionVertical {
    #[xml(attr = "relativeFrom")]
    pub relative_from: Option<RelativeFromV>,
    #[xml(flatten_text = "wp:posOffset")]
    pub pos_offset: Option<isize>,
}

__define_enum! {
    RelativeFromH {
        Margin= "margin",	//Page Margin
        Page = "page",//	Page Edge
        Column= "column",//	Column
        Character= "character",//	Character
        LeftMargin= "leftMargin",//	Left Margin
        RightMargin= "rightMargin",//	Right Margin
        InsideMargin= "insideMargin",//	Inside Margin
        OUtsideMargin= "outsideMargin",//	Outside Margin
    }
}

__define_enum! {
    RelativeFromV {
        Margin= "margin",	//Page Margin
        Page = "page",//	Page Edge
        Paragraph= "paragraph",//	Paragraph
        Line= "line",//	Line
        TopMargin= "topMargin",//	Left Margin
        BottomMargin= "bottomMargin",//	Right Margin
        InsideMargin= "insideMargin",//	Inside Margin
        OUtsideMargin= "outsideMargin",//	Outside Margin
    }
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:inline")]
pub struct Inline<'a> {
    #[xml(attr = "distT")]
    pub dist_t: Option<isize>,
    #[xml(attr = "distB")]
    pub dist_b: Option<isize>,
    #[xml(attr = "distL")]
    pub dist_l: Option<isize>,
    #[xml(attr = "distR")]
    pub dist_r: Option<isize>,
    #[xml(attr = "simplePossimplePos")]
    pub simple_pos_attr: Option<isize>,
    #[xml(attr = "relativeHeight")]
    pub relative_height: Option<isize>,
    #[xml(attr = "behindDoc")]
    pub behind_doc: Option<bool>,
    #[xml(attr = "locked")]
    pub locked: Option<bool>,
    #[xml(attr = "layoutInCell")]
    pub layout_in_cell: Option<bool>,
    #[xml(attr = "allowOverlap")]
    pub allow_overlap: Option<bool>,

    #[xml(child = "wp:simplePos")]
    pub simple_pos: Option<SimplePos>,
    #[xml(child = "wp:positionH")]
    pub position_horizontal: Option<PositionHorizontal>,
    #[xml(child = "wp:positionV")]
    pub position_vertical: Option<PositionVertical>,
    #[xml(child = "wp:extent")]
    pub extent: Option<Extent>,
    #[xml(child = "wp:docPr")]
    pub doc_property: DocPr<'a>,
    #[xml(child = "a:graphic")]
    pub graphic: Option<Graphic<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:docPr")]
pub struct DocPr<'a> {
    #[xml(attr = "id")]
    pub id: Option<isize>,
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,
    #[xml(attr = "descr")]
    pub descr: Option<Cow<'a, str>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:graphic")]
pub struct Graphic<'a> {
    #[xml(default, attr = "xmlns:a")]
    pub a: Cow<'a, str>,
    #[xml(default, child = "a:graphicData")]
    pub data: GraphicData<'a>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:graphicData")]
pub struct GraphicData<'a> {
    #[xml(default, attr = "uri")]
    pub uri: Cow<'a, str>,
    #[xml(child = "pic:pic")]
    pub pic: Picture<'a>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "pic:pic")]
pub struct Picture<'a> {
    #[xml(default, attr = "xmlns:pic")]
    pub a: Cow<'a, str>,
    #[xml(child = "pic:nvPicPr")]
    pub nv_pic_pr: NvPicPr<'a>,
    #[xml(child = "pic:blipFill")]
    pub fill: BlipFill<'a>,
    #[xml(child = "pic:spPr")]
    pub sp_pr: SpPr<'a>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "pic:spPr")]
pub struct SpPr<'a> {
    #[xml(child = "a:xfrm")]
    pub xfrm: Option<Xfrm>,
    #[xml(child = "a:prstGeom")]
    pub prst_geom: Option<PrstGeom<'a>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:prstGeom")]
pub struct PrstGeom<'a> {
    #[xml(attr = "prst")]
    pub prst: Option<Cow<'a, str>>,
    #[xml(child = "a:avLst")]
    pub av_lst: Option<AvList>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:avLst")]
pub struct AvList {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:xfrm")]
pub struct Xfrm {
    #[xml(child = "a:off")]
    pub offset: Option<Offset>,
    #[xml(child = "a:ext")]
    pub ext: Option<Ext>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:ext")]
pub struct Ext {
    #[xml(attr = "cx")]
    pub cx: Option<isize>,
    #[xml(attr = "cy")]
    pub cy: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:off")]
pub struct Offset {
    #[xml(attr = "x")]
    pub x: Option<isize>,
    #[xml(attr = "y")]
    pub y: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:simplePos")]
pub struct SimplePos {
    #[xml(attr = "x")]
    pub x: Option<isize>,
    #[xml(attr = "y")]
    pub y: Option<isize>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "pic:nvPicPr")]
pub struct NvPicPr<'a> {
    #[xml(child = "pic:cNvPr")]
    pub c_nv_pr: Option<CNvPr<'a>>,
    #[xml(child = "pic:cNvPicPr")]
    pub c_nv_pic_pr: Option<CNvPicPr>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "pic:cNvPr")]
pub struct CNvPr<'a> {
    #[xml(attr = "id")]
    pub id: Option<isize>,
    #[xml(attr = "name")]
    pub name: Option<Cow<'a, str>>,
    #[xml(attr = "descr")]
    pub descr: Option<Cow<'a, str>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "pic:cNvPicPr")]
pub struct CNvPicPr {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "pic:blipFill")]
pub struct BlipFill<'a> {
    #[xml(default, child = "a:blip")]
    pub blip: Blip<'a>,
    #[xml(child = "a:stretch")]
    pub stretch: Option<Stretch>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:blip")]
pub struct Blip<'a> {
    #[xml(default, attr = "r:embed")]
    pub embed: Cow<'a, str>,
    #[xml(default, attr = "cstate")]
    pub cstate: Option<Cow<'a, str>>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:stretch")]
pub struct Stretch {
    #[xml(child = "a:fillRect")]
    pub fill_rect: Option<FillRect>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "a:fillRect")]
pub struct FillRect {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "wp:extent")]
pub struct Extent {
    #[xml(default, attr = "cx")]
    pub cx: u64,

    #[xml(default, attr = "cy")]
    pub cy: u64,
}
