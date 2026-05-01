//! Inline picture helper.
//!
//! Wraps the tedium of constructing the
//! `Drawing → Inline → Graphic → Picture → BlipFill → Blip` XML chain
//! for the common case of placing an image in a run.
//!
//! The helper does NOT register image bytes with the document; callers
//! do that via [`Docx::add_image`](crate::Docx::add_image) to obtain
//! the relationship id used here.
//!
//! ```no_run
//! use docx_rust::Docx;
//! use docx_rust::document::{Paragraph, Run};
//! use docx_rust::media::{MediaType, Pic};
//!
//! # let png_bytes: Vec<u8> = vec![];
//! let mut docx = Docx::default();
//! let rid = docx.add_image("cat.png", MediaType::Image, &png_bytes);
//! let drawing = Pic::new(rid).size_px(160, 120).into_drawing();
//! let para = Paragraph::default()
//!     .push(Run::default().push_image(drawing));
//! docx.document.push(para);
//! ```

use std::borrow::Cow;

use crate::document::{
    Blip, BlipFill, CNvPicPr, CNvPr, DocPr, Drawing, Ext, Extent, FillRect, Graphic, GraphicData,
    Inline, NvPicPr, Offset, Picture, PrstGeom, SpPr, Stretch, Xfrm,
};
use crate::schema::{SCHEMA_DRAWINGML, SCHEMA_PIC};

/// English Metric Units per pixel at 96 DPI.
pub const EMU_PER_PIXEL: u64 = 9_525;

/// English Metric Units per inch.
pub const EMU_PER_INCH: u64 = 914_400;

/// Builder for an inline picture.
///
/// Sizes default to 16x16 pixels at 96 DPI; override with
/// [`Pic::size_px`] or [`Pic::size_emu`].
pub struct Pic<'a> {
    rid: Cow<'a, str>,
    name: Cow<'a, str>,
    descr: Option<Cow<'a, str>>,
    width_emu: u64,
    height_emu: u64,
    doc_id: isize,
}

impl<'a> Pic<'a> {
    /// Create a builder from a relationship id previously returned by
    /// [`Docx::add_image`](crate::Docx::add_image).
    pub fn new(rid: impl Into<Cow<'a, str>>) -> Self {
        Self {
            rid: rid.into(),
            name: Cow::Borrowed("image"),
            descr: None,
            width_emu: 16 * EMU_PER_PIXEL,
            height_emu: 16 * EMU_PER_PIXEL,
            doc_id: 1,
        }
    }

    /// Set the picture display name (appears in `wp:docPr` / `pic:cNvPr`).
    pub fn name(mut self, name: impl Into<Cow<'a, str>>) -> Self {
        self.name = name.into();
        self
    }

    /// Set the alt-text description.
    pub fn descr(mut self, descr: impl Into<Cow<'a, str>>) -> Self {
        self.descr = Some(descr.into());
        self
    }

    /// Set size in pixels at 96 DPI. Internally converts to EMU.
    pub fn size_px(mut self, w: u64, h: u64) -> Self {
        self.width_emu = w * EMU_PER_PIXEL;
        self.height_emu = h * EMU_PER_PIXEL;
        self
    }

    /// Set size directly in English Metric Units.
    pub fn size_emu(mut self, w: u64, h: u64) -> Self {
        self.width_emu = w;
        self.height_emu = h;
        self
    }

    /// Set the document-wide unique picture id. Defaults to 1; bump
    /// when embedding multiple images in the same document.
    pub fn doc_id(mut self, id: isize) -> Self {
        self.doc_id = id;
        self
    }

    /// Build the [`Drawing`] ready to push into a [`Run`](crate::document::Run).
    pub fn into_drawing(self) -> Drawing<'a> {
        let Pic {
            rid,
            name,
            descr,
            width_emu,
            height_emu,
            doc_id,
        } = self;

        let cx = width_emu as isize;
        let cy = height_emu as isize;

        let picture = Picture {
            a: Cow::Borrowed(SCHEMA_PIC),
            nv_pic_pr: NvPicPr {
                c_nv_pr: Some(CNvPr {
                    id: Some(doc_id),
                    name: Some(name.clone()),
                    descr: descr.clone(),
                }),
                c_nv_pic_pr: Some(CNvPicPr::default()),
            },
            fill: BlipFill {
                blip: Blip {
                    embed: rid,
                    cstate: None,
                },
                stretch: Some(Stretch {
                    fill_rect: Some(FillRect::default()),
                }),
            },
            sp_pr: SpPr {
                xfrm: Some(Xfrm {
                    offset: Some(Offset {
                        x: Some(0),
                        y: Some(0),
                    }),
                    ext: Some(Ext {
                        cx: Some(cx),
                        cy: Some(cy),
                    }),
                }),
                prst_geom: Some(PrstGeom {
                    prst: Some(Cow::Borrowed("rect")),
                    av_lst: None,
                }),
            },
        };

        let graphic = Graphic {
            a: Cow::Borrowed(SCHEMA_DRAWINGML),
            data: GraphicData {
                uri: Cow::Borrowed(SCHEMA_PIC),
                children: vec![picture],
            },
        };

        let inline = Inline {
            extent: Some(Extent {
                cx: width_emu,
                cy: height_emu,
            }),
            doc_property: DocPr {
                id: Some(doc_id),
                name: Some(name),
                descr,
            },
            graphic: Some(graphic),
            ..Inline::default()
        };

        Drawing {
            anchor: None,
            inline: Some(inline),
        }
    }
}
