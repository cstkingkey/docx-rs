//! Inline picture helper.
//!
//! Wraps the tedium of constructing the
//! `Drawing → Inline → Graphic → Picture → BlipFill → Blip` XML chain
//! for the common case of placing an image in a run.
//!
//! The helper does NOT register image bytes with the document; callers
//! do that via [`Docx::add_image`](crate::Docx::add_image) (raster) or
//! [`Docx::add_svg`](crate::Docx::add_svg) (vector + raster fallback)
//! to obtain the relationship id used here.
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
    Blip, BlipExt, BlipExtList, BlipFill, CNvPicPr, CNvPr, DocPr, Drawing, Ext, Extent, FillRect,
    Graphic, GraphicData, Inline, NvPicPr, Offset, Picture, PrstGeom, SpPr, Stretch, SvgBlip, Xfrm,
};
use crate::schema::{SCHEMA_ASVG, SCHEMA_DRAWINGML, SCHEMA_PIC, SVG_BLIP_EXT_URI};

/// English Metric Units per pixel at 96 DPI.
pub const EMU_PER_PIXEL: u64 = 9_525;

/// English Metric Units per inch.
pub const EMU_PER_INCH: u64 = 914_400;

/// Identifiers for the two media parts behind a single SVG-with-fallback
/// image. Returned by [`Docx::add_svg`](crate::Docx::add_svg) and
/// consumed by [`Pic::with_svg`].
#[derive(Debug, Clone)]
pub struct SvgImageIds {
    /// Relationship id of the SVG part. Modern Word renders this
    /// vectorially via the `asvg:svgBlip` extension element.
    pub svg_rid: String,
    /// Relationship id of the PNG (or other raster) fallback part.
    /// Legacy Word reads this; modern Word reads it too if the
    /// extension element is missing.
    pub png_rid: String,
}

/// Builder for an inline picture.
///
/// Sizes default to 16x16 pixels at 96 DPI; override with
/// [`Pic::size_px`] or [`Pic::size_emu`].
pub struct Pic<'a> {
    rid: Cow<'a, str>,
    svg_rid: Option<Cow<'a, str>>,
    name: Cow<'a, str>,
    descr: Option<Cow<'a, str>>,
    width_emu: u64,
    height_emu: u64,
    doc_id: u32,
}

impl<'a> Pic<'a> {
    /// Create a builder from a relationship id previously returned by
    /// [`Docx::add_image`](crate::Docx::add_image).
    pub fn new(rid: impl Into<Cow<'a, str>>) -> Self {
        Self {
            rid: rid.into(),
            svg_rid: None,
            name: Cow::Borrowed("image"),
            descr: None,
            width_emu: 16 * EMU_PER_PIXEL,
            height_emu: 16 * EMU_PER_PIXEL,
            doc_id: 1,
        }
    }

    /// Create a builder for an SVG image with a raster fallback. The
    /// resulting drawing carries both the standard `<a:blip>` (PNG,
    /// for legacy Word and as a fallback) and the
    /// `<asvg:svgBlip>` extension (SVG, for Office 2016+).
    ///
    /// `ids` is the value returned by
    /// [`Docx::add_svg`](crate::Docx::add_svg).
    pub fn with_svg(ids: SvgImageIds) -> Self {
        Self {
            rid: Cow::Owned(ids.png_rid),
            svg_rid: Some(Cow::Owned(ids.svg_rid)),
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
    ///
    /// Saturates at `u64::MAX` if the multiplication overflows, so a
    /// pathological caller cannot panic the builder. The resulting
    /// EMU value will exceed Word's max image dimension and the
    /// document will simply render that picture clamped to the page,
    /// which is the right failure mode for unrealistic sizes.
    pub fn size_px(mut self, w: u64, h: u64) -> Self {
        self.width_emu = w.saturating_mul(EMU_PER_PIXEL);
        self.height_emu = h.saturating_mul(EMU_PER_PIXEL);
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
    ///
    /// `wp:docPr/@id` and `pic:cNvPr/@id` are both unsigned in the
    /// OOXML schema; `u32` matches that contract.
    pub fn doc_id(mut self, id: u32) -> Self {
        self.doc_id = id;
        self
    }

    /// Build the [`Drawing`] ready to push into a [`Run`](crate::document::Run).
    pub fn into_drawing(self) -> Drawing<'a> {
        let Pic {
            rid,
            svg_rid,
            name,
            descr,
            width_emu,
            height_emu,
            doc_id,
        } = self;

        // The XML uses `Option<isize>` for these fields. Saturate
        // rather than wrap if a caller passed an absurdly large EMU
        // value: an overflow-wrap would emit a negative cx/cy and
        // Word would treat the picture as zero-sized or invalid.
        let cx = isize::try_from(width_emu).unwrap_or(isize::MAX);
        let cy = isize::try_from(height_emu).unwrap_or(isize::MAX);
        // doc_id is u32 from the public API; widening to isize for
        // the XML attribute is always lossless on 64-bit, and on
        // 32-bit isize::MAX > u32::MAX/2 so values up to 2^31-1 fit.
        let id = isize::try_from(doc_id).unwrap_or(isize::MAX);

        let ext_lst = svg_rid.map(|svg| BlipExtList {
            exts: vec![BlipExt {
                uri: Cow::Borrowed(SVG_BLIP_EXT_URI),
                svg_blip: Some(SvgBlip {
                    xmlns_asvg: Cow::Borrowed(SCHEMA_ASVG),
                    embed: svg,
                }),
            }],
        });

        let picture = Picture {
            a: Cow::Borrowed(SCHEMA_PIC),
            nv_pic_pr: NvPicPr {
                c_nv_pr: Some(CNvPr {
                    id: Some(id),
                    name: Some(name.clone()),
                    descr: descr.clone(),
                }),
                c_nv_pic_pr: Some(CNvPicPr::default()),
            },
            fill: BlipFill {
                blip: Blip {
                    embed: rid,
                    cstate: None,
                    ext_lst,
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
                id: Some(id),
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
