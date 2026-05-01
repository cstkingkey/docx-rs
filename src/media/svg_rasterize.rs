//! Optional SVG rasterisation, gated on the `svg-rasterize` cargo
//! feature. Provides [`rasterize_svg`] which turns SVG bytes into a
//! PNG of a requested pixel size, suitable for use as the raster
//! fallback in [`Docx::add_svg`](crate::Docx::add_svg).
//!
//! The default build does not include this module's dependencies —
//! callers that already have a rasteriser (browser, ImageMagick,
//! cairo, pre-baked assets) should keep the feature off and supply
//! the PNG bytes themselves.

use resvg::tiny_skia::{Pixmap, Transform};
use usvg::{Options, Tree};

/// Rasterise SVG bytes to a PNG of `(width_px, height_px)` size.
///
/// The SVG's viewBox is scaled (with aspect ratio preserved) to fit
/// the requested pixel canvas; any leftover area is transparent. The
/// returned bytes are a complete `image/png` blob ready to hand to
/// [`Docx::add_svg`](crate::Docx::add_svg) as the fallback.
pub fn rasterize_svg(
    svg_bytes: &[u8],
    width_px: u32,
    height_px: u32,
) -> Result<Vec<u8>, SvgRasterizeError> {
    let opts = Options::default();
    let tree = Tree::from_data(svg_bytes, &opts).map_err(SvgRasterizeError::Parse)?;

    let mut pixmap = Pixmap::new(width_px, height_px).ok_or_else(|| {
        SvgRasterizeError::Pixmap(format!("cannot allocate {}x{} pixmap", width_px, height_px))
    })?;

    let svg_size = tree.size();
    let scale_x = width_px as f32 / svg_size.width();
    let scale_y = height_px as f32 / svg_size.height();
    let scale = scale_x.min(scale_y);
    let transform = Transform::from_scale(scale, scale);

    resvg::render(&tree, transform, &mut pixmap.as_mut());

    pixmap
        .encode_png()
        .map_err(|e| SvgRasterizeError::Encode(e.to_string()))
}

/// Errors from [`rasterize_svg`].
#[derive(Debug)]
pub enum SvgRasterizeError {
    /// SVG bytes failed to parse.
    Parse(usvg::Error),
    /// The pixmap allocator rejected the requested size.
    Pixmap(String),
    /// PNG encoding of the rasterised pixmap failed. The string is
    /// the upstream encoder's error rendered with `Display`; the
    /// concrete type is an implementation detail of `tiny-skia`.
    Encode(String),
}

impl std::fmt::Display for SvgRasterizeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Parse(e) => write!(f, "SVG parse error: {}", e),
            Self::Pixmap(s) => write!(f, "pixmap allocation: {}", s),
            Self::Encode(s) => write!(f, "PNG encode error: {}", s),
        }
    }
}

impl std::error::Error for SvgRasterizeError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Parse(e) => Some(e),
            Self::Pixmap(_) | Self::Encode(_) => None,
        }
    }
}
