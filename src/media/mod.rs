use crate::schema::SCHEMA_IMAGE;

mod pic;
pub use pic::{Pic, EMU_PER_INCH, EMU_PER_PIXEL};

/// Specifies the type of a media file
///
#[derive(Debug, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum MediaType {
    Image,
}

pub fn get_media_type_relation_type(mt: &MediaType) -> &'static str {
    match mt {
        MediaType::Image => SCHEMA_IMAGE,
    }
}

pub fn get_media_type(filename: &str) -> Option<MediaType> {
    if filename.ends_with("png")
        | filename.ends_with("jpg")
        | filename.ends_with("jpeg")
        | filename.ends_with("bmp")
    {
        Some(MediaType::Image)
    } else {
        None
    }
}
