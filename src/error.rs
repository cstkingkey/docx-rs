use std::io::Error as IOError;

use hard_xml::XmlError;
use thiserror::Error;
use zip::result::ZipError;

/// Error type of docx-rs
#[derive(Debug, Error)]
pub enum DocxError {
    #[error("IO error: {0}")]
    IO(#[from] IOError),
    #[error("malformed XML: {0}")]
    Xml(#[from] XmlError),
    #[error("unable to unpack file: {0}")]
    Zip(#[from] ZipError),
    #[cfg(feature = "async")]
    #[error("unable to unpack file: {0}")]
    AsyncZip(#[from] async_zip::error::ZipError),
}

/// Specialized `Result` which the error value is `DocxError`.
pub type DocxResult<T> = Result<T, DocxError>;
