//! WebSettings part
//!
//! The corresponding ZIP item is `/word/webSettings.xml`.
//!

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::io::Write;

use crate::__xml_test_suites;
use crate::schema::{SCHEMA_MAIN, SCHEMA_WORDML_14};

/// The root element of the main document part.
#[derive(Debug, Default, XmlRead, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:webSettings")]
pub struct WebSettings {
    #[xml(child = "w:optimizeForBrowser")]
    pub optimize_for_browser: Option<OptimizeForBrowser>,
    #[xml(child = "w:relyOnVML")]
    pub rely_on_vml: Option<RelyOnVml>,
    #[xml(child = "w:allowPNG")]
    pub allow_png: Option<AllowPNG>,
    #[xml(child = "w:doNotSaveAsSingleFile")]
    pub do_not_save_as_single_file: Option<DoNotSaveAsSingleFile>,
}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:optimizeForBrowser")]
pub struct OptimizeForBrowser {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:relyOnVML")]
pub struct RelyOnVml {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:allowPNG")]
pub struct AllowPNG {}

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:doNotSaveAsSingleFile")]
pub struct DoNotSaveAsSingleFile {}

impl XmlWrite for WebSettings {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let WebSettings {
            optimize_for_browser,
            rely_on_vml,
            allow_png,
            do_not_save_as_single_file,
        } = self;

        log::debug!("[WebSettings] Started writing.");
        let _ = write!(writer.inner, "{}", crate::schema::SCHEMA_XML);

        writer.write_element_start("w:webSettings")?;

        writer.write_attribute("xmlns:w", SCHEMA_MAIN)?;

        writer.write_attribute("xmlns:w14", SCHEMA_WORDML_14)?;

        writer.write_element_end_open()?;

        write_attr(optimize_for_browser, writer)?;

        write_attr(rely_on_vml, writer)?;

        write_attr(allow_png, writer)?;

        write_attr(do_not_save_as_single_file, writer)?;

        writer.write_element_end_close("w:webSettings")?;

        log::debug!("[webSettings] Finished writing.");

        Ok(())
    }
}

fn write_attr<W: Write, T: XmlWrite>(
    element: &Option<T>,
    writer: &mut XmlWriter<W>,
) -> Result<(), hard_xml::XmlError> {
    if let Some(e) = element {
        e.to_writer(writer)?;
    };
    Ok(())
}

__xml_test_suites!(
    WebSettings,
    WebSettings::default(),
    format!(
        r#"{}<w:webSettings xmlns:w="{}" xmlns:w14="{}"></w:webSettings>"#,
        crate::schema::SCHEMA_XML,
        SCHEMA_MAIN,
        SCHEMA_WORDML_14
    )
    .as_str(),
);
