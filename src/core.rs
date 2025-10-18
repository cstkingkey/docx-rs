//! Core File Properties part
//!
//! The corresponding ZIP item is `/docProps/core.xml`.

use hard_xml::{XmlRead, XmlResult, XmlWrite, XmlWriter};
use std::borrow::Cow;
use std::io::Write;

use crate::schema::{SCHEMA_CORE_2, SCHEMA_DC, SCHEMA_XML, SCHEMA_DCTERMS};

#[derive(Debug, XmlRead, XmlWrite, Clone)]
pub enum Core<'a> {
    #[xml(tag = "cp:coreProperties")]
    CoreNamespace(CoreNamespace<'a>),
    #[xml(tag = "coreProperties")]
    CoreNoNamespace(CoreNoNamespace<'a>),
}

#[derive(Debug, Default, XmlRead, Clone)]
#[xml(tag = "cp:coreProperties")]
pub struct CoreNamespace<'a> {
    #[xml(flatten_text = "dc:title")]
    pub title: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:subject")]
    pub subject: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:creator")]
    pub creator: Option<Cow<'a, str>>,
    #[xml(flatten_text = "cp:keywords")]
    pub keywords: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:description")]
    pub description: Option<Cow<'a, str>>,
    #[xml(flatten_text = "cp:lastModifiedBy")]
    pub last_modified_by: Option<Cow<'a, str>>,
    #[xml(flatten_text = "cp:revision")]
    pub revision: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dcterms:created")]
    pub created: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dcterms:modified")]
    pub modified: Option<Cow<'a, str>>,
    #[xml(flatten_text = "cp:contentStatus")]
    pub content_status: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:language")]
    pub language: Option<Cow<'a, str>>,
    #[xml(flatten_text = "cp:category")]
    pub category: Option<Cow<'a, str>>,
}

#[derive(Debug, Default, XmlRead, Clone)]
#[xml(tag = "coreProperties")]
pub struct CoreNoNamespace<'a> {
    #[xml(flatten_text = "dc:title")]
    pub title: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:subject")]
    pub subject: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:creator")]
    pub creator: Option<Cow<'a, str>>,
    #[xml(flatten_text = "keywords")]
    pub keywords: Option<Cow<'a, str>>,
    #[xml(flatten_text = "dc:description")]
    pub description: Option<Cow<'a, str>>,
    #[xml(flatten_text = "lastModifiedBy")]
    pub last_modified_by: Option<Cow<'a, str>>,
    #[xml(flatten_text = "revision")]
    pub revision: Option<Cow<'a, str>>,
    #[xml(flatten_text = "created")]
    pub created: Option<Cow<'a, str>>,
    #[xml(flatten_text = "modified")]
    pub modified: Option<Cow<'a, str>>,
    #[xml(flatten_text = "contentStatus")]
    pub content_status: Option<Cow<'a, str>>,
    #[xml(flatten_text = "language")]
    pub language: Option<Cow<'a, str>>,
    #[xml(flatten_text = "category")]
    pub category: Option<Cow<'a, str>>,
}

impl<'a> XmlWrite for CoreNamespace<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let CoreNamespace {
            title,
            subject,
            creator,
            keywords,
            description,
            last_modified_by,
            revision,
            created,
            modified,
            content_status,
            language,
            category,
        } = self;

        log::debug!("[Core] Started writing.");
        let _ = write!(writer.inner, "{}", SCHEMA_XML);

        writer.write_element_start("cp:coreProperties")?;

        writer.write_attribute("xmlns:cp", SCHEMA_CORE_2)?;

        writer.write_attribute("xmlns:dc", SCHEMA_DC)?;

        writer.write_attribute(
            "xmlns:dcterms",
            SCHEMA_DCTERMS,
        )?;

        if title.is_none()
            && subject.is_none()
            && creator.is_none()
            && keywords.is_none()
            && description.is_none()
            && last_modified_by.is_none()
            && revision.is_none()
            && created.is_none()
            && modified.is_none()
            && content_status.is_none()
            && language.is_none()
            && category.is_none()
        {
            writer.write_element_end_empty()?;
        } else {
            writer.write_element_end_open()?;
            if let Some(val) = title {
                writer.write_flatten_text("dc:title", val, false)?;
            }
            if let Some(val) = subject {
                writer.write_flatten_text("dc:subject", val, false)?;
            }
            if let Some(val) = creator {
                writer.write_flatten_text("dc:creator", val, false)?;
            }
            if let Some(val) = keywords {
                writer.write_flatten_text("cp:keywords", val, false)?;
            }
            if let Some(val) = description {
                writer.write_flatten_text("dc:description", val, false)?;
            }
            if let Some(val) = last_modified_by {
                writer.write_flatten_text("cp:lastModifiedBy", val, false)?;
            }
            if let Some(val) = revision {
                writer.write_flatten_text("cp:revision", val, false)?;
            }
            if let Some(val) = revision {
                writer.write_flatten_text("cp:revision", val, false)?;
            }
            if let Some(val) = created {
                writer.write_flatten_text("dcterms:created", val, false)?;
            }
            if let Some(val) = modified {
                writer.write_flatten_text("dcterms:modified", val, false)?;
            }
            if let Some(val) = content_status {
                writer.write_flatten_text("cp:contentStatus", val, false)?;
            }
            if let Some(val) = language {
                writer.write_flatten_text("dc:language", val, false)?;
            }
            if let Some(val) = category {
                writer.write_flatten_text("cp:category", val, false)?;
            }
            writer.write_element_end_close("cp:coreProperties")?;
        }

        log::debug!("[Core] Finished writing.");

        Ok(())
    }
}

impl<'a> XmlWrite for CoreNoNamespace<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let CoreNoNamespace {
            title,
            subject,
            creator,
            keywords,
            description,
            last_modified_by,
            revision,
            created,
            modified,
            content_status,
            language,
            category,
        } = self;

        log::debug!("[Core] Started writing.");
        let _ = write!(writer.inner, "{}", SCHEMA_XML);

        writer.write_element_start("cp:coreProperties")?;

        writer.write_attribute("xmlns:cp", SCHEMA_CORE_2)?;

        writer.write_attribute("xmlns:dc", SCHEMA_DC)?;

        if title.is_none()
            && subject.is_none()
            && creator.is_none()
            && keywords.is_none()
            && description.is_none()
            && last_modified_by.is_none()
            && revision.is_none()
        {
            writer.write_element_end_empty()?;
        } else {
            writer.write_element_end_open()?;
            if let Some(val) = title {
                writer.write_flatten_text("dc:title", val, false)?;
            }
            if let Some(val) = subject {
                writer.write_flatten_text("dc:subject", val, false)?;
            }
            if let Some(val) = creator {
                writer.write_flatten_text("dc:creator", val, false)?;
            }
            if let Some(val) = keywords {
                writer.write_flatten_text("cp:keywords", val, false)?;
            }
            if let Some(val) = description {
                writer.write_flatten_text("dc:description", val, false)?;
            }
            if let Some(val) = last_modified_by {
                writer.write_flatten_text("cp:lastModifiedBy", val, false)?;
            }
            if let Some(val) = revision {
                writer.write_flatten_text("cp:revision", val, false)?;
            }
            if let Some(val) = created {
                writer.write_flatten_text("dcterms:created", val, false)?;
            }
            if let Some(val) = modified {
                writer.write_flatten_text("dcterms:modified", val, false)?;
            }
            if let Some(val) = content_status {
                writer.write_flatten_text("cp:contentStatus", val, false)?;
            }
            if let Some(val) = language {
                writer.write_flatten_text("dc:language", val, false)?;
            }
            if let Some(val) = category {
                writer.write_flatten_text("cp:category", val, false)?;
            }
            writer.write_element_end_close("cp:coreProperties")?;
        }

        log::debug!("[Core] Finished writing.");

        Ok(())
    }
}
