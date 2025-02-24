use hard_xml::{XmlRead, XmlWrite, XmlWriter};
use std::borrow::Cow;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;
use zip::{result::ZipError, CompressionMethod, ZipArchive, ZipWriter};

use crate::document::{Comments, EndNotes, FootNotes, Footer, Header, Numbering, Theme};
use crate::media::MediaType;
use crate::schema::{
    SCHEMA_COMMENTS, SCHEMA_ENDNOTES, SCHEMA_FOOTNOTES, SCHEMA_HEADER, SCHEMA_NUMBERING,
    SCHEMA_SETTINGS, SCHEMA_THEME, SCHEMA_WEB_SETTINGS,
};
use crate::settings::Settings;
use crate::web_settings::WebSettings;
use crate::{
    app::App,
    content_type::ContentTypes,
    core::Core,
    document::Document,
    error::DocxResult,
    font_table::FontTable,
    rels::Relationships,
    schema::{
        SCHEMA_CORE, SCHEMA_FONT_TABLE, SCHEMA_OFFICE_DOCUMENT, SCHEMA_REL_EXTENDED, SCHEMA_STYLES,
    },
    styles::Styles,
};

/// A WordprocessingML package
#[derive(Debug, Default, Clone)]
pub struct Docx<'a> {
    /// Specifies package-level properties part
    pub app: Option<App<'a>>,
    /// Specifies core properties part
    pub core: Option<Core<'a>>,
    /// Specifies the content type of relationship parts and the main document part.
    pub content_types: ContentTypes<'a>,
    /// Specifies the main document part.
    pub document: Document<'a>,
    /// Specifies the font table part
    pub font_table: Option<FontTable<'a>>,
    /// Specifies the style definitions part
    pub styles: Styles<'a>,
    /// Specifies the package-level relationship to the main document part
    pub rels: Relationships<'a>,
    /// Specifies the part-level relationship to the main document part
    pub document_rels: Option<Relationships<'a>>,
    pub settings_rels: Option<Relationships<'a>>,
    pub headers: HashMap<String, Header<'a>>,
    pub footers: HashMap<String, Footer<'a>>,
    pub themes: HashMap<String, Theme<'a>>,
    pub media: HashMap<String, (MediaType, &'a Vec<u8>)>,
    pub footnotes: Option<FootNotes<'a>>,
    pub endnotes: Option<EndNotes<'a>>,
    pub settings: Option<Settings<'a>>,
    pub web_settings: Option<WebSettings>,
    pub comments: Option<Comments<'a>>,
    pub numbering: Option<Numbering<'a>>,
    pub custom_xml: HashMap<String, Cow<'a, [u8]>>,
}

impl<'a> Docx<'a> {
    pub fn write<W: Write + Seek>(&'a mut self, writer: W) -> DocxResult<W> {
        let mut writer = XmlWriter::new(ZipWriter::new(writer));

        let opt = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .unix_permissions(0o755);

        // ==== Add Relationships ====

        if self.app.is_some() {
            self.rels.add_rel(SCHEMA_REL_EXTENDED, "docProps/app.xml");
        }

        if self.core.is_some() {
            self.rels.add_rel(SCHEMA_CORE, "docProps/core.xml");
        }

        self.rels
            .add_rel(SCHEMA_OFFICE_DOCUMENT, "word/document.xml");

        self.document_rels
            .get_or_insert(Relationships::default())
            .add_rel(SCHEMA_STYLES, "styles.xml");

        if self.font_table.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_FONT_TABLE, "fontTable.xml");
        }

        if self.footnotes.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_FOOTNOTES, "footnotes.xml");
        }

        if self.endnotes.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_ENDNOTES, "endnotes.xml");
        }

        if self.settings.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_SETTINGS, "settings.xml");
        }

        if self.web_settings.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_WEB_SETTINGS, "webSettings.xml");
        }

        if self.comments.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_COMMENTS, "comments.xml");
        }

        if self.numbering.is_some() {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_NUMBERING, "numbering.xml");
        }

        for hd in &self.headers {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_HEADER, hd.0);
        }

        for ft in &self.footers {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_HEADER, ft.0);
        }

        for theme in &self.themes {
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(SCHEMA_THEME, theme.0);
        }

        for media in &self.media {
            let rel = crate::media::get_media_type_relation_type(&media.1 .0);
            self.document_rels
                .get_or_insert(Relationships::default())
                .add_rel(rel, media.0);
        }

        // ==== Write Zip Item ====

        macro_rules! write_xml {
            (Some($xml:expr) => $name:tt) => {
                if let Some(ref xml) = $xml {
                    write_xml!(xml => $name);
                }
            };
            (Some($xml:expr) => $name:tt $($rest:tt)*) => {
                write_xml!(Some($xml) => $name);
                write_xml!($($rest)*);
            };
            ($xml:expr => $name:tt) => {
                writer.inner.start_file($name, opt)?;
                $xml.to_writer(&mut writer)?;
            };
            ($xml:expr => $name:tt $($rest:tt)*) => {
                write_xml!($xml => $name);
                write_xml!($($rest)*);
            };
        }

        write_xml!(
            self.content_types        => "[Content_Types].xml"
            Some(self.app)            => "docProps/app.xml"
            Some(self.core)           => "docProps/core.xml"
            self.rels                 => "_rels/.rels"
            self.document             => "word/document.xml"
            self.styles               => "word/styles.xml"
            Some(self.font_table)     => "word/fontTable.xml"
            Some(self.footnotes)      => "word/footnotes.xml"
            Some(self.endnotes)       => "word/endnotes.xml"
            Some(self.settings)       => "word/settings.xml"
            Some(self.web_settings)   => "word/webSettings.xml"
            Some(self.comments)       => "word/comments.xml"
            Some(self.numbering)      => "word/numbering.xml"
            Some(self.document_rels)  => "word/_rels/document.xml.rels"
            Some(self.settings_rels)  => "word/_rels/settings.xml.rels"
        );

        for hd in self.headers.iter() {
            let file_path = format!("word/{}", hd.0);
            let content = hd.1;
            write_xml!(
                content => file_path
            );
        }

        for hd in self.footers.iter() {
            let file_path = format!("word/{}", hd.0);
            let content = hd.1;
            write_xml!(
                content => file_path
            );
        }

        for theme in self.themes.iter() {
            let file_path = format!("word/{}", theme.0);
            let content = theme.1;
            write_xml!(
                content => file_path
            );
        }

        for media in self.media.iter() {
            let file_path = format!("word/{}", media.0);
            writer.inner.start_file(file_path, opt)?;
            writer.inner.write_all(media.1 .1)?;
        }

        for (file_path, content) in &self.custom_xml {
            writer.inner.start_file(file_path.clone(), opt)?;
            writer.inner.write_all(content)?;
        }

        Ok(writer.inner.finish()?)
    }

    pub fn write_file<P: AsRef<Path>>(&'a mut self, path: P) -> DocxResult<File> {
        if let Some(p) = path.as_ref().parent() {
            std::fs::create_dir_all(p)?;
        }
        let file = File::create(path)?;
        self.write(file)
    }
}

/// An extracted docx file
pub struct DocxFile {
    app: Option<String>,
    content_types: String,
    core: Option<String>,
    document: String,
    document_rels: Option<String>,
    settings_rels: Option<String>,
    font_table: Option<String>,
    rels: String,
    styles: Option<String>,
    settings: Option<String>,
    web_settings: Option<String>,
    headers: Vec<(String, String)>,
    footers: Vec<(String, String)>,
    themes: Vec<(String, String)>,
    medias: Vec<(String, Vec<u8>)>,
    footnotes: Option<String>,
    endnotes: Option<String>,
    comments: Option<String>,
    numbering: Option<String>,
    custom_xml: Vec<(String, Vec<u8>)>,
}

impl DocxFile {
    /// Extracts from reader
    pub fn from_reader<T: Read + Seek>(reader: T) -> DocxResult<Self> {
        let mut zip = ZipArchive::new(reader)?;

        macro_rules! read {
            ($xml:tt, $name:expr) => {{
                let mut file = zip.by_name($name)?;
                let mut buffer = String::new();
                file.read_to_string(&mut buffer)?;
                buffer
            }};
        }

        macro_rules! option_read {
            ($xml:tt, $name:expr) => {
                match zip.by_name($name) {
                    Err(ZipError::FileNotFound) => None,
                    Err(e) => return Err(e.into()),
                    Ok(mut file) => {
                        let mut buffer = String::new();
                        file.read_to_string(&mut buffer)?;
                        Some(buffer)
                    }
                }
            };
        }

        macro_rules! option_read_multiple {
            ($xml:tt, $name:expr) => {{
                let names: Vec<_> = zip.file_names().map(|x| x.to_string()).collect();
                let name_and_value: Vec<_> = names
                    .iter()
                    .filter(|n| n.contains($name))
                    .filter_map(|f| {
                        zip.by_name(f).ok().and_then(|mut file| {
                            let mut buffer = String::new();
                            file.read_to_string(&mut buffer).ok()?;
                            Some((f.to_string(), buffer))
                        })
                    })
                    .collect();
                name_and_value
            }};
        }

        macro_rules! option_read_multiple_files {
            ($xml:tt, $name:expr) => {{
                let names: Vec<_> = zip.file_names().map(|x| x.to_string()).collect();
                let name_and_value: Vec<_> = names
                    .iter()
                    .filter(|n| n.contains($name))
                    .filter_map(|f| {
                        zip.by_name(f).ok().and_then(|mut file| {
                            let mut buffer = Vec::new();
                            file.read_to_end(&mut buffer).ok()?;
                            Some((f.to_string(), buffer))
                        })
                    })
                    .collect();
                name_and_value
            }};
        }

        let app = option_read!(App, "docProps/app.xml");
        let content_types = read!(ContentTypes, "[Content_Types].xml");
        let core = option_read!(Core, "docProps/core.xml");
        let document_rels = option_read!(Relationships, "word/_rels/document.xml.rels");
        let settings_rels = option_read!(Relationships, "word/_rels/settings.xml.rels");
        let document = read!(Document, "word/document.xml");
        let font_table = option_read!(FontTable, "word/fontTable.xml");
        let rels = read!(Relationships, "_rels/.rels");
        let styles = option_read!(Styles, "word/styles.xml");
        let settings = option_read!(Settings, "word/settings.xml");
        let web_settings = option_read!(WebSettings, "word/webSettings.xml");
        let footnotes = option_read!(Footnotes, "word/footnotes.xml");
        let endnotes = option_read!(Endnotes, "word/endnotes.xml");
        let comments = option_read!(Comments, "word/comments.xml");
        let numbering = option_read!(Numbering, "word/numbering.xml");

        let headers = option_read_multiple!(Headers, "word/header");
        let footers = option_read_multiple!(Footers, "word/footer");
        let themes = option_read_multiple!(Themes, "word/theme/theme");
        let medias = option_read_multiple_files!(Medias, "word/media");
        let custom_xml = option_read_multiple_files!(_, "custom");

        Ok(DocxFile {
            app,
            content_types,
            core,
            document_rels,
            settings_rels,
            document,
            font_table,
            rels,
            styles,
            settings,
            web_settings,
            headers,
            footers,
            themes,
            medias,
            footnotes,
            endnotes,
            comments,
            numbering,
            custom_xml,
        })
    }

    /// Extracts from file
    #[inline]
    pub fn from_file<P: AsRef<Path>>(path: P) -> DocxResult<Self> {
        Self::from_reader(File::open(path)?)
    }

    /// Parses content into `Docx` struct
    pub fn parse(&self) -> DocxResult<Docx<'_>> {
        let app = if let Some(content) = &self.app {
            Some(App::from_str(content)?)
        } else {
            None
        };

        let document = Document::from_str(&self.document)?;

        let mut headers = HashMap::new();
        for f in self.headers.iter() {
            let hd = Header::from_str(&f.1)?;
            let name = f.0.replace("word/", "");
            headers.insert(name, hd);
        }

        let mut footers = HashMap::new();
        for f in self.footers.iter() {
            let ft = Footer::from_str(&f.1)?;
            let name = f.0.replace("word/", "");
            footers.insert(name, ft);
        }

        let mut media = HashMap::new();
        for m in self.medias.iter() {
            let mt = crate::media::get_media_type(&m.0);
            if let Some(mt) = mt {
                let name = m.0.replace("word/", "");
                let m = (mt, &m.1);
                media.insert(name, m);
            }
        }

        let mut themes = HashMap::new();
        // turn off for now
        for t in self.themes.iter() {
            let th = Theme::from_str(&t.1)?;
            let name = t.0.replace("word/", "");
            themes.insert(name, th);
        }

        let content_types = ContentTypes::from_str(&self.content_types)?;

        let core = if let Some(content) = &self.core {
            Some(Core::from_str(content)?)
        } else {
            None
        };

        let document_rels: Option<Relationships> = if let Some(content) = &self.document_rels {
            Some(Relationships::from_str(content)?)
        } else {
            None
        };
        let document_rels = document_rels.map(|rel: Relationships| {
            let rrr: Vec<_> = rel
                .relationships
                .iter()
                .filter(|r2| {
                    matches!(
                        r2.ty.to_string().as_str(),
                        crate::schema::SCHEMA_HEADER
                            | crate::schema::SCHEMA_FOOTER
                            | crate::schema::SCHEMA_THEME
                            | crate::schema::SCHEMA_FONT_TABLE
                            | crate::schema::SCHEMA_STYLES
                            | crate::schema::SCHEMA_FOOTNOTES
                            | crate::schema::SCHEMA_ENDNOTES
                            | crate::schema::SCHEMA_SETTINGS
                            | crate::schema::SCHEMA_WEB_SETTINGS
                            | crate::schema::SCHEMA_COMMENTS
                            | crate::schema::SCHEMA_IMAGE
                            | crate::schema::SCHEMA_HYPERLINK
                            | crate::schema::SCHEMA_NUMBERING
                    )
                })
                .map(|d| d.to_owned())
                .collect();
            Relationships { relationships: rrr }
        });

        let settings_rels = self
            .settings_rels
            .as_deref()
            .map(Relationships::from_str)
            .transpose()?;

        let font_table = if let Some(content) = &self.font_table {
            Some(FontTable::from_str(content)?)
        } else {
            None
        };

        let footnotes = if let Some(content) = &self.footnotes {
            Some(FootNotes::from_str(content)?)
        } else {
            None
        };

        let endnotes = if let Some(content) = &self.endnotes {
            Some(EndNotes::from_str(content)?)
        } else {
            None
        };

        let settings = if let Some(content) = &self.settings {
            Some(Settings::from_str(content)?)
        } else {
            None
        };

        let web_settings = if let Some(content) = &self.web_settings {
            Some(WebSettings::from_str(
                &content.replace("ns0:", "w:").to_string(),
            )?)
        } else {
            None
        };

        let comments = if let Some(content) = &self.comments {
            Some(Comments::from_str(content)?)
        } else {
            None
        };

        let numbering = if let Some(content) = &self.numbering {
            Some(Numbering::from_str(content)?)
        } else {
            None
        };

        let rels = Relationships::from_str(&self.rels)?;
        let rels = {
            let rrr: Vec<_> = rels
                .relationships
                .iter()
                .filter(|r2| {
                    matches!(
                        r2.ty.to_string().as_str(),
                        crate::schema::SCHEMA_CORE
                            | crate::schema::SCHEMA_REL_EXTENDED
                            | crate::schema::SCHEMA_OFFICE_DOCUMENT
                    )
                })
                .map(|d| d.to_owned())
                .collect();
            Relationships { relationships: rrr }
        };

        let styles = self
            .styles
            .as_ref()
            .map(|content| Styles::from_str(content))
            .transpose()?
            .unwrap_or_default();

        let custom_xml = self
            .custom_xml
            .iter()
            .map(|(name, content)| (name.to_string(), Cow::Borrowed(content.as_slice())))
            .collect();

        Ok(Docx {
            app,
            content_types,
            core,
            document,
            document_rels,
            settings_rels,
            font_table,
            rels,
            styles,
            headers,
            footers,
            themes,
            media,
            footnotes,
            endnotes,
            settings,
            web_settings,
            comments,
            numbering,
            custom_xml,
        })
    }
}
