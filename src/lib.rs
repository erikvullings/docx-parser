use std::collections::HashMap;
use std::env;
use std::fs::{create_dir_all, File};
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::str::FromStr;

use docx_rust::document::BodyContent::{Paragraph, Sdt, SectionProperty, Table, TableCell};
use docx_rust::document::{ParagraphContent, RunContent};
use docx_rust::formatting::{NumberFormat, ParagraphProperty};
use docx_rust::media::MediaType;
use docx_rust::styles::StyleType;
use docx_rust::DocxFile;

fn save_image_to_file(path: &str, image_data: &[u8]) -> io::Result<()> {
    // Get the current working directory
    let current_dir = env::current_dir()?;

    // Concatenate the file path to the current working directory
    let full_path = current_dir.join(path);

    // Create the directory if it doesn't exist
    if let Some(parent) = full_path.parent() {
        create_dir_all(parent)?;
    }

    // Convert the path to a PathBuf
    let mut file_path = PathBuf::new();
    file_path.push(full_path);

    // Create a file at the specified path
    let mut file = File::create(&file_path)?;

    // Write the image data to the file
    file.write_all(image_data)?;

    Ok(())
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct BlockStyle {
    pub bold: bool,
    pub italics: bool,
    pub underline: bool,
    pub strike: bool,
    /// Size is specified in points x 2, so size 19 is equal to 9.5pt
    pub size: Option<isize>,
}

impl BlockStyle {
    pub fn new() -> Self {
        BlockStyle {
            bold: false,
            italics: false,
            underline: false,
            strike: false,
            size: None,
        }
    }

    pub fn combine_with(&mut self, other: &BlockStyle) {
        self.bold = other.bold;
        self.italics = other.italics;
        self.underline = other.underline;
        self.strike = other.strike;
        if let Some(size) = other.size {
            self.size = Some(size);
        }
    }
}

#[derive(Debug, Clone)]
pub struct MarkdownNumbering {
    pub id: Option<isize>,
    pub indent_level: Option<isize>,
    pub format: Option<String>, // NumberFormat
    pub level_text: Option<String>,
}

#[derive(Debug, Default, Clone)]
pub struct ParagraphStyle {
    pub style_id: Option<String>,
    pub outline_lvl: Option<isize>,
    pub numbering: Option<MarkdownNumbering>,
    pub page_break_before: Option<bool>,
    pub style: Option<BlockStyle>,
}

impl ParagraphStyle {
    pub fn new() -> Self {
        ParagraphStyle {
            style_id: None,
            outline_lvl: None,
            numbering: None,
            page_break_before: None,
            style: None,
        }
    }

    pub fn combine_with(&mut self, other: &ParagraphStyle) {
        self.style_id = self.style_id.clone().or_else(|| other.style_id.clone());
        self.outline_lvl = self.outline_lvl.or_else(|| other.outline_lvl);
        self.page_break_before = self.page_break_before.or_else(|| other.page_break_before);
        if self.numbering.is_none() {
            self.numbering = other.numbering.clone()
        }
        if let Some(ref mut style) = self.style {
            if let Some(ref other_style) = other.style {
                style.combine_with(other_style);
            }
        } else {
            self.style = other.style.clone();
        }
    }
}

impl<'a> From<&'a ParagraphProperty<'a>> for ParagraphStyle {
    fn from(paragraph_property: &'a ParagraphProperty) -> Self {
        // Extract properties from ParagraphProperty and create a new ParagraphStyle
        let mut paragraph_style = ParagraphStyle::new();
        if let Some(style_id) = &paragraph_property.style_id {
            paragraph_style.style_id = Some(style_id.value.to_string());
        }
        if let Some(outline_lvl) = &paragraph_property.outline_lvl {
            paragraph_style.outline_lvl = Some(outline_lvl.value);
        }
        if let Some(page_break_before) = &paragraph_property.page_break_before {
            paragraph_style.page_break_before = page_break_before.value;
        }
        if let Some(numbering) = &paragraph_property.numbering {
            paragraph_style.numbering = Some(MarkdownNumbering {
                id: numbering.id.as_ref().map(|ni| ni.value),
                indent_level: numbering.level.as_ref().map(|level| level.value),
                format: None,
                level_text: None,
            });
        }
        if paragraph_property.r_pr.len() > 0 {
            let mut block_style = BlockStyle::new();
            paragraph_property
                .r_pr
                .iter()
                .for_each(|character_property| {
                    if let Some(size) = &character_property.size {
                        block_style.size = Some(size.value);
                    }
                    if character_property.bold.is_some() {
                        block_style.bold = true;
                    }
                    if character_property.underline.is_some() {
                        block_style.underline = true;
                    }
                    if character_property.italics.is_some() || character_property.emphasis.is_some()
                    {
                        block_style.italics = true;
                    }
                    if character_property.strike.is_some() || character_property.dstrike.is_some() {
                        block_style.strike = true;
                    }
                });
            paragraph_style.style = Some(block_style);
        }
        paragraph_style
    }
}

#[derive(Debug, PartialEq, Eq)]
pub enum TextType {
    Text,
    Image,
    Link,
    Code,
    Quote,
    List,
    Table,
    Header,
    HorizontalRule,
    BlockQuote,
    CodeBlock,
    HeaderBlock,
    BookmarkLink,
}

#[derive(Debug, PartialEq, Eq)]
pub struct TextBlock {
    pub text_type: TextType,
    pub style: Option<BlockStyle>,
    pub text: String,
}

impl TextBlock {
    pub fn new(text: String, style: Option<BlockStyle>, text_type: TextType) -> Self {
        TextBlock {
            style,
            text,
            text_type,
        }
    }

    pub fn to_markdown(&self, paragraph_style: &ParagraphStyle) -> String {
        let mut markdown = self.text.clone();

        let mut style = if self.style.is_some() {
            self.style.as_ref().unwrap().clone()
        } else {
            BlockStyle::new()
        };

        if let Some(block_style) = &paragraph_style.style {
            style.combine_with(block_style);
        };

        // Add bold formatting if enabled
        if style.bold {
            markdown = format!("**{markdown}**");
        }

        // Add italic formatting if enabled
        if style.italics {
            markdown = format!("*{markdown}*");
        }

        // Add underline formatting if enabled
        if style.underline {
            markdown = format!("__{markdown}__");
        }

        // Add strike-through formatting if enabled
        if style.strike {
            markdown = format!("~~{markdown}~~");
        }
        markdown
    }
}

#[derive(Debug)]
pub struct MarkdownParagraph {
    pub style: Option<ParagraphStyle>,
    pub blocks: Vec<TextBlock>,
}

impl MarkdownParagraph {
    pub fn new() -> Self {
        MarkdownParagraph {
            style: None,
            blocks: vec![],
        }
    }

    pub fn to_markdown(
        &self,
        styles: &HashMap<String, ParagraphStyle>,
        numberings: &mut HashMap<isize, usize>,
        doc: &MarkdownDocument,
    ) -> String {
        let mut markdown = String::new();

        let mut style = if self.style.is_some() {
            self.style.as_ref().unwrap().clone()
        } else {
            ParagraphStyle::default()
        };

        if let Some(style_id) = &style.style_id {
            if let Some(doc_style) = styles.get(style_id) {
                style.combine_with(doc_style);
            }
            // markdown += &format!("[{}]", style_id);
        };

        // Add outline level if available
        if let Some(outline_lvl) = style.outline_lvl {
            // Convert outline level to appropriate Markdown heading level
            let heading_level = match outline_lvl {
                0 => "# ",
                1 => "## ",
                2 => "### ",
                3 => "#### ",
                4 => "##### ",
                _ => "###### ", // Use the smallest heading level for higher levels
            };
            markdown += heading_level;
        }

        // Add numbering if available
        if let Some(numbering) = &style.numbering {
            if let Some(level) = numbering.indent_level {
                if level > 0 {
                    markdown += &"    ".repeat(level as usize); // Start numbering from 1
                }
            }
            if let Some(id) = numbering.id {
                let format = match &doc.numberings[&id].format {
                    Some(entry) => NumberFormat::from_str(entry).unwrap_or(NumberFormat::Decimal),
                    None => NumberFormat::Decimal,
                };
                let count = numberings.entry(id).or_insert(0); // Start numbering from 1
                let numbering_symbol = match format {
                    NumberFormat::UpperRoman => format!("{}.", ((*count) as u8 + b'I') as char),
                    NumberFormat::LowerRoman => format!("{}.", ((*count) as u8 + b'i') as char),
                    NumberFormat::UpperLetter => format!("{}.", ((*count) as u8 + b'A') as char),
                    NumberFormat::LowerLetter => format!("{}.", ((*count) as u8 + b'a') as char),
                    NumberFormat::Bullet => match &doc.numberings[&id].level_text {
                        Some(level_text) if level_text.trim().is_empty() => " ".to_string(),
                        _ => "-".to_string(),
                    },
                    _ => format!("{}.", *count + 1),
                };
                *count += 1;
                markdown += &format!("{numbering_symbol} ");
            }
        }

        for block in &self.blocks {
            markdown += &block.to_markdown(&style);
        }
        markdown += "\n";

        // Add page break before if available
        // if let Some(page_break_before) = style.page_break_before {
        //     markdown += &format!("{{page_break_before: {}}}", page_break_before);
        // }

        markdown
    }
}

#[derive(Debug)]
pub struct MarkdownDocument {
    pub creator: Option<String>,
    pub last_editor: Option<String>,
    pub company: Option<String>,
    pub title: Option<String>,
    pub description: Option<String>,
    pub subject: Option<String>,
    pub keywords: Option<String>,
    pub paragraphs: Vec<MarkdownParagraph>,
    pub styles: HashMap<String, ParagraphStyle>,
    pub numberings: HashMap<isize, MarkdownNumbering>,
    pub images: HashMap<String, Vec<u8>>,
}

impl MarkdownDocument {
    pub fn new() -> Self {
        MarkdownDocument {
            creator: None,
            last_editor: None,
            company: None,
            title: None,
            description: None,
            subject: None,
            keywords: None,
            paragraphs: vec![],
            styles: HashMap::new(),
            numberings: HashMap::new(),
            images: HashMap::new(),
        }
    }

    pub fn from_file<P: AsRef<Path>>(path: P) -> Self {
        let mut markdown_doc = MarkdownDocument::new();

        let docx = match DocxFile::from_file(path) {
            Ok(docx_file) => docx_file,
            Err(err) => {
                panic!("Error processing file: {:?}", err)
            }
        };
        let docx = match docx.parse() {
            Ok(docx) => docx,
            Err(err) => {
                panic!("Exiting: {:?}", err);
            }
        };

        // println!("{:?}", &docx);

        if let Some(app) = docx.app {
            if let Some(company) = app.company {
                if !company.is_empty() {
                    markdown_doc.company = Some(company.to_string());
                }
            }
        }

        if let Some(core) = docx.core {
            if let Some(title) = core.title {
                if !title.is_empty() {
                    markdown_doc.title = Some(title.to_string());
                }
            }
            if let Some(subject) = core.subject {
                if !subject.is_empty() {
                    markdown_doc.subject = Some(subject.to_string());
                }
            }
            if let Some(keywords) = core.keywords {
                if !keywords.is_empty() {
                    markdown_doc.keywords = Some(keywords.to_string());
                }
            }
            if let Some(description) = core.description {
                if !description.is_empty() {
                    markdown_doc.description = Some(description.to_string());
                }
            }
            if let Some(creator) = core.creator {
                if !creator.is_empty() {
                    markdown_doc.creator = Some(creator.to_string());
                }
            }
            if let Some(last_modified_by) = core.last_modified_by {
                if !last_modified_by.is_empty() {
                    markdown_doc.last_editor = Some(last_modified_by.to_string());
                }
            }
        }

        if let Some(numbering) = docx.numbering {
            numbering.numberings.iter().for_each(|n| {
                if let Some(id) = n.num_id {
                    if let Some(details) = numbering.numbering_details(id) {
                        markdown_doc.numberings.insert(
                            id,
                            MarkdownNumbering {
                                id: Some(id),
                                indent_level: None,
                                format: details.levels[0]
                                    .number_format
                                    .as_ref()
                                    .map(|i| i.value.to_string()),
                                level_text: details.levels[0]
                                    .level_text
                                    .as_ref()
                                    .map(|i| i.value.to_string()),
                            },
                        );
                        ()
                    }
                }
            })
        }

        for (id, (MediaType::Image, media_data)) in docx.media {
            markdown_doc.images.insert(id, media_data.to_vec());
        }

        for style in &docx.styles.styles {
            match style.ty {
                Some(StyleType::Paragraph) => {
                    if let Some(paragraph_property) = &style.paragraph {
                        let paragraph_style: ParagraphStyle = paragraph_property.into();
                        markdown_doc
                            .styles
                            .insert(style.style_id.to_string(), paragraph_style);
                    }
                }
                _ => (),
            }
        }

        for content in docx.document.body.content {
            match content {
                Paragraph(paragraph) => {
                    let mut markdown_paragraph = MarkdownParagraph::new();
                    if let Some(paragraph_property) = &paragraph.property {
                        let paragraph_style: ParagraphStyle = paragraph_property.into();
                        markdown_paragraph.style = Some(paragraph_style);
                    }
                    for paragraph_content in paragraph.content {
                        match paragraph_content {
                            ParagraphContent::Run(run) => {
                                let block_style = match run.property {
                                    Some(character_property) => {
                                        let mut block_style = BlockStyle::new();
                                        if let Some(size) = character_property.size {
                                            block_style.size = Some(size.value);
                                        }
                                        if character_property.bold.is_some() {
                                            block_style.bold = true;
                                        }
                                        if character_property.underline.is_some() {
                                            block_style.underline = true;
                                        }
                                        if character_property.italics.is_some()
                                            || character_property.emphasis.is_some()
                                        {
                                            block_style.italics = true;
                                        }
                                        if character_property.strike.is_some()
                                            || character_property.dstrike.is_some()
                                        {
                                            block_style.strike = true;
                                        }
                                        Some(block_style)
                                    }
                                    None => None,
                                };

                                let is_same_style =
                                    |style: &Option<BlockStyle>| style == &block_style;

                                for run_content in run.content {
                                    match run_content {
                                        RunContent::Text(text) => {
                                            let text = text.text.to_string();
                                            let mut could_extend_text = false;
                                            if let Some(prev_block) =
                                                markdown_paragraph.blocks.last_mut()
                                            {
                                                if is_same_style(&prev_block.style) {
                                                    prev_block.text.push_str(&text);
                                                    could_extend_text = true
                                                }
                                            };
                                            if !could_extend_text {
                                                let text_block = TextBlock::new(
                                                    text,
                                                    block_style.clone(),
                                                    TextType::Text,
                                                );
                                                markdown_paragraph.blocks.push(text_block);
                                            }
                                        }
                                        RunContent::Drawing(drawing) => {
                                            if let Some(inline) = drawing.inline {
                                                if let Some(graphic) = inline.graphic {
                                                    let id = graphic
                                                        .data
                                                        .pic
                                                        .fill
                                                        .blip
                                                        .embed
                                                        .to_string();
                                                    if let Some(relationships) = &docx.document_rels
                                                    {
                                                        if let Some(target) =
                                                            relationships.get_target(&id)
                                                        {
                                                            let descr =
                                                                match inline.doc_property.descr {
                                                                    Some(descr) => descr.into(),
                                                                    None => "".to_string(),
                                                                };
                                                            let img_text = format!(
                                                                "![{}](./{})",
                                                                descr, target
                                                            );
                                                            let text_block = TextBlock::new(
                                                                img_text,
                                                                None,
                                                                TextType::Image,
                                                            );
                                                            markdown_paragraph
                                                                .blocks
                                                                .push(text_block);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        _ => (),
                                    }
                                }
                            }
                            ParagraphContent::Link(link) => {
                                let descr = link.content.content.first();
                                let target = match link.anchor {
                                    Some(anchor) => Some(format!("#{}", anchor.to_string())),
                                    None => match link.id {
                                        Some(id) => match &docx.document_rels {
                                            Some(doc_relationships) => doc_relationships
                                                .relationships
                                                .iter()
                                                .find_map(|r| {
                                                    if r.id == id {
                                                        Some(r.target.to_string())
                                                    } else {
                                                        None
                                                    }
                                                }),
                                            None => None,
                                        },
                                        None => None,
                                    },
                                };
                                if let (Some(RunContent::Text(descr)), Some(target)) =
                                    (descr, target)
                                {
                                    let link = format!("[{}]({})", descr.text, target);
                                    let text_block = TextBlock::new(link, None, TextType::Link);
                                    markdown_paragraph.blocks.push(text_block);
                                }
                            }
                            ParagraphContent::BookmarkStart(bookmark_start) => {
                                if let Some(name) = bookmark_start.name {
                                    let bookmark = format!(r#"<a name="{}"></a>"#, name);
                                    let text_block =
                                        TextBlock::new(bookmark, None, TextType::BookmarkLink);
                                    markdown_paragraph.blocks.push(text_block);
                                }
                            }
                            _ => (),
                        }
                    }
                    if markdown_paragraph.blocks.len() > 0 {
                        markdown_doc.paragraphs.push(markdown_paragraph);
                    }
                }
                Table(table) => {
                    println!("Table: {:?}", table);
                }
                Sdt(_) => {
                    // println!("Sdt");
                }
                SectionProperty(_sp) => {
                    // println!("SectionProperty: {:?}", sp);
                }
                TableCell(tc) => {
                    println!("TableCell: {:?}", tc);
                }
            }
        }

        markdown_doc
    }

    pub fn to_markdown(&self, export_images: bool) -> String {
        let mut markdown = String::new();

        if let Some(title) = &self.title {
            markdown += &format!("# {}\n\n", title);
        }

        let mut numberings: HashMap<isize, usize> = HashMap::new();

        for (index, paragraph) in self.paragraphs.iter().enumerate() {
            markdown += &paragraph.to_markdown(&self.styles, &mut numberings, &self);
            if index != self.paragraphs.len() - 1 {
                markdown += "\n";
            }
        }

        if export_images {
            for (image, data) in &self.images {
                match save_image_to_file(image, data) {
                    Ok(_) => (),
                    Err(err) => eprintln!("{err}"),
                };
            }
        }

        markdown
    }
}

#[cfg(test)]
mod tests {
    use std::fs;

    // Import necessary items
    use super::*;

    #[test]
    fn test_headers() {
        let markdown_pandoc = fs::read_to_string("./test/headers.md").unwrap();
        let markdown_doc = MarkdownDocument::from_file("./test/headers.docx");
        let markdown = markdown_doc.to_markdown(false);
        assert_eq!(markdown_pandoc, markdown);
    }

    #[test]
    fn test_bullets() {
        let markdown_pandoc = fs::read_to_string("./test/lists.md").unwrap();
        let markdown_doc = MarkdownDocument::from_file("./test/lists.docx");
        let markdown = markdown_doc.to_markdown(false);
        assert_eq!(markdown_pandoc, markdown);
    }

    #[test]
    fn test_images() {
        let markdown_pandoc = fs::read_to_string("./test/image.md").unwrap();
        let markdown_doc = MarkdownDocument::from_file("./test/image.docx");
        let markdown = markdown_doc.to_markdown(false);
        assert_eq!(markdown_pandoc, markdown);
    }

    #[test]
    fn test_links() {
        let markdown_pandoc = fs::read_to_string("./test/links.md").unwrap();
        let markdown_doc = MarkdownDocument::from_file("./test/links.docx");
        let markdown = markdown_doc.to_markdown(false);
        assert_eq!(markdown_pandoc, markdown);
    }

    // More test functions can be defined here
}
