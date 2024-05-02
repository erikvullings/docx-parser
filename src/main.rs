use std::collections::HashMap;

use docx_rust::document::BodyContent::{Paragraph, Sdt, SectionProperty, Table, TableCell};
use docx_rust::document::{ParagraphContent, RunContent};
use docx_rust::formatting::ParagraphProperty;
use docx_rust::styles::StyleType;
use docx_rust::DocxFile;

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
pub struct ParagraphStyle {
    pub style_id: Option<String>,
    pub outline_lvl: Option<isize>,
    pub numbering: Option<bool>,
    pub page_break_before: Option<bool>,
    pub style: Option<BlockStyle>,
}

impl Default for ParagraphStyle {
    fn default() -> Self {
        ParagraphStyle::new()
    }
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
        self.numbering = self.numbering.or_else(|| other.numbering);
        self.page_break_before = self.page_break_before.or_else(|| other.page_break_before);
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
        if let Some(_) = paragraph_property.numbering {
            paragraph_style.numbering = Some(true);
        }
        if let Some(character_property) = &paragraph_property.r_pr {
            let mut block_style = BlockStyle::new();
            if let Some(size) = &character_property.size {
                block_style.size = Some(size.value);
            }
            if character_property.bold.is_some() {
                block_style.bold = true;
            }
            if character_property.underline.is_some() {
                block_style.underline = true;
            }
            if character_property.italics.is_some() || character_property.emphasis.is_some() {
                block_style.italics = true;
            }
            if character_property.strike.is_some() || character_property.dstrike.is_some() {
                block_style.strike = true;
            }
            paragraph_style.style = Some(block_style);
        }
        paragraph_style
    }
}

#[derive(Debug, PartialEq, Eq)]
pub struct TextBlock {
    pub style: Option<BlockStyle>,
    pub text: String,
}

impl TextBlock {
    pub fn new(text: String, style: Option<BlockStyle>) -> Self {
        TextBlock { style, text }
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

    pub fn to_markdown(&self, styles: &HashMap<String, ParagraphStyle>) -> String {
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
        if let Some(numbering) = style.numbering {
            if numbering {
                markdown += "1. "; // Start numbering from 1
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
        }
    }

    pub fn to_markdown(&self) -> String {
        let mut markdown = String::new();

        if let Some(title) = &self.title {
            markdown += &format!("# {}\n\n", title);
        }

        for paragraph in &self.paragraphs {
            markdown += &paragraph.to_markdown(&self.styles);
            markdown += "\n";
        }

        markdown
    }
}

fn main() {
    let docx = match DocxFile::from_file("./test/headers.docx") {
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

    let mut markdown_doc = MarkdownDocument::new();

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

    for media in docx.media {
        println!("Media: {media:?}",);
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
                            if let Some(text) =
                                run.content
                                    .into_iter()
                                    .find_map(|run_content| match run_content {
                                        RunContent::Text(text) => Some(text.text.to_string()),
                                        _ => None,
                                    })
                            {
                                let could_extend_text = if let Some(prev_block) =
                                    markdown_paragraph.blocks.last_mut()
                                {
                                    if prev_block.style == block_style {
                                        prev_block.text.push_str(&text);
                                        true
                                    } else {
                                        false
                                    }
                                } else {
                                    false
                                };
                                if !could_extend_text {
                                    let text_block = TextBlock::new(text, block_style);
                                    markdown_paragraph.blocks.push(text_block);
                                }
                            };
                        }
                        ParagraphContent::Link(link) => {
                            println!("  Link: {:?}", link);
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
    println!("\n\n{}", markdown_doc.to_markdown());
}
