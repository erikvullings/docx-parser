use doc_parser::MarkdownDocument;

fn main() {
    let markdown_doc = MarkdownDocument::from_file("./test/headers.docx");
    println!("\n\n{}", markdown_doc.to_markdown());
}
