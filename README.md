# DOXC-PARSER

This package uses the [docx-rs](https://crates.io/crates/docx-rust) crate to parse docx files. It subsequently converts the parsed docx file into Markdown format. Alternatively, it can also be used to convert docx files into JSON format, where only the structure relevant for creating Markdown documents is kept.

```bash
$ docx-parser -h

Processes a DOCX file and outputs as Markdown or JSON

Usage: docx-parser [OPTIONS] <FILE>

Arguments:
  <FILE>  The input DOCX file

Options:
  -o, --output <OUTPUT>  Sets the output destination. Default is console
  -f, --format <FORMAT>  Sets the output format. Default is markdown. Options: md, json
  -h, --help             Print help
  -V, --version          Print version
```

## Example

```rust
use docx_parser::MarkdownDocument;

let markdown_doc = MarkdownDocument::from_file("./test/tables.docx");
let markdown = markdown_doc.to_markdown(true);
let json = markdown_doc.to_json(true);

println!("\n\n{}", markdown);
println!("\n\n{}", json);
```

## Development

```bash
cargo update
cargo test
cargo build --release
```