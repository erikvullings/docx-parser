use clap::{command, Parser};
use docx_parser::MarkdownDocument;
use std::{fs, path::PathBuf};

#[derive(Parser)]
#[command(name = "docx-parser")]
#[command(version = "0.1.0")]
#[command(author = "Erik Vullings <erik.vullings@gmail.com>")]
#[command(about = "Processes a DOCX file and outputs as Markdown or JSON", long_about = None)]
struct Cli {
    /// The input DOCX file
    #[arg(short, long, value_name = "FILE", required = true)]
    input: PathBuf,

    /// Sets the output destination. Default is console.
    #[arg(short, long)]
    output: Option<String>,

    /// Sets the output format. Default is markdown. Options: md, json.
    #[arg(short, long)]
    format: Option<String>,
}

fn main() {
    let cli = Cli::parse();

    println!("File: {:?}", cli.input);

    let output = match cli.output {
        Some(file) => file,
        None => "console".to_string(),
    };

    let format = match cli.format {
        Some(format) => {
            if format == "json" || format == "md" {
                format
            } else {
                "md".to_string()
            }
        }
        None => "md".to_string(),
    };

    if format != "md" && format != "json" {
        eprintln!(
            "Unsupported format: {}. Supported formats are md and json.",
            format
        );
        std::process::exit(1);
    }

    let mut input_file = cli.input.to_string_lossy().trim().to_string();

    if !input_file.ends_with(".docx") {
        input_file = format!("{}.docx", input_file);
    }

    if !file_exists_and_readable(&input_file) {
        eprintln!(
            "Input file does not exist or cannot be read: {:?}",
            input_file
        );
        std::process::exit(1);
    }

    println!("Processing file: {:?}", input_file);
    println!("Output destination: {}", output);
    println!("Output format: {}", format);

    let markdown_doc = MarkdownDocument::from_file(input_file);
    let result = if format == "md" {
        markdown_doc.to_markdown(true)
    } else {
        markdown_doc.to_json()
    };
    if output == "console" {
        println!("{result}");
    } else {
        fs::write(output, result).expect("Could not write output");
    }
}

fn file_exists_and_readable(path: &str) -> bool {
    fs::metadata(path)
        .map(|metadata| metadata.is_file())
        .unwrap_or(false)
}

// fn test() {
//     let markdown_doc = MarkdownDocument::from_file("./test/tables.docx");
//     println!("\n\n{}", markdown_doc.to_markdown(true));
//     println!("\n\n{}", markdown_doc.to_json());
// }
