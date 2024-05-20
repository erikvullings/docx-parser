use base64::prelude::*;
use serde::ser::SerializeMap;
use serde::Serializer;
use std::{
    collections::HashMap,
    env,
    fs::{create_dir_all, File},
    io::{self, Write},
    path::PathBuf,
};

pub fn max_lengths_per_column(
    table_with_simple_cells: &Vec<(bool, Vec<String>)>,
    min_width: usize,
) -> Vec<usize> {
    // Check if the table is empty
    if table_with_simple_cells.is_empty() {
        return vec![];
    }

    // Determine the number of columns by inspecting the first row
    let num_columns = table_with_simple_cells[0].1.len();

    // Initialize a vector to store the maximum length of each column
    let mut max_lengths = vec![min_width; num_columns];

    // Iterate over each row
    for (_, row) in table_with_simple_cells {
        // Iterate over each cell in the row
        for (i, cell) in row.iter().enumerate() {
            // The first row may have merged cells.
            if i == max_lengths.len() {
                max_lengths.push(0);
            }
            // Update the max length for the current column
            if cell.len() > max_lengths[i] {
                max_lengths[i] = cell.len();
            }
        }
    }

    max_lengths
}

pub fn pad_left(s: &str, width: &usize) -> String {
    let mut padded = String::new();
    // If the string is already long enough, return it unchanged.
    if *width <= s.len() {
        return s.to_string();
    }
    let padding = width - s.len();
    // Add padding to the left of the string.
    padded.push_str(s);
    padded.push_str(&" ".repeat(padding));
    padded
}

pub fn table_row_to_markdown(column_lengths: &Vec<usize>, row: &Vec<String>) -> String {
    let mut table_row_in_markdown = "".to_string();
    column_lengths.iter().enumerate().for_each(|(j, width)| {
        let cell = if j < row.len() { &row[j] } else { "" };
        table_row_in_markdown.push_str(&format!("| {} ", pad_left(cell, width)));
    });
    table_row_in_markdown.push_str("|\n");
    table_row_in_markdown
}

#[test]
fn test_pad_left() {
    let text = "This is a test".to_string();
    let width = 10;
    let padded = pad_left(&text, &width);
    assert_eq!(padded, "This is a test");
    let width = 20;
    let padded = pad_left(&text, &width);
    assert_eq!(padded, "This is a test      ");
}

#[test]
fn test_table_row_to_markdown() {
    let column_lengths = vec![10, 15, 20];
    let row = vec![
        "This is".to_string(),
        "This is a".to_string(),
        "This is a test".to_string(),
    ];
    let table_row_in_markdown = table_row_to_markdown(&column_lengths, &row);
    assert_eq!(
        table_row_in_markdown,
        "| This is    | This is a       | This is a test       |\n",
    );
}

pub fn save_image_to_file(path: &str, image_data: &[u8]) -> io::Result<()> {
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

fn get_mime_type(filename: &str) -> Option<&'static str> {
    let extension = filename.split('.').last()?;
    match extension.to_lowercase().as_str() {
        "png" => Some("image/png"),
        "jpg" | "jpeg" => Some("image/jpeg"),
        "gif" => Some("image/gif"),
        "bmp" => Some("image/bmp"),
        "tiff" => Some("image/tiff"),
        _ => None,
    }
}

pub fn serialize_images<S>(
    images: &HashMap<String, Vec<u8>>,
    serializer: S,
) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    let mut map = serializer.serialize_map(Some(images.len()))?;
    for (key, value) in images {
        let encoded = BASE64_STANDARD.encode(value);

        let prefix = match get_mime_type(key) {
            Some(mime_type) => format!("data:{};base64,", mime_type),
            None => "data:application/octet-stream;base64,".to_string(),
        };
        let base64_string = format!("{}{}", prefix, encoded);
        map.serialize_entry(key, &base64_string)?;
    }
    map.end()
}
