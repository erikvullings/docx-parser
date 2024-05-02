#!/bin/bash

# Specify the input folder containing the .docx files
input_folder="./test"

# Specify the media folder where media files will be extracted
media_folder="${input_folder}"

# Loop through all .docx files in the input folder
for file in "$input_folder"/*.docx; do
    # Check if the file is a regular file
    if [ -f "$file" ]; then
        # Extract the file name without extension
        filename=$(basename -- "$file")
        filename_no_ext="${filename%.*}"
        
        # Convert the .docx file to Markdown using Pandoc. See also https://stackoverflow.com/a/74654058/319711 for heading style
        pandoc -s "$file" --wrap=none --reference-links --atx-headers --extract-media="$input_folder" -t markdown -o "${input_folder}/${filename_no_ext}.md"
        
        echo "Converted $file to Markdown."
    fi
done

