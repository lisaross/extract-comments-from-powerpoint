# PowerPoint Comment Extractor

This Python script extracts comments from all PowerPoint (.pptx) files in the current directory and displays them in order of appearance.

## Features

- Processes all .pptx files in the current directory
- Extracts comments from each PowerPoint file
- Sorts comments by slide number
- Displays comments without slide numbers, author IDs, or dates

## Requirements

- Python 3.6 or higher
- No additional libraries required (uses only built-in Python modules)

## Usage

1. Clone this repository or download the `extract_comments_all_pptx.py` file.
2. Place the script in the same directory as your PowerPoint files.
3. Open a terminal or command prompt.
4. Navigate to the directory containing the script and PowerPoint files.
5. Run the script:

```python extract_comments_all_pptx.py```

6. The script will process all .pptx files in the directory and display the extracted comments for each file.

## Output

The script will display:
- The name of each PowerPoint file being processed
- The total number of comments extracted from each file
- The text of each comment, separated by dashed lines

## Limitations

- This script is designed to work with modern PowerPoint files (.pptx format).
- It may not extract comments from older PowerPoint file formats (.ppt).
- The script assumes comments are stored in a specific XML structure within the .pptx file.

## Contributing

Contributions, issues, and feature requests are welcome. Feel free to check [issues page](https://github.com/yourusername/your-repo-name/issues) if you want to contribute.

## License

[MIT](https://choosealicense.com/licenses/mit/)

## Author

Lisa Ross

