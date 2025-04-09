# Text Searcher From Files

A Python utility that performs text searches across multiple file types in specified directories.

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/Tertho1/Text-Searcher-From-Files.git
cd Text-Searcher-From-Files
```

2. Install required packages:

```bash
pip install -r requirements.txt
```

## Features

- Search through multiple file types:
  - PDF (.pdf)
  - Text files (.txt)
  - Word documents (.docx)
  - PowerPoint presentations (.pptx)
  - Excel spreadsheets (.xlsx)
- Support for system-wide or specific directory searches
- File extension filtering
- Automatic encoding detection for text files
- Logging system for successful and unsuccessful operations
- Cross-platform support (Windows and Linux)

## Supported File Types

| File Type      | Extension | Notes                             |
| -------------- | --------- | --------------------------------- |
| PDF files      | .pdf      | Uses PyMuPDF for text extraction  |
| Text files     | .txt      | Supports multiple encodings       |
| Word documents | .docx     | Microsoft Word 2007+ format       |
| PowerPoint     | .pptx     | Microsoft PowerPoint 2007+ format |
| Excel          | .xlsx     | Microsoft Excel 2007+ format      |

## Requirements

```
fitz (PyMuPDF)
python-docx
python-pptx
openpyxl
chardet
```

## Usage

1. Run the script:

```bash
python main.py
```

Or

```bash
python gui-main.py
```

2. Enter your search criteria when prompted:

   - Text to search for
   - Starting path(s) (comma-separated, or leave blank for system-wide search)
   - File extension filter (optional)

3. The script will display matching files and create two log files:
   - `success.log`: Successfully processed files
   - `unsuccessful.log`: Files that encountered errors during processing

## Example

```bash
Enter the text to search: Hello World
Enter starting path: C:/Documents, D:/Projects
Filter by file extension: .txt
```

## Logging

The script maintains two log files:

- `success.log`: Records successfully processed files
- `unsuccessful.log`: Records files that encountered errors during processing

Both logs are automatically cleared at the start of each new search operation.

## Error Handling

The script includes robust error handling:

- Corrupted files are skipped and logged
- Inaccessible directories are ignored
- Invalid file encodings are handled gracefully

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
