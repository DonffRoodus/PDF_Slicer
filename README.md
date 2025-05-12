# PDF Slicer

PDF Slicer is a simple tool for extracting specific pages from a PDF file and saving them as a new PDF. It is built with Python's Tkinter for the interface and uses PyPDF2 for PDF manipulation.

## Features
- Select an input PDF file.
- Specify pages to extract using a flexible format (e.g., `1, 3-5, 7`).
- Choose an output file location for the new PDF.

## Requirements
- Python 3.x
- PyPDF2

## Installation
1. Clone or download this repository.
2. Install the required package:
   ```sh
   pip install PyPDF2
   ```

## Usage
1. Run the script:
   ```sh
   python pdf_slicer.py
   ```
2. In the GUI:
   - Click **Select** next to "Input PDF" to choose your source file.
   - Enter the pages you want to extract in the format `1, 3-5, 7` (pages start from 1).
   - Click **Select** next to "Output PDF" to choose where to save the new file.
   - Click **Extract** to create the new PDF with the selected pages.

## Page Selection Format
- Use commas to separate pages or ranges: `1, 3-5, 7`
- Ranges are inclusive (e.g., `3-5` extracts pages 3, 4, and 5).
- Pages are 1-based (the first page is 1).

## License
This project is provided as-is for personal use.
