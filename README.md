# PDF Slicer

PDF Slicer is a small tool for extracting specific pages from a PDF file and saving them as a new PDF. It is built with Python's Tkinter for the interface, PyPDF2 for PDF manipulation, and supports visual preview of PDF pages.

## Features
- Modern and intuitive user interface
- Select an input PDF file
- Specify pages to extract using a flexible format (e.g., `1, 3-5, 7`)
- Choose an output file location for the new PDF
- Preview functionality - see the actual PDF pages before extracting

## Requirements
- Python 3.x
- PyPDF2
- pdf2image (for preview functionality)
- Pillow (PIL Fork)
- Poppler (required by pdf2image as an external dependency)

## Installation
1. Clone or download this repository.
2. Install the required packages:
   ```sh
   pip install PyPDF2 pdf2image pillow
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
   - Click **Preview** to see the selected pages before extracting.
   - Click **Extract** to create the new PDF with the selected pages.

## Page Selection Format
- Use commas to separate pages or ranges: `1, 3-5, 7`
- Ranges are inclusive (e.g., `3-5` extracts pages 3, 4, and 5)
- Pages are 1-based (the first page is 1)
- The order of pages in the output PDF will match the order specified in the input
- Repeated pages are allowed (e.g., `1, 1, 2` will include page 1 twice)

## Preview Features
- View PDF pages as they will appear in the extracted document
- Scrollable interface for navigating through multiple pages
- Mouse wheel support for easy navigation
- Loading indicator to show progress during preview generation
- Smart handling of large PDFs (warning for 30-50 pages, restriction for >50 pages)

## License
This project is provided as-is for personal use.
