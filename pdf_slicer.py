import tkinter as tk
from tkinter import filedialog
import PyPDF2
import os


class PDFExtractor:
    def __init__(self, root):
        """Initialize the GUI window and variables."""
        self.root = root
        self.input_path = None
        self.output_path = None
        self.page_selection = tk.StringVar()

        # Set up GUI elements
        tk.Label(root, text="Input PDF:").grid(row=0, column=0, padx=5, pady=5)
        self.input_label = tk.Label(root, text="No file selected")
        self.input_label.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Select", command=self.select_input).grid(
            row=0, column=2, padx=5, pady=5
        )
        # Add label to show valid page range after input is selected
        self.page_range_label = tk.Label(root, text="", fg="blue")
        self.page_range_label.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        tk.Label(root, text="Page selection:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(root, textvariable=self.page_selection).grid(
            row=1, column=1, columnspan=2, padx=5, pady=5
        )
        # Add instruction label for page selection format
        tk.Label(root, text="Format: 1, 3-5, 7 (pages start from 1)", fg="gray").grid(
            row=1, column=3, padx=5, pady=5, sticky="w"
        )

        tk.Label(root, text="Output PDF:").grid(row=2, column=0, padx=5, pady=5)
        self.output_label = tk.Label(root, text="No file selected")
        self.output_label.grid(row=2, column=1, padx=5, pady=5)
        tk.Button(root, text="Select", command=self.select_output).grid(
            row=2, column=2, padx=5, pady=5
        )

        tk.Button(root, text="Extract", command=self.extract_pages).grid(
            row=3, column=1, pady=10
        )

        self.status_label = tk.Label(root, text="")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

    def select_input(self):
        """Open a dialog to select the input PDF file."""
        self.input_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.input_path:
            # Show only the file name, not the full path
            self.input_label.config(text=os.path.basename(self.input_path))
            # Show valid page range
            try:
                with open(self.input_path, "rb") as f:
                    reader = PyPDF2.PdfReader(f)
                    total_pages = len(reader.pages)
                self.page_range_label.config(text=f"Pages: 1 - {total_pages}")
            except Exception as e:
                self.page_range_label.config(text="Error reading PDF")
        else:
            self.input_label.config(text="No file selected")
            self.page_range_label.config(text="")

    def select_output(self):
        """Open a dialog to specify the output PDF file."""
        self.output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if self.output_path:
            # Show only the file name, not the full path
            self.output_label.config(text=os.path.basename(self.output_path))
        else:
            self.output_label.config(text="No file selected")

    def extract_pages(self):
        """Extract selected pages from the input PDF and save to the output PDF."""
        # Validate inputs
        if not self.input_path:
            self.status_label.config(text="Please select input PDF")
            return
        if not self.output_path:
            self.status_label.config(text="Please select output PDF")
            return
        selection_str = self.page_selection.get()
        if not selection_str:
            self.status_label.config(text="Please enter page selection")
            return

        try:
            # Open and process the PDF
            with open(self.input_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                pages_to_extract = self.parse_page_selection(selection_str, total_pages)
                writer = PyPDF2.PdfWriter()
                # Add each selected page to the new PDF
                for p in pages_to_extract:
                    writer.add_page(reader.pages[p])
                # Save the new PDF
                with open(self.output_path, "wb") as out_f:
                    writer.write(out_f)
            self.status_label.config(text="Done!", fg="green")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg="red")

    def parse_page_selection(self, selection_str, total_pages):
        """
        Parse the page selection string into a list of 0-based page indices.
        Example input: "1-2, 5-7, 11, 13"
        Returns: [0, 1, 4, 5, 6, 10, 12]
        """
        pages = []
        parts = selection_str.split(",")
        for part in parts:
            part = part.strip()
            if not part:  # Skip empty parts
                continue
            if "-" in part:
                start_end = part.split("-")
                if len(start_end) != 2:
                    raise ValueError(f"Invalid range: {part}")
                start, end = start_end
                start = int(start.strip())
                end = int(end.strip())
                if start > end:
                    raise ValueError(f"Invalid range: {start} > {end}")
                for p in range(start, end + 1):
                    if p < 1 or p > total_pages:
                        raise ValueError(f"Page {p} is out of range (1-{total_pages})")
                    pages.append(p - 1)  # Convert to 0-based index
            else:
                p = int(part)
                if p < 1 or p > total_pages:
                    raise ValueError(f"Page {p} is out of range (1-{total_pages})")
                pages.append(p - 1)  # Convert to 0-based index
        return pages


if __name__ == "__main__":
    root = tk.Tk()
    root.title("PDF Page Extractor")
    app = PDFExtractor(root)
    root.mainloop()
