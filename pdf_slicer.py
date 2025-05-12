import tkinter as tk
from tkinter import filedialog
import PyPDF2
import os
from pdf2image import convert_from_path
from PIL import ImageTk, Image


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
            row=2, column=1, padx=5, pady=5, sticky="w"
        )

        tk.Label(root, text="Output PDF:").grid(row=3, column=0, padx=5, pady=5)
        self.output_label = tk.Label(root, text="No file selected")
        self.output_label.grid(row=3, column=1, padx=5, pady=5)
        tk.Button(root, text="Select", command=self.select_output).grid(
            row=3, column=2, padx=5, pady=5
        )

        tk.Button(root, text="Extract", command=self.extract_pages, bg="green", fg="white").grid(
            row=4, column=1, pady=10, sticky="e"
        )
        # Add Preview button next to Extract
        tk.Button(root, text="Preview", command=self.preview_pages, bg="lightblue", fg="black").grid(
            row=4, column=2, pady=10, sticky="w"
        )

        self.status_label = tk.Label(root, text="")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

    def select_input(self):
        """Open a dialog to select the input PDF file."""
        self.input_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.input_path:
            # Show only the file name, not the full path
            filename = os.path.basename(self.input_path)
            if len(filename) > 20:
                filename = filename[:17] + "..."
            self.input_label.config(text=filename)
            # Show valid page range
            try:
                with open(self.input_path, "rb") as f:
                    reader = PyPDF2.PdfReader(f)
                    total_pages = len(reader.pages)
                self.page_range_label.config(text=f"Pages: 1 - {total_pages}")
            except Exception as e:
                self.page_range_label.config(text=f"Error reading PDF: {str(e)}")
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
            # Show only the file name, truncated to 20 characters if necessary
            filename = os.path.basename(self.output_path)
            if len(filename) > 20:
                filename = filename[:17] + "..."
            self.output_label.config(text=filename)
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

    def preview_pages(self):
        """Show a preview of the actual PDF pages as images based on the selection."""
        if not self.input_path:
            self.status_label.config(text="Please select input PDF")
            return
        selection_str = self.page_selection.get()
        if not selection_str:
            self.status_label.config(text="Please enter page selection")
            return
        try:
            with open(self.input_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                pages_to_extract = self.parse_page_selection(selection_str, total_pages)
            # pdf2image uses 1-based page numbers
            page_numbers = [p + 1 for p in pages_to_extract]
            images = convert_from_path(self.input_path, first_page=min(page_numbers), last_page=max(page_numbers), fmt='ppm')
            # Map page_numbers to images (pdf2image returns all pages in the range)
            page_to_img = {num: img for num, img in zip(range(min(page_numbers), max(page_numbers)+1), images)}
            preview_win = tk.Toplevel(self.root)
            preview_win.title("Preview PDF Pages")
            canvas = tk.Canvas(preview_win, width=600, height=800)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar = tk.Scrollbar(preview_win, orient="vertical", command=canvas.yview)
            scrollbar.pack(side="right", fill="y")
            canvas.configure(yscrollcommand=scrollbar.set)
            frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=frame, anchor="nw")
            self._preview_imgs = []  # Prevent garbage collection
            for p in page_numbers:
                img = page_to_img.get(p)
                if img is not None:
                    img = img.resize((500, int(img.height * 500 / img.width)), Image.LANCZOS)
                    tk_img = ImageTk.PhotoImage(img)
                    self._preview_imgs.append(tk_img)
                    label = tk.Label(frame, text=f"Page {p}", font=("Arial", 14, "bold"))
                    label.pack(pady=(20, 5))
                    img_label = tk.Label(frame, image=tk_img)
                    img_label.pack(pady=(0, 10))
            def on_configure(event):
                canvas.configure(scrollregion=canvas.bbox("all"))
            frame.bind("<Configure>", on_configure)
            tk.Button(preview_win, text="OK", command=preview_win.destroy).pack(pady=10)
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
