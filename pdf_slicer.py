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

        # Set a modern theme background
        root.configure(bg="#f4f6fa")

        # Set up GUI elements with improved style
        label_font = ("Segoe UI", 11)
        entry_font = ("Segoe UI", 11)
        button_font = ("Segoe UI", 11, "bold")
        status_font = ("Segoe UI", 10, "italic")

        tk.Label(root, text="Input PDF:", font=label_font, bg="#f4f6fa").grid(
            row=0, column=0, padx=8, pady=8, sticky="e"
        )
        self.input_label = tk.Label(
            root, text="No file selected", font=label_font, bg="#f4f6fa", fg="#888"
        )
        self.input_label.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        tk.Button(
            root,
            text="Select",
            command=self.select_input,
            font=button_font,
            bg="#4f8cff",
            fg="white",
            activebackground="#357ae8",
            activeforeground="white",
            bd=0,
            relief="flat",
        ).grid(row=0, column=2, padx=8, pady=8, sticky="w")
        self.page_range_label = tk.Label(
            root, text="", fg="#4f8cff", font=label_font, bg="#f4f6fa"
        )
        self.page_range_label.grid(row=0, column=3, padx=8, pady=8, sticky="w")

        tk.Label(root, text="Page selection:", font=label_font, bg="#f4f6fa").grid(
            row=1, column=0, padx=8, pady=8, sticky="e"
        )
        entry = tk.Entry(
            root,
            textvariable=self.page_selection,
            font=entry_font,
            width=20,
            bd=2,
            relief="groove",
        )
        entry.grid(row=1, column=1, columnspan=2, padx=8, pady=8, sticky="w")
        # Move instruction label below the entry box, spanning columns 1-3
        tk.Label(
            root,
            text="Format: 1, 3-5, 7 (pages start from 1)",
            fg="#888",
            font=("Segoe UI", 9),
            bg="#f4f6fa",
        ).grid(row=2, column=1, columnspan=3, padx=8, pady=(0, 8), sticky="w")
        # Shift output row and below down by 1
        tk.Label(root, text="Output PDF:", font=label_font, bg="#f4f6fa").grid(
            row=3, column=0, padx=8, pady=8, sticky="e"
        )
        self.output_label = tk.Label(
            root, text="No file selected", font=label_font, bg="#f4f6fa", fg="#888"
        )
        self.output_label.grid(row=3, column=1, padx=8, pady=8, sticky="w")
        tk.Button(
            root,
            text="Select",
            command=self.select_output,
            font=button_font,
            bg="#4f8cff",
            fg="white",
            activebackground="#357ae8",
            activeforeground="white",
            bd=0,
            relief="flat",
        ).grid(row=3, column=2, padx=8, pady=8, sticky="w")
        # Button frame for better alignment
        btn_frame = tk.Frame(root, bg="#f4f6fa")
        btn_frame.grid(row=4, column=0, columnspan=4, pady=16)
        tk.Button(
            btn_frame,
            text="Extract",
            command=self.extract_pages,
            font=button_font,
            bg="#43b77a",
            fg="white",
            activebackground="#2e8b57",
            activeforeground="white",
            bd=0,
            relief="flat",
            width=12,
            height=1,
        ).pack(side="left", padx=10)
        tk.Button(
            btn_frame,
            text="Preview",
            command=self.preview_pages,
            font=button_font,
            bg="#f7c948",
            fg="#333",
            activebackground="#f4b400",
            activeforeground="#333",
            bd=0,
            relief="flat",
            width=12,
            height=1,
        ).pack(side="left", padx=10)

        self.status_label = tk.Label(
            root, text="", font=status_font, bg="#f4f6fa", fg="#d9534f"
        )
        self.status_label.grid(row=5, column=0, columnspan=4, pady=8)

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
            # Show loading indicator before starting image conversion
            preview_win = tk.Toplevel(self.root)
            preview_win.title("Preview PDF Pages")
            loading_label = tk.Label(preview_win, text="Loading preview... Please wait.", font=("Segoe UI", 12, "italic"), fg="#4f8cff", bg="#f4f6fa", pady=40, padx=40)
            loading_label.pack(expand=True, fill="both")
            preview_win.update()

            with open(self.input_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                pages_to_extract = self.parse_page_selection(selection_str, total_pages)
                number_of_pages = len(pages_to_extract)
                if number_of_pages >= 30 and number_of_pages <= 50:
                    self.status_label.config(
                    text="Warning: Preview may take a while for large PDFs."
                    )
                elif number_of_pages > 50:
                    self.status_label.config(
                        text="PDF too large to preview", fg="red"
                    )
                    preview_win.destroy()
                    return
            # pdf2image uses 1-based page numbers
            page_numbers = [p + 1 for p in pages_to_extract]
            images = convert_from_path(
                self.input_path,
                first_page=min(page_numbers),
                last_page=max(page_numbers),
                fmt="ppm",
            )
            # Remove loading window after images are loaded
            preview_win.destroy()
            # Map page_numbers to images (pdf2image returns all pages in the range)
            page_to_img = {
                num: img
                for num, img in zip(
                    range(min(page_numbers), max(page_numbers) + 1), images
                )
            }
            preview_win = tk.Toplevel(self.root)
            preview_win.title("Preview PDF Pages")
            canvas = tk.Canvas(preview_win, width=600, height=800)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar = tk.Scrollbar(
                preview_win, orient="vertical", command=canvas.yview
            )
            scrollbar.pack(side="right", fill="y")
            canvas.configure(yscrollcommand=scrollbar.set)
            frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=frame, anchor="nw")
            self._preview_imgs = []  # Prevent garbage collection
            for p in page_numbers:
                img = page_to_img.get(p)
                if img is not None:
                    img = img.resize(
                        (500, int(img.height * 500 / img.width)), Image.LANCZOS
                    )
                    tk_img = ImageTk.PhotoImage(img)
                    self._preview_imgs.append(tk_img)
                    label = tk.Label(
                        frame, text=f"Page {p}", font=("Arial", 14, "bold")
                    )
                    label.pack(pady=(20, 5))
                    img_label = tk.Label(frame, image=tk_img)
                    img_label.pack(pady=(0, 10))

            def on_configure(event):
                canvas.configure(scrollregion=canvas.bbox("all"))

            frame.bind("<Configure>", on_configure)

            # Enable mouse wheel scrolling
            def _on_mousewheel(event):
                if event.delta:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:  # For Linux systems
                    canvas.yview_scroll(-3, "units")
                elif event.num == 5:
                    canvas.yview_scroll(3, "units")

            # Windows and MacOS
            preview_win.bind_all("<MouseWheel>", _on_mousewheel)
            # Linux (X11)
            preview_win.bind_all("<Button-4>", _on_mousewheel)
            preview_win.bind_all("<Button-5>", _on_mousewheel)

            tk.Button(
                preview_win,
                text="OK",
                command=preview_win.destroy,
                font=("Segoe UI", 11, "bold"),
                bg="#4f8cff",
                fg="white",
                activebackground="#357ae8",
                activeforeground="white",
                bd=0,
                relief="flat",
                width=10,
                height=1,
            ).pack(pady=16)
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
