import tkinter as tk
from tkinter import filedialog
import PyPDF2
import os
from pdf2image import convert_from_path
from PIL import ImageTk, Image
import threading  # Added import


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
        """Show a preview of the actual PDF pages as images based on the selection, with threading."""
        if not self.input_path:
            self.status_label.config(text="Please select input PDF")
            return
        selection_str = self.page_selection.get()
        if not selection_str:
            self.status_label.config(text="Please enter page selection")
            return

        # Create loading window
        loading_win = tk.Toplevel(self.root)
        loading_win.title("Loading Preview")
        loading_win.transient(self.root)  # Make it stay on top of root
        loading_win.grab_set()  # Make it modal
        loading_win.resizable(False, False)  # Prevent resizing

        loading_label = tk.Label(
            loading_win,
            text="Loading preview...",
            font=("Segoe UI", 12, "italic"),
            fg="#4f8cff",
            bg="#f4f6fa",
            pady=40,
            padx=40,
        )
        loading_label.pack(expand=True, fill="both")

        animation_chars = ["⢿", "⣻", "⣽", "⣾", "⣷", "⣯", "⣟", "⡿"]
        animation_idx = 0
        self._animation_after_id = None  # Store 'after' ID

        def update_loading_text():
            nonlocal animation_idx
            # Check if the widgets are still valid at the beginning of the call
            if not loading_win.winfo_exists() or not loading_label.winfo_exists():
                self._animation_after_id = None # Animation stops if widgets are gone
                return

            loading_label.config(
                text=f"Loading preview... {animation_chars[animation_idx % len(animation_chars)]}"
            )
            animation_idx += 1
            
            # Re-check loading_win's existence immediately before scheduling the next 'after'
            if loading_win.winfo_exists():
                self._animation_after_id = loading_win.after(100, update_loading_text)
            else:
                # loading_win was destroyed between the outer check and here, or by another thread/event.
                self._animation_after_id = None 

        update_loading_text()
        loading_win.update_idletasks()  # Ensure window is drawn

        def _destroy_loading_window():
            # Attempt to cancel the animation if it's scheduled and loading_win still exists.
            if self._animation_after_id and loading_win.winfo_exists():
                try:
                    loading_win.after_cancel(self._animation_after_id)
                except tk.TclError:
                    # This can happen if the ID is no longer valid (e.g., already fired and not rescheduled)
                    # or if loading_win is in a state where its Tcl command is already gone.
                    # print(f"DEBUG: TclError cancelling animation_id {self._animation_after_id} for loading_win in _destroy_loading_window")
                    pass 
            self._animation_after_id = None # Always clear the ID after attempting/assuming cancellation.

            if loading_win.winfo_exists():
                try:
                    loading_win.grab_release()
                except tk.TclError:
                    # print(f"DEBUG: TclError releasing grab for loading_win in _destroy_loading_window")
                    pass # May not have had grab, or window is problematic.
                
                try:
                    # Process any pending Tcl events for loading_win. This might help Tcl
                    # clean up its internal state regarding commands associated with loading_win.
                    loading_win.update_idletasks() 
                    loading_win.destroy() # This is the line from the traceback (pdf_slicer.py:257)
                except tk.TclError as e:
                    print(f"ERROR_DEBUG: TclError during loading_win.destroy(): {e}. This occurs after winfo_exists, after_cancel, and update_idletasks.")
            # else:
                # print(f"DEBUG: loading_win did not exist at the point of destruction in _destroy_loading_window.")

        def _process_conversion_and_display():
            try:
                with open(self.input_path, "rb") as f:
                    reader = PyPDF2.PdfReader(f)
                    total_pages = len(reader.pages)
                    pages_to_extract_indices = self.parse_page_selection(
                        selection_str, total_pages
                    )

                if not pages_to_extract_indices:
                    self.root.after(
                        0,
                        lambda: [
                            _destroy_loading_window(),
                            self.status_label.config(
                                text="No valid pages selected for preview.", fg="red"
                            ),
                        ],
                    )
                    return

                number_of_pages = len(pages_to_extract_indices)
                if number_of_pages >= 30 and number_of_pages <= 50:
                    self.root.after(
                        0,
                        lambda: self.status_label.config(
                            text="Warning: Preview may take a while for many pages."
                        ),
                    )
                elif number_of_pages > 50:
                    self.root.after(
                        0,
                        lambda: [
                            _destroy_loading_window(),
                            self.status_label.config(
                                text="Too many pages to preview (>50). Please select fewer.",
                                fg="red",
                            ),
                        ],
                    )
                    return

                page_numbers_for_pdf2image = sorted(
                    list(set(p + 1 for p in pages_to_extract_indices))
                )  # 1-based, unique, sorted

                if not page_numbers_for_pdf2image:  # Should not happen if pages_to_extract_indices is not empty
                    self.root.after(
                        0,
                        lambda: [
                            _destroy_loading_window(),
                            self.status_label.config(
                                text="No pages to render for preview.", fg="red"
                            ),
                        ],
                    )
                    return

                images_from_pdf = convert_from_path(
                    self.input_path,
                    first_page=min(page_numbers_for_pdf2image),
                    last_page=max(page_numbers_for_pdf2image),
                    fmt="ppm",
                )

                if not images_from_pdf:
                    self.root.after(
                        0,
                        lambda: [
                            _destroy_loading_window(),
                            self.status_label.config(
                                text="No images generated from PDF.", fg="red"
                            ),
                        ],
                    )
                    return

                # Map the returned images (which are for the whole range) to their 1-based page numbers
                # in that range.
                converted_range_page_numbers = list(
                    range(min(page_numbers_for_pdf2image), max(page_numbers_for_pdf2image) + 1)
                )
                page_to_img_map = {
                    num: img
                    for num, img in zip(converted_range_page_numbers, images_from_pdf)
                }

                final_images_to_display = []
                for p_num_1_based in page_numbers_for_pdf2image:  # Iterate through originally selected pages
                    img = page_to_img_map.get(p_num_1_based)
                    if img:
                        final_images_to_display.append((p_num_1_based, img))

                if not final_images_to_display:
                    self.root.after(
                        0,
                        lambda: [
                            _destroy_loading_window(),
                            self.status_label.config(
                                text="Selected pages could not be rendered.", fg="red"
                            ),
                        ],
                    )
                    return

                self.root.after(
                    0, lambda: _setup_actual_preview_window(final_images_to_display)
                )

            except Exception as e:
                self.root.after(
                    0,
                    lambda: [
                        _destroy_loading_window(),
                        self.status_label.config(
                            text=f"Error during preview: {str(e)}", fg="red"
                        ),
                    ],
                )

        def _setup_actual_preview_window(images_with_page_numbers):
            _destroy_loading_window()

            actual_preview_win = None  # Initialize
            try:
                actual_preview_win = tk.Toplevel(self.root)
                actual_preview_win.title("Preview PDF Pages")

                # Set initial size and allow resizing
                actual_preview_win.geometry("650x850")
                actual_preview_win.minsize(400, 300)

                canvas = tk.Canvas(actual_preview_win)
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar = tk.Scrollbar(
                    actual_preview_win, orient="vertical", command=canvas.yview
                )
                scrollbar.pack(side="right", fill="y")
                canvas.configure(yscrollcommand=scrollbar.set)

                frame_in_canvas = tk.Frame(canvas, bg="#e0e0e0")
                # Create the window item ONCE and store its ID
                frame_in_canvas_id = canvas.create_window(
                    (0, 0), window=frame_in_canvas, anchor="nw"
                )

                actual_preview_win._preview_photo_images = []

                # Determine correct resampling filter based on Pillow version
                try:
                    resample_filter = Image.Resampling.LANCZOS
                except AttributeError:
                    resample_filter = Image.LANCZOS  # Fallback for older Pillow

                if not images_with_page_numbers:
                    # This case should ideally be caught before calling this function,
                    # but as a safeguard:
                    self.status_label.config(text="Preview error: No images to display.", fg="red")
                    if actual_preview_win and actual_preview_win.winfo_exists():
                        actual_preview_win.destroy()
                    return

                for p_num_1_based, pil_image in images_with_page_numbers:
                    try:
                        resized_img = pil_image.resize(
                            (500, int(pil_image.height * 500 / pil_image.width)),
                            resample_filter,
                        )
                        tk_img = ImageTk.PhotoImage(resized_img)
                        actual_preview_win._preview_photo_images.append(tk_img)

                        page_label = tk.Label(
                            frame_in_canvas,
                            text=f"Page {p_num_1_based}",
                            font=("Arial", 14, "bold"),
                            bg=frame_in_canvas.cget("bg"),
                        )
                        page_label.pack(pady=(20, 5))

                        img_display_label = tk.Label(
                            frame_in_canvas,
                            image=tk_img,
                            bg=frame_in_canvas.cget("bg"),
                        )
                        img_display_label.pack(pady=(0, 10))
                    except Exception as img_err:
                        error_label = tk.Label(
                            frame_in_canvas,
                            text=f"Error loading page {p_num_1_based}: {img_err}",
                            fg="red",
                            bg=frame_in_canvas.cget("bg"),
                        )
                        error_label.pack(pady=(20, 10))
                
                # This function is called when frame_in_canvas's size changes
                def on_frame_configure(event):
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    # Optional: Make the frame width match the canvas width if you don't want horizontal scroll for narrow content
                    # canvas.itemconfig(frame_in_canvas_id, width=canvas.winfo_width()) 
                    # Default: canvas item width will be frame's natural width.
                    # canvas.itemconfig(frame_in_canvas_id, width=event.width) # event.width is frame's width

                frame_in_canvas.bind("<Configure>", on_frame_configure)

                # Call update_idletasks to ensure widgets are sized before getting bbox
                actual_preview_win.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))

                def _on_mousewheel_scroll(event):
                    if event.num == 4: # Linux scroll up
                        canvas.yview_scroll(-1, "units")
                    elif event.num == 5: # Linux scroll down
                        canvas.yview_scroll(1, "units")
                    else: # Windows/MacOS
                        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

                def _set_focus_on_enter(event): # Ensure canvas gets focus on mouse enter
                    canvas.focus_set()

                canvas.bind("<MouseWheel>", _on_mousewheel_scroll)
                canvas.bind("<Button-4>", _on_mousewheel_scroll) # For Linux scroll up
                canvas.bind("<Button-5>", _on_mousewheel_scroll) # For Linux scroll down
                canvas.bind("<Enter>", _set_focus_on_enter) # Added binding for focus on enter
                canvas.focus_set() # Set initial focus

                actual_preview_win.protocol("WM_DELETE_WINDOW", actual_preview_win.destroy)
                actual_preview_win.focus_set()

            except Exception as e_setup:
                error_msg = f"Failed to display preview: {str(e_setup)}"
                print(f"ERROR_DEBUG: Caught in e_setup of _setup_actual_preview_window: {error_msg}") 
                try:
                    # Determine parent for messagebox
                    msg_parent = self.root
                    # Check if actual_preview_win was created and still exists to be a parent
                    # However, if it's in an error state, self.root might be safer.
                    # Let's use self.root to ensure the messagebox is always displayed.
                    
                    tk.messagebox.showerror("Preview Setup Error", error_msg, parent=self.root)

                    if hasattr(self, 'status_label') and self.status_label.winfo_exists():
                        self.status_label.config(text=error_msg, fg="red")
                except Exception as final_e: # Error trying to report error
                    print(f"ERROR_DEBUG: Original error during preview setup: {error_msg}")
                    print(f"ERROR_DEBUG: Error during error reporting: {final_e}")
                
                if actual_preview_win and actual_preview_win.winfo_exists():
                    actual_preview_win.destroy()

        conversion_thread = threading.Thread(
            target=_process_conversion_and_display, daemon=True
        )
        conversion_thread.start()

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
