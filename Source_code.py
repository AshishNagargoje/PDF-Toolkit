import sys
import os
import subprocess
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from pdf2image import convert_from_path
import pytesseract
from pypdf import PdfReader, PdfWriter
from docx import Document 
from docx.shared import Pt 
from docx2pdf import convert as convert_docx

# --- MAGIC FUNCTION FOR PYINSTALLER ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
# --------------------------------------

class DocumentToolkitGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Document & OCR Toolkit")
        # --- Add this line (replace with your exact .ico file name) ---
        self.root.iconbitmap(resource_path("favicon.ico"))
        self.root.geometry("550x560")
        self.root.resizable(False, False)

        # --- DYNAMIC PATHS FOR BUNDLED SOFTWARE ---
        self.tesseract_dir = resource_path("Tesseract-OCR")
        pytesseract.pytesseract.tesseract_cmd = os.path.join(self.tesseract_dir, "tesseract.exe")
        # Tell Tesseract exactly where the language files are:
        os.environ["TESSDATA_PREFIX"] = os.path.join(self.tesseract_dir, "tessdata")
        
        # Path to bundled Poppler
        self.poppler_path = resource_path(os.path.join("poppler", "bin"))
        # ------------------------------------------

        self.input_files = [] 
        self.output_dir = ""

        self.setup_ui()
        
        

    def setup_ui(self):
        tk.Label(self.root, text="1. Select Input File(s):", font=("Arial", 10, "bold")).pack(pady=(15, 5))
        
        file_frame = tk.Frame(self.root)
        file_frame.pack()
        
        self.file_label = tk.Label(file_frame, text="No files selected", width=40, relief="sunken", anchor="w")
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        tk.Button(file_frame, text="Browse...", command=self.browse_file).pack(side=tk.LEFT)

        tk.Label(self.root, text="2. Choose Operation:", font=("Arial", 10, "bold")).pack(pady=(15, 5))
        
        self.operation_var = tk.StringVar(value="jpg_to_pdf")
        
        ops_frame = tk.Frame(self.root)
        ops_frame.pack()
        
        tk.Radiobutton(ops_frame, text="JPG(s) to PDF", variable=self.operation_var, value="jpg_to_pdf").grid(row=0, column=0, sticky="w", padx=10)
        tk.Radiobutton(ops_frame, text="PDF to JPG", variable=self.operation_var, value="pdf_to_jpg").grid(row=0, column=1, sticky="w", padx=10)
        tk.Radiobutton(ops_frame, text="Scanned PDF to DOCX (OCR)", variable=self.operation_var, value="pdf_to_docx").grid(row=1, column=0, sticky="w", padx=10)
        tk.Radiobutton(ops_frame, text="Extract Text (OCR)", variable=self.operation_var, value="ocr").grid(row=1, column=1, sticky="w", padx=10)
        tk.Radiobutton(ops_frame, text="Word (DOCX) to PDF", variable=self.operation_var, value="docx_to_pdf").grid(row=2, column=0, sticky="w", padx=10)
        tk.Radiobutton(ops_frame, text="Merge Multiple PDFs", variable=self.operation_var, value="merge_pdfs").grid(row=2, column=1, sticky="w", padx=10)
        tk.Radiobutton(ops_frame, text="Visual PDF Organizer (1 File)", variable=self.operation_var, value="visual_organize").grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(5,0))

        tk.Button(self.root, text="PROCESS FILE(S)", command=self.process_file, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=20).pack(pady=20)

        tk.Label(self.root, text="Status / OCR Output:", font=("Arial", 10, "bold")).pack(pady=(5, 5))
        self.output_text = tk.Text(self.root, height=8, width=60)
        self.output_text.pack()

    def log_status(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update()

    def browse_file(self):
        filetypes = (
            ("All Supported Files", "*.pdf *.jpg *.jpeg *.png *.docx"),
            ("PDF Files", "*.pdf"),
            ("Word Documents", "*.docx"),
            ("Image Files", "*.jpg *.jpeg *.png"),
            ("All Files", "*.*")
        )
        filenames = filedialog.askopenfilenames(title="Select file(s)", filetypes=filetypes)
        if filenames:
            self.input_files = filenames
            if len(filenames) == 1:
                self.file_label.config(text=os.path.basename(filenames[0]))
            else:
                self.file_label.config(text=f"{len(filenames)} files selected")

    def open_file_location(self, filepath):
        try:
            if platform.system() == "Windows":
                subprocess.Popen(f'explorer /select,"{os.path.normpath(filepath)}"')
            elif platform.system() == "Darwin":
                subprocess.run(["open", "-R", filepath])
            else:
                subprocess.run(["xdg-open", os.path.dirname(filepath)])
        except Exception as e:
            self.log_status(f"Could not automatically open file location: {e}")

    def open_visual_organizer(self, pdf_path):
        try:
            reader = PdfReader(pdf_path)
            num_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read PDF: {e}")
            return

        org_window = tk.Toplevel(self.root)
        org_window.title(f"Organize: {os.path.basename(pdf_path)}")
        org_window.geometry("700x550")
        org_window.transient(self.root)
        org_window.grab_set()

        tk.Label(org_window, text="Drag & Drop list on the left. Live preview on the right.", font=("Arial", 10, "bold")).pack(pady=10)

        main_frame = tk.Frame(org_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        left_frame = tk.Frame(main_frame, width=200)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))

        scrollbar = tk.Scrollbar(left_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.page_listbox = tk.Listbox(left_frame, selectmode=tk.SINGLE, yscrollcommand=scrollbar.set, font=("Arial", 12), width=15)
        self.page_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.page_listbox.yview)

        for i in range(num_pages):
            self.page_listbox.insert(tk.END, f"Page {i + 1}")

        right_frame = tk.Frame(main_frame, relief="sunken", borderwidth=2)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.preview_label = tk.Label(right_frame, text="Select a page to preview", bg="gray20", fg="white")
        self.preview_label.pack(fill=tk.BOTH, expand=True)

        self.current_preview_image = None 

        def update_preview(event=None):
            pos = self.page_listbox.curselection()
            if not pos:
                self.preview_label.config(image='', text="No page selected")
                return
            
            self.preview_label.config(image='', text="Loading preview...")
            org_window.update() 

            item_text = self.page_listbox.get(pos[0])
            actual_page_num = int(item_text.split(" ")[1])

            try:
                # USING DYNAMIC POPPLER PATH HERE
                images = convert_from_path(pdf_path, first_page=actual_page_num, last_page=actual_page_num, size=(400, None), poppler_path=self.poppler_path)
                if images:
                    img = images[0]
                    self.current_preview_image = ImageTk.PhotoImage(img)
                    self.preview_label.config(image=self.current_preview_image, text="")
            except Exception as e:
                self.preview_label.config(image='', text=f"Preview unavailable\n{str(e)}")

        self.page_listbox.bind('<<ListboxSelect>>', update_preview)

        def on_drag_start(event):
            widget = event.widget
            widget._drag_start_index = widget.nearest(event.y)

        def on_drag_motion(event):
            widget = event.widget
            y = event.y
            if y < 20: 
                widget.yview_scroll(-1, "units")
            elif y > widget.winfo_height() - 20:
                widget.yview_scroll(1, "units")
            
            i = widget.nearest(event.y)
            if i < widget._drag_start_index:
                x = widget.get(i)
                widget.delete(i)
                widget.insert(i+1, x)
                widget._drag_start_index = i
            elif i > widget._drag_start_index:
                x = widget.get(i)
                widget.delete(i)
                widget.insert(i-1, x)
                widget._drag_start_index = i

        def on_drag_stop(event):
            self.page_listbox.selection_clear(0, tk.END)
            self.page_listbox.selection_set(event.widget._drag_start_index)
            update_preview()

        self.page_listbox.bind("<Button-1>", on_drag_start)
        self.page_listbox.bind("<B1-Motion>", on_drag_motion)
        self.page_listbox.bind("<ButtonRelease-1>", on_drag_stop)

        btn_frame = tk.Frame(org_window)
        btn_frame.pack(pady=10)

        def delete_page():
            pos = self.page_listbox.curselection()
            if not pos: return
            self.page_listbox.delete(pos[0])
            update_preview() 

        def save_organized_pdf():
            items = self.page_listbox.get(0, tk.END)
            if not items:
                messagebox.showwarning("Warning", "No pages left to save!")
                return

            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_path = filedialog.asksaveasfilename(
                title="Save Organized PDF As...",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=f"{base_name}_organized.pdf"
            )
            
            if not out_path:
                self.log_status("Save cancelled.")
                return

            new_order = [int(item.split(" ")[1]) - 1 for item in items]
            writer = PdfWriter()
            for idx in new_order:
                writer.add_page(reader.pages[idx])

            with open(out_path, "wb") as f:
                writer.write(f)
                
            self.log_status(f"Success! Organized PDF saved to:\n{out_path}")
            org_window.destroy()
            self.open_file_location(out_path)

        tk.Button(btn_frame, text="Delete Selected Page", command=delete_page, fg="red", width=20).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="SAVE NEW PDF", command=save_organized_pdf, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), width=20).grid(row=0, column=1, padx=10)

        if num_pages > 0:
            self.page_listbox.selection_set(0)
            update_preview()

    def process_file(self):
        if not self.input_files:
            messagebox.showwarning("Warning", "Please select input file(s) first.")
            return

        operation = self.operation_var.get()
        self.output_text.delete(1.0, tk.END)
        self.log_status(f"Starting {operation}...")

        try:
            if operation == "visual_organize":
                if len(self.input_files) > 1:
                    messagebox.showwarning("Warning", "Please select only ONE PDF file for the visual organizer.")
                    return
                if os.path.splitext(self.input_files[0])[1].lower() != '.pdf':
                    messagebox.showwarning("Warning", "Selected file must be a PDF.")
                    return
                
                self.open_visual_organizer(self.input_files[0])
                return

            elif operation == "docx_to_pdf":
                docx_files = [f for f in self.input_files if f.lower().endswith('.docx')]
                if not docx_files:
                    messagebox.showwarning("Warning", "Please select at least one Word (.docx) file.")
                    return
                
                for file in docx_files:
                    base_name = os.path.splitext(os.path.basename(file))[0]
                    out_path = filedialog.asksaveasfilename(
                        title=f"Save PDF for {base_name}",
                        defaultextension=".pdf",
                        filetypes=[("PDF files", "*.pdf")],
                        initialfile=f"{base_name}_converted.pdf"
                    )
                    
                    if not out_path:
                        continue
                    
                    self.log_status(f"Converting {base_name}.docx to PDF...")
                    try:
                        convert_docx(file, out_path)
                        self.log_status(f"Success! Saved to:\n{out_path}")
                        self.open_file_location(out_path)
                    except Exception as e:
                        self.log_status(f"Error: {str(e)}")
                        self.log_status("(Requires Microsoft Word to be installed on this computer)")

            elif operation == "merge_pdfs":
                pdf_files = [f for f in self.input_files if f.lower().endswith('.pdf')]
                if len(pdf_files) < 2:
                    messagebox.showwarning("Warning", "Please select at least TWO PDF files to merge.")
                    return
                
                out_path = filedialog.asksaveasfilename(
                    title="Save Merged PDF As...",
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    initialfile="Merged_Document.pdf"
                )
                
                if not out_path:
                    return
                
                writer = PdfWriter()
                for file in pdf_files:
                    self.log_status(f"Adding {os.path.basename(file)}...")
                    reader = PdfReader(file)
                    for page in reader.pages:
                        writer.add_page(page)
                
                with open(out_path, "wb") as f:
                    writer.write(f)
                
                self.log_status(f"Success! Merged {len(pdf_files)} PDFs.")
                
                if messagebox.askyesno("Organize?", "Would you like to open the visual organizer to adjust the final page order?"):
                    self.open_visual_organizer(out_path)
                else:
                    self.open_file_location(out_path)


            elif operation == "jpg_to_pdf":
                image_list = []
                for file in self.input_files:
                    ext = os.path.splitext(file)[1].lower()
                    if ext in ['.jpg', '.jpeg', '.png']:
                        img = Image.open(file).convert('RGB')
                        image_list.append(img)

                if not image_list:
                    raise ValueError("No valid image files selected.")

                out_path = filedialog.asksaveasfilename(
                    title="Save Combined PDF As...",
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    initialfile="Combined_Images.pdf"
                )
                
                if not out_path: return

                image_list[0].save(out_path, save_all=True, append_images=image_list[1:])
                self.log_status(f"Success! Saved:\n{out_path}")

                if messagebox.askyesno("Organize PDF?", "Would you like to open the visual organizer to reorder or delete pages?"):
                    self.open_visual_organizer(out_path)
                else:
                    self.open_file_location(out_path)

            elif operation == "pdf_to_jpg":
                out_dir = filedialog.askdirectory(title="Select Output Folder")
                if not out_dir: return
                
                for file in self.input_files:
                    if os.path.splitext(file)[1].lower() != '.pdf': continue
                    
                    self.log_status(f"Processing {os.path.basename(file)}...")
                    # USING DYNAMIC POPPLER PATH HERE
                    images = convert_from_path(file, poppler_path=self.poppler_path)
                    base_name = os.path.splitext(os.path.basename(file))[0]
                    
                    for i, image in enumerate(images):
                        image.save(os.path.join(out_dir, f'{base_name}_page_{i+1}.jpg'), 'JPEG')
                self.log_status(f"Success! Images saved to:\n{out_dir}")
                self.open_file_location(os.path.join(out_dir, "temp")) 

            elif operation == "pdf_to_docx":
                for file in self.input_files:
                    if os.path.splitext(file)[1].lower() != '.pdf': continue
                    
                    base_name = os.path.splitext(os.path.basename(file))[0]
                    out_path = filedialog.asksaveasfilename(
                        title=f"Save Word Document for {os.path.basename(file)}",
                        defaultextension=".docx",
                        filetypes=[("Word Documents", "*.docx")],
                        initialfile=f"{base_name}_OCR_converted.docx"
                    )
                    
                    if not out_path: continue
                    
                    self.log_status(f"Converting {os.path.basename(file)} to Word using OCR...")
                    
                    # USING DYNAMIC POPPLER PATH HERE
                    images = convert_from_path(file, poppler_path=self.poppler_path)
                    doc = Document() 
                    
                    style = doc.styles['Normal']
                    font = style.font
                    font.name = 'Arial' 
                    font.size = Pt(12)  
                    
                    for i, image in enumerate(images):
                        self.log_status(f"  -> Reading text from page {i+1}...")
                        page_text = pytesseract.image_to_string(image, lang='mar+eng')
                        doc.add_paragraph(page_text)
                        if i < len(images) - 1:
                            doc.add_page_break()

                    doc.save(out_path)
                    self.log_status(f"Success! Saved {os.path.basename(out_path)}")
                    self.open_file_location(out_path)

            elif operation == "ocr":
                for file in self.input_files:
                    if os.path.splitext(file)[1].lower() not in ['.jpg', '.jpeg', '.png']: continue
                    
                    self.log_status(f"Reading image (Marathi + English)...")
                    image = Image.open(file)
                    text = pytesseract.image_to_string(image, lang='mar+eng')
                    self.log_status(f"\n--- OCR Results for {os.path.basename(file)} ---")
                    self.log_status(text if text.strip() else "[No text found in image]")

        except Exception as e:
            self.log_status(f"ERROR: {str(e)}")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentToolkitGUI(root)
    root.mainloop()