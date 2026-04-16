# PDF-Toolkit
A powerful, all-in-one desktop application built with Python and Tkinter for manipulating documents and extracting text. This toolkit features a user-friendly graphical interface, a visual drag-and-drop PDF organizer, and built-in OCR support for English, Marathi and Hindi.

## ✨ Features

* **Visual PDF Organizer:** A split-view interface to visually preview, drag-and-drop to reorder, and delete specific pages from a PDF.
* **JPG to PDF:** Combine one or multiple image files (`.jpg`, `.png`) into a single PDF document.
* **PDF to JPG:** Extract every page of a PDF into high-quality, individual image files.
* **Merge PDFs:** Combine multiple PDF files into a single, continuous document.
* **Word to PDF:** Convert Microsoft Word documents (`.docx`) into PDF files while maintaining exact formatting.
* **Scanned PDF to Word (OCR):** Automatically read scanned PDFs using Optical Character Recognition (OCR) and convert the extracted text into an editable Word Document (`.docx`). *Supports English and Marathi out-of-the-box.*
* **Extract Text (OCR):** Extract text directly from image files.
* **Smart File Management:** Prompts for custom file names/save locations and automatically opens the destination folder upon completion.

## 🛠️ Prerequisites

Because this application handles complex PDF rendering and OCR, it relies on a few external system tools:

1. **Python 3.8+**
2. **Microsoft Word:** Required on the host machine for the *Word to PDF* functionality to work.
3. **Tesseract OCR:** * Download and install [Tesseract for Windows](https://github.com/UB-Mannheim/tesseract/wiki).
   * Default installation path should be `C:\Program Files\Tesseract-OCR`.
   * **Marathi Support:** Download the [`mar.traineddata`](https://github.com/tesseract-ocr/tessdata/raw/main/mar.traineddata) file and place it inside the `tessdata` folder in your Tesseract directory.
4. **Poppler:**
   * Download the latest [Poppler for Windows binaries](https://github.com/oschwartz10612/poppler-windows/releases/).
   * Extract the folder. The app expects the `poppler` folder (containing the `bin` directory) to be placed in the root directory of this project.

## 📦 Installation

1. Clone this repository to your local machine:
   ```bash
   git clone [https://github.com/YourUsername/Document-OCR-Toolkit.git](https://github.com/YourUsername/Document-OCR-Toolkit.git)
   cd Document-OCR-Toolkit

2. Install the required Python dependencies:

````bash
pip install Pillow pdf2image pytesseract pypdf python-docx docx2pdf
````
3.  Ensure your project directory looks like this:

Plaintext
PDF-Toolkit/
├── poppler/                 # Extracted Poppler folder (must contain /bin)
├── Tesseract-OCR/           # Copied Tesseract folder (must contain tessdata)
├── app_icon.ico             # (Optional) Icon file for the application
└── Source_Code.py           # Main application script


Here is a complete, professional README.md file tailored specifically for the application you just built. It includes all the features, installation instructions, and the exact steps for others to build the .exe file.

You can copy this text directly into your GitHub repository!

Markdown
# 📄 PDF-Toolkit

A powerful, all-in-one desktop application built with Python and Tkinter for manipulating documents and extracting text. This toolkit features a user-friendly graphical interface, a visual drag-and-drop PDF organizer, and built-in OCR support for English and Marathi.

## ✨ Features

* **Visual PDF Organizer:** A split-view interface to visually preview, drag-and-drop to reorder, and delete specific pages from a PDF.
* **JPG to PDF:** Combine one or multiple image files (`.jpg`, `.png`) into a single PDF document.
* **PDF to JPG:** Extract every page of a PDF into high-quality, individual image files.
* **Merge PDFs:** Combine multiple PDF files into a single, continuous document.
* **Word to PDF:** Convert Microsoft Word documents (`.docx`) into PDF files while maintaining exact formatting.
* **Scanned PDF to Word (OCR):** Automatically read scanned PDFs using Optical Character Recognition (OCR) and convert the extracted text into an editable Word Document (`.docx`). *Supports English and Marathi out-of-the-box.*
* **Extract Text (OCR):** Extract text directly from image files.
* **Smart File Management:** Prompts for custom file names/save locations and automatically opens the destination folder upon completion.

## 🛠️ Prerequisites

Because this application handles complex PDF rendering and OCR, it relies on a few external system tools:

1. **Python 3.8+**
2. **Microsoft Word:** Required on the host machine for the *Word to PDF* functionality to work.
3. **Tesseract OCR:** * Download and install [Tesseract for Windows](https://github.com/UB-Mannheim/tesseract/wiki).
   * Default installation path should be `C:\Program Files\Tesseract-OCR`.
   * **Marathi Support:** Download the [`mar.traineddata`](https://github.com/tesseract-ocr/tessdata/raw/main/mar.traineddata) file and place it inside the `tessdata` folder in your Tesseract directory.
4. **Poppler:**
   * Download the latest [Poppler for Windows binaries](https://github.com/oschwartz10612/poppler-windows/releases/).
   * Extract the folder. The app expects the `poppler` folder (containing the `bin` directory) to be placed in the root directory of this project.

## 📦 Installation

1. Clone this repository to your local machine:
   ```bash
   git clone [https://github.com/YourUsername/PDF-Toolkit.git](https://github.com/AshishNagargoje/PDF-Toolkit.git)
   cd PDF-Toolkit
Install the required Python dependencies:

Bash
pip install Pillow pdf2image pytesseract pypdf python-docx docx2pdf
Ensure your project directory looks like this:

Plaintext
PDF-Toolkit/
├── poppler/                 # Extracted Poppler folder (must contain /bin)
├── Tesseract-OCR/           # Copied Tesseract folder (must contain tessdata)
├── app_icon.ico             # (Optional) Icon file for the application
└── doc_converter_gui.py     # Main application script
🚀 Usage
To run the application via Python, simply execute:

````bash
python doc_converter_gui.py
````
🏗️ Building a Standalone Executable (.exe)
You can package this application into a single .exe file that can be shared and run on other Windows machines without requiring them to install Python, Poppler, or Tesseract manually.

Install PyInstaller:

````bash
pip install pyinstaller
````
Run the following build command in your terminal from the root of the project:

````bash
python -m PyInstaller --noconsole --onefile --icon="app_icon.ico" --hidden-import docx2pdf --add-data "poppler;poppler" --add-data "Tesseract-OCR;Tesseract-OCR" Source_code.py
````
Once the build is complete, you will find your standalone application inside the newly created dist/ folder.

(Note: The host machine will still need Microsoft Word installed if they wish to use the DOCX to PDF feature).

📝 License
This project is open-source and available under the MIT License.
