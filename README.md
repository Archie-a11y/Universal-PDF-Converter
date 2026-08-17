# 📄 Universal PDF Converter

A lightweight, multi-lingual, and highly extensible desktop document converter built with Python and Tkinter. This tool allows seamless bi-directional file format transitions between PDF and major Microsoft Office or graphic formats, featuring integrated custom themes, real-time operational logging diagnostics, and automated standalone executable compilation configurations.

---

## ✨ Features

- **🔄 Bi-Directional Formats Processing**: Convert PDF files to Word, Excel, PPTX, image folders, or plaintext; and easily convert native formats (DOCX, XLSX, PPTX, TXT, PNG/JPG) back to standard PDFs.
- **🌐 Seamless Multi-Language Support**: Complete native English and Simplified Chinese interface switching at runtime with persistent localization mappings.
- **🌓 Adaptive Theme Selection**: Features optimized Light and Dark custom visual palettes designed directly in the Tkinter frame components to prevent user eye strain.
- **⚙️ Native COM Office Integration**: Automatically automates local Microsoft Office applications via `pywin32` and `docx2pdf` fallbacks for precise document layout formatting.
- **📋 Diagnostics & Event Logger**: An integrated internal operational terminal logging display monitors background threads and reports conversion warnings, errors, and output pathways in real-time.
- **📦 Executable Deployment Ready**: Includes pre-configured PyInstaller specification blueprints to compile the source script into a standalone `.exe` with a single command.

---

## 🏗️ Tech Stack

- **GUI Core**: Python 3.10+ / Standard Tkinter / Custom ttk widget configurations
- **Data & Slide Engines**: `pandas` (Excel export tables), `python-pptx` (presentation layouts), `Pillow` (image pre-processing)
- **Document Rendering**: `PyMuPDF` (MuPDF fitz bindings), `pdf2docx` (layout recognition), `img2pdf` (image rendering), `fpdf` (text cells)
- **COM Windows Service**: `pywin32` / `pythoncom` (Microsoft Office API bridges)
- **Asynchronous Execution**: Native standard `threading` implementation for interface responsiveness during conversions
- **Compiling Framework**: PyInstaller v6.x+

---

## 🖼️ Project Screenshots & User Guide

### 📍 Step 1: Launch and Display Customization
Start the application on your computer. Use the top configuration menu to choose your preferred interface language (English/Chinese) and GUI color palette (Light/Dark).
<img width="1366" height="727" alt="image" src="https://github.com/user-attachments/assets/c5c7461d-4167-4ace-af76-ea54487961c9" />


### 📍 Step 2: Conversion Parameters Adjustment
Select your conversion direction (e.g., "PDF to" or "To PDF") using the direction radio buttons. Then, select your target document format from the interactive type list dropdown.
<img width="1366" height="728" alt="image" src="https://github.com/user-attachments/assets/8bfd0f89-9e76-45a8-8895-4d1815e25f6f" />


### 📍 Step 3: Input & Output Selection
Click the browse buttons to specify the source document path and select your destination directory. The system automatically restricts file dialog filters based on your chosen conversion type.
<img width="1366" height="727" alt="image" src="https://github.com/user-attachments/assets/aece73f0-1255-4e8c-838f-dbeaf27ca294" />


### 📍 Step 4: Dark Theme Execution & Operations Log
Switch to Dark Mode for comfortable viewing. Click the conversion button to trigger the process inside a background thread. You can monitor progress and success outputs in the real-time logging panel.
<img width="1366" height="728" alt="image" src="https://github.com/user-attachments/assets/33b448c1-9d69-4db2-843a-e59b8e82fa2d" />


---

## 🛠️ Project Setup

### Prerequisites
- **Operating System**: Windows OS (highly recommended, as `pywin32` COM automation features require a local Microsoft Office installation to process Word/Excel/PPT to PDF conversions natively).
- **Python**: version 3.10 or higher.
- **Dependencies**: Microsoft Word/Excel/Powerpoint installations, and `Poppler` (required for PDF-to-image conversions; ensure its bin directory is added to your system's PATH).

### Installation Steps

1. **Clone the Repository**
   ```bash
   git clone https://github.com/Archie-a11y/Universal-PDF-Converter.git
   cd Universal-PDF-Converter
   ```

2. **Install Dependencies**
   Install the required libraries using pip:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**
   Launch the converter interface directly using python:
   ```bash
   python pdf_converter.py
   ```

---

## 📦 Building a Standalone Executable

You can compile this Python application into a standalone Windows executable (`.exe`) that runs without a Python installation using the provided `pdf_converter.spec` configuration.

1. **Install PyInstaller**
   ```bash
   pip install pyinstaller
   ```

2. **Execute the Build Action**
   Run PyInstaller targeting the specification file:
   ```bash
   pyinstaller pdf_converter.spec
   ```

3. **Access Your Executable**
   Once compilation is complete, navigate to the newly generated `dist/` directory inside your project folder to find `pdf_converter.exe`. 
   *(Note: Remember to keep the `build/` and `dist/` folders ignored in your `.gitignore` to prevent repository bloat).*
