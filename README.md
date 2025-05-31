# 🧾 PDF Text Extractor with OCR (Python)

This script allows you to extract text from PDF documents using either direct text extraction or Optical Character Recognition (OCR). It's especially useful for scanned PDFs or image-based documents that don't contain selectable text.

---

## 📌 Features

- Attempts **fast text extraction** using PyMuPDF
- Falls back to **OCR using Tesseract** if needed
- Converts pages to images at 2x zoom for better OCR accuracy
- Saves output as `.docx` (Word) or `.txt` (plain text)
- Offers a user-friendly command-line interface
- Prints progress with helpful messages per page

---

## 🛠️ Requirements

Before running the script, ensure you install the following:

### ✅ Install Tesseract OCR

1. Go to: [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)
2. Download and install the latest version (e.g., `tesseract-ocr-w64-setup-5.3.4.20240606.exe`)

> 💡 After installation, the script attempts to auto-detect Tesseract on Windows. No need to configure paths manually unless necessary.

---

### ✅ Install Python Dependencies

Use `pip` to install required libraries:

```bash
pip install PyMuPDF pillow pytesseract python-docx
```

---

## 🚀 How to Use

### 📁 Step 1: Run the Script

You can run the script via command line:

```bash
python pdf_text_extractor.py path/to/your/file.pdf
```

If no file path is provided as an argument, you will be prompted to input it.

### 💾 Step 2: Choose Output Format

After extraction, choose your preferred format:
- Word document (`.docx`)
- Plain text (`.txt`)
- Both

---

## 🧠 How It Works

1. **Simple extraction** is attempted using PyMuPDF's native `get_text()` method.
2. If no text is found:
   - The script falls back to **OCR** using Tesseract.
   - Each page is rendered as a high-resolution PNG and passed through Tesseract for text recognition.
3. You are prompted to save the results in the desired format.

---

## 📂 Output Files

Output files will be saved in the same directory as your input PDF:

- `yourfile_extracted.docx`
- `yourfile_extracted.txt`

---

## ❗ Notes

- OCR can be slower and depends on the quality of the scanned text.
- For best results, make sure your PDF pages contain high-resolution scans.

---

## 📃 License

MIT License

## 🙏 Acknowledgments
Made with ❤️ by Denis (BeforeMyCompileFails) — 2025
