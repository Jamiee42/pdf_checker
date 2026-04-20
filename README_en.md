<div align="center">

[![🇨🇳 中文](https://img.shields.io/badge/语言-中文-blue.svg)](README.md)

</div>

# PDF Checker

A desktop tool built with Python for quickly analyzing PDF page types and generating an Excel-based statistical report.

## Key features

- **Page analysis** - Analyze PDF files and classify each page into one of the following types: Editable / Non-editable / Blank / Unrecognizable (e.g., encrypted or corrupted).

- **Data export** - Generate an Excel report with detailed statistics, including page-level breakdowns.

- **Confidence indicators** - Flag special cases (e.g., vectorized text) with review suggestions to assist manual validation.

- **User-friendly GUI** - Clean and intuitive interface with real-time progress display.

## Download

[⬇️ Download Windows version](https://github.com/Jamiee42/pdf_checker/releases/latest)

## Quick start

1. **Launch** - Double-click `PDFChecker.exe`.

2. **Select file** - In **"1. Select PDF file"**, click **[Browse]** and choose the PDF you want to analyze.

3. **Select output location** - In **"2. Select output location"**, click **[Browse]** and choose where to save the Excel report.

4. **Start analysis** - Click **[Start analysis]** and wait for the process to complete. Do not close the application during processing.

5. **Preview results** - After completion, a summary will be displayed under **"3. Analysis results"**.

6. **Export Excel** - To save the report locally, click **[Export Excel]**.

## Screenshots

- Main interface  
![Main interface, including three sections: top - input area (PDF path and report output path selection), middle - analysis area (**[Start analysis]** button and progress bar), and bottom - output area (analysis summary and **[Export Excel]** button)](/screenshot/1.png)

- Statistical report  
![Excel report with columns including File Name, Total Pages, Editable/Non-editable/Blank/Unrecognizable page counts, and Notes (review suggestions)](/screenshot/2.png)

## Tech stack

- Python + CustomTkinter + pdfplumber + PyPDF2 + openpyxl

## Known limitations

| Limitation | Description | Recommendation |
|------------|------------|----------------|
| Large file performance | Files over 500 pages or 50MB may take longer to process | Please be patient and avoid force closing |
| Header/Footer detection | PDF Checker may misclassify pages with minimal text (e.g., cover pages) as "Blank" | Check the **"Notes"** column in Excel |
| Image content | PDF Checker cannot determine whether images require processing (e.g., photos vs text screenshots). It marks image-based pages as "Non-editable" | Manual review recommended |
| Encrypted PDFs | PDF Checker does not support decryption. It only indicates encryption | Please decrypt the file before processing |
| Layered PDFs | PDF Checker may misclassify pages with image + hidden OCR (Optical Character Recognition) text as "Editable" | Manual review recommended |
| Form-only pages | PDF Checker may misclassify pages containing only form fields as "Blank" | Manual review recommended |

## Roadmap

- [ ] Batch processing
- [ ] Form field detection
- [ ] Page range selection
- [ ] Customizable thresholds
- [ ] Password input prompt for encrypted files

## Need help?

For detailed instructions, result interpretation, and FAQs, check the [User Guide](docs/guide.md).