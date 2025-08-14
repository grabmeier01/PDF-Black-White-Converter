# PDF to Black & White Converter

[![Python 3.7+](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/downloads/)

A professional Python GUI application that converts PDF files to optimized black & white or grayscale documents with customizable settings and batch processing capabilities.

## Features

- **Dual Conversion Modes**: 
  - Crisp black & white output with adjustable threshold
  - Compact grayscale output with quality control
- **Batch Processing**: Convert multiple PDFs simultaneously
- **Precision Controls**:
  - Adjustable DPI (100-600+)
  - Custom page range selection (e.g., "1-5,8,10-12")
  - JPEG compression quality for grayscale
- **Smart Output Handling**:
  - Custom filename suffixes
  - Optional timestamp addition
  - Output directory selection
  - Overwrite options (ask, overwrite, skip)
- **Comprehensive Monitoring**:
  - Real-time progress tracking
  - Detailed logging system
  - Exportable conversion history
- **Cross-Platform**: Windows, macOS, and Linux support

## Requirements

- Python 3.7+
- Required packages:
  ```bash
  PyMuPDF (fitz)
  Pillow (PIL)
  img2pdf

# Installation
  Clone the repository:
  ```bash
  git clone https://github.com/yourusername/pdf-to-bw-converter.git
  cd pdf-to-bw-converter
