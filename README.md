# certificate-automation

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange)
![License](https://img.shields.io/badge/License-Educational-green)
![Status](https://img.shields.io/badge/Status-Active-success)

A Python desktop application that bulk-generates professional PDF certificates from a PowerPoint template and a list of names — in minutes.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Input Formats](#input-formats)
- [Template Setup](#template-setup)
- [Font Sizing Logic](#font-sizing-logic)
- [Project Structure](#project-structure)
- [Limitations](#limitations)
- [Roadmap](#roadmap)
- [Author](#author)

---

## Overview

certificate-automation automates the repetitive task of creating individual certificates for events, workshops, competitions, and academic programs. You provide a `.pptx` template and a list of names — the app handles the rest, replacing placeholder text, adjusting font sizes, and exporting polished PDFs.

---

## Features

- **GUI-based** — clean Tkinter interface, no command-line needed
- **Flexible input** — accepts name lists from Excel (`.xlsx`) or PDF files
- **Smart font sizing** — automatically scales text based on name length
- **PPT → PDF conversion** — fully automated via Microsoft PowerPoint
- **Live progress tracking** — real-time log console and progress bar
- **Non-blocking** — runs generation in a background thread, keeping the UI responsive
- **Auto-cleanup** — removes temporary `.pptx` files after PDF export
- **Duplicate removal** — deduplicates and cleans name lists before processing

---

## Requirements

- **OS:** Windows only
- **Python:** 3.8 or higher
- **Microsoft PowerPoint** must be installed (required for PPT → PDF conversion)

### Python Dependencies

```bash
pip install pandas pdfplumber python-pptx pywin32
```

| Package | Purpose |
|---|---|
| `pandas` | Reading Excel name lists |
| `pdfplumber` | Reading PDF name lists |
| `python-pptx` | Editing the PowerPoint template |
| `pywin32` | Automating PowerPoint for PDF export |

---

## Installation

```bash
git clone https://github.com/your-username/certificate-automation.git
cd certificate-automation
pip install pandas pdfplumber python-pptx pywin32
```

---

## Usage

```bash
python main.py
```

Once the GUI opens, follow these steps:

1. **Select template** — choose your `.pptx` certificate template
2. **Select names file** — choose an `.xlsx` or `.pdf` file containing names
3. **Select output folder** — where the generated PDFs will be saved
4. **Configure settings** — set the placeholder text, font name, and font size range
5. **Click "Load Names"** — preview and verify the extracted names
6. **Click "Generate Certificates"** — sit back while the app does the work

---

## Input Formats

**Excel (`.xlsx`)**

The app auto-detects a column named `name` (case-insensitive). If no such column exists, it falls back to the first column.

| Name |
|---|
| John Doe |
| Alice Smith |

**PDF (`.pdf`)**

Names should be listed one per line:

```
John Doe
Alice Smith
Michael Johnson
```

---

## Template Setup

1. Create your certificate design in PowerPoint (`.pptx`)
2. On the slide, add a text box containing exactly:

```
names
```

3. Style that text box with your desired font, color, and position — the app will inherit those styles and only replace the text content

> **Tip:** Fonts like *Great Vibes*, *Playfair Display*, or *Lavonia Classy* give certificates a premium look. Make sure the font is installed on the machine running the app.

---

## Font Sizing Logic

Font size is automatically adjusted based on name length to prevent overflow:

| Name Length | Font Size Applied |
|---|---|
| ≤ 12 characters | Maximum size (as configured) |
| 13 – 20 characters | 80% of maximum size |
| > 20 characters | Minimum size (as configured) |

You can set the max and min font sizes in the GUI settings panel.

---

## Project Structure

```
certificate-automation/
├── main.py            # Application entry point
├── template.pptx      # Your certificate template (user-provided)
├── names.xlsx         # Input names file (user-provided)
└── output/            # Generated PDF certificates
```

---

## Limitations

- **Windows only** — PDF conversion relies on `win32com` and requires Microsoft PowerPoint
- **Single-slide templates only** — multi-slide templates are not currently supported
- **PowerPoint must be installed** — LibreOffice is not supported as a substitute

---

## Roadmap

- [ ] Linux / macOS support (LibreOffice-based conversion)
- [ ] Multi-slide template support
- [ ] Drag-and-drop file input
- [ ] PNG / JPG image export
- [ ] Custom text positioning via GUI
- [ ] Preview pane before generation

---

## Author

**Accel**

Contributions, issues, and pull requests are welcome. If you find this project useful, consider giving it a ⭐.

---

## License

This project is intended for educational and personal use only.
