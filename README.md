# 🎓 certificate-automation

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange)
![License](https://img.shields.io/badge/License-Educational-green)
![Status](https://img.shields.io/badge/Status-Active-success)

> 🚀 **Generate 100+ professional certificates in minutes** — just bring a template and a name list, and let the app do the rest!

Perfect for 🎓 colleges, 🏆 competitions, 🛠️ workshops, and 📋 events of any size.

---

## 📚 Table of Contents

- [✨ Overview](#-overview)
- [🌟 Features](#-features)
- [⚙️ Requirements](#️-requirements)
- [📦 Installation](#-installation)
- [▶️ Usage](#️-usage)
- [📄 Input Formats](#-input-formats)
- [🎨 Template Setup](#-template-setup)
- [🔠 Font Sizing Logic](#-font-sizing-logic)
- [📂 Project Structure](#-project-structure)
- [⚠️ Limitations](#️-limitations)
- [🔮 Roadmap](#-roadmap)
- [👨‍💻 Author](#-author)

---

## ✨ Overview

**certificate-automation** takes the pain out of certificate creation. Instead of manually editing a PowerPoint for every single participant, just:

1. 🎨 Design your template **once**
2. 📋 Provide a **list of names**
3. ⚡ Hit generate and get **polished PDFs instantly**

No command-line knowledge needed — everything runs through a clean, friendly GUI!

---

## 🌟 Features

| Feature | Description |
|---|---|
| 🖥️ **Easy GUI** | Clean Tkinter interface — no terminal required |
| 📊 **Excel support** | Import names directly from `.xlsx` files |
| 📄 **PDF support** | Also reads name lists from `.pdf` files |
| 🔠 **Smart font sizing** | Auto-adjusts text size so names always fit |
| 🔄 **PPT → PDF** | Fully automated conversion via PowerPoint |
| 📈 **Live progress** | Real-time log + progress bar while generating |
| 🧵 **Multithreaded** | Background processing — UI never freezes |
| 🧹 **Auto-cleanup** | Temp files removed after export |
| 🔍 **Deduplication** | Cleans and deduplicates names automatically |

---

## ⚙️ Requirements

Before you get started, make sure you have:

- 🪟 **Windows OS** — required for PowerPoint automation
- 🐍 **Python 3.8 or higher**
- 💼 **Microsoft PowerPoint** — must be installed for PPT → PDF conversion

### 📦 Python Dependencies

```bash
pip install pandas pdfplumber python-pptx pywin32
```

| Package | What it does |
|---|---|
| `pandas` | 📊 Reads Excel name lists |
| `pdfplumber` | 📄 Reads PDF name lists |
| `python-pptx` | 🎨 Edits the PowerPoint template |
| `pywin32` | 🔄 Automates PowerPoint for PDF export |

---

## 📦 Installation

```bash
# 1️⃣ Clone the repository
git clone https://github.com/your-username/certificate-automation.git

# 2️⃣ Move into the project folder
cd certificate-automation

# 3️⃣ Install dependencies
pip install pandas pdfplumber python-pptx pywin32
```

That's it — you're ready to go! 🎉

---

## ▶️ Usage

### Launch the app

```bash
python main.py
```

### 🖱️ Steps in the GUI

| Step | Action |
|---|---|
| 1️⃣ | Select your `.pptx` certificate template |
| 2️⃣ | Select your names file (`.xlsx` or `.pdf`) |
| 3️⃣ | Choose an output folder for the PDFs |
| 4️⃣ | Configure placeholder text, font name, and font size range |
| 5️⃣ | Click **"Load Names"** to preview and verify the list |
| 6️⃣ | Click **"Generate Certificates"** and watch the magic happen ✨ |

---

## 📄 Input Formats

### 📊 Excel (`.xlsx`)

The app looks for a column named `name` (case-insensitive). If it doesn't find one, it uses the first column automatically.

| Name |
|---|
| John Doe |
| Alice Smith |
| Michael Johnson |

### 📋 PDF (`.pdf`)

List names one per line — simple and clean:

```
John Doe
Alice Smith
Michael Johnson
```

---

## 🎨 Template Setup

Setting up your template is quick and easy:

1. 🖌️ Design your certificate in PowerPoint (`.pptx`)
2. 📝 Add a text box to the slide with **exactly** this placeholder:

```
names
```

3. 💅 Style the text box however you like — font, color, size, position. The app replaces only the text, keeping your formatting intact.

> 💡 **Font tip:** These fonts make certificates look stunning:
> - *Great Vibes* — elegant cursive
> - *Playfair Display* — classy serif
> - *Lavonia Classy* — decorative and premium
>
> Just make sure the font is **installed on your machine** before running the app!

---

## 🔠 Font Sizing Logic

Names come in all lengths — so the app automatically adjusts the font size to keep everything looking clean:

| 📏 Name Length | 🔡 Font Size Used |
|---|---|
| ≤ 12 characters | ✅ Maximum size (as configured) |
| 13 – 20 characters | 🔽 80% of maximum size |
| > 20 characters | ⬇️ Minimum size (as configured) |

You can set your preferred **max** and **min** sizes in the settings panel.

---

## 📂 Project Structure

```
certificate-automation/
├── 🐍 main.py            # App entry point — run this!
├── 🎨 template.pptx      # Your certificate template (you provide this)
├── 📊 names.xlsx         # Your names list (you provide this)
└── 📁 output/            # Generated PDFs land here
```

---

## ⚠️ Limitations

A few things to keep in mind:

- 🪟 **Windows only** — the PPT → PDF step uses `win32com`, which needs PowerPoint on Windows
- 📑 **Single-slide templates only** — multi-slide support is not yet available
- 💼 **PowerPoint required** — LibreOffice cannot be used as a substitute

---

## 🔮 Roadmap

Here's what's planned for future versions:

- [ ] 🐧 Linux / macOS support via LibreOffice
- [ ] 📑 Multi-slide template support
- [ ] 🖱️ Drag-and-drop file input
- [ ] 🖼️ PNG / JPG image export
- [ ] 📐 Custom text positioning in GUI
- [ ] 👁️ Certificate preview before bulk generation

Got an idea? Open an issue or submit a PR — contributions are always welcome! 🙌

---

## 👨‍💻 Author

Made with ❤️ by **Accel**

If this project saved you time, consider giving it a ⭐ — it means a lot!

- 🐛 Found a bug? [Open an issue](../../issues)
- 💡 Have an idea? [Start a discussion](../../discussions)
- 🍴 Want to contribute? Fork it and send a PR!

---

## 📜 License

This project is intended for **educational and personal use** only.
