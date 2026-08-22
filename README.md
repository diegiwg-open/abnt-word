# 📑 ABNT Automatic Formatter for Microsoft Word (2021 / 2024 / 365)

Official extension and add-in developed for **Microsoft Word 2021**, **Word 2024**, **Word 365** and **Word on the Web** to automate, format, audit, and generate academic works (TCC, Monographs, Scientific Articles, Dissertations and Theses) in strict compliance with **ABNT** standards.

[![Download](https://img.shields.io/badge/Download-v1.0.0-blue)](https://github.com/diegiwg-open/abnt-word/archive/refs/heads/master.zip)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Word](https://img.shields.io/badge/Word-2021%2F2024%2F365-orange)](https://www.microsoft.com/word)

---

## 🚀 Quick Installation

### Method 1: Auto Installer (Recommended)

1. **Clone the project**:
   ```bash
   git clone https://github.com/diegiwg-open/abnt-word.git
   cd abnt-word
   ```

2. **Run the installer**:
   - Windows: Double-click `INSTALL.bat`
   - Or run manually: `.\ABNT.ps1`

3. **Done!** Microsoft Word will automatically open with the add-in loaded.

### Method 2: Direct Download

1. [Download the latest ZIP](https://github.com/diegiwg-open/abnt-word/archive/refs/heads/master.zip)
2. Extract the contents to a folder
3. Double-click `ABNT.bat`
4. Word will open with the ABNT add-in

### Requirements

- Microsoft Word 2021, 2024, 365, or Word on the Web
- Windows 10 or higher
- Node.js (for development)
- PowerShell (included in Windows)

---

## ⚡ How to Use

### 1. Create a New Document:
```powershell
.\ABNT.ps1
```

### 2. Open an Existing Document with the Add-in:
```powershell
.\ABNT.ps1 "path\to\your_document.docx"
```

### 3. Use in Word:
- Open Microsoft Word
- Click the **"ABNT Standards"** tab at the top
- Choose the desired functionality

*(In Windows Explorer, you can also simply drag any `.docx` file onto `ABNT.bat`!)*

---

## 🔧 Local Installation

If the auto-installer does not work, follow these steps to install locally:

### 1. Install Dependencies
```bash
npm install
```

### 2. Configure SSL Certificates
```bash
npx office-addin-dev-certs install
```

### 3. Start the Server
```bash
npm start
```

### 4. Load in Word
- Open Word
- Go to **File > Options > Add-ins > Trust Center**
- Click **Trust Center Settings**
- Under **Add-in Locations**, add the project directory
- Click **Manage Add-ins** and select **ABNT**

---

## 🛠️ Development

### Project Structure
```
abnt-word/
├── src/
│   ├── js/              # JavaScript logic
│   │   ├── app.js      # Main application
│   │   ├── wordApi.js  # Word API abstraction
│   │   ├── auditors/   # ABNT auditing
│   │   ├── formatters/ # Formatting
│   │   ├── generators/ # Content generators
│   │   └── utils/      # Utilities
│   ├── css/            # Styles
│   └── index.html      # Interface
├── server.js           # Local server
├── manifest.xml        # Add-in manifest
└── ABNT.ps1            # Startup script
```

### Available Scripts
```bash
npm start              # Start the local server
npm test               # Run tests
npm run validate       # Validate the manifest
npm run build:template # Generate templates
```

---

## 🌟 Key Features

### 🚀 1. One-Click General Formatting
- **Official Margins (NBR 14724):**
  - Top: 3.0 cm | Left: 3.0 cm
  - Bottom: 2.0 cm | Right: 2.0 cm
- **Typography & Paragraphs:**
  - Choice between **Arial** or **Times New Roman**
  - Body text at 12 pt, standardized black color
  - 1.5 line spacing
  - First-line indent of 1.25 cm (1 tab)
  - Justified alignment and 0 pt before/after spacing
  - Automatic A4 paper setup (21 x 29.7 cm)

---

### 🔍 2. Real-Time ABNT Auditor (Linter & Diagnostic Tool)
- Analyzes the entire document in real time and calculates the **ABNT Compliance Score (0 to 100%)**.
- Instantly detects:
  - Incorrect margins
  - Mixed or invalid fonts
  - Paragraphs without 1.25 cm indent
  - Incorrect spacing (e.g., single or double line spacing in body text)
  - Section titles with erroneous dots after the number (e.g., `1. INTRODUCTION` → `1 INTRODUCTION`)
  - Double spaces and excessive blank lines
- **"⚡ Fix All Errors"** button for instant autofix of all non-conformities!

---

### 📐 3. Quick Selection & Title Formatter (NBR 6024)
Apply ABNT styles with a single click on any selected paragraph:
- **📄 Body Text:** 12pt, 1.5 line spacing, 1.25cm indent, justified.
- **❝ Long Direct Quote (> 3 lines):** 4.0 cm left margin indent, 10pt, single spacing, no quotes.
- **🏷️ 1 PRIMARY SECTION:** 12pt, Bold, ALL CAPS.
- **📑 1.1 Secondary Section:** 12pt, No Bold, ALL CAPS.
- **📌 1.1.1 Tertiary Section:** 12pt, Bold, Only 1st Letter Capitalized.
- **🔹 1.1.1.1 Quaternary Section:** 12pt, Normal, Only 1st Letter Capitalized.
- **🖼️ Figure/Table Caption:** 10pt, Centered, Single.
- **📝 Source / Footnote:** 10pt, Left-aligned, Single.
- **📚 Reference Item:** 12pt, Left-aligned, Single, 6pt after paragraph.

---

### 📚 4. Smart Citations & References Generator (NBR 10520:2023 & NBR 6023)
- Intuitive form to generate references and automatic citations:
  - 📘 Books and Book Chapters
  - 📄 Journal / Scientific Articles
  - 🎓 Theses, Dissertations, and TCCs
  - 🌐 Websites and Online Articles
  - ⚖️ Legislation, Laws, and Decrees
- Automatically generates:
  - In-text citation: `(SILVA, 2024, p. 45)` and `Silva (2024, p. 45)`
  - Complete bibliographic reference formatted per NBR 6023
- **🔤 Alphabetical Organizer (A-Z):** Select your reference list in Word and sort them in strict alphabetical order with one click!

---

### 📑 5. Structured Template Insertion
Insert pre-formatted templates with one click:
- **📘 ABNT Cover Page**
- **📄 Title Page** (with nature-of-work note indented 8 cm to the right)
- **📝 Abstract & Keywords** (single paragraph of 150 to 500 words)
- **🌐 Abstract & Keywords**
- **📋 Table of Contents Model**
- **🖼️ Figure Model** (with Top and Base per IBGE/ABNT)
- **📊 Table Model**

---

### 🧹 6. Document Sanitizer
- Removes double spaces and excessive repeated spaces.
- Corrects section punctuation.
- Cleans up excessive paragraph breaks.

---

## 🧪 Automated Tests

To run the ABNT rule validation test suite:
```bash
npm test
```
Results: 13 tests covering all current standards, metric conversions, citation generators, references, and audit diagnostics with 100% pass rate.
