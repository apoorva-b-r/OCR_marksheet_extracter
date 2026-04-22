# OCR_marksheet_extracter
> Automated extraction of student marksheet data from images using **Google Cloud Vision OCR**, with structured output directly into **Google Sheets** - all powered by **Google Apps Script**.

---

## 📌 Overview

Manual data entry from marksheets is slow, tedious, and error-prone - especially at scale. This project eliminates that bottleneck by automating the entire pipeline:

- Upload marksheet images to Google Drive
- OCR processes and extracts raw text via Google Cloud Vision API
- Board-specific parsing logic interprets the extracted content
- Structured academic data is written directly into Google Sheets

Supports **CBSE** and **ICSE** board formats. No local setup, no servers - runs entirely inside Google Apps Script.

---

## 🚀 Features

| Feature | Description |
|---|---|
| 🔍 OCR Extraction | Powered by Google Cloud Vision API for accurate text recognition |
| 🏫 Multi-Board Support | Dedicated parsers for CBSE and ICSE marksheet layouts |
| 📊 Google Sheets Output | Auto-populates structured rows in your spreadsheet |
| 🔐 Secure API Key Handling | API key stored safely in Script Properties (never hardcoded) |
| ☁️ Fully Cloud-Based | Runs entirely in Google Apps Script - no local setup required |
| ⚡ Lightweight & Fast | No external dependencies or complex infrastructure |

---

## 🗂️ Project Structure

```
OCR_MARKSHEET_EXTRACTOR/
│
├── code.js                  # OCR parsing logic for CBSE and State marksheets
│
└── README.md                # Project documentation
```

---

## 📋 Extracted Information

For each marksheet image processed, the following fields are captured:

- **Student Name**
- **Board Name** (CBSE / ICSE)
- **Board Roll Number**
- **Mother's Name**
- **Father's Name**
- **Examination Year**
- **Subject Names**
- **Subject-wise Marks** (Theory + Practical)
- **Total Marks**
- **Percentage** *(calculated by us)*
- **Result** (Pass / Fail)

---

## ⚙️ Technologies Used

- [Google Apps Script](https://developers.google.com/apps-script) - Scripting environment
- [Google Cloud Vision API](https://cloud.google.com/vision) - OCR engine
- [Google Sheets API](https://developers.google.com/sheets/api) - Output destination
- [Google Drive API](https://developers.google.com/drive) - Image source

---

## 🔄 Workflow

```
┌──────────────────────┐
│  Upload marksheet    │
│  images to Drive     │
└────────┬─────────────┘
         │
         ▼
┌──────────────────────┐
│  Add image IDs/URLs  │
│  to Google Sheet     │
└────────┬─────────────┘
         │
         ▼
┌──────────────────────┐
│  Apps Script sends   │
│  image to Vision API │
└────────┬─────────────┘
         │
         ▼
┌──────────────────────┐
│  Raw OCR text is     │
│  extracted           │
└────────┬─────────────┘
         │
         ▼
┌──────────────────────┐
│  Board-specific      │
│  parser processes    │
│  the text            │
└────────┬─────────────┘
         │
         ▼
┌──────────────────────┐
│  Structured data     │
│  written to Sheet    │
└──────────────────────┘
```

---

## 🧩 Prerequisites

Before running the project, ensure you have the following ready:

1. A **Google Account** with access to Google Drive, Sheets, and Apps Script
2. A **Google Cloud Project** with billing enabled
3. **Google Cloud Vision API** enabled in your project
4. A valid **Google Cloud Vision API Key** generated
5. A **Google Sheet** prepared for input/output data

---

## 🔧 Setup Instructions

### Step 1 - Prepare Your Google Sheet

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Add a column for image IDs or public Drive URLs

### Step 2 - Open Apps Script

1. In the spreadsheet, go to **Extensions → Apps Script**
2. This opens the Apps Script IDE

### Step 3 - Create Script Files

Inside the Apps Script IDE:

1. Click **+ (Add a file)** and create a new Script file named `CBSE`
2. Paste the contents of `CBSE.js` into this file
3. Repeat - create another file named `ICSE`
4. Paste the contents of `ICSE.js` into that file
5. Click **Save**

### Step 4 - Add Your API Key Securely

1. In the Apps Script IDE, go to **Project Settings** (⚙️ gear icon)
2. Scroll to **Script Properties**
3. Click **Add script property**
4. Set the following:

| Property | Value |
|---|---|
| `API_KEY` | `YOUR_GOOGLE_CLOUD_VISION_API_KEY` |

5. Click **Save script properties**

> ⚠️ **Never hardcode your API key directly in the script.** Always use Script Properties for security.

---

## ▶️ How to Use

1. **Upload** your marksheet images (CBSE or ICSE) to **Google Drive**
2. **Copy** the image file ID or shareable URL from Google Drive
3. **Paste** the image ID or URL into your input Google Sheet
4. Open **Apps Script** and run the extraction function:
   - Click **Run** ▶️ on the relevant function (`extractCBSE` or `extractICSE`)
5. On first run, **grant the required permissions** when prompted (Drive, Sheets, external requests)
6. The extracted data will automatically **populate the output columns** in your Sheet

---

## 🔑 Getting Your Google Cloud Vision API Key

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or select an existing one)
3. Navigate to **APIs & Services → Library**
4. Search for **"Cloud Vision API"** and click **Enable**
5. Go to **APIs & Services → Credentials**
6. Click **Create Credentials → API Key**
7. Copy the generated key and add it to Script Properties as described in Setup Step 4

> 💡 **Tip:** Restrict your API key to the Cloud Vision API only in the Credentials settings for better security.

---

## 📸 Sample Marksheet

The project was tested with real CBSE Secondary School Examination marksheets (Class 10), which include:

- Student details (name, roll number, school, date of birth)
- Subject-wise marks in theory and internal assessment
- Positional grades per subject
- Final result (Pass/Fail)

---

## ⚠️ Limitations & Notes

- Image quality directly affects OCR accuracy. Use clear, well-lit scans or photographs.
- Heavily skewed or rotated images may reduce parsing accuracy.
- The parsing logic is tuned for standard CBSE and ICSE layouts; custom or older formats may need adjustments.
- Google Cloud Vision API usage is subject to [Google's pricing and quotas](https://cloud.google.com/vision/pricing).

---

## 📄 License

This project is intended for **educational and administrative use** only. Please ensure compliance with your institution's data privacy policies when handling student information.

---

## 👤 Author

Built for automating academic data extraction in educational environments.  
For questions or feedback, feel free to open an issue on the repository.
