# 📋 Marathi → English Voter Name Transliteration Tool

Converts Marathi (Devanagari) voter names in Excel files to English transliteration automatically.

---

## ✨ What It Does

- Reads `.xlsx` files from the `input/` folder
- Transliterates **"मतदाराचे पूर्ण नाव"** (Voter Name) and **"नातेवाईकाचे पूर्ण नाव"** (Relative Name) into English
- Adds **2 new columns** at the end of the sheet — all original data stays untouched
- Saves translated files to the `output/` folder (prefixed with `EN_`)



## 🖥️ Setup on a New PC

### Step 1: Install Python

1. Download Python from [python.org](https://www.python.org/downloads/) (version 3.8 or higher)
2. During installation, **check the box** that says **"Add Python to PATH"** — this is important!
3. Click **Install Now** and wait for it to finish

### Step 2: Download This Project

**Option A — Using Git:**
```bash
git clone https://github.com/AshwetKini/MARATHItoEnglishBJPTranslate.git
cd MARATHItoEnglishBJPTranslate
```

**Option B — Without Git:**
1. Go to [https://github.com/AshwetKini/MARATHItoEnglishBJPTranslate](https://github.com/AshwetKini/MARATHItoEnglishBJPTranslate)
2. Click the green **"Code"** button → **"Download ZIP"**
3. Extract the ZIP to a folder on your PC

### Step 3: Install Dependencies

Open **Command Prompt** or **PowerShell** in the project folder and run:

```bash
pip install -r requirements.txt
```

> **Note:** If `pip` is not recognized, try `python -m pip install -r requirements.txt`

### Step 4: Verify Installation

Run this to confirm everything is installed:

```bash
python -c "import openpyxl; from indic_transliteration import sanscript; print('All good!')"
```

You should see: `All good!`

---

## 🚀 How to Use

### Every time you want to translate a new file:

1. **Place** your `.xlsx` file(s) into the `input/` folder
2. **Open** Command Prompt / PowerShell in the project folder
3. **Run:**
   ```bash
   python translate.py
   ```
4. **Collect** the translated file(s) from the `output/` folder

> 💡 **Tip:** You can process multiple files at once — just drop them all into `input/` before running the script.

---

## 📁 Folder Structure

```
MARATHItoEnglishBJPTranslate/
├── input/                ← Place your Excel files here
├── output/               ← Translated files appear here (prefixed with EN_)
├── translate.py          ← Main script
├── requirements.txt      ← Python dependencies
└── README.md             ← This file
```

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| `python is not recognized` | Reinstall Python and make sure to check **"Add Python to PATH"** |
| `pip is not recognized` | Use `python -m pip install -r requirements.txt` instead |
| `No .xlsx files found` | Make sure your files are in the `input/` folder and are `.xlsx` format (not `.xls`) |
| `Target columns not found` | The script looks for columns named exactly **"मतदाराचे पूर्ण नाव"** and **"नातेवाईकाचे पूर्ण नाव"** — make sure your Excel has these headers |
| Script runs but output looks wrong | Make sure the Excel file is not open in another program while running the script |

---

## 📌 Requirements

- **Python** 3.8 or higher
- **Libraries:** `openpyxl`, `indic-transliteration` (installed via `requirements.txt`)
- **OS:** Windows / Mac / Linux
- **No internet needed** — works fully offline after setup
