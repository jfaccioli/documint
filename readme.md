# 📄 DocuMint – Bulk Word Document Generator from Excel Data

**DocuMint** is a simple yet powerful Flask web app that automates the creation of personalized Word documents using a single template and an Excel spreadsheet.

[🟢 Live Demo – Try DocuMint Now](https://documint.onrender.com/)

---

### 🔥 Why DocuMint?

Microsoft Word's built-in **Mail Merge** lets you personalise letters and contracts — but it **only merges into a single `.docx` file with multiple pages**, one per recipient. This means **you still need to split the file manually** if you want individual documents.

**DocuMint solves that.**  
It automates the entire process — taking a Word template and an Excel spreadsheet, and generating a **separate document for each row**, ready to send, sign, or upload.

---


## 🚀 Features

- 📝 Upload a Word `.docx` template with placeholders or merge fields
- 📊 Upload an Excel file with matching column headers
- 🔁 Choose a column to name the generated documents
- 📂 Merge fields into both **paragraphs** and **tables**
- 🗂️ Download a ZIP file with all generated documents
- 🧹 Auto-cleans temporary files after download

---

## 📁 Project Structure

documint/
├── main.py # Flask app
├── requirements.txt # Python dependencies
├── .gitignore # Git ignored files/folders
├── README.md # Project description
├── templates/ # HTML templates
│ ├── index.html
│ └── choose_column.html
├── static/ # Static assets (e.g. logo)
│ └── generated-icon.png
├── uploads/ # Temporary folder for uploaded files
├── output/ # Temporary folder for processed files

---

## ⚙️ Requirements

- Python 3.8+
- `pip install -r requirements.txt` (includes Flask, pandas, python-docx, openpyxl)

---

## ▶️ How to Use

### 1. Clone the repository
```bash
git clone https://github.com/yourusername/documint.git
cd documint

2. Install dependencies
bash
Copy
Edit
pip install -r requirements.txt

3. Run the app
bash
Copy
Edit
python main.py

4. Open your browser
Go to: http://localhost:5000


🧠 How Placeholders Work
DocuMint supports two methods for embedding merge data into your .docx template:

✅ Option 1: Custom Placeholders (Default)
Type placeholders manually in your Word document using «..._...». Replace spaces with underscores to match Excel column names.

Example in Word:
Dear «First_Name» «Last_Name»,
Your plan starts on «Start_Date».

Matching Excel headers:
First Name | Last Name | Start Date

✅ Option 2: Microsoft Word “Insert Merge Field” (Advanced)
For users familiar with Microsoft Word’s Mail Merge:

Open your template in Word.

Go to Mailings > Select Recipients > Use an Existing List…

Load your Excel file.

Use Mailings > Insert Merge Field to place data placeholders.

Save the template .docx file.

DocuMint detects and replaces both typed placeholders and Word merge fields within paragraphs or tables.


💼 Example Use Cases
NDIS SIL and SDA contracts
Letters of offer or employment
Bulk tenancy agreements (e.g. Form 1AA)
Certificates or registration confirmations
School or healthcare document mail-outs

🔐 Security Notes
Uploaded and generated files are stored only temporarily.
All files are deleted immediately after ZIP download.
