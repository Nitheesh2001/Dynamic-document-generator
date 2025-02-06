# DOCX Template Generator

A Python script to automatically fill placeholders in a Word document (`.docx`) with values from a JSON file. This tool is useful for automating document generation where specific information needs to be dynamically inserted into a structured format.

---

## **Features**

- Reads a `.docx` template and replaces placeholders with values from a JSON file.
- Handles text formatting, including **bold**, _underline_, and font size adjustments.
- Inserts exhibits dynamically, formatting them properly within the document.
- Supports multiple placeholder formats (e.g., `[[placeholder]]`, `[[Placeholder]]`, `[[PLACEHOLDER]]`).

---

## **Installation**

1. Clone the repository:
   ```sh
   git clone https://github.com/yourusername/Dynamic-document-generator.git
   cd Dynamic-document-generator
   ```

2. Create a virtual environment (optional but recommended):
   ```sh
   python -m venv venv
   source venv/bin/activate  # On macOS/Linux
   venv\Scripts\activate     # On Windows
   ```

3. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

---

## **Usage**

Run the script to generate a document:
```sh
python generate_doc.py
```

This will take an input `.docx` template and a `data.json` file, process them, and output a new `.docx` file with all placeholders replaced.

---

## **Example Input and Output**

### **1. Example `data.json` (Input JSON)**  
Create a `data.json` file in the project directory:

```json
{
  "case_type": "H1B",
  "Petitioner": "XYZ Corp",
  "Beneficiary": "John Doe",
  "exhibits": [
    { "A": "Beneficiary's Identity, Status Documents, and Pay Statements" },
    { "B": "I-140 Approval and Visa Bulletin" },
    { "C": "Company Background Information" },
    { "D": "OOH Excerpt: 17-1521.00" },
    { "E": "O*NET Excerpt: Software Developers" },
    { "F": "Beneficiary's Credentials" },
    { "G": "Beneficiary's Resume" }
  ]
}
```

### **2. Example Output (`output.docx`)**

```
TABLE OF CONTENTS

Form I-129, Petitioner for Nonimmigrant Worker (H1B)

Petitioner: XYZ Corp
Beneficiary: John Doe

----------------------------
Exhibit A  Beneficiary's Identity, Status Documents, and Pay Statements
Exhibit B  I-140 Approval and Visa Bulletin
Exhibit C  Company Background Information
Exhibit D  OOH Excerpt: 17-1521.00
Exhibit E  O*NET Excerpt: Software Developers
Exhibit F  Beneficiary's Credentials
Exhibit G  Beneficiary's Resume
```

---

## **Project Structure**
```
Dynamic-document-generator/
├── input.docx             # Template Word file
├── output.docx            # Generated document
├── data.json              # JSON file with dynamic values
├── generate_doc.py        # Python script to process the document
├── requirements.txt       # Dependencies
├── README.md              # Project documentation
└── .gitignore             # Files to ignore in Git
```

---

## **Contributing**
Feel free to fork the repository and submit pull requests with improvements or bug fixes!

---

## **License**
This project is licensed under the MIT License. See `LICENSE` for details.

---

## **Author**
Developed by [Your Name].