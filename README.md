# SSAS / Power BI Semantic Model → Excel Documentation (Python)

This project provides a Python script that **automatically extracts metadata from an SSAS / Power BI Tabular model JSON (.bim)** and exports it into **structured Excel files**.

It is designed to help BI and data teams **document, audit, and govern semantic models** without manual effort.

---

## 🚀 Features

The script extracts and exports:

- ✅ Tables, columns, and measures  
- ✅ Calculated tables and calculated columns  
- ✅ Full DAX expressions  
- ✅ Model relationships  
- ✅ Translations (captions) for tables, columns, and measures  

All outputs are generated as **separate Excel files** for easy review and sharing.

---

## 📂 Output Files

After running the script, the following Excel files are created:

| File | Description |
|-----|------------|
| `table_data.xlsx` | Tables, columns, measures, calculation type, and DAX expressions |
| `relationships.xlsx` | Model relationships (from table/column → to table/column) |
| `table_translations.xlsx` | Table captions / translations |
| `column_translations.xlsx` | Column captions / translations |
| `measure_translations.xlsx` | Measure captions / translations |

---

## 🛠 Prerequisites

- Python **3.8+**
- Required libraries:
  ```bash
  pip install pandas
