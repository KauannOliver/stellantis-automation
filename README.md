# **Stellantis Invoice Automation - Excel Processor**

This project was developed to automate the decision-making process for invoice generation based on transport operation data. It reads, analyzes, and categorizes financial and logistical information from structured Excel files exported by Stellantis, applying complex business rules to determine which records are ready for invoicing and which require verification.

---

## **KEY FEATURES**

### 1. **Automated Business Logic Processing**
- Applies multiple custom validation rules such as:
  - Route-based lead times (e.g., GOIANA = 17 days, BETIM = 9 days).
  - Differentiates between document types (`CT-e` vs others).
  - Adds grace periods based on description (`+30` or `+50` days if “spot” is present).
  - Only allows invoicing if deadlines are met and lot rules are respected.

### 2. **Advanced Lot Analysis**
- Validates whether all documents within a batch refer to either “IDA” or “RETORNO”.
- Blocks invoice generation if a single batch has mixed or undefined directions.
- Applies special logic for cases with complementary documents (`CT-e Complementar`).

### 3. **Excel Integration**
- Reads from a shared network file (`.xlsx`) containing operational data.
- Generates two structured outputs:
  - `A Faturar`: entries eligible for immediate invoicing.
  - `Verificar`: entries needing manual review due to validation errors.

### 4. **CSV Export for Integration**
- Automatically generates a `.csv` version of the "A Faturar" sheet.
- Ideal for system imports or further processing in ERPs or Power BI.

### 5. **Optional Excel Data Refresh**
- Option to automatically open the Excel file and refresh data connections using COM automation (via `win32com.client`).
- Ensures the latest data before processing.

---

## **TECHNOLOGIES USED**

### 1. **Python**
- Main language used for business logic, date handling, and data transformation.

### 2. **Pandas**
- Powerful library for data analysis and Excel reading/writing.

### 3. **Openpyxl**
- Used by `pandas` as the backend for writing Excel files with multiple sheets.

### 4. **Win32com**
- Optional module for Excel automation (refreshing data connections inside `.xlsx` files).

### 5. **Datetime & OS**
- Handle date calculations, file system paths, and file naming logic.

---

## **HOW IT WORKS**

1. User places the input file (`00. Base Stellantis.xlsx`) in the shared folder.
2. The script reads the sheets `queryStellantis` and `queryFaturados`.
3. Applies all rules to determine:
   - If the operation is ready to be invoiced.
   - If the lot structure is valid.
   - If the document types meet requirements.
4. Outputs two Excel files:
   - One for ready-to-invoice entries.
   - One for entries requiring verification.
5. A CSV is also generated for the invoice-ready entries.

---

## **CONCLUSION**

The **Stellantis Invoice Automation** tool transforms manual, error-prone invoice tracking into a robust, auditable, and rule-driven process. It’s designed for logistics and financial teams dealing with strict compliance and large volumes of transactional data.

By automating validation and classification, this tool reduces human error, ensures compliance with business rules, and saves hours of manual work—streamlining the invoice cycle from transport to billing.
