# Python EDI Order Automation  
### Excel-to-EDI Process Automation for ERP Integration  

This project automates the processing of Excel-based order files and generates standardized EDI transaction files for ERP integration. Built using Python, the automation eliminates repetitive manual data entry, improves transaction accuracy, and significantly reduces processing time in a manufacturing operations environment.

---

## Business Problem  

Order spreadsheets were being manually reviewed, filtered, and converted into EDI transaction files before being uploaded into an ERP system. This process was:

- Time-consuming
- Prone to manual entry errors
- Difficult to scale
- Dependent on repetitive daily processing

The goal was to fully automate the workflow while maintaining data accuracy and compliance with EDI formatting requirements.

---

## Solution  

Developed a Python automation pipeline that:

- Automatically locates the newest order spreadsheet
- Detects spreadsheet headers dynamically
- Extracts and validates order data
- Groups orders by product family
- Parses order numbers and line-item references
- Generates standardized EDI transaction files
- Creates date-based output folders automatically
- Exports files ready for ERP upload

---

## Key Features  

### Automated File Detection  
Scans an input directory and automatically selects the newest valid Excel order file.

### Dynamic Header Recognition  
Identifies required spreadsheet headers automatically, making the script resilient to formatting changes.

### Data Validation  
Validates:

- Order numbers
- Delivery dates
- Quantities
- Product classifications

Invalid records are identified before transaction generation.

### EDI Transaction Generation  
Creates structured EDI transaction files including:

- Header segments
- Shipping schedules
- Forecast quantities
- Reference numbers
- Transaction totals

### Automated Folder Management  
Generates date-based output folders for organized archival and traceability.

---

## Technologies Used  

- Python
- `openpyxl`
- `csv`
- `glob`
- `re`
- `datetime`

---

## Skills Demonstrated  

- Process Automation
- Supply Chain Analytics
- ERP Integration
- EDI Transaction Processing
- File System Automation
- Data Validation
- Manufacturing Operations

---

## Project Structure  

```bash
├── input/
├── output/
├── edi_automation.py
└── README.md
```

---

## Business Impact  

This automation:

- Reduced manual processing from minutes to seconds
- Improved transaction accuracy
- Eliminated repetitive data entry
- Standardized EDI file generation
- Increased scalability for future order volume growth

---

## Author  

Developed by Eva Karakostas  
M.S. Data Science | Operations Supervisor | Python Automation
