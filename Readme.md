# Result Analysis Automation

This project automates the extraction of academic result data from university PDF result sheets and converts it into a clean, structured Excel file for analysis.

## Tech Stack
- Python
- FastAPI
- pdfplumber
- pandas
- openpyxl

## Features
- Upload digital result PDF
- Extract student details and subject-wise data
- Extract SGPA and credit summary
- Generate downloadable Excel file

## How to Run
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
