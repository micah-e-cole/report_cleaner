# File Cleaner

Project: Automating ETL For Messy Data

Purpose: Automatically convert raw organizational exports into a clean, standardized, analytics‑ready Excel workbook.

AI Acknowledgement: This project was conducted with the assistance of Copilot.

---

## 1. System Overview

The excel_cleaner project automates the cleaning and formatting of exported utilization reports.  
It removes header noise, page breaks, timestamp rows, and merged‑cell artifacts while producing a uniform dataset suitable for Registrar reporting and downstream tools (Power BI, Excel pivots, dashboards, etc.).

### The tool performs:

- Import of raw building information report data (.xls, .xlsx, .csv)
- Detection and forward-fill of building names  
- Extraction of room‑level utilization rows  
- Parsing/normalization of numeric and percentage values  
- Removal of repeated headers and footer clutter  
- Final export to a styled Excel sheet with:
  - Headers at row 1  
  - Filters applied  
  - Formatted number/percent columns  
  - Date stamp at **N1**  

Users interact with the tool via a simple Tkinter GUI that requires **no coding experience**.
