# Classroom Utilization Cleaner

Technical Implementation Plan & User Guide

Project: **EMS Classroom Utilization → Clean Excel Transformation Tool**  
Purpose: Automatically convert raw EMS classroom utilization CSV exports into a clean, standardized, analytics‑ready Excel workbook.

---

## 1. System Overview

The Classroom Utilization Cleaner automates the cleaning and formatting of EMS‑exported utilization reports.  
It removes header noise, page breaks, timestamp rows, and merged‑cell artifacts while producing a uniform dataset suitable for Registrar reporting and downstream tools (Power BI, Excel pivots, dashboards, etc.).

### The tool performs:

- Import of raw EMS CSV reports  
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