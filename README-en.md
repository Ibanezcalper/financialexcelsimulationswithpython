# Financial Excel Simulations with Python

## Overview
This repository contains a set of Python scripts designed to automate the processing, analysis, and consolidation of financial data related to real estate properties.

## The Problem it Solves
Manual processing of large volumes of financial data in Excel is error-prone and time-consuming. This project solves these issues by:
1. Automating data ingestion from a flat file (CSV) and inserting it directly into complex financial Excel templates.
2. Mass generating individual financial models for each property or record.
3. Consolidating key financial metrics (such as IRR, NPV, ROI, among others) extracted from multiple output Excel files, grouping them into a single summary report ready for analysis.

## Technologies Used
- **Python 3.x**: Main development language.
- **openpyxl**: Reading, writing, and manipulating Excel files (`.xlsx`).
- **xlsxwriter**: Creation and advanced formatting of Excel documents.
- **pywin32 (win32com.client)**: Interface with the Excel desktop application to force the recalculation of complex formulas before data extraction.
- **tqdm**: Console progress bar visualization to monitor task progression.
- **csv / os**: Python standard libraries for data reading and file system management.

## Project Structure and Workflow

The project is designed to run in two main phases or workflows:

### 1. Individual Financial Processing
The core scripts (such as `B_NRM_REHAB.py` or `B_NRM_REHAB_SCOM.py`) read a CSV file containing a list of properties. They apply validation filters (for example, checking the conservation status) and distribute numerical data into specific sheets of a base Excel template (sheets like `CalendarioInv`, `PresupuestoCost`, `PresupuestoIng`, and `Flujo`). Once completed, the code saves an independent Excel file for each processed property.

### 2. Consolidated Summary Generator (`summary_loadbar.py`)
This script scans an output directory looking for the financial models generated in the previous phase. It uses the Windows COM interface to open each Excel file, safely recalculate all its internal formulas, and extract key financial indicators. The entire dataset is then dumped into a final consolidated summary file.

## Installation and Prerequisites

1. Clone this repository to your local environment.
2. Ensure Python 3 is installed on your system.
3. Install the required dependencies using `pip`:
   ```bash
   pip install openpyxl tqdm XlsxWriter pywin32
   ```
4. Have Microsoft Excel installed on the machine (an absolute requirement for the summary script).

## Usage Guide

1. **Configure Paths**: Open the scripts and configure the file paths (`input_file_a`, `template_file_b`, `output_folder`, `summary_file`) to point to your local directories.
2. **Adjust Variables**: Modify environment variables or parameters according to the operational needs (e.g., `filtro_estatus` or `apreciacion_values`).
3. **Execute Phase 1**: Run the main processing script from your terminal.
   ```bash
   python B_NRM_REHAB.py
   ```
4. **Execute Phase 2 (Consolidation)**: Run the summary script to extract the results.
   ```bash
   python summary_loadbar.py
   ```

*Important Note: For the execution of the summary script (`summary_loadbar.py`), it is strictly necessary to run the code in a Windows environment and ensure that all target Excel files are closed. The script controls the Excel application in the background, so open files may cause read/write blocks.*
