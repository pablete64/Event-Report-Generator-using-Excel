# Excel Event Analyzer

## Overview
**Excel Event Analyzer** is a Python tool that processes event data from Excel files. The program uses a Tkinter GUI for file and date range selection, and then generates detailed reports based on the analysis of events. It calculates event frequencies, checks for specific conditions, and applies conditional formatting for better readability in the generated Excel report.

## Features
- **GUI for File and Date Selection**: Select files and specify date ranges via a simple user interface.
- **Event Analysis**: Automatically calculates key metrics like the number of events, durations exceeding 1 hour or 1 day, and the number of days with specific events.
- **Excel Report Generation**: Generates an Excel file with detailed sheets for each entry and a summary sheet that includes analysis results.
- **Conditional Formatting**: Applies color formatting to cells in the Excel file to highlight important results (e.g., exceeding 1 hour or 1 day).

## Requirements
- Python 3.8 or higher
- pandas
- openpyxl
- tkinter
- xlsxwriter

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/excel-event-analyzer.git
   cd excel-event-analyzer

2. Install the required dependencies:
    ```bash
   pip install -r requirements.txt

## Run the program:

    ```bash
      python3 ESTRUCTURACION_DE_DATOS.py
