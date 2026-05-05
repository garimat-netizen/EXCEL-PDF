
# PDF Report Generator
## Overview
This project implements an end-to-end pipeline to transform raw Excel or CSV data into a structured PDF report.

The system:

* Extracts tables from complex Excel files
* Performs statistical and data quality analysis
* Uses an LLM to generate business insights
* Produces a formatted PDF report

It is designed to handle real-world Excel files, including:

* Multiple sheets
* Merged cells
* Irregular formatting
* Partially structured data

---

## Features

### Input Support

* `.xlsx`, `.xlsm`, `.xls`, `.csv`

### Data Extraction

* Detects formatted tables using Excel styles (borders, bold, fill, merged cells)
* Fallback detection for unformatted data blocks
* Handles multi-sheet workbooks
* Combines extracted tables into a single DataFrame

### Data Analysis

* Descriptive statistics:

  * Mean, median, standard deviation
  * Percentiles (5th, 25th, 75th, 95th)
* Missing value analysis
* Duplicate detection
* Correlation analysis (Pearson)
* Outlier detection:

  * IQR method
  * Z-score method
* Distribution analysis:

  * Skewness
  * Kurtosis

### LLM Integration

* Generates human-readable insights from computed data
* Strictly grounded in precomputed statistics
* No hallucinated values
* Dynamic section selection based on dataset characteristics

### Report Generation

* Built using ReportLab
* Structured layout with:

  * Metadata summary
  * Analytical sections
  * Narrative insights
* Automatically selects 10 most relevant sections from ~20 options

---

## Project Structure

```
preprocess_file.py   # Data loading, Excel parsing, table detection
pdfgen.py            # Analysis, LLM integration, PDF generation
requirements.txt
README.md
```

---

## Architecture

### 1. Data Ingestion

**Function:** `load_data()`
**Location:** `preprocess_file.py`

* Reads input file (Excel or CSV)
* Detects tables using:

  * Formatting-based heuristics
  * Content-based fallback detection
* Extracts tables into pandas DataFrames
* Combines all tables into a single dataset

---

### 2. Column Type Detection

**Function:** `detect_numeric_columns()`

* Uses sample-based heuristic (first few values)
* Handles:

  * Currency symbols
  * Commas
  * Percentages
* Identifies numeric vs categorical columns before analysis

---

### 3. Analysis Engine

**Function:** `analyze_data()`
**Location:** `pdfgen.py`

Performs:

* Numeric analysis (stats, distribution, outliers)
* Categorical analysis (value counts, dominance)
* Correlation detection
* Missing data summary

All outputs are stored in a structured dictionary used by the LLM.

---

### 4. LLM Integration

#### Section Selection

**Function:** `select_sections()`

* Chooses 10 relevant sections based on dataset profile
* Uses LLM with strict JSON output requirement

#### Narrative Generation

**Function:** `llm_narrative()`

* Sends only precomputed data to LLM
* Enforces:

  * No hallucination
  * No external assumptions
* Output is cleaned to extract `<think>` content only

---

### 5. Report Generation

**Function:** `build_pdf()`

* Uses ReportLab
* Generates:

  * Title page
  * Metadata summary table
  * Section-wise content
* Supports both:

  * Text sections (LLM)
  * Structured sections (non-LLM)

---

## Section System

The system maintains a catalogue of ~20 sections, including:

* Executive Summary
* Dataset Overview
* Missing Value Analysis
* Statistical Summary
* Trend Analysis
* Correlation Analysis
* Outlier Detection
* Distribution Analysis
* KPI Analysis
* Revenue Analysis
* Cost Efficiency
* Risk Assessment
* Opportunities
* Forecasting
* Strategic Recommendations
* Key Questions

### Selection Logic

* LLM selects 10 most relevant sections
* Always includes:

  * Executive Summary
  * Strategic Recommendations

---

## Non-LLM Sections

These are generated directly from data:

* Dataset Overview
* Missing Value Analysis
* Statistical Summary

All other sections use LLM-generated narratives.

---

## Requirements

Install dependencies:

```bash
pip install -r requirements.txt
```

### Dependencies

```
numpy>=1.23
pandas>=1.5
matplotlib>=3.7
scipy>=1.10
openpyxl>=3.1
xlrd>=2.0
reportlab>=4.0
requests>=2.31
python-dotenv>=1.0
```

---

## Environment Configuration

Create a `.env` file:

```
TRITON_NEMOTRON_URL=""
```

This is used by the `call_llm` function.

---

## Usage

Update file paths in `pdfgen.py`:

```python
INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "report.pdf"
```

Run the script:

```bash
python pdfgen.py
```

---

## Output

The system generates a PDF containing:

* Dataset metadata
* Selected analytical sections
* Statistical tables
* LLM-generated insights
* Business recommendations

---

## Entry Point

```python
if __name__ == "__main__":
    df = load_data(INPUT_FILE)
    report = PDFGenerationReport(df)
    report.run(OUTPUT_FILE)
```


