# Intelligent Excel Data Mapping & Reconciliation Automation

## Overview

This project automates the reconciliation, mapping, and population of financial reporting templates using multiple Excel data sources. It leverages intelligent fuzzy matching techniques to identify corresponding schemes across datasets and automatically update reporting workbooks while preserving existing formatting, styles, and workbook structure.

The solution significantly reduces manual effort involved in financial reporting, data validation, and scheme performance updates.

---

## Features

* Multi-workbook Excel integration
* Intelligent fuzzy matching using similarity scoring
* Automated scheme reconciliation
* Dynamic template population
* Preservation of workbook formatting and styles
* Automated benchmark and performance data updates
* Confidence-based record matching
* Detailed audit logging and validation support
* Financial reporting workflow automation

---

## Technologies Used

* Python
* OpenPyXL
* Difflib (SequenceMatcher)
* OS Module
* Copy Module

---

## Workflow

1. Load input, mapping, and output template workbooks.
2. Extract scheme records from source files.
3. Compare scheme names using fuzzy matching.
4. Calculate similarity scores.
5. Identify best matching records.
6. Populate reporting template automatically.
7. Preserve existing workbook formatting and layouts.
8. Update performance and benchmark data.
9. Generate audit logs for validation.
10. Save the final reconciled workbook.

---

## Core Functionalities

### Intelligent Scheme Matching

* Uses fuzzy matching to compare scheme names.
* Handles naming inconsistencies across datasets.
* Selects the highest-confidence match automatically.
* Supports scalable reconciliation workflows.

### Data Population Engine

* Updates scheme-related information dynamically.
* Fills missing fields automatically.
* Updates performance metrics across reporting templates.
* Maintains consistency across worksheets.

### Formatting Preservation

* Retains workbook structure.
* Preserves cell formatting, colors, borders, and styles.
* Maintains merged cells and report layouts.
* Prevents template corruption during updates.

### Performance Data Automation

Automatically updates:

* 1-Year Performance
* 1-Year Benchmark
* 1-Year Additional Benchmark
* 3-Year Performance
* 3-Year Benchmark
* 3-Year Additional Benchmark
* 5-Year Performance
* 5-Year Benchmark
* 5-Year Additional Benchmark

### Audit Logging

* Generates detailed matching logs.
* Records similarity scores.
* Tracks updated records.
* Supports verification and compliance workflows.

---

## Business Applications

* Mutual Fund Reporting
* Asset Management Operations
* Financial Data Reconciliation
* Investment Performance Reporting
* Benchmark Mapping
* Wealth Management Reporting
* Regulatory Reporting
* Fund Factsheet Preparation

---

## Benefits

* Reduces manual reconciliation effort
* Improves reporting accuracy
* Accelerates monthly reporting cycles
* Preserves reporting templates automatically
* Supports large-scale financial operations
* Provides complete auditability through logs

---

## Installation

```bash
pip install openpyxl
```

## Run

```bash
python main.py
```

---

## Output

The automation generates:

* Updated Excel reporting workbook
* Reconciled scheme performance data
* Preserved report formatting and layouts
* Detailed matching log file for validation
