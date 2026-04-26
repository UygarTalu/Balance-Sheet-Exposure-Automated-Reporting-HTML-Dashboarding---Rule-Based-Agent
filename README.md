# Balance-Sheet-Exposure-Automated-Reporting-HTML-Dashboarding---Rule-Based-Agent

## Overview

This repository contains a Python-based automation engine developed for Treasury Reporting to streamline the **Balance Sheet Exposure reporting process**.

The solution is designed to reduce manual work in month-end reporting by automating the consolidation and update of **revaluation master data**, integrating monthly source files, enriching account-level records with **Chart of Accounts** mappings, applying **analysis classifications**, and generating a reporting-ready Excel output.

The project supports a more controlled, traceable, and efficient reporting workflow from raw source files to final reporting output.

---

## What This Project Does

The Balance Sheet Exposure engine automates the end-to-end preparation of the Balance Sheet Exposure report.

The script is designed to:

- read the current **revaluation master data**
- integrate both **current-period** and **out-of-date** source files
- assign the correct **company** and **reporting date** to imported records
- align all imported files to the master schema
- enrich records using **Chart of Accounts** mappings
- classify accounts into **CONSIDERED** and **NOT_CONSIDERED**
- update the reporting workbook automatically
- generate a refreshed **masterdata export**
- produce process statistics for validation and control

---

## Key Features

### Master Data Update
- Reads the current revaluation master data file
- Supports integration of both:
  - outdated files that were not previously added
  - current files used for the latest reporting cycle
- Preserves the master schema while aligning new files to the same structure

### Company and Date Assignment
- Detects company labels from file names
- Adds reporting dates based on month-end logic
- Standardizes date handling for consistent downstream processing

### Chart of Accounts Enrichment
- Reads the Chart of Accounts input file
- Maps account-level records to:
  - `ACCOUNT_GROUP_DESCRIPTION`
  - `ACCOUNT_TYPE`
- Improves reporting usability and analytical categorization

### Analysis Classification
- Applies `IS_IN_ANALYSIS` logic
- Supports classification such as:
  - `CONSIDERED`
  - `NOT_CONSIDERED`
- Allows manual override for account-level review where needed

### Excel Reporting Automation
- Updates the target reporting workbook
- Populates the `REVAL_MASTER_DATA_1` structure
- Writes a refreshed Chart of Accounts sheet
- Generates a dedicated statistics sheet for process review
- Saves an updated reporting-ready workbook

### Control and Traceability
- Creates process statistics for:
  - row counts
  - company-level totals
  - date ranges
  - COA enrichment coverage
  - final master data summary
- Improves auditability and reporting confidence

---

## End-to-End Reporting Flow

The reporting workflow is designed as follows:

1. The user provides the current **revaluation master data** file.
2. The user provides the current **reporting workbook**.
3. The user provides the **Chart of Accounts** file.
4. Optional outdated files can be added to backfill missing historical master data.
5. Current-period files are added for the active reporting cycle.
6. All files are standardized and aligned to the master schema.
7. The reporting workbook is updated with enriched and classified data.
8. A refreshed **masterdata export** is produced for future use.

In short:

**Source files → Master Data Consolidation → COA Enrichment → Analysis Classification → Reporting Workbook Update → New Masterdata Export**

---

## Main Inputs

The solution typically requires the following inputs:

- current revaluation master data file
- current reporting workbook
- Chart of Accounts file
- optional outdated source files
- current-period source files

These files are typically provided as Excel workbooks.

---

## Main Outputs

The script generates the following outputs:

- an updated **Balance Sheet Exposure reporting workbook**
- an exported **masterdata file** for future reporting cycles
- a **Chart of Accounts** sheet inside the workbook
- a **data process statistics** sheet for validation and control

This makes the solution suitable for both operational reporting and structured analytical review.

---

## Reporting Logic in Business Terms

From a Treasury Reporting perspective, this solution helps standardize the preparation of the Balance Sheet Exposure report by ensuring that:

- all relevant revaluation data is consolidated into a single master structure
- missing historical files can be incorporated retrospectively
- current reporting files are added consistently
- account records are enriched with business classifications
- only relevant accounts are included in the analysis logic
- the final reporting workbook is generated in a controlled and repeatable way

This reduces dependency on manual Excel handling and improves reporting consistency across reporting periods.

---

## Repository Purpose

This repository is intended to document and store the automation logic behind the **Balance Sheet Exposure reporting workflow**.

It serves as a central place for:

- Python automation scripts
- reporting documentation
- process standardization
- future maintenance and enhancement
- controlled versioning of the reporting engine

---
