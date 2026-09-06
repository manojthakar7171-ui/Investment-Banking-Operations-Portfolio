# Manoj Thakar | Investment Banking Operations Portfolio

Welcome to my technical portfolio. This repository contains a series of automated middle-office and treasury engines built using Advanced Excel and VBA. These projects demonstrate my ability to manage high-volume institutional data, enforce strict trade validation logic, and mitigate financial settlement risk.

## 📁 Portfolio Projects

### 1. Corporate Action Settlement & Tax Engine
*   **Focus:** Ex-Date Logic, Fractional Bonus Shares, TDS Withholding
*   **Tech:** VBA String Parsing (`Split` functions), PivotCaches
*   **Summary:** An end-to-end settlement tool that extracts unstructured market feeds to calculate client dividend eligibility, bonus share ratios, and net cash payouts.

### 2. Trade Reconciliation & Exception Management System
*   **Focus:** Front-to-Back Matching, Exception Handling, Settlement Risk
*   **Tech:** Multi-Pass VBA Memory Loops, Data Cleansing (`Trim`, `Val`)
*   **Summary:** A dual-pass reconciliation algorithm that cross-references FO ledgers against prime broker statements to hunt down unbooked trades and quantify capital exposure.

### 3. Trade Life Cycle (TLC) Enrichment & Validation Engine
*   **Focus:** Reference Data Management (RDM), STP Rates, Limit Breaches
*   **Tech:** VBA Automation, Conditional Formatting, Interactive Dashboards
*   **Summary:** A high-volume TLC processing engine that enriches raw FO tickets with correct Standard Settlement Instructions (SSIs) and strictly validates execution metrics prior to regulatory reporting.

---
*Feel free to download the `.xlsm` files in this repository to test the VBA macros, or read the attached PDF reports for a deep dive into the business logic and technical architecture.*
