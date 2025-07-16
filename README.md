# PROSPR Google Sheets Automation

> **Advanced Google Apps Script Automation for PROSPR Financial Planning Template**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Key Features](#key-features)

   * [Admin Menu with Access Control](#admin-menu-with-access-control)
   * [Monthly Comparative Report: Dynamic Analysis & Output](#monthly-comparative-report-dynamic-analysis--output)
3. [Bulk Deployment Strategy (Bonus)](#bulk-deployment-strategy-bonus)
4. [Development Assumptions](#development-assumptions)
5. [Setup Instructions](#setup-instructions)
6. [License](#license)

---

## Project Overview

This repository provides a robust Google Apps Script automation suite for the **PROSPR financial planning template**. Designed for client-facing scalability, the solution automates in-depth monthly budget analysis, administrative access management, and delivers polished, actionable reports—both as formatted Google Sheet tabs and Gmail draft summaries.

Developed as part of a technical assignment for the role of Part-Time Technical Systems Manager & Automation Developer at PROSPR.

---

## Key Features

### Admin Menu with Access Control

* **Custom Menu Injection:** Seamlessly adds an "Admin" menu to the Google Sheet interface on open.
* **Secure Access Control:** Menu is password-protected (per Google account) and can be locked/unlocked dynamically.
* **Per-User Admin Codes:** Codes are stored using `PropertiesService.getUserProperties()`, ensuring every user sets and manages their own admin credentials (default is `PROSPR2025`).

  * **Initial Setup:** First use will prompt the user to enter the default code, then require them to set a new one.
  * **Code Management:** Unlocked menus allow secure admin code updates or relocking the menu at any time.
* **UX-First Design:** User prompts are clear, all logic is modular and thoroughly commented, supporting easy extensibility.

### Monthly Comparative Report: Dynamic Analysis & Output

* **Automated Parsing:** Dynamically detects and processes all categories and line items in the "Monthly Budget" tab—no hardcoded ranges required. Designed to handle blank rows, category blocks, and totals robustly.

* **Deviation Analysis:**

  * Calculates both absolute and percentage deviations for each main budget category and all underlying items.
  * Applies a configurable threshold (default: 20%) to flag significant deviations.
  * Explicitly handles missing, blank, or zero values for real-world spreadsheets.

* **Tabular Report Output (Sheet):**

  * Generates a new tab `[Month] Budget Comparison` (e.g., "May Budget Comparison").
  * Polished formatting: category highlighting, bold/underlined headers, conditional status colors (Over/Under/OK), indented item breakdowns, and visual row separation.
  * Example Table Layout:

    | Category        | Item Description | Actual     | Planned    | Deviation (\$) | Deviation (%) | Status |
    | --------------- | ---------------- | ---------- | ---------- | -------------- | ------------- | ------ |
    | Shelter         |                  | \$4,001.02 | \$5,409.94 | -1,408.92      | -26.0%        | Under  |
    |                 | Mortgage         | \$0.00     | \$700.00   | -700.00        | -100.0%       |        |
    |                 | Pool Maintenance | \$193.88   | \$700.00   | -506.12        | -72.3%        |        |
    | Food & Supplies |           | \$2,193.05 | \$3,500.00 | -1,306.95      | -37.3%        | Under       |
    |  | Grocery          | \$2,193.05 | \$3,500.00 | -1,306.95      | -37.3%        |        |
    | ...             | ...              | ...        | ...        | ...            | ...           |        |

* **Executive Summary Output (Gmail Draft):**

  * Creates a summary as a Gmail draft using `GmailApp.createDraft()`, ready to review or forward to the client.
  * Clear, actionable language with easy-to-read breakdowns. Major deviations and responsible items are highlighted.
  * **Example Output:**

    ```
    Monthly Budget Deviation Report
    Period: Jun 2025 (06/01/2025 - 06/30/2025)
    Generated on: July 15, 2025

    Person 1: Under budget by -100.0% ($0.00 vs. $20000.00)
       Key Items:
         1. Salary 1 - After Tax: $0.00 (Actual) vs $20000.00 (Planned) (Diff: -20000.00, -100.0%)
    Person 2: Over budget by 100.0% ($2000.00 vs. $0.00)
       Key Items:
         2. Draw from Business: $2000.00 (Actual) vs $0.00 (Planned) (Diff: +2000.00, 100.0%)
    ...
    ```

* **Professional Quality:** Both outputs are designed for client presentation, suitable for direct delivery to stakeholders.

* **Extensive Comments & Clean Structure:** All code is modular, maintainable, and thoroughly documented for real-world handoff.

---

## Bulk Deployment Strategy (Bonus)

For agencies or consultants managing multiple client copies:

* **Deployment via Library:** Deploy the core functionality as a protected Apps Script library, then inject lightweight setup scripts into each client Sheet (using a list of URLs).
* **Automated Linking:** Client files simply include the library; updates to core logic are instantly available across all deployments, protecting your IP and ensuring maintainability.
* **For details:** See `deploy-report.pdf`, which provides a conceptual deployment plan leveraging Google Apps Script Libraries and APIs for scalable roll-out.

---

## Development Assumptions

* **Data Mapping:** "Budget" columns are mapped to "Planned" throughout all reports and calculations.
* **Blank/Empty Handling:** Blank, empty, or non-numeric cells are interpreted as zero to prevent calculation errors and accurately flag missing/uncategorized data.
* **Consistent Structure:** Assumes that the "Monthly Budget" tab follows a standard block/category structure, but is robust to extra rows and category order changes.
* **Per-Account Security:** Admin codes are managed per Google account for decentralized security.
* **Gmail Permissions:** Script will prompt for Gmail permissions as needed for draft generation.

---

## Setup Instructions

To enable this automation in your Google Sheet:

1. **Open your Google Sheet** where you'd like to add these features.
2. **Go to Extensions > Apps Script** in the menu bar.
3. **Copy & Paste:** Replace (or add) the code in your `Code.gs` file with the code from this repository.
4. **Save the project.**
5. **Reload your Sheet:** The "Admin" menu should now appear at the top.
6. **Unlock Admin Menu:** Click `Admin > Unlock Admin Menu`. The first time, use the code `PROSPR2025`. Afterwards, you may set your own code.

---

## License

MIT License. See the [LICENSE](LICENSE) file for details.

---

**Questions? Want to discuss the implementation or deployment plan? Reach out anytime.**
