# Invoice Eligibility Rules Processor

A Windows desktop workflow that reads logistics data from Excel, applies route and document-type rules, and exports records for invoice review.

> Private repository: the rule set reflects a company-specific operation. Do not publish customer data or treat the output as an invoice authorization without business validation.

## Workflow

1. Load the source workbook.
2. Apply lead-time, document-type, and description rules.
3. Group eligible records and export the result.

## Stack

Python, Pandas, `pywin32`, and Microsoft Excel COM automation.

## Run

Use Windows with Microsoft Excel installed. Install dependencies from `requirements.txt`, place the authorized source workbook where the application expects it, and run `python main.py`.
