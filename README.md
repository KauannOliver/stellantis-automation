# Invoice Eligibility Rules Processor

A Windows desktop workflow that reads logistics data from Excel, applies route and document-type rules, and exports records for invoice review.

> Public source edition. Configure credentials locally and use empty or synthetic inputs. Company and institution names identify the original integration context; this repository does not claim affiliation or endorsement.

## Workflow

1. Load the source workbook.
2. Apply lead-time, document-type, and description rules.
3. Group eligible records and export the result.

## Stack

Python, Pandas, `pywin32`, and Microsoft Excel COM automation.

## Run

Use Windows with Microsoft Excel installed. Install dependencies from `requirements.txt`, place the authorized source workbook where the application expects it, and run `python main.py`.
