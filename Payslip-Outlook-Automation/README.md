# Payslip Outlook Automation

A Python-based automation tool that prepares employee payslip emails through Microsoft Outlook.

The automation reads site and email information from an Excel file, identifies the corresponding payslip PDFs, groups them by recipient, creates a ZIP attachment, and prepares an Outlook email for review.

## Features

- Reads site mapping information from Excel
- Supports recipient email and CC mapping
- Skips records marked as `HOLD`
- Finds payslip PDFs using site codes
- Groups multiple payslips by recipient
- Creates ZIP files for grouped payslips
- Generates an email body with the relevant site names
- Adds CC recipients automatically
- Creates Outlook email drafts for manual review

## Workflow

```text
Excel Mapping
      ↓
Read Site & Email Details
      ↓
Skip HOLD Records
      ↓
Find Matching Payslip PDFs
      ↓
Group Payslips by Email
      ↓
Create ZIP Attachment
      ↓
Create Outlook Email
      ↓
Manual Review
