"""
Payslip Outlook Automation
---------------------------
Automates the preparation of payslip emails using Microsoft Outlook.

Features:
- Reads site/email mapping from Excel
- Skips records marked as HOLD
- Finds payslip PDFs using site codes
- Groups PDFs by recipient email
- Creates ZIP attachments
- Adds CC recipients
- Creates Outlook email drafts for manual review

Note:
This project is intended for demonstration/portfolio purposes.
Use sample or authorized data only.
"""

import os
import zipfile
import pandas as pd
import win32com.client as win32


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PDF_FOLDER = os.path.join(BASE_DIR, "sample_payslips")
EXCEL_PATH = os.path.join(
    BASE_DIR,
    "sample_data",
    "site_mapping_sample.xlsx"
)

# Change this when running the automation
PAYSLIP_MONTH = "August"
PAYSLIP_YEAR = "2026"


# ============================================================
# STEP 1: LOAD EXCEL DATA
# ============================================================

if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(
        f"Excel mapping file not found:\n{EXCEL_PATH}"
    )

df = pd.read_excel(EXCEL_PATH)

# Clean column names
df.columns = df.columns.str.strip().str.title()


# ============================================================
# VALIDATE REQUIRED COLUMNS
# ============================================================

required_columns = ["Code", "Name", "Mail"]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    raise ValueError(
        f"Missing required Excel column(s): "
        f"{', '.join(missing_columns)}"
    )


# ============================================================
# CLEAN DATA
# ============================================================

df["Code"] = df["Code"].astype(str).str.strip()
df["Name"] = df["Name"].astype(str).str.strip()
df["Mail"] = df["Mail"].fillna("").astype(str).str.strip()

if "Cc" not in df.columns:
    df["Cc"] = ""

if "Remarks" not in df.columns:
    df["Remarks"] = ""

df["Cc"] = df["Cc"].fillna("").astype(str).str.strip()
df["Remarks"] = (
    df["Remarks"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)


# ============================================================
# STEP 2: SKIP HOLD RECORDS
# ============================================================

df = df[df["Remarks"] != "HOLD"]

# Remove records without an email address
df = df[df["Mail"] != ""]


# ============================================================
# STEP 3: BUILD MAPPINGS
# ============================================================

site_to_email = df.set_index("Code")["Mail"].to_dict()
site_to_name = df.set_index("Code")["Name"].to_dict()
site_to_cc = df.set_index("Code")["Cc"].to_dict()


# ============================================================
# STEP 4: GROUP PDF FILES BY EMAIL
# ============================================================

email_to_pdfs = {}
email_to_ccs = {}

for code, email in site_to_email.items():

    filename = f"Payslip_{code}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, filename)

    if os.path.exists(pdf_path):
        email_to_pdfs.setdefault(email, []).append(
            (code, pdf_path)
        )
    else:
        print(f"PDF not found: {filename}")

    cc_ids = site_to_cc.get(code, "")

    if cc_ids:
        cc_list = [
            item.strip()
            for item in cc_ids.split(",")
            if item.strip()
        ]

        email_to_ccs.setdefault(
            email,
            set()
        ).update(cc_list)


# ============================================================
# STEP 5: CONNECT TO MICROSOFT OUTLOOK
# ============================================================

outlook = win32.Dispatch("Outlook.Application")


# ============================================================
# STEP 6: CREATE EMAIL DRAFTS
# ============================================================

for email in set(site_to_email.values()):

    attached_items = email_to_pdfs.get(email, [])

    # Skip recipient when no PDF is available
    if not attached_items:
        print(f"No payslips available for: {email}")
        continue

    pdf_site_codes = [
        code for code, _ in attached_items
    ]

    site_names = [
        site_to_name.get(
            code,
            f"Site_{code}"
        )
        for code in pdf_site_codes
    ]


    # --------------------------------------------------------
    # CREATE ZIP FILE
    # --------------------------------------------------------

    safe_email = (
        email
        .replace("@", "_")
        .replace(".", "_")
    )

    zip_name = f"Payslips_{safe_email}.zip"

    temp_zip_path = os.path.join(
        PDF_FOLDER,
        zip_name
    )

    with zipfile.ZipFile(
        temp_zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        for code, pdf_path in attached_items:

            site_name = site_to_name.get(
                code,
                f"Site_{code}"
            )

            friendly_name = f"{site_name}.pdf"

            zipf.write(
                pdf_path,
                arcname=friendly_name
            )


    # --------------------------------------------------------
    # CREATE EMAIL BODY
    # --------------------------------------------------------

    site_list = "\n".join(
        f"{index + 1}. {name}"
        for index, name
        in enumerate(sorted(site_names))
    )

    body = (
        "Dear Team,\n\n"
        "Please find attached the payslips "
        "for the following site(s):\n\n"
        f"{site_list}\n\n"
        "Regards,\n"
        "Payroll Team"
    )


    # --------------------------------------------------------
    # CREATE OUTLOOK EMAIL
    # --------------------------------------------------------

    mail = outlook.CreateItem(0)

    mail.To = email

    mail.Subject = (
        f"PAYSLIPS FOR THE MONTH OF "
        f"{PAYSLIP_MONTH.upper()} {PAYSLIP_YEAR}"
    )

    mail.Body = body

    mail.Attachments.Add(temp_zip_path)


    # --------------------------------------------------------
    # ADD CC
    # --------------------------------------------------------

    if email in email_to_ccs:

        mail.CC = ";".join(
            sorted(email_to_ccs[email])
        )


    # --------------------------------------------------------
    # DISPLAY EMAIL FOR MANUAL REVIEW
    # --------------------------------------------------------

    mail.Display()

    print(
        f"Draft created for: {email} "
        f"({len(attached_items)} payslip(s))"
    )


    # --------------------------------------------------------
    # REMOVE TEMPORARY ZIP
    # --------------------------------------------------------

    if os.path.exists(temp_zip_path):
        os.remove(temp_zip_path)


print(
    "\nAll available payslip emails "
    "have been prepared in Outlook."
)