import json
import re
import os
import base64
import hashlib
from io import BytesIO
from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st
from openai import OpenAI

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image as RLImage,
)


# =========================================================
# ACCOUNTRA BRAND ASSETS
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "accountra_logo_clean.png"
CREATOR_NAME = "Rohan A."


def export_widget_key(kind):
    """Create a stable, upload-specific Streamlit widget key."""
    token = str(st.session_state.get("file_token") or "no_file")
    digest = hashlib.sha256(token.encode("utf-8")).hexdigest()[:12]
    return f"{kind}_{digest}"


# =========================================================
# DISPLAY / DATA HELPERS
# =========================================================


def indian_currency(amount):
    amount = float(amount or 0)
    sign = "-" if amount < 0 else ""
    amount = abs(amount)
    integer_part = int(amount)
    decimal_part = round(amount - integer_part, 2)
    number = str(integer_part)

    if len(number) > 3:
        last_three = number[-3:]
        remaining = number[:-3]
        groups = []
        while len(remaining) > 2:
            groups.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            groups.insert(0, remaining)
        number = ",".join(groups) + "," + last_three

    if decimal_part:
        decimals = f"{decimal_part:.2f}"[1:]
        return f"₹{sign}{number}{decimals}"
    return f"₹{sign}{number}"


APPROVED_ASSET_HEADS = {
    "PPE",
    "Capital Work-in-Progress",
    "Intangible Assets",
    "Intangible Assets Under Development",
    "Investment Property",
    "Investments",
    "Inventories",
    "Trade Receivables",
    "Cash & Cash Equivalents",
    "Other Current Assets",
    "Other Non-current Assets",
}

APPROVED_LIABILITY_HEADS = {
    "Borrowings",
    "Current Borrowings",
    "Non-current Borrowings",
    "Trade Payables",
    "Provisions",
    "Other Current Liabilities",
    "Other Non-current Liabilities",
}

APPROVED_EQUITY_HEADS = {
    "Share Capital",
    "Other Equity",
    "Capital Account",
}

APPROVED_INCOME_HEADS = {
    "Revenue from Operations",
    "Other Income",
}

APPROVED_EXPENSE_HEADS = {
    "Cost of Materials Consumed",
    "Purchases",
    "Changes in Inventories",
    "Employee Benefits Expense",
    "Finance Costs",
    "Depreciation & Amortisation",
    "Other Expenses",
    "Tax Expense",
}

APPROVED_HEADS = (
    APPROVED_ASSET_HEADS
    | APPROVED_LIABILITY_HEADS
    | APPROVED_EQUITY_HEADS
    | APPROVED_INCOME_HEADS
    | APPROVED_EXPENSE_HEADS
)



def make_result(
    nature,
    classification,
    statement,
    reason,
    ambiguous=False,
    confidence=1.0,
    missing_information=None,
):
    return {
        "nature": nature,
        "classification": classification,
        "statement": statement,
        "ambiguous": ambiguous,
        "confidence": float(confidence),
        "reason": reason,
        "missing_information": missing_information,
    }



def clean_number_series(series):
    """Convert Excel/CSV money values to numeric safely."""
    return (
        series.astype(str)
        .str.replace("₹", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace("\\u00a0", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
        .replace(
            {
                "": "0",
                "-": "0",
                "—": "0",
                "–": "0",
                "nan": "0",
                "None": "0",
            }
        )
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )



# =========================================================
# EXPORT / REPORTING HELPERS
# =========================================================


def make_schedule3_excel(
    company_name,
    cin,
    reporting_date,
    results_df,
    pnl_rows,
    bs_rows,
    validation_rows,
    notes_rows,
):
    """Create a professional Excel working-paper/report package."""
    wb = Workbook()

    ws_tb = wb.active
    ws_tb.title = "Classified TB"

    headers = [
        "Account",
        "Debit",
        "Credit",
        "Nature",
        "Classification",
        "Statement",
        "Ambiguous",
        "Confidence",
        "Reason",
        "Missing Information",
    ]

    ws_tb.append(headers)

    for cell in ws_tb[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")

    for _, row in results_df.iterrows():
        ws_tb.append([
            row.get("Account", ""),
            float(row.get("Debit", 0) or 0),
            float(row.get("Credit", 0) or 0),
            row.get("Nature", ""),
            row.get("Classification", ""),
            row.get("Statement", ""),
            bool(row.get("Ambiguous", False)),
            float(row.get("Confidence", 0) or 0),
            row.get("Reason", ""),
            row.get("Missing Information", "") or "",
        ])

    for row in ws_tb.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            cell.number_format = '#,##0.00'

    for row in ws_tb.iter_rows(min_row=2, min_col=8, max_col=8):
        for cell in row:
            cell.number_format = '0.00%'

    widths = {
        1: 30, 2: 16, 3: 16, 4: 14, 5: 34,
        6: 20, 7: 12, 8: 13, 9: 65, 10: 45
    }
    for col, width in widths.items():
        ws_tb.column_dimensions[get_column_letter(col)].width = width
    ws_tb.freeze_panes = "A2"
    ws_tb.auto_filter.ref = ws_tb.dimensions

    # Explicit numeric formatting for the working paper.
    for row in ws_tb.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            cell.number_format = '₹#,##,##0.00'


    def add_report_sheet(title, columns, rows):
        ws = wb.create_sheet(title)
        ws.append([company_name])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
        ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="17365D")
        ws.append([f"CIN: {cin}" if cin else ""] + [None] * (len(columns) - 1))
        ws.append([f"Reporting date: {reporting_date.strftime('%d %B %Y')}"] + [None] * (len(columns) - 1))
        ws.append([])
        ws.append(columns)

        for cell in ws[5]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center")

        for row in rows:
            ws.append(list(row))

        # Make numeric amounts visible as Indian Rupee values.
        if len(columns) >= 2:
            for r in range(6, ws.max_row + 1):
                cell = ws.cell(r, 2)
                if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    cell.number_format = '₹#,##,##0.00'
                    cell.alignment = Alignment(horizontal="right")

        # Highlight Schedule III section/total rows.
        for r in range(6, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "")
            if (label.startswith(("I.", "II.", "1.", "2.", "3."))
                    or label.startswith("Total ")):
                for c in range(1, len(columns) + 1):
                    ws.cell(r, c).font = Font(bold=True)
                    if label.startswith(("I.", "II.")):
                        ws.cell(r, c).fill = PatternFill("solid", fgColor="D9EAF7")

        widths = [62, 24, 24] if len(columns) == 2 else [34, 16, 38]
        for col, width in enumerate(widths[:len(columns)], start=1):
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes = "A6"
        return ws

    add_report_sheet(
        "Profit & Loss",
        ["Particulars", "Amount"],
        pnl_rows,
    )

    add_report_sheet(
        "Balance Sheet",
        ["Particulars", "Amount"],
        bs_rows,
    )

    add_report_sheet(
        "Validation",
        ["Check", "Status", "Detail"],
        validation_rows,
    )

    add_report_sheet(
        "Notes",
        ["Note", "Detail"],
        notes_rows,
    )

    # Cover sheet
    cover = wb.create_sheet("Report Cover", 0)

    if LOGO_PATH.exists():
        try:
            logo = XLImage(str(LOGO_PATH))
            logo.width = 260
            logo.height = 194
            cover.add_image(logo, "A1")
        except Exception:
            pass
    cover["A12"] = company_name
    cover["A12"].font = Font(size=20, bold=True, color="FFFFFF")
    cover["A12"].fill = PatternFill("solid", fgColor="17365D")
    cover.merge_cells("A12:D13")
    cover["A15"] = "Accountra — Schedule III-style Financial Statements"
    cover["A4"].font = Font(size=14, bold=True)
    cover["A16"] = f"Reporting date: {reporting_date.strftime('%d %B %Y')}"
    cover["A17"] = f"CIN: {cin}" if cin else "CIN: Not provided"
    cover["A19"] = "Important"
    cover["A19"].font = Font(bold=True)
    cover["A20"] = (
        "This is an AI-assisted accounting preparation report. "
        "Review classifications, notes and statutory disclosures before filing."
    )
    cover["A22"] = f"Built by {CREATOR_NAME} • Accountra"
    cover["A22"].font = Font(size=9, italic=True, color="7F7F7F")
    cover.column_dimensions["A"].width = 100

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _register_pdf_fonts():
    """Register Unicode fonts so the Indian Rupee symbol renders correctly."""
    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
         "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf"),
        ("C:/Windows/Fonts/DejaVuSans.ttf",
         "C:/Windows/Fonts/DejaVuSans-Bold.ttf"),
        ("C:/Windows/Fonts/NotoSans-Regular.ttf",
         "C:/Windows/Fonts/NotoSans-Bold.ttf"),
    ]

    for regular, bold in candidates:
        if Path(regular).exists() and Path(bold).exists():
            try:
                pdfmetrics.registerFont(TTFont("AccountingAI-Regular", regular))
                pdfmetrics.registerFont(TTFont("AccountingAI-Bold", bold))
                return "AccountingAI-Regular", "AccountingAI-Bold"
            except Exception:
                pass

    return "Helvetica", "Helvetica-Bold"


def make_schedule3_pdf(
    company_name,
    cin,
    reporting_date,
    pnl_rows,
    bs_rows,
    validation_rows,
    notes_rows,
):
    """Create a clean PDF presentation of the generated statements."""
    output = BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
        title=f"{company_name} - Financial Statements",
    )

    regular_font, bold_font = _register_pdf_fonts()

    styles = getSampleStyleSheet()
    for style in styles.byName.values():
        try:
            style.fontName = regular_font
        except Exception:
            pass

    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=17,
        fontName=bold_font,
        spaceAfter=6,
    )
    subtitle_style = ParagraphStyle(
        "ReportSubtitle",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=10,
        fontName=regular_font,
        textColor=colors.grey,
        spaceAfter=16,
    )
    section_style = ParagraphStyle(
        "Section",
        parent=styles["Heading2"],
        fontSize=12,
        fontName=bold_font,
        spaceBefore=12,
        spaceAfter=7,
    )

    story = []

    if LOGO_PATH.exists():
        try:
            logo = RLImage(str(LOGO_PATH), width=150, height=112)
            logo.hAlign = "CENTER"
            story.append(logo)
            story.append(Spacer(1, 5))
        except Exception:
            pass

    story.extend([
        Paragraph(company_name, title_style),
        Paragraph(
            f"Schedule III-style Financial Statements<br/>"
            f"As at / for the period ended {reporting_date.strftime('%d %B %Y')}"
            + (f"<br/>CIN: {cin}" if cin else ""),
            subtitle_style,
        ),
    ])

    def add_table(title, rows):
        story.append(Paragraph(title, section_style))
        data = [["Particulars", "Amount"]]
        for label, amount in rows:
            if isinstance(amount, (int, float)) and not isinstance(amount, bool):
                amount_text = indian_currency(amount)
            else:
                amount_text = str(amount or "")
            data.append([str(label), amount_text])

        table = Table(data, colWidths=[350, 150], repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), bold_font),
            ("FONTNAME", (0, 1), (-1, -1), regular_font),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F7F9FC")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(table)

    add_table("Statement of Profit & Loss", pnl_rows)
    story.append(PageBreak())
    add_table("Balance Sheet", bs_rows)
    story.append(PageBreak())

    story.append(Paragraph("Final Validation", section_style))
    validation_data = [["Check", "Status", "Detail"]]
    validation_data.extend([
        [str(a), str(b), str(c)]
        for a, b, c in validation_rows
    ])
    vt = Table(validation_data, colWidths=[230, 70, 200], repeatRows=1)
    vt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(vt)

    story.append(Spacer(1, 14))
    story.append(Paragraph("Presentation Notes", section_style))
    for note, detail in notes_rows:
        story.append(Paragraph(f"<b>{note}</b>: {detail}", styles["BodyText"]))
        story.append(Spacer(1, 5))

    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "AI-assisted report — review classifications and statutory disclosures "
        "before relying on the statements for statutory filing.",
        styles["Italic"],
    ))
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        f"Built by {CREATOR_NAME} • Accountra",
        ParagraphStyle(
            "CreatorCredit",
            parent=styles["Normal"],
            alignment=TA_CENTER,
            fontSize=8,
            fontName=regular_font,
            textColor=colors.grey,
        ),
    ))

    doc.build(story)
    output.seek(0)
    return output.getvalue()



# =========================================================
# DETERMINISTIC CLASSIFICATION ENGINE
# =========================================================


def classify_account(account, debit=0, credit=0):
    """
    Deterministic first-pass accounting classification.

    The order is intentional: specific payable/receivable/adjustment
    phrases are checked before broad words such as salary, capital,
    stock or loan.
    """
    name = str(account).strip().lower()
    name = re.sub(r"\s+", " ", name)

    # -----------------------------------------------------
    # YEAR-END / INVENTORY ADJUSTMENTS
    # -----------------------------------------------------
    if any(x in name for x in [
        "closing stock",
        "closing inventory",
        "closing inventories",
        "stock in trade",
        "stock-in-trade",
    ]):
        return make_result(
            "Asset",
            "Inventories",
            "Balance Sheet",
            "Closing inventory is presented as an inventory asset.",
        )

    if any(x in name for x in [
        "opening stock",
        "opening inventory",
        "opening inventories",
    ]):
        return make_result(
            "Expense",
            "Changes in Inventories",
            "Profit & Loss",
            "Opening inventory is included in determining the period's cost of goods consumed/sold.",
        )

    # -----------------------------------------------------
    # CONTRA / ADJUSTMENT TERMS THAT MUST BE REVIEWED
    # -----------------------------------------------------
    if any(x in name for x in [
        "sales return",
        "sales returns",
        "returns inward",
        "return inward",
    ]):
        return make_result(
            "Income",
            "Revenue from Operations",
            "Profit & Loss",
            "Sales returns are a reduction of operating revenue.",
        )

    if any(x in name for x in [
        "purchase return",
        "purchase returns",
        "returns outward",
        "return outward",
    ]):
        return make_result(
            "Expense",
            "Purchases",
            "Profit & Loss",
            "Purchase returns reduce purchases.",
        )

    # -----------------------------------------------------
    # SPECIFIC ASSET / LIABILITY PAIRS
    # -----------------------------------------------------
    if any(x in name for x in [
        "capital advance",
        "capital advances",
        "advance for capital goods",
        "advance for fixed asset",
        "advance against fixed asset",
    ]):
        return make_result(
            "Asset",
            "Other Non-current Assets",
            "Balance Sheet",
            "A capital advance represents an amount paid toward a capital asset and is an asset until adjusted against the asset cost.",
        )

    if any(x in name for x in [
        "advance to supplier",
        "supplier advance",
        "supplier advances",
        "advance to vendor",
        "vendor advance",
        "vendor advances",
    ]):
        return make_result(
            "Asset",
            "Other Current Assets",
            "Balance Sheet",
            "An advance paid to a supplier/vendor is an asset until the related goods or services are received.",
        )

    if any(x in name for x in [
        "advance from customer",
        "advance from customers",
        "customer advance",
        "customer advances",
        "advance received from customer",
        "income received in advance",
        "unearned income",
        "deferred revenue",
    ]):
        return make_result(
            "Liability",
            "Other Current Liabilities",
            "Balance Sheet",
            "An amount received before the related goods or services are delivered represents an obligation.",
        )

    if any(x in name for x in [
        "salary payable",
        "salaries payable",
        "wages payable",
        "employee payable",
        "bonus payable",
        "bonus payable to employees",
    ]):
        return make_result(
            "Liability",
            "Other Current Liabilities",
            "Balance Sheet",
            "Amounts payable to employees are liabilities, not current-period employee expense accounts.",
        )

    if any(x in name for x in [
        "expense payable",
        "expenses payable",
        "outstanding expense",
        "outstanding expenses",
        "accrued expense",
        "accrued expenses",
        "expense accrual",
    ]):
        return make_result(
            "Liability",
            "Other Current Liabilities",
            "Balance Sheet",
            "An outstanding/accrued expense represents an amount payable.",
        )

    if any(x in name for x in [
        "interest payable",
        "interest accrued payable",
    ]):
        return make_result(
            "Liability",
            "Other Current Liabilities",
            "Balance Sheet",
            "Interest payable is a current liability unless the facts indicate otherwise.",
        )

    if any(x in name for x in [
        "interest receivable",
        "interest accrued",
        "accrued income",
        "income accrued",
    ]):
        return make_result(
            "Asset",
            "Other Current Assets",
            "Balance Sheet",
            "Income earned but not yet received is a receivable/current asset unless facts indicate otherwise.",
        )

    # -----------------------------------------------------
    # TAX / GST / STATUTORY ASSETS AND LIABILITIES
    # -----------------------------------------------------
    if any(x in name for x in [
        "deferred tax liability",
        "deferred tax liabilities",
        "dtl",
    ]):
        return make_result(
            "Liability",
            "Other Non-current Liabilities",
            "Balance Sheet",
            "Deferred tax liabilities are treated as non-current liabilities in this model.",
        )

    if any(x in name for x in [
        "deferred tax asset",
        "deferred tax assets",
        "dta",
    ]):
        return make_result(
            "Asset",
            "Other Non-current Assets",
            "Balance Sheet",
            "A deferred tax asset is a non-current tax-related asset in this model.",
        )

    if any(x in name for x in [
        "gst receivable",
        "gst input",
        "input gst",
        "input tax credit",
        "input tax credit receivable",
        "cgst receivable",
        "sgst receivable",
        "igst receivable",
        "tds receivable",
        "tds recoverable",
        "tax receivable",
        "income tax receivable",
        "advance income tax",
        "advance tax",
    ]):
        return make_result(
            "Asset",
            "Other Current Assets",
            "Balance Sheet",
            "Tax/GST recoverable or input credits represent amounts recoverable from the tax authorities.",
        )

    if any(x in name for x in [
        "gst payable",
        "output gst",
        "output tax",
        "cgst payable",
        "sgst payable",
        "igst payable",
        "tds payable",
        "tds deducted payable",
        "tax payable",
        "professional tax payable",
        "pf payable",
        "esi payable",
        "statutory dues payable",
    ]):
        return make_result(
            "Liability",
            "Other Current Liabilities",
            "Balance Sheet",
            "Statutory taxes and dues payable represent current liabilities.",
        )

    # -----------------------------------------------------
    # PROVISIONS / PAYABLES
    # -----------------------------------------------------
    if "provision" in name:
        return make_result(
            "Liability",
            "Provisions",
            "Balance Sheet",
            "A provision represents an estimated obligation and is classified under provisions.",
        )

    if any(x in name for x in [
        "trade payable",
        "trade payables",
        "accounts payable",
        "sundry creditor",
        "sundry creditors",
        "creditor",
        "creditors",
        "supplier payable",
        "supplier payables",
    ]):
        return make_result(
            "Liability",
            "Trade Payables",
            "Balance Sheet",
            "Amounts owed to suppliers are trade payables.",
        )

    # -----------------------------------------------------
    # BORROWINGS — SPECIFIC TERMS FIRST
    # -----------------------------------------------------
    if any(x in name for x in [
        "bank overdraft",
        "bank overdrafts",
        "cash credit",
        "cash-credit",
        "working capital loan",
        "working capital borrowing",
        "short term loan",
        "short-term loan",
        "short term borrowing",
        "short-term borrowing",
        "current borrowing",
        "current borrowings",
    ]):
        return make_result(
            "Liability",
            "Current Borrowings",
            "Balance Sheet",
            "The account wording indicates a current/short-term borrowing.",
        )

    if any(x in name for x in [
        "long term loan",
        "long-term loan",
        "long term borrowing",
        "long-term borrowing",
        "non current borrowing",
        "non-current borrowing",
        "non current borrowings",
        "non-current borrowings",
    ]):
        return make_result(
            "Liability",
            "Non-current Borrowings",
            "Balance Sheet",
            "The account wording indicates a non-current/long-term borrowing.",
        )

    if any(x in name for x in [
        "loan",
        "loans",
        "borrowing",
        "borrowings",
        "term loan",
        "debenture",
        "debentures",
        "bonds payable",
        "secured loan",
        "unsecured loan",
    ]):
        return make_result(
            "Liability",
            "Borrowings",
            "Balance Sheet",
            "A loan/borrowing account represents a borrowing; current versus non-current requires maturity information.",
            ambiguous=True,
            confidence=0.90,
            missing_information="Loan maturity/repayment terms are required to determine current or non-current classification.",
        )

    # -----------------------------------------------------
    # CASH / BANK
    # -----------------------------------------------------
    if any(x in name for x in [
        "bank overdraft",
        "cash credit",
    ]):
        return make_result(
            "Liability",
            "Current Borrowings",
            "Balance Sheet",
            "Bank overdrafts/cash-credit facilities are borrowings when presented as credit balances.",
        )

    if any(x in name for x in [
        "cash",
        "petty cash",
        "cash in hand",
    ]):
        return make_result(
            "Asset",
            "Cash & Cash Equivalents",
            "Balance Sheet",
            "Cash represents cash and cash equivalents.",
        )

    if any(x in name for x in [
        "bank",
        "current account",
        "savings account",
        "fixed deposit",
        "term deposit",
        "bank deposit",
    ]):
        return make_result(
            "Asset",
            "Cash & Cash Equivalents",
            "Balance Sheet",
            "A normal debit bank/deposit balance represents cash or a cash equivalent in this model.",
        )

    # -----------------------------------------------------
    # TRADE / OTHER RECEIVABLES
    # -----------------------------------------------------
    if any(x in name for x in [
        "trade receivable",
        "trade receivables",
        "accounts receivable",
        "sundry debtor",
        "sundry debtors",
        "debtor",
        "debtors",
        "customer receivable",
        "customer receivables",
    ]):
        return make_result(
            "Asset",
            "Trade Receivables",
            "Balance Sheet",
            "Amounts due from customers are trade receivables.",
        )

    # -----------------------------------------------------
    # PREPAIDS / DEPOSITS / OTHER ASSETS
    # -----------------------------------------------------
    if any(x in name for x in [
        "prepaid",
        "prepaid expense",
        "prepaid expenses",
        "advance insurance",
        "insurance paid in advance",
    ]):
        return make_result(
            "Asset",
            "Other Current Assets",
            "Balance Sheet",
            "Prepaid expenses represent future economic benefits and are current assets in this model.",
        )

    if any(x in name for x in [
        "security deposit",
        "security deposits",
        "deposit with",
        "deposit paid",
        "refundable deposit",
    ]):
        return make_result(
            "Asset",
            "Other Non-current Assets",
            "Balance Sheet",
            "A refundable security deposit is an asset; its current/non-current presentation depends on expected recovery timing.",
            ambiguous=True,
            confidence=0.85,
            missing_information="Expected recovery period is required to determine current/non-current presentation.",
        )

    # -----------------------------------------------------
    # INVESTMENTS
    # -----------------------------------------------------
    if any(x in name for x in [
        "investment",
        "investments",
        "mutual fund",
        "mutual funds",
        "shares held",
        "bonds held",
        "securities",
    ]):
        return make_result(
            "Asset",
            "Investments",
            "Balance Sheet",
            "Investments are classified under investments in the Balance Sheet.",
        )

    # -----------------------------------------------------
    # PPE / INTANGIBLES / CAPITAL PROJECTS
    # -----------------------------------------------------
    if any(x in name for x in [
        "capital work in progress",
        "capital work-in-progress",
        "capital wip",
        "cwip",
    ]):
        return make_result(
            "Asset",
            "Capital Work-in-Progress",
            "Balance Sheet",
            "Capital expenditure on an incomplete project is classified as capital work-in-progress.",
        )

    if any(x in name for x in [
        "software under development",
        "intangible under development",
        "intangible assets under development",
        "development of software",
    ]):
        return make_result(
            "Asset",
            "Intangible Assets Under Development",
            "Balance Sheet",
            "The wording indicates an intangible asset that is still under development.",
        )

    if any(x in name for x in [
        "goodwill",
        "patent",
        "patents",
        "trademark",
        "trademarks",
        "copyright",
        "copyrights",
        "licence",
        "license",
        "franchise",
        "computer software",
        "software",
        "intangible asset",
        "intangible assets",
    ]):
        return make_result(
            "Asset",
            "Intangible Assets",
            "Balance Sheet",
            "The account represents an identifiable intangible asset.",
        )

    if any(x in name for x in [
        "land",
        "building",
        "plant",
        "machinery",
        "machine",
        "equipment",
        "furniture",
        "fixture",
        "fixtures",
        "vehicle",
        "vehicles",
        "motor car",
        "motor vehicle",
        "office equipment",
        "computer equipment",
        "computer",
        "factory building",
    ]):
        return make_result(
            "Asset",
            "PPE",
            "Balance Sheet",
            "The account represents a tangible fixed asset classified under PPE.",
        )

    if any(x in name for x in [
        "investment property",
        "investment properties",
    ]):
        return make_result(
            "Asset",
            "Investment Property",
            "Balance Sheet",
            "The account is explicitly identified as investment property.",
        )

    # -----------------------------------------------------
    # EQUITY
    # -----------------------------------------------------
    if any(x in name for x in [
        "share capital",
        "equity share capital",
        "preference share capital",
        "paid up capital",
        "paid-up capital",
    ]):
        return make_result(
            "Equity",
            "Share Capital",
            "Balance Sheet",
            "Share capital is classified under equity.",
        )

    if any(x in name for x in [
        "retained earnings",
        "retained profit",
        "surplus",
        "general reserve",
        "capital reserve",
        "securities premium",
        "share premium",
        "other reserves",
        "reserve fund",
        "reserves and surplus",
        "other equity",
    ]):
        return make_result(
            "Equity",
            "Other Equity",
            "Balance Sheet",
            "Reserves, retained earnings and similar balances are components of other equity.",
        )

    if name in {"capital", "capital account", "proprietor capital", "owner capital"}:
        return make_result(
            "Equity",
            "Capital Account",
            "Balance Sheet",
            "The capital account represents owner/proprietor equity.",
        )

    # -----------------------------------------------------
    # REVENUE / OTHER INCOME
    # -----------------------------------------------------
    if any(x in name for x in [
        "sales",
        "sale of goods",
        "sales revenue",
        "revenue from operations",
        "turnover",
        "service revenue",
        "service income",
        "consulting income",
        "consultancy income",
        "professional income",
        "operating revenue",
    ]):
        return make_result(
            "Income",
            "Revenue from Operations",
            "Profit & Loss",
            "The account represents operating revenue.",
        )

    if any(x in name for x in [
        "interest received",
        "interest income",
        "dividend income",
        "rent received",
        "rental income",
        "profit on sale of asset",
        "gain on sale of asset",
        "commission received",
        "commission income",
        "discount received",
        "other income",
        "miscellaneous income",
        "foreign exchange gain",
        "forex gain",
    ]):
        return make_result(
            "Income",
            "Other Income",
            "Profit & Loss",
            "The account represents non-operating/other income.",
        )

    # -----------------------------------------------------
    # EMPLOYEE / OPERATING EXPENSES
    # -----------------------------------------------------
    if any(x in name for x in [
        "salary",
        "salaries",
        "wages",
        "staff welfare",
        "employee benefit",
        "employee benefits",
        "bonus expense",
        "gratuity expense",
        "leave encashment expense",
        "provident fund expense",
        "pf expense",
        "esi expense",
    ]):
        return make_result(
            "Expense",
            "Employee Benefits Expense",
            "Profit & Loss",
            "Salaries, wages and employee-related costs are employee benefit expenses.",
        )

    if any(x in name for x in [
        "depreciation",
        "amortisation",
        "amortization",
    ]):
        return make_result(
            "Expense",
            "Depreciation & Amortisation",
            "Profit & Loss",
            "Depreciation/amortisation is a period expense.",
        )

    if any(x in name for x in [
        "interest expense",
        "interest paid",
        "finance cost",
        "finance costs",
        "bank charges",
        "borrowing cost",
        "borrowing costs",
        "loan interest",
    ]):
        return make_result(
            "Expense",
            "Finance Costs",
            "Profit & Loss",
            "Interest and borrowing-related charges are finance costs.",
        )

    if any(x in name for x in [
        "purchase",
        "purchases",
        "cost of goods purchased",
        "cost of materials purchased",
    ]):
        return make_result(
            "Expense",
            "Purchases",
            "Profit & Loss",
            "Purchases represent goods/materials acquired for resale or production.",
        )

    if any(x in name for x in [
        "freight inward",
        "carriage inward",
        "transport inward",
        "inward freight",
        "inward carriage",
        "direct material freight",
    ]):
        return make_result(
            "Expense",
            "Cost of Materials Consumed",
            "Profit & Loss",
            "Freight/carriage incurred to bring purchased materials into the business is treated as a material cost in this model.",
        )

    if any(x in name for x in [
        "cost of materials consumed",
        "materials consumed",
        "raw material consumed",
        "raw materials consumed",
    ]):
        return make_result(
            "Expense",
            "Cost of Materials Consumed",
            "Profit & Loss",
            "The account directly represents material consumption.",
        )

    if any(x in name for x in [
        "rent expense",
        "rent",
        "electricity",
        "power charges",
        "water charges",
        "telephone expense",
        "internet expense",
        "repairs",
        "repair expense",
        "advertising",
        "advertisement",
        "professional fees",
        "audit fees",
        "legal fees",
        "printing",
        "stationery",
        "office expenses",
        "travelling",
        "travel expense",
        "insurance expense",
        "subscription expense",
        "postage",
        "courier",
        "commission expense",
        "selling expenses",
        "miscellaneous expenses",
        "bad debts",
        "bad debt expense",
        "discount allowed",
        "foreign exchange loss",
        "forex loss",
    ]):
        return make_result(
            "Expense",
            "Other Expenses",
            "Profit & Loss",
            "The account represents an operating expense.",
        )

    if any(x in name for x in [
        "income tax expense",
        "income tax",
        "tax expense",
        "current tax",
        "current tax expense",
    ]):
        return make_result(
            "Expense",
            "Tax Expense",
            "Profit & Loss",
            "Income tax expense is presented as tax expense in Profit & Loss.",
        )

    # -----------------------------------------------------
    # UNKNOWN → AI FALLBACK
    # -----------------------------------------------------
    return None


# =========================================================
# AI FALLBACK
# =========================================================


def get_openai_client():
    """Return an OpenAI client using Streamlit secrets or environment variables."""
    api_key = None
    try:
        api_key = st.secrets.get("OPENAI_API_KEY")
    except Exception:
        pass
    api_key = api_key or os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


def openai_json(prompt, model=None):
    client = get_openai_client()
    if client is None:
        return None
    model = model or os.getenv("OPENAI_MODEL", "gpt-5")
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "Return valid JSON only. Do not use markdown fences."
                },
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        return json.loads(response.choices[0].message.content)
    except Exception as exc:
        print(f"OpenAI error: {exc}")
        return None


def classify_account_ai(account, debit, credit):
    prompt = f"""
You are an accounting classification assistant for Indian companies.
Classify this ledger account for financial statement preparation.

Account: {account}
Debit: ₹{debit}
Credit: ₹{credit}

Approved classifications only:
{sorted(APPROVED_HEADS)}

Rules:
- Never invent facts.
- If current/non-current information is missing for a borrowing or asset, mark ambiguous=true.
- Return only one approved classification.
- Confidence must be between 0 and 1.

Return exactly:
{{
  "nature": "Asset/Liability/Equity/Income/Expense",
  "classification": "approved classification",
  "statement": "Balance Sheet/Profit & Loss",
  "ambiguous": false,
  "confidence": 0.00,
  "reason": "brief explanation",
  "missing_information": null
}}
"""
    return openai_json(prompt)


def extract_pdf_text(file_bytes):
    if PdfReader is None:
        raise RuntimeError("PDF support requires the pypdf package.")
    reader = PdfReader(BytesIO(file_bytes))
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n\n".join(pages).strip()


def extract_source_text(uploaded_file):
    """Turn a source document into text/table context for the AI TB builder."""
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()

    if name.endswith(".pdf"):
        return extract_pdf_text(data)

    if name.endswith(".csv"):
        df = pd.read_csv(BytesIO(data))
        return df.to_csv(index=False)

    if name.endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(BytesIO(data), sheet_name=None)
        chunks = []
        for sheet_name, sheet_df in sheets.items():
            chunks.append(f"SHEET: {sheet_name}\n{sheet_df.fillna('').to_csv(index=False)}")
        return "\n\n".join(chunks)

    if name.endswith((".txt", ".md")):
        return data.decode("utf-8", errors="replace")

    raise ValueError("Unsupported source format.")


def build_ai_trial_balance(source_text):
    """Use AI to reconstruct an aggregated TB without inventing missing sides."""
    if not source_text or not source_text.strip():
        return None, "No readable data was found in the uploaded/pasted source."

    # Keep very large documents bounded while preserving the beginning and end.
    max_chars = 120000
    if len(source_text) > max_chars:
        source_text = (
            source_text[:90000]
            + "\n\n[...middle of source omitted for model context...]\n\n"
            + source_text[-30000:]
        )

    prompt = f"""
You are an Indian accounting data reconstruction assistant.
Your task is to convert raw accounting data into an aggregated Trial Balance for user verification.

SOURCE DATA:
{source_text}

IMPORTANT RULES:
1. Identify transactions, ledger balances, and accounting facts in the source.
2. Aggregate repeated transactions into the same account.
3. Use standard account names such as Purchases, Sales, Rent Expense, Cash, Bank, Trade Receivables, Trade Payables, etc.
4. Preserve the exact monetary amounts you can support from the source.
5. Do NOT create a balancing figure merely to make Debit equal Credit.
6. Do NOT invent cash/bank/credit settlement when the source does not support it.
7. When wording clearly supports a normal double-entry (for example, a credit purchase from a named supplier), you may infer the counterpart, but mark the entry confidence appropriately and include the evidence.
8. If a transaction cannot be completed without missing information, mark it ambiguous and explain exactly what the user needs to clarify.
9. Combine duplicate account entries after extraction.
10. Debit and credit amounts must be numeric and non-negative.
11. The final Trial Balance may be unbalanced. Never force it to balance.

Return JSON in exactly this shape:
{{
  "accounts": [
    {{
      "account": "Purchases",
      "debit": 40000,
      "credit": 0,
      "confidence": 0.95,
      "ambiguous": false,
      "evidence": "Purchased goods from ABC Traders ₹40,000",
      "missing_information": null
    }}
  ],
  "clarifications": [
    {{
      "item": "short description",
      "question": "What information is required?"
    }}
  ],
  "source_summary": "brief summary"
}}
"""

    payload = openai_json(prompt)
    if not payload or not isinstance(payload.get("accounts"), list):
        return None, "The AI could not construct a Trial Balance from the supplied data."

    rows = []
    for item in payload["accounts"]:
        if not isinstance(item, dict):
            continue
        account = str(item.get("account", "")).strip()
        if not account:
            continue
        debit = float(item.get("debit", 0) or 0)
        credit = float(item.get("credit", 0) or 0)
        if debit < 0 or credit < 0:
            continue
        rows.append({
            "Account": account,
            "Debit": debit,
            "Credit": credit,
            "Confidence": float(item.get("confidence", 0) or 0),
            "Ambiguous": bool(item.get("ambiguous", False)),
            "Evidence": str(item.get("evidence", "") or ""),
            "Missing Information": item.get("missing_information") or "",
        })

    if not rows:
        return None, "No usable accounting accounts were extracted."

    tb = pd.DataFrame(rows)
    tb["Debit"] = clean_number_series(tb["Debit"])
    tb["Credit"] = clean_number_series(tb["Credit"])

    # Aggregate repeated account names while preserving the highest-risk flags.
    tb["_key"] = tb["Account"].astype(str).str.strip().str.lower()
    grouped = []
    for _, group in tb.groupby("_key", sort=False):
        first = group.iloc[0]
        grouped.append({
            "Account": first["Account"],
            "Debit": group["Debit"].sum(),
            "Credit": group["Credit"].sum(),
            "Confidence": group["Confidence"].min(),
            "Ambiguous": bool(group["Ambiguous"].any()),
            "Evidence": " | ".join(x for x in group["Evidence"].astype(str) if x.strip())[:2000],
            "Missing Information": " | ".join(x for x in group["Missing Information"].astype(str) if x.strip())[:1000],
        })

    result = pd.DataFrame(grouped)
    return result, payload.get("clarifications", [])


# =========================================================
# COMPARATIVE / RESET HELPERS
# =========================================================


def classify_comparative_tb(comparative_df):
    """Classify a previous-year TB with the same deterministic rules used by the current TB."""
    rows = []
    for _, row in comparative_df.iterrows():
        account = str(row.get("Account", "")).strip()
        if not account:
            continue
        debit = float(row.get("Debit", 0) or 0)
        credit = float(row.get("Credit", 0) or 0)
        result = classify_account(account, debit, credit)
        if result is None:
            # For comparative information, avoid AI guessing; keep only what can be classified.
            continue
        classification = result.get("classification", "")
        if classification not in APPROVED_HEADS:
            continue
        rows.append({
            "Account": account,
            "Debit": debit,
            "Credit": credit,
            "Classification": classification,
            "Nature": result.get("nature", "Unknown"),
        })
    return pd.DataFrame(rows)


def clear_accounting_session():
    """Clear all generated/upload-derived accounting state."""
    keep_keys = {
        "input_mode", "company_name", "cin", "reporting_date", "business_nature",
        "materiality_threshold", "reset_nonce"
    }
    for key in list(st.session_state.keys()):
        if key not in keep_keys and (
            key.startswith("override_")
            or key.startswith("trial_balance_upload")
            or key.startswith("comparative_trial_balance_upload")
            or key.startswith("ai_source_upload")
            or key.startswith("generated_tb_editor")
            or key.startswith("prepare_fs_button")
            or key.startswith("confirm_ai_tb")
            or key.startswith("download_")
            or key in {
                "file_token", "prepared", "results", "comparative_results",
                "generated_tb", "generated_tb_clarifications", "generated_tb_confirmed",
                "comparative_loaded"
            }
        ):
            del st.session_state[key]
    st.session_state["file_token"] = None
    st.session_state["prepared"] = False
    st.session_state.pop("results", None)
    st.session_state.pop("comparative_results", None)
    st.session_state.pop("comparative_loaded", None)
    st.session_state["generated_tb_confirmed"] = False
    st.session_state["reset_nonce"] = int(st.session_state.get("reset_nonce", 0)) + 1

# =========================================================
# STREAMLIT APP — SCHEDULE III PRESENTATION
# =========================================================

st.set_page_config(
    page_title="Accountra",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.6rem;
        padding-bottom: 3rem;
        max-width: 1450px;
    }

    .app-hero {
        padding: 1.25rem 1.5rem;
        border-radius: 18px;
        border: 1px solid rgba(128,128,128,.18);
        background: linear-gradient(
            135deg,
            rgba(99,102,241,.12),
            rgba(14,165,233,.08)
        );
        margin-bottom: 1.2rem;
    }

    .app-hero h1 {
        margin: 0 0 .25rem 0;
        font-size: 2.1rem;
    }

    .app-hero p {
        margin: 0;
        opacity: .78;
        font-size: 1rem;
    }

    .section-card {
        padding: 1rem 1.2rem;
        border: 1px solid rgba(128,128,128,.18);
        border-radius: 14px;
        margin: .6rem 0 1rem 0;
    }

    .fs-title {
        font-size: 1.35rem;
        font-weight: 700;
        margin-top: .5rem;
    }

    .fs-subtitle {
        font-size: .92rem;
        opacity: .7;
        margin-bottom: .8rem;
    }

    .schedule-note {
        font-size: .82rem;
        opacity: .7;
        padding: .6rem .8rem;
        border-left: 3px solid rgba(99,102,241,.55);
        margin: .6rem 0 1rem 0;
    }

    .statement-card {
        border: 1px solid rgba(128,128,128,.22);
        border-radius: 16px;
        overflow: hidden;
        margin: .75rem 0 1.4rem 0;
        background: rgba(255,255,255,.02);
        box-shadow: 0 8px 28px rgba(0,0,0,.06);
    }
    .statement-head {
        padding: 1rem 1.15rem .8rem 1.15rem;
        border-bottom: 1px solid rgba(128,128,128,.18);
    }
    .statement-title {
        font-size: 1.25rem;
        font-weight: 800;
        letter-spacing: .01em;
    }
    .statement-subtitle {
        margin-top: .2rem;
        font-size: .88rem;
        opacity: .72;
    }
    .statement-scroll {
        overflow-x: auto;
    }
    .statement-table {
        width: 100%;
        border-collapse: collapse;
        min-width: 760px;
    }
    .statement-table th {
        padding: .72rem .85rem;
        border-bottom: 1px solid rgba(128,128,128,.25);
        font-size: .78rem;
        text-transform: uppercase;
        letter-spacing: .03em;
        opacity: .72;
        text-align: left;
        white-space: nowrap;
    }
    .statement-table td {
        padding: .62rem .85rem;
        border-bottom: 1px solid rgba(128,128,128,.10);
        font-size: .91rem;
        vertical-align: middle;
    }
    .statement-table .note {
        width: 90px;
        text-align: center;
    }
    .statement-table .amount {
        width: 180px;
        text-align: right;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
    }
    .statement-table .section td {
        font-weight: 800;
        background: rgba(99,102,241,.09);
        border-top: 1px solid rgba(99,102,241,.18);
    }
    .statement-table .subsection td {
        font-weight: 700;
    }
    .statement-table .indent .particular {
        padding-left: 1.8rem;
    }
    .statement-table .total td {
        font-weight: 750;
        border-top: 1px solid rgba(128,128,128,.25);
    }
    .statement-table .subtotal td {
        font-weight: 750;
        border-top: 1px dashed rgba(128,128,128,.28);
    }
    .statement-table .grand-total td {
        font-weight: 850;
        border-top: 2px solid rgba(128,128,128,.40);
        border-bottom: 2px double rgba(128,128,128,.40);
    }

    .validation-pass {
        padding: .7rem .9rem;
        border-radius: 10px;
        border: 1px solid rgba(34,197,94,.25);
        margin: .35rem 0;
    }

    .landing-shell {
        max-width: 1180px;
        margin: 2rem auto 0 auto;
    }

    .landing-hero {
        padding: 3.4rem 3.2rem;
        border-radius: 28px;
        border: 1px solid rgba(99,102,241,.18);
        background:
            radial-gradient(circle at 85% 18%, rgba(14,165,233,.18), transparent 28%),
            radial-gradient(circle at 15% 90%, rgba(99,102,241,.16), transparent 30%),
            linear-gradient(135deg, rgba(99,102,241,.12), rgba(14,165,233,.07));
        box-shadow: 0 20px 60px rgba(30,41,59,.10);
    }

    .landing-eyebrow {
        display: inline-block;
        padding: .42rem .78rem;
        border-radius: 999px;
        border: 1px solid rgba(99,102,241,.22);
        background: rgba(255,255,255,.55);
        font-size: .82rem;
        font-weight: 700;
        letter-spacing: .03em;
        margin-bottom: 1rem;
    }

    .landing-title {
        font-size: clamp(2.6rem, 5vw, 4.5rem);
        line-height: 1.02;
        font-weight: 850;
        letter-spacing: -.04em;
        margin: 0;
    }

    .landing-title span {
        background: linear-gradient(90deg, #4f46e5, #0891b2);
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent;
    }

    .landing-copy {
        max-width: 650px;
        margin-top: 1.1rem;
        font-size: 1.12rem;
        line-height: 1.65;
        opacity: .76;
    }

    .landing-visual {
        min-height: 360px;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .landing-feature {
        height: 100%;
        padding: 1.15rem 1.2rem;
        border-radius: 18px;
        border: 1px solid rgba(128,128,128,.16);
        background: rgba(255,255,255,.48);
    }

    .landing-feature-title {
        font-weight: 800;
        font-size: 1rem;
        margin-bottom: .35rem;
    }

    .landing-feature-text {
        font-size: .88rem;
        line-height: 1.5;
        opacity: .68;
    }

    .landing-note {
        text-align: center;
        margin-top: 1.4rem;
        font-size: .82rem;
        opacity: .58;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# LANDING PAGE
# =========================================================

if st.session_state.get("app_page", "landing") != "workspace":
    if LOGO_PATH.exists():
        logo_left, logo_right = st.columns([0.16, 0.84], vertical_alignment="center")
        with logo_left:
            st.image(str(LOGO_PATH), width=125)
        with logo_right:
            st.markdown(
                '<div style="font-size:1.15rem;font-weight:750;opacity:.7;">Classify • Prepare • Validate</div>',
                unsafe_allow_html=True,
            )

    st.markdown(
        """
        <div class="landing-shell">
            <div class="landing-hero">
                <div class="landing-eyebrow">📊 AI-POWERED ACCOUNTING</div>
                <div class="landing-title">
                    Welcome to <span>Accountra</span>
                </div>
                <div class="landing-copy">
                    Your AI assistant for preparing clear, structured and validated
                    financial statements from a Trial Balance — without the tedious
                    manual classification work.
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    visual_col, text_col = st.columns([1.05, 1], gap="large")

    with visual_col:
        st.markdown(
            """
            <div class="landing-visual">
                <svg viewBox="0 0 620 430" width="100%" role="img"
                     aria-label="Accounting dashboard illustration">
                    <defs>
                        <linearGradient id="card" x1="0" x2="1">
                            <stop offset="0%" stop-color="#eef2ff"/>
                            <stop offset="100%" stop-color="#ecfeff"/>
                        </linearGradient>
                        <linearGradient id="screen" x1="0" x2="1">
                            <stop offset="0%" stop-color="#4f46e5"/>
                            <stop offset="100%" stop-color="#0891b2"/>
                        </linearGradient>
                    </defs>
                    <rect x="65" y="55" width="490" height="315" rx="28"
                          fill="url(#card)" stroke="#cbd5e1" stroke-width="2"/>
                    <rect x="105" y="92" width="410" height="235" rx="18"
                          fill="#ffffff" stroke="#dbeafe" stroke-width="2"/>
                    <rect x="132" y="120" width="190" height="18" rx="9"
                          fill="url(#screen)"/>
                    <rect x="132" y="154" width="115" height="10" rx="5"
                          fill="#cbd5e1"/>
                    <rect x="132" y="180" width="160" height="86" rx="12"
                          fill="#f8fafc" stroke="#e2e8f0"/>
                    <line x1="148" y1="247" x2="148" y2="203"
                          stroke="#4f46e5" stroke-width="12" stroke-linecap="round"/>
                    <line x1="180" y1="247" x2="180" y2="220"
                          stroke="#0891b2" stroke-width="12" stroke-linecap="round"/>
                    <line x1="212" y1="247" x2="212" y2="194"
                          stroke="#6366f1" stroke-width="12" stroke-linecap="round"/>
                    <line x1="244" y1="247" x2="244" y2="211"
                          stroke="#06b6d4" stroke-width="12" stroke-linecap="round"/>
                    <line x1="276" y1="247" x2="276" y2="187"
                          stroke="#4f46e5" stroke-width="12" stroke-linecap="round"/>
                    <rect x="315" y="180" width="170" height="86" rx="12"
                          fill="#f8fafc" stroke="#e2e8f0"/>
                    <text x="335" y="205" font-size="12" font-weight="700"
                          fill="#334155">FINANCIAL STATEMENTS</text>
                    <text x="335" y="229" font-size="11" fill="#64748b">Profit &amp; Loss</text>
                    <text x="335" y="248" font-size="11" fill="#64748b">Balance Sheet</text>
                    <rect x="165" y="285" width="260" height="16" rx="8"
                          fill="#e2e8f0"/>
                    <rect x="165" y="285" width="178" height="16" rx="8"
                          fill="url(#screen)"/>
                    <circle cx="500" cy="90" r="33" fill="#ffffff"
                            stroke="#cbd5e1" stroke-width="2"/>
                    <path d="M485 91 L496 102 L516 78"
                          fill="none" stroke="#16a34a" stroke-width="8"
                          stroke-linecap="round" stroke-linejoin="round"/>
                    <rect x="25" y="245" width="125" height="90" rx="18"
                          fill="#ffffff" stroke="#cbd5e1" stroke-width="2"/>
                    <text x="45" y="272" font-size="12" font-weight="700"
                          fill="#334155">TRIAL BALANCE</text>
                    <text x="45" y="297" font-size="11" fill="#64748b">Debit</text>
                    <text x="105" y="297" font-size="11" fill="#4f46e5">₹</text>
                    <text x="45" y="317" font-size="11" fill="#64748b">Credit</text>
                    <text x="105" y="317" font-size="11" fill="#0891b2">₹</text>
                    <circle cx="535" cy="315" r="45" fill="#eef2ff"/>
                    <path d="M515 315 h40 M535 295 v40"
                          stroke="#4f46e5" stroke-width="7" stroke-linecap="round"/>
                    <path d="M520 345 h30" stroke="#0891b2" stroke-width="6"
                          stroke-linecap="round"/>
                </svg>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with text_col:
        st.markdown("### From numbers to statements — in one workflow.")
        st.write(
            "Upload your Trial Balance, let Accountra classify the accounts, "
            "review anything ambiguous, and generate Schedule III-style "
            "financial statements with validation checks."
        )

        f1, f2 = st.columns(2)
        with f1:
            st.markdown(
                """
                <div class="landing-feature">
                    <div class="landing-feature-title">📂 Upload</div>
                    <div class="landing-feature-text">
                        Start with your Excel or CSV Trial Balance.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with f2:
            st.markdown(
                """
                <div class="landing-feature">
                    <div class="landing-feature-title">🤖 Classify</div>
                    <div class="landing-feature-text">
                        AI-assisted account classification with review flags.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        f3, f4 = st.columns(2)
        with f3:
            st.markdown(
                """
                <div class="landing-feature">
                    <div class="landing-feature-title">📑 Generate</div>
                    <div class="landing-feature-text">
                        Prepare Profit &amp; Loss and Balance Sheet statements.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with f4:
            st.markdown(
                """
                <div class="landing-feature">
                    <div class="landing-feature-title">✅ Validate</div>
                    <div class="landing-feature-text">
                        Check balances and review accounting classifications.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button(
            "🚀 Get Started",
            type="primary",
            use_container_width=True,
            key="landing_get_started",
        ):
            st.session_state["app_page"] = "workspace"
            st.rerun()

    st.markdown(
        '<div class="landing-note">Built for accounting workflows • Always review outputs before statutory use.</div>',
        unsafe_allow_html=True,
    )

    st.stop()

if LOGO_PATH.exists():
    logo_col, name_col = st.columns([0.18, 0.82], vertical_alignment="center")
    with logo_col:
        st.image(str(LOGO_PATH), width=115)
    with name_col:
        st.markdown(
            """
            <div style="padding-top:.25rem;">
                <div style="font-size:1.85rem;font-weight:850;letter-spacing:-.03em;">
                    Accountra
                </div>
                <div style="font-size:.92rem;opacity:.68;">
                    Classify • Prepare • Validate
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
else:
    st.markdown(
        """
        <div class="app-hero">
            <h1>📊 Accountra</h1>
            <p>Trial Balance → AI classification → Schedule III-style financial statements → validation</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# SIDEBAR — REPORTING INFORMATION
# =========================================================

with st.sidebar:
    if st.button("← Home", use_container_width=True, key="back_home"):
        st.session_state["app_page"] = "landing"
        st.rerun()

    st.header("📋 Report Details")

    company_name = st.text_input(
        "Company / Entity Name",
        value="ABC Private Limited",
        key="company_name",
    )

    cin = st.text_input(
        "CIN / Registration No. (optional)",
        value="",
        key="cin",
    )

    materiality_threshold = st.number_input(
        "Comparative movement review threshold (%)",
        min_value=1.0,
        max_value=100.0,
        value=20.0,
        step=5.0,
        key="materiality_threshold",
        help="Used only when a previous-year Trial Balance is supplied.",
    )

    reporting_date = st.date_input(
        "Reporting Date",
        value=date.today(),
        key="reporting_date",
    )

    financial_year = (
        f"{reporting_date.year - 1}-{str(reporting_date.year)[-2:]}"
        if reporting_date.month <= 3
        else f"{reporting_date.year}-{str(reporting_date.year + 1)[-2:]}"
    )

    st.caption(
        f"Financial year: {financial_year}"
    )

    st.caption(
        "These details are used for the report heading and exports."
    )

    st.divider()
    st.markdown("### Workflow")
    st.markdown(
        """
        **1.** Upload Trial Balance  
        **2.** Validate Debit = Credit  
        **3.** Classify accounts  
        **4.** Review ambiguous items  
        **5.** Generate P&L + Balance Sheet  
        **6.** Run final validation
        """
    )

# STEP 1 — CHOOSE INPUT WORKFLOW
# =========================================================

st.markdown("## 1️⃣ Choose how you want to start")

input_mode = "📊 I have a Trial Balance"

source_df = None
source_ready = False
source_is_generated = False

st.caption(
    "Upload an Excel or CSV Trial Balance containing Account, Debit and Credit columns."
)

tb_col, soon_col = st.columns([1.45, 1])

with tb_col:
    st.markdown(
        """
        <div class="section-card">
            <div class="fs-title">📊 I have a Trial Balance</div>
            <div class="fs-subtitle">
                Upload your Trial Balance and turn it into validated financial statements.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "reset_nonce" not in st.session_state:
        st.session_state["reset_nonce"] = 0

    uploaded_file = st.file_uploader(
        "Choose a Trial Balance file",
        type=["xlsx", "xls", "csv"],
        help="Supported formats: .xlsx, .xls and .csv",
        key=f"trial_balance_upload_v3_{st.session_state['reset_nonce']}",
    )

with soon_col:
    st.markdown(
        """
        <div class="section-card">
            <div class="fs-title">🤖 Don't have a Trial Balance?</div>
            <div class="fs-subtitle">
                This option is coming soon.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------
# PATH A — EXISTING TRIAL BALANCE WORKFLOW
# ---------------------------------------------------------

if uploaded_file:
        reset_col, compare_col = st.columns([1, 2])
        with reset_col:
            if st.button(
                "🗑️ Remove file & reset",
                key="reset_uploaded_accounting_v3",
                use_container_width=True,
            ):
                clear_accounting_session()
                st.rerun()

        st.markdown("### 📊 Comparative Information (Optional)")
        st.caption(
            "Have a previous-year Trial Balance? Upload it to show previous-period amounts and movement analysis. "
            "If you do not provide one, the previous-period column remains — and no movement is calculated."
        )
        comparative_file = st.file_uploader(
            "Upload previous-year Trial Balance (optional)",
            type=["xlsx", "xls", "csv"],
            help="Optional previous-year Trial Balance used for comparative presentation and movement analysis.",
            key=f"comparative_trial_balance_upload_v3_{st.session_state['reset_nonce']}",
        )

        file_token = f"tb_{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get("file_token") != file_token:
            st.session_state["file_token"] = file_token
            st.session_state["prepared"] = False
            st.session_state.pop("results", None)
            st.session_state.pop("comparative_results", None)
            st.session_state.pop("comparative_loaded", None)
            for key in list(st.session_state.keys()):
                if key.startswith("override_"):
                    del st.session_state[key]

        st.success(f"Uploaded: **{uploaded_file.name}**")

        if uploaded_file.size and uploaded_file.size > 15 * 1024 * 1024:
            st.error("File is larger than 15 MB. Please upload a smaller Trial Balance.")
            st.stop()

        try:
            if uploaded_file.name.lower().endswith(".csv"):
                source_df = pd.read_csv(uploaded_file)
            else:
                source_df = pd.read_excel(uploaded_file)
        except Exception as upload_error:
            st.error("Could not read the uploaded Trial Balance. Please check the file.")
            st.exception(upload_error)
            st.stop()

        comparative_df = None
        if comparative_file is not None:
            if comparative_file.size and comparative_file.size > 15 * 1024 * 1024:
                st.error("Comparative file is larger than 15 MB.")
                st.stop()
            try:
                if comparative_file.name.lower().endswith(".csv"):
                    comparative_df = pd.read_csv(comparative_file)
                else:
                    comparative_df = pd.read_excel(comparative_file)

                comparative_df.columns = (
                    comparative_df.columns.astype(str)
                    .str.strip().str.lower()
                    .str.replace("₹", "", regex=False)
                    .str.replace("(", "", regex=False)
                    .str.replace(")", "", regex=False)
                    .str.strip()
                )
                comparative_df = comparative_df.rename(columns={
                    "account": "Account", "account name": "Account", "ledger": "Account",
                    "ledger account": "Account", "particulars": "Account",
                    "debit": "Debit", "debits": "Debit", "dr": "Debit",
                    "credit": "Credit", "credits": "Credit", "cr": "Credit",
                })
                required_comparative = {"Account", "Debit", "Credit"}
                if not required_comparative.issubset(comparative_df.columns):
                    st.warning(
                        "Comparative file was uploaded but does not contain Account, Debit and Credit columns. "
                        "Comparative information will be skipped."
                    )
                    comparative_df = None
                else:
                    total_row_names = {"total", "grand total", "trial balance total", "subtotal", "total trial balance"}
                    comparative_df = comparative_df[
                        ~comparative_df["Account"].astype(str).str.strip().str.lower().isin(total_row_names)
                    ].copy()
                    comparative_df["Debit"] = clean_number_series(comparative_df["Debit"])
                    comparative_df["Credit"] = clean_number_series(comparative_df["Credit"])
                    st.success(f"Comparative TB loaded: **{comparative_file.name}**")
            except Exception as comparative_error:
                st.warning("Could not read the comparative Trial Balance. The current-year workflow will continue without comparative information.")
                st.exception(comparative_error)
                comparative_df = None

        source_ready = True
        source_is_generated = False

# ---------------------------------------------------------
# COMMON TRIAL BALANCE NORMALISATION + VALIDATION
# ---------------------------------------------------------

if source_ready and source_df is not None:

    df = source_df.copy()

    # Normalise columns so both paths enter the same existing engine.
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace("₹", "", regex=False)
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )

    df = df.rename(columns={
        "account": "Account",
        "account name": "Account",
        "ledger": "Account",
        "ledger account": "Account",
        "particulars": "Account",
        "debit": "Debit",
        "debits": "Debit",
        "dr": "Debit",
        "credit": "Credit",
        "credits": "Credit",
        "cr": "Credit",
    })

    required_columns = {"Account", "Debit", "Credit"}
    if not required_columns.issubset(df.columns):
        st.error(
            "The file/data must contain Account, Debit and Credit columns. "
            f"Detected: {df.columns.tolist()}"
        )
        st.stop()

    total_row_names = {
        "total", "grand total", "trial balance total", "subtotal", "total trial balance"
    }
    df = df[
        ~df["Account"].astype(str).str.strip().str.lower().isin(total_row_names)
    ].copy()

    df["Debit"] = clean_number_series(df["Debit"])
    df["Credit"] = clean_number_series(df["Credit"])

    st.markdown("### 📄 Trial Balance Preview")
    preview_col1, preview_col2, preview_col3 = st.columns(3)
    with preview_col1:
        st.metric("Accounts", f"{len(df):,}")
    with preview_col2:
        st.metric("Total Debit", indian_currency(df["Debit"].sum()))
    with preview_col3:
        st.metric("Total Credit", indian_currency(df["Credit"].sum()))

    st.dataframe(df, width="stretch", hide_index=True)

    total_debit = float(df["Debit"].sum())
    total_credit = float(df["Credit"].sum())
    difference = total_debit - total_credit

    st.markdown("### ⚖️ Trial Balance Validation")
    if abs(difference) < 0.01:
        st.success(
            f"**Trial Balance is balanced ✅**  Debit: {indian_currency(total_debit)} | Credit: {indian_currency(total_credit)}"
        )
    else:
        st.error(
            f"**Trial Balance is NOT balanced ❌**  Debit: {indian_currency(total_debit)} | Credit: {indian_currency(total_credit)} | Difference: {indian_currency(abs(difference))}"
        )

    st.divider()
    st.markdown("## 2️⃣ Prepare Financial Statements")

    if st.button(
        "📄 Prepare Financial Statements",
        key="prepare_fs_button_v2",
        type="primary",
        use_container_width=True,
    ):
        if abs(difference) >= 0.01:
            st.error("Cannot prepare financial statements because the Trial Balance is not balanced.")
        elif source_is_generated and not st.session_state.get("generated_tb_confirmed", False):
            st.error("Please confirm the AI-generated Trial Balance first.")
        else:
            st.session_state["prepared"] = True
            st.session_state.pop("results", None)
            for key in list(st.session_state.keys()):
                if key.startswith("override_"):
                    del st.session_state[key]

            results = []
            progress = st.progress(0)
            total_rows = max(len(df), 1)

            for index, row in df.iterrows():
                account = str(row["Account"]).strip()
                debit = float(row["Debit"] or 0)
                credit = float(row["Credit"] or 0)

                result = classify_account(account, debit, credit)
                if result is None:
                    st.write(f"🤖 AI review required: **{account}**")
                    try:
                        result = classify_account_ai(account, debit, credit)
                    except Exception:
                        result = None

                if result is None:
                    result = make_result(
                        "Unknown", "NEEDS_REVIEW", "NEEDS_REVIEW",
                        "Unable to classify the account.",
                        ambiguous=True, confidence=0,
                        missing_information="Manual review required.",
                    )

                nature = result.get("nature", "Unknown")
                classification = result.get("classification", "NEEDS_REVIEW")
                statement = result.get("statement", "NEEDS_REVIEW")
                ambiguous = bool(result.get("ambiguous", True))
                confidence = float(result.get("confidence", 0) or 0)
                reason = result.get("reason", "") or ""
                missing_information = result.get("missing_information")

                if classification not in APPROVED_HEADS:
                    classification = "NEEDS_REVIEW"
                    statement = "NEEDS_REVIEW"
                    ambiguous = True
                    confidence = min(confidence, 0.50)
                    reason = "AI returned an unapproved classification."
                    missing_information = "Manual classification required."
                elif classification in APPROVED_EXPENSE_HEADS | APPROVED_INCOME_HEADS:
                    statement = "Profit & Loss"
                else:
                    statement = "Balance Sheet"

                results.append({
                    "Account": account,
                    "Debit": debit,
                    "Credit": credit,
                    "Nature": nature,
                    "Classification": classification,
                    "Statement": statement,
                    "Ambiguous": ambiguous,
                    "Confidence": confidence,
                    "Reason": reason,
                    "Missing Information": missing_information,
                })
                progress.progress((index + 1) / total_rows)

            results_df = pd.DataFrame(results)
            results_df["Debit"] = clean_number_series(results_df["Debit"])
            results_df["Credit"] = clean_number_series(results_df["Credit"])
            st.session_state["results"] = results_df.to_dict("records")

            if comparative_df is not None and len(comparative_df):
                comparative_results_df = classify_comparative_tb(comparative_df)
                st.session_state["comparative_results"] = comparative_results_df.to_dict("records")
                st.session_state["comparative_loaded"] = True
            else:
                st.session_state.pop("comparative_results", None)
                st.session_state.pop("comparative_loaded", None)

            st.success("AI analysis completed successfully! ✅")


# LOAD PREPARED RESULTS
# =========================================================

if st.session_state.get("prepared", False):

    if "results" not in st.session_state:

        st.error(
            "Prepared results are missing. "
            "Please prepare the financial statements again."
        )
        st.stop()

    results_df = pd.DataFrame(
        st.session_state["results"]
    )

    results_df["Debit"] = clean_number_series(
        results_df["Debit"]
    )
    results_df["Credit"] = clean_number_series(
        results_df["Credit"]
    )

    # =====================================================
    # MANUAL OVERRIDES
    # =====================================================

    for key, override in list(
        st.session_state.items()
    ):

        if not key.startswith("override_"):
            continue

        account_name = key.replace(
            "override_",
            "",
            1,
        )

        mask = (
            results_df["Account"]
            == account_name
        )

        results_df.loc[
            mask,
            "Classification"
        ] = override

        if override in (
            APPROVED_INCOME_HEADS
            | APPROVED_EXPENSE_HEADS
        ):

            results_df.loc[
                mask,
                "Statement"
            ] = "Profit & Loss"

        else:

            results_df.loc[
                mask,
                "Statement"
            ] = "Balance Sheet"

        results_df.loc[
            mask,
            "Ambiguous"
        ] = False

        results_df.loc[
            mask,
            "Confidence"
        ] = 1.0

    # =====================================================
    # ACCOUNT CLASSIFICATION TABLE
    # =====================================================

    st.divider()
    st.markdown("## 3️⃣ AI Classification Results")

    display_columns = [
        "Account",
        "Debit",
        "Credit",
        "Nature",
        "Classification",
        "Statement",
        "Ambiguous",
        "Confidence",
    ]

    st.dataframe(
        results_df[display_columns],
        width="stretch",
        hide_index=True,
    )

    # =====================================================
    # REVIEW & EXCEPTIONS
    # =====================================================

    review_df = results_df[
        results_df["Ambiguous"] == True
    ].copy()

    st.divider()
    st.markdown("## 4️⃣ Review & Exceptions")

    if len(review_df) > 0:

        st.warning(
            f"⚠️ {len(review_df)} account(s) require review."
        )

        st.dataframe(
            review_df[
                [
                    "Account",
                    "Classification",
                    "Reason",
                    "Missing Information",
                    "Confidence",
                ]
            ],
            width="stretch",
            hide_index=True,
        )

        st.markdown("### 🛠️ Manual Classification Override")

        selected_account = st.selectbox(
            "Select an account to review",
            review_df["Account"].tolist(),
            key="selected_review_account_v2",
        )

        selected_row = review_df[
            review_df["Account"]
            == selected_account
        ].iloc[0]

        st.info(
            f"**Current classification:** "
            f"{selected_row['Classification']}\n\n"
            f"**Reason:** {selected_row['Reason']}"
        )

        selected_nature = selected_row["Nature"]
        selected_classification = (
            selected_row["Classification"]
        )

        if selected_nature == "Liability":

            if selected_classification == "Borrowings":

                override_options = [
                    "Current Borrowings",
                    "Non-current Borrowings",
                ]

            else:

                override_options = [
                    "Trade Payables",
                    "Provisions",
                    "Other Current Liabilities",
                    "Other Non-current Liabilities",
                ]

        elif selected_nature == "Asset":

            override_options = [
                "PPE",
                "Capital Work-in-Progress",
                "Intangible Assets",
                "Intangible Assets Under Development",
                "Investment Property",
                "Investments",
                "Inventories",
                "Trade Receivables",
                "Cash & Cash Equivalents",
                "Other Current Assets",
                "Other Non-current Assets",
            ]

        elif selected_nature == "Equity":

            override_options = [
                "Share Capital",
                "Other Equity",
                "Capital Account",
            ]

        elif selected_nature == "Income":

            override_options = [
                "Revenue from Operations",
                "Other Income",
            ]

        elif selected_nature == "Expense":

            override_options = [
                "Cost of Materials Consumed",
                "Purchases",
                "Changes in Inventories",
                "Employee Benefits Expense",
                "Finance Costs",
                "Depreciation & Amortisation",
                "Other Expenses",
                "Tax Expense",
            ]

        else:

            override_options = []

        if override_options:

            new_classification = st.selectbox(
                "Choose the correct classification",
                override_options,
                key=f"classification_override_{selected_account}",
            )

            if st.button(
                "✅ Apply Classification",
                key="apply_override_v2",
            ):

                st.session_state[
                    f"override_{selected_account}"
                ] = new_classification

                st.success(
                    f"{selected_account} updated to "
                    f"{new_classification}. ✅"
                )

                st.rerun()

    else:

        st.success(
            "No accounts require manual review. ✅"
        )

        # =====================================================
        # SCHEDULE III — MODERN STATEMENT DISPLAY
    # =====================================================

    st.divider()
    st.markdown("## 5️⃣ Financial Statements")

    st.caption(
        "Schedule III-style presentation with a separate amount column. "
        "Previous-period figures are shown only when comparative data is supplied; otherwise they remain —."
    )

    pnl_df = results_df[results_df["Statement"] == "Profit & Loss"].copy()

    revenue_rows = pnl_df[pnl_df["Classification"].isin(APPROVED_INCOME_HEADS)]
    revenue_summary = (
        revenue_rows.groupby("Classification")[["Debit", "Credit"]].sum()
        if len(revenue_rows)
        else pd.DataFrame(columns=["Debit", "Credit"])
    )
    if len(revenue_summary):
        revenue_summary["Net"] = revenue_summary["Credit"] - revenue_summary["Debit"]

    revenue_ops = (
        float(revenue_summary.loc["Revenue from Operations", "Credit"]
              - revenue_summary.loc["Revenue from Operations", "Debit"])
        if "Revenue from Operations" in revenue_summary.index else 0.0
    )
    other_income = (
        float(revenue_summary.loc["Other Income", "Credit"]
              - revenue_summary.loc["Other Income", "Debit"])
        if "Other Income" in revenue_summary.index else 0.0
    )
    total_revenue = float(revenue_summary["Net"].sum()) if len(revenue_summary) else 0.0

    pre_tax_expense_heads = APPROVED_EXPENSE_HEADS - {"Tax Expense"}
    expense_rows = pnl_df[pnl_df["Classification"].isin(pre_tax_expense_heads)]
    expense_summary = (
        expense_rows.groupby("Classification")[["Debit", "Credit"]].sum()
        if len(expense_rows)
        else pd.DataFrame(columns=["Debit", "Credit"])
    )
    if len(expense_summary):
        expense_summary["Net"] = expense_summary["Debit"] - expense_summary["Credit"]

    total_expenses = float(expense_summary["Net"].sum()) if len(expense_summary) else 0.0

    tax_rows = pnl_df[pnl_df["Classification"] == "Tax Expense"]
    tax_summary = (
        tax_rows[["Debit", "Credit"]].sum()
        if len(tax_rows)
        else pd.Series({"Debit": 0.0, "Credit": 0.0})
    )
    tax_expense = max(0.0, float(tax_summary["Debit"] - tax_summary["Credit"]))
    profit_before_tax = total_revenue - total_expenses
    profit = profit_before_tax - tax_expense


    def amount_from_summary(summary, head, sign="net"):
        if head not in summary.index:
            return 0.0
        if sign == "income":
            return float(summary.loc[head, "Credit"] - summary.loc[head, "Debit"])
        if sign == "expense":
            return float(summary.loc[head, "Debit"] - summary.loc[head, "Credit"])
        return float(summary.loc[head, "Net"])


    def render_statement_table(title, subtitle, rows):
        html_rows = []
        for row in rows:
            kind = row.get("kind", "line")
            label = row.get("label", "")
            note = row.get("note", "")
            current = row.get("current")
            previous = row.get("previous")

            if current is None:
                current_html = ""
            else:
                current_html = indian_currency(current)

            if previous is None:
                previous_html = "—"
            else:
                previous_html = indian_currency(previous)

            cls = f"fs-row {kind}"
            html_rows.append(
                f"<tr class='{cls}'>"
                f"<td class='particular'>{label}</td>"
                f"<td class='note'>{note}</td>"
                f"<td class='amount'>{current_html}</td>"
                f"<td class='amount'>{previous_html}</td>"
                f"</tr>"
            )

        st.markdown(
            f"""
            <div class='statement-card'>
                <div class='statement-head'>
                    <div class='statement-title'>{title}</div>
                    <div class='statement-subtitle'>{company_name} · {subtitle}</div>
                </div>
                <div class='statement-scroll'>
                    <table class='statement-table'>
                        <thead>
                            <tr>
                                <th>Particulars</th>
                                <th class='note'>Note No.</th>
                                <th class='amount'>Current Period ₹</th>
                                <th class='amount'>Previous Period ₹</th>
                            </tr>
                        </thead>
                        <tbody>{''.join(html_rows)}</tbody>
                    </table>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


    # -----------------------------------------------------
    # COMPARATIVE INFORMATION / MOVEMENT ANALYSIS
    # -----------------------------------------------------

    comparative_previous = {}
    if st.session_state.get("comparative_results"):
        comparative_results_df = pd.DataFrame(st.session_state["comparative_results"])
        if len(comparative_results_df):
            prev_group = comparative_results_df.groupby("Classification")[["Debit", "Credit"]].sum()
            prev_group["Net"] = prev_group["Credit"] - prev_group["Debit"]
            comparative_previous = prev_group["Net"].to_dict()

            current_group = results_df.groupby("Classification")[["Debit", "Credit"]].sum()
            current_group["Net"] = current_group["Credit"] - current_group["Debit"]
            movement_rows = []
            all_heads = sorted(set(current_group.index) | set(prev_group.index))
            for head in all_heads:
                cur = float(current_group.loc[head, "Net"]) if head in current_group.index else 0.0
                prev = float(prev_group.loc[head, "Net"]) if head in prev_group.index else 0.0
                change = cur - prev
                if abs(prev) < 0.005:
                    change_pct = None
                else:
                    change_pct = (change / abs(prev)) * 100.0
                material = bool(change_pct is not None and abs(change_pct) >= float(materiality_threshold))
                movement_rows.append({
                    "Classification": head,
                    "Current Period": cur,
                    "Previous Period": prev,
                    "Change": change,
                    "Change %": change_pct,
                    "Material Movement": material,
                })

            st.divider()
            st.markdown("## 📈 Comparative Information")
            st.caption(
                f"Previous-year information supplied. Material movement threshold: {materiality_threshold:.0f}%."
            )
            movement_df = pd.DataFrame(movement_rows)
            display_movement = movement_df.copy()
            display_movement["Current Period"] = display_movement["Current Period"].map(indian_currency)
            display_movement["Previous Period"] = display_movement["Previous Period"].map(indian_currency)
            display_movement["Change"] = display_movement["Change"].map(indian_currency)
            display_movement["Change %"] = display_movement["Change %"].apply(
                lambda x: "—" if pd.isna(x) else f"{x:+.1f}%"
            )
            display_movement = display_movement.rename(columns={
                "Current Period": "Current Period ₹",
                "Previous Period": "Previous Period ₹",
                "Change": "Change ₹",
            })
            st.dataframe(display_movement, width="stretch", hide_index=True)

            material_rows = movement_df[movement_df["Material Movement"]].copy()
            if len(material_rows):
                st.warning(f"⚠️ {len(material_rows)} classification(s) crossed the {materiality_threshold:.0f}% movement review threshold.")
            else:
                st.success("No classification-level movements crossed the selected review threshold.")

    # -----------------------------------------------------
    # P&L
    # -----------------------------------------------------

    def previous_amount(classification, income=False, expense=False):
        value = comparative_previous.get(classification)
        if value is None:
            return None
        return float(value)

    pnl_rows_display = [
        {"label": "I. Revenue from Operations", "kind": "section"},
        {"label": "Revenue from Operations", "note": "1", "current": amount_from_summary(revenue_summary, "Revenue from Operations", "income"), "previous": previous_amount("Revenue from Operations"), "kind": "line indent"},
        {"label": "II. Other Income", "kind": "section"},
        {"label": "Other Income", "note": "2", "current": amount_from_summary(revenue_summary, "Other Income", "income"), "previous": previous_amount("Other Income"), "kind": "line indent"},
        {"label": "III. Total Income", "current": total_revenue, "previous": (sum(v for k,v in comparative_previous.items() if k in APPROVED_INCOME_HEADS) if comparative_previous else None), "kind": "total"},
        {"label": "IV. Expenses", "kind": "section"},
    ]

    pnl_order = [
        "Cost of Materials Consumed",
        "Purchases",
        "Changes in Inventories",
        "Employee Benefits Expense",
        "Finance Costs",
        "Depreciation & Amortisation",
        "Other Expenses",
    ]

    note_no = 3
    for head in pnl_order:
        amount = amount_from_summary(expense_summary, head, "expense")
        if abs(amount) > 0.005:
            pnl_rows_display.append({
                "label": head,
                "note": str(note_no),
                "current": amount,
                "previous": previous_amount(head),
                "kind": "line indent",
            })
            note_no += 1

    pnl_rows_display.extend([
        {"label": "Total Expenses", "current": total_expenses, "previous": (sum(v for k,v in comparative_previous.items() if k in APPROVED_EXPENSE_HEADS and k != "Tax Expense") if comparative_previous else None), "kind": "total"},
        {"label": "Profit Before Tax", "current": profit_before_tax, "kind": "subtotal"},
        {"label": "Tax Expense", "note": str(note_no), "current": tax_expense, "previous": previous_amount("Tax Expense"), "kind": "line indent"},
        {"label": "Profit for the Period", "current": profit, "kind": "grand-total"},
    ])

    st.markdown("### 📑 Statement of Profit & Loss")
    render_statement_table(
        "STATEMENT OF PROFIT & LOSS",
        f"For the period ended {reporting_date.strftime('%d %B %Y')}",
        pnl_rows_display,
    )

    # -----------------------------------------------------
    # BALANCE SHEET SUMMARIES
    # -----------------------------------------------------

    asset_rows = results_df[results_df["Classification"].isin(APPROVED_ASSET_HEADS)]
    asset_summary = (
        asset_rows.groupby("Classification")[["Debit", "Credit"]].sum()
        if len(asset_rows)
        else pd.DataFrame(columns=["Debit", "Credit"])
    )
    if len(asset_summary):
        asset_summary["Net"] = asset_summary["Debit"] - asset_summary["Credit"]

    total_assets = float(asset_summary["Net"].sum()) if len(asset_summary) else 0.0

    liability_rows = results_df[results_df["Classification"].isin(APPROVED_LIABILITY_HEADS)]
    liability_summary = (
        liability_rows.groupby("Classification")[["Debit", "Credit"]].sum()
        if len(liability_rows)
        else pd.DataFrame(columns=["Debit", "Credit"])
    )
    if len(liability_summary):
        liability_summary["Net"] = liability_summary["Credit"] - liability_summary["Debit"]

    total_liabilities = float(liability_summary["Net"].sum()) if len(liability_summary) else 0.0

    equity_rows = results_df[results_df["Classification"].isin(APPROVED_EQUITY_HEADS)]
    equity_summary = (
        equity_rows.groupby("Classification")[["Debit", "Credit"]].sum()
        if len(equity_rows)
        else pd.DataFrame(columns=["Debit", "Credit"])
    )
    if len(equity_summary):
        equity_summary["Net"] = equity_summary["Credit"] - equity_summary["Debit"]

    total_equity = float(equity_summary["Net"].sum()) if len(equity_summary) else 0.0

    total_equity_and_liabilities = total_equity + profit + total_liabilities

    non_current_asset_groups = [
        ("Property, Plant and Equipment", ["PPE"]),
        ("Intangible Assets", ["Intangible Assets"]),
        ("Capital Work-in-Progress", ["Capital Work-in-Progress"]),
        ("Intangible Assets under Development", ["Intangible Assets Under Development"]),
        ("Investment Property", ["Investment Property"]),
        ("Non-current Investments", ["Investments"]),
        ("Other Non-current Assets", ["Other Non-current Assets"]),
    ]
    current_asset_groups = [
        ("Inventories", ["Inventories"]),
        ("Trade Receivables", ["Trade Receivables"]),
        ("Cash and Cash Equivalents", ["Cash & Cash Equivalents"]),
        ("Other Current Assets", ["Other Current Assets"]),
    ]
    non_current_liability_groups = [
        ("Long-term Borrowings", ["Non-current Borrowings"]),
        ("Other Long-term Liabilities", ["Other Non-current Liabilities"]),
    ]
    current_liability_groups = [
        ("Short-term Borrowings", ["Current Borrowings"]),
        ("Trade Payables", ["Trade Payables"]),
        ("Other Current Liabilities", ["Other Current Liabilities"]),
        ("Short-term Provisions", ["Provisions"]),
    ]

    def group_amount(summary, classifications):
        total = 0.0
        for classification in classifications:
            if classification in summary.index:
                total += float(summary.loc[classification, "Net"])
        return total

    share_capital = group_amount(equity_summary, ["Share Capital"])
    other_equity = group_amount(equity_summary, ["Other Equity", "Capital Account"])
    shareholders_funds = share_capital + other_equity

    non_current_liability_total = sum(group_amount(liability_summary, c) for _, c in non_current_liability_groups)
    current_liability_total = sum(group_amount(liability_summary, c) for _, c in current_liability_groups)
    non_current_asset_total = sum(group_amount(asset_summary, c) for _, c in non_current_asset_groups)
    current_asset_total = sum(group_amount(asset_summary, c) for _, c in current_asset_groups)

    # -----------------------------------------------------
    # BALANCE SHEET
    # -----------------------------------------------------

    bs_rows_display = [
        {"label": "I. EQUITY AND LIABILITIES", "kind": "section"},
        {"label": "1. Shareholders' Funds", "kind": "subsection"},
        {"label": "Share Capital", "note": "1", "current": share_capital, "previous": previous_amount("Share Capital"), "kind": "line indent"},
        {"label": "Other Equity", "note": "2", "current": other_equity, "previous": (previous_amount("Other Equity") if previous_amount("Other Equity") is not None else previous_amount("Capital Account")), "kind": "line indent"},
        {"label": "2. Non-current Liabilities", "kind": "subsection"},
    ]

    note_no = 3
    for label, classifications in non_current_liability_groups:
        amount = group_amount(liability_summary, classifications)
        if abs(amount) > 0.005:
            bs_rows_display.append({"label": label, "note": str(note_no), "current": amount, "previous": (sum(previous_amount(c) or 0.0 for c in classifications) if comparative_previous else None), "kind": "line indent"})
            note_no += 1

    bs_rows_display.append({"label": "3. Current Liabilities", "kind": "subsection"})
    for label, classifications in current_liability_groups:
        amount = group_amount(liability_summary, classifications)
        if abs(amount) > 0.005:
            bs_rows_display.append({"label": label, "note": str(note_no), "current": amount, "previous": (sum(previous_amount(c) or 0.0 for c in classifications) if comparative_previous else None), "kind": "line indent"})
            note_no += 1

    bs_rows_display.append({"label": "Total Equity and Liabilities", "current": total_equity_and_liabilities, "kind": "grand-total"})
    bs_rows_display.append({"label": "II. ASSETS", "kind": "section"})
    bs_rows_display.append({"label": "1. Non-current Assets", "kind": "subsection"})

    for label, classifications in non_current_asset_groups:
        amount = group_amount(asset_summary, classifications)
        if abs(amount) > 0.005:
            bs_rows_display.append({"label": label, "note": str(note_no), "current": amount, "previous": (sum(previous_amount(c) or 0.0 for c in classifications) if comparative_previous else None), "kind": "line indent"})
            note_no += 1

    bs_rows_display.append({"label": "2. Current Assets", "kind": "subsection"})
    for label, classifications in current_asset_groups:
        amount = group_amount(asset_summary, classifications)
        if abs(amount) > 0.005:
            bs_rows_display.append({"label": label, "note": str(note_no), "current": amount, "previous": (sum(previous_amount(c) or 0.0 for c in classifications) if comparative_previous else None), "kind": "line indent"})
            note_no += 1

    bs_rows_display.append({"label": "Total Assets", "current": total_assets, "kind": "grand-total"})

    st.markdown("### 📊 Balance Sheet")
    render_statement_table(
        "BALANCE SHEET",
        f"As at {reporting_date.strftime('%d %B %Y')}",
        bs_rows_display,
    )

    balance_difference = total_assets - total_equity_and_liabilities
    if abs(balance_difference) < 0.01:
        st.success(f"**Balance Sheet Tally ✅** · Difference: {indian_currency(0)}")
    else:
        st.error(
            f"**Balance Sheet Tally Failed ❌** · Difference: {indian_currency(abs(balance_difference))}"
        )

    # -----------------------------------------------------
    # NOTES TO ACCOUNTS
    # -----------------------------------------------------

    st.divider()
    st.markdown("## 📚 Notes to Accounts")
    st.caption("Account-level composition behind the Schedule III-style line items shown above.")

    note_definitions = [
        ("1", "Share Capital", ["Share Capital"]),
        ("2", "Other Equity", ["Other Equity", "Capital Account"]),
    ]

    for label, classifications in non_current_liability_groups:
        amount = group_amount(liability_summary, classifications)
        if abs(amount) > 0.005:
            note_definitions.append((str(len(note_definitions) + 1), label, classifications))
    for label, classifications in current_liability_groups:
        amount = group_amount(liability_summary, classifications)
        if abs(amount) > 0.005:
            note_definitions.append((str(len(note_definitions) + 1), label, classifications))
    for label, classifications in non_current_asset_groups:
        amount = group_amount(asset_summary, classifications)
        if abs(amount) > 0.005:
            note_definitions.append((str(len(note_definitions) + 1), label, classifications))
    for label, classifications in current_asset_groups:
        amount = group_amount(asset_summary, classifications)
        if abs(amount) > 0.005:
            note_definitions.append((str(len(note_definitions) + 1), label, classifications))

    # P&L notes are displayed first in the P&L, so provide their detailed composition too.
    pnl_note_definitions = [
        ("1", "Revenue from Operations", ["Revenue from Operations"]),
        ("2", "Other Income", ["Other Income"]),
    ]
    pnl_note_no = 3
    for head in pnl_order:
        if abs(amount_from_summary(expense_summary, head, "expense")) > 0.005:
            pnl_note_definitions.append((str(pnl_note_no), head, [head]))
            pnl_note_no += 1
    pnl_note_definitions.append((str(pnl_note_no), "Tax Expense", ["Tax Expense"]))

    def render_note_table(note_no, title, classifications):
        subset = results_df[results_df["Classification"].isin(classifications)].copy()
        if not len(subset):
            return
        rows = []
        for _, r in subset.iterrows():
            net = float(r["Credit"] - r["Debit"]) if r["Classification"] in (APPROVED_INCOME_HEADS | APPROVED_LIABILITY_HEADS | APPROVED_EQUITY_HEADS) else float(r["Debit"] - r["Credit"])
            rows.append({"Account": r["Account"], "Amount": net})
        total = sum(r["Amount"] for r in rows)
        st.markdown(f"**Note {note_no} — {title}**")
        note_df = pd.DataFrame(rows)
        note_df["Amount ₹"] = note_df["Amount"].map(indian_currency)
        note_df = note_df[["Account", "Amount ₹"]]
        st.dataframe(note_df, width="stretch", hide_index=True)
        st.markdown(f"**Total — {indian_currency(total)}**")

    for note_no, title, classifications in pnl_note_definitions:
        render_note_table(note_no, title, classifications)
    for note_no, title, classifications in note_definitions:
        render_note_table(note_no, title, classifications)

    # -----------------------------------------------------
    # PRESENTATION NOTES
    # -----------------------------------------------------

    st.divider()
    st.markdown("## 📝 Presentation Notes")

    presentation_notes = []
    if "Investments" in results_df["Classification"].values:
        presentation_notes.append(
            "Investments are displayed under Non-current Investments by default. Current/non-current classification should be confirmed when the Trial Balance does not provide the required facts."
        )
    if "Provisions" in results_df["Classification"].values:
        presentation_notes.append(
            "Provisions are displayed under Short-term Provisions by default because the current classification engine does not yet capture a separate long-term provision head."
        )
    if "Capital Account" in results_df["Classification"].values:
        presentation_notes.append(
            "Capital Account is presented within Shareholders' Funds / Other Equity for this Schedule III-style layout."
        )
    if not presentation_notes:
        presentation_notes.append(
            "No additional Schedule III presentation assumptions were detected."
        )
    for note in presentation_notes:
        st.info(f"ℹ️ {note}")

    # 7️⃣ REVIEW & AUDIT INSIGHTS
        # =====================================================

        st.divider()
        st.markdown("## 🔎 Review & Audit Insights")

        low_confidence = results_df[
            pd.to_numeric(
                results_df["Confidence"],
                errors="coerce",
            ).fillna(0) < 0.80
        ].copy()

        missing_info = results_df[
            results_df["Missing Information"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
        ].copy()

        unusual_balance_rows = results_df[
            (
                results_df["Nature"].isin(
                    {"Liability", "Equity", "Income"}
                )
            )
            &
            (
                results_df["Debit"].abs() > 0.01
            )
        ].copy()

        insight_cols = st.columns(3)

        with insight_cols[0]:
            st.metric(
                "Accounts Reviewed",
                len(results_df),
            )

        with insight_cols[1]:
            st.metric(
                "Low Confidence",
                len(low_confidence),
            )

        with insight_cols[2]:
            st.metric(
                "Missing Information",
                len(missing_info),
            )

        if len(low_confidence):
            st.warning(
                f"{len(low_confidence)} account(s) have confidence below 80%. "
                "Review these before final use."
            )

        if len(missing_info):
            st.info(
                f"{len(missing_info)} account(s) have missing-information flags."
            )

        if len(unusual_balance_rows):
            st.info(
                f"{len(unusual_balance_rows)} account(s) have debit balances "
                "despite being classified as liability/equity/income heads. "
                "These may represent contra or adjustment balances."
            )

        # =====================================================
        # 8️⃣ EXPORTS
        # =====================================================

        st.divider()
        st.markdown("## 📥 Export Financial Statements")

        st.caption(
            "Generate an Excel working-paper package and a PDF presentation "
            "from the current validated results."
        )

        revenue_ops = amount_from_summary(revenue_summary, "Revenue from Operations", "income")
        other_income = amount_from_summary(revenue_summary, "Other Income", "income")

        export_pnl_rows = [
            ("I. Revenue from Operations", revenue_ops),
            ("II. Other Income", other_income),
            ("III. Total Revenue", total_revenue),
        ]

        for head in pnl_order:
            amount = (
                float(expense_summary.loc[head, "Net"])
                if head in expense_summary.index
                else 0.0
            )
            if abs(amount) > 0.005:
                export_pnl_rows.append(
                    (head, indian_currency(amount))
                )

        export_pnl_rows.extend([
            ("Total Expenses", total_expenses),
            ("Profit Before Tax", profit_before_tax),
            ("Tax Expense", tax_expense),
            ("Profit for the Period", profit),
        ])

        export_bs_rows = [
            ("I. EQUITY AND LIABILITIES", ""),
            ("1. Shareholders' Funds", ""),
            ("Share Capital", indian_currency(share_capital)),
            ("Reserves and Surplus / Other Equity", indian_currency(other_equity)),
            ("Total Shareholders' Funds", indian_currency(shareholders_funds)),
            ("2. Non-current Liabilities", ""),
        ]

        for label, classifications in non_current_liability_groups:
            amount = group_amount(liability_summary, classifications)
            if abs(amount) > 0.005:
                export_bs_rows.append(
                    (label, amount)
                )

        export_bs_rows.extend([
            ("Total Non-current Liabilities",
             non_current_liability_total),
            ("3. Current Liabilities", ""),
        ])

        for label, classifications in current_liability_groups:
            amount = group_amount(liability_summary, classifications)
            if abs(amount) > 0.005:
                export_bs_rows.append(
                    (label, amount)
                )

        export_bs_rows.extend([
            ("Total Current Liabilities",
             current_liability_total),
            ("Total Equity and Liabilities",
             total_equity_and_liabilities),
            ("II. ASSETS", ""),
            ("1. Non-current Assets", ""),
        ])

        for label, classifications in non_current_asset_groups:
            amount = group_amount(asset_summary, classifications)
            if abs(amount) > 0.005:
                export_bs_rows.append(
                    (label, amount)
                )

        export_bs_rows.extend([
            ("Total Non-current Assets",
             non_current_asset_total),
            ("2. Current Assets", ""),
        ])

        for label, classifications in current_asset_groups:
            amount = group_amount(asset_summary, classifications)
            if abs(amount) > 0.005:
                export_bs_rows.append(
                    (label, amount)
                )

        export_bs_rows.extend([
            ("Total Current Assets",
             current_asset_total),
            ("Total Assets", total_assets),
        ])

        # =====================================================
        # NOTES TO ACCOUNTS
        # =====================================================

        notes_rows = []

        notes_rows.append((
            "Basis of preparation",
            "Financial statements are generated from the uploaded Trial Balance "
            "using deterministic accounting rules with AI fallback for "
            "unrecognized accounts."
        ))

        notes_rows.append((
            "Schedule III presentation",
            "The primary statements use a Schedule III-style hierarchy. "
            "Required statutory notes and disclosures should be reviewed "
            "before filing."
        ))

        if len(low_confidence):
            notes_rows.append((
                "Low-confidence classifications",
                f"{len(low_confidence)} account(s) have confidence below 80%."
            ))

        if len(missing_info):
            notes_rows.append((
                "Missing information",
                f"{len(missing_info)} account(s) contain missing-information flags."
            ))

        if "Inventories" in results_df["Classification"].values:
            notes_rows.append((
                "Inventories",
                "Inventory balances are presented under Current Assets. "
                "Opening/closing inventory treatment should be reviewed against "
                "the entity's accounting policy and adjusted Trial Balance."
            ))

        if "Borrowings" in results_df["Classification"].values:
            notes_rows.append((
                "Borrowings",
                "Current/non-current presentation depends on contractual "
                "maturity and repayment terms where those facts are not "
                "available from the Trial Balance."
            ))

        notes_rows.append((
            "Statutory review",
            "This application is an AI-assisted preparation tool and does not "
            "replace professional review, applicable accounting standards, "
            "company-specific disclosures or statutory filing requirements."
        ))

        # Validation rows are built after the checks below.


        # =====================================================
        # FINAL VALIDATION
        # =====================================================

        st.divider()
        st.markdown("## 9️⃣ Final Validation")

        tb_balanced = (
            abs(difference) < 0.01
        )

        bs_balanced = (
            abs(balance_difference) < 0.01
        )

        approved_classifications = (
            results_df["Classification"]
            .isin(APPROVED_HEADS)
        )

        invalid_classification_count = int(
            (~approved_classifications).sum()
        )

        review_count = int(
            results_df["Ambiguous"]
            .fillna(True)
            .sum()
        )

        numeric_data_ok = (
            pd.api.types.is_numeric_dtype(
                results_df["Debit"]
            )
            and
            pd.api.types.is_numeric_dtype(
                results_df["Credit"]
            )
        )

        pnl_reconciles = abs(
            total_revenue
            - total_expenses
            - tax_expense
            - profit
        ) < 0.01

        tax_pbt_review_required = (
            profit_before_tax < -0.01
            and tax_expense > 0.01
        )

        validation_checks = [
            (
                "Trial Balance balances",
                tb_balanced,
            ),
            (
                "All account classifications are approved",
                invalid_classification_count == 0,
            ),
            (
                "Debit and Credit values are numeric",
                numeric_data_ok,
            ),
            (
                "Profit & Loss reconciles",
                pnl_reconciles,
            ),
            (
                "Balance Sheet Tally",
                bs_balanced,
            ),
            (
                "No accounts require manual review",
                review_count == 0,
            ),
        ]

        passed_count = sum(
            bool(passed)
            for _, passed in validation_checks
        )

        st.metric(
            "Validation Score",
            f"{passed_count}/{len(validation_checks)} checks passed",
        )

        validation_cols = st.columns(2)

        for index, (
            check_name,
            passed,
        ) in enumerate(validation_checks):

            with validation_cols[index % 2]:

                if passed:

                    st.success(
                        f"✅ {check_name}"
                    )

                else:

                    st.warning(
                        f"⚠️ {check_name}"
                    )

        if invalid_classification_count > 0:

            st.error(
                f"{invalid_classification_count} account(s) "
                "have an unapproved classification."
            )

        if tax_pbt_review_required:
            st.warning(
                "PBT is negative while Tax Expense is present. "
                "Tax Expense is correctly shown below PBT, but review whether "
                "the amount represents deferred tax or another applicable tax adjustment."
            )

        if review_count > 0:

            st.info(
                f"{review_count} account(s) still require manual review. "
                "The statements can be viewed, but the validation is "
                "not fully complete."
            )

        all_core_checks_pass = (
            tb_balanced
            and bs_balanced
            and pnl_reconciles
            and invalid_classification_count == 0
            and numeric_data_ok
        )

        if all_core_checks_pass and review_count == 0:

            st.success(
                "🎉 Financial statements passed all validation checks."
            )

        elif all_core_checks_pass:

            st.warning(
                "Financial statements balance, but manual review "
                "is still pending."
            )


        # =====================================================
        # BUILD EXPORT VALIDATION DATA
        # =====================================================

        validation_rows = []

        for check_name, passed in validation_checks:
            validation_rows.append(
                (
                    check_name,
                    "PASS" if passed else "REVIEW",
                    "Passed" if passed else "Requires attention",
                )
            )

        excel_bytes = make_schedule3_excel(
            company_name=company_name,
            cin=cin,
            reporting_date=reporting_date,
            results_df=results_df,
            pnl_rows=export_pnl_rows,
            bs_rows=export_bs_rows,
            validation_rows=validation_rows,
            notes_rows=notes_rows,
        )

        pdf_bytes = make_schedule3_pdf(
            company_name=company_name,
            cin=cin,
            reporting_date=reporting_date,
            pnl_rows=export_pnl_rows,
            bs_rows=export_bs_rows,
            validation_rows=validation_rows,
            notes_rows=notes_rows,
        )

        st.divider()

        export_col1, export_col2 = st.columns(2)

        safe_company = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            company_name.strip() or "AccountingAI"
        ).strip("_")

        # Use literal, upload-specific keys so Streamlit can never confuse
        # the two download widgets, even when the same workbook triggers
        # multiple reruns.
        export_token = hashlib.sha256(
            f"{uploaded_file.name}_{uploaded_file.size}_{file_token}".encode("utf-8")
        ).hexdigest()[:16]

        with export_col1:
            st.download_button(
                "📊 Download Excel Working Paper",
                data=excel_bytes,
                file_name=f"{safe_company}_Schedule_III_Working_Paper.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key=f"download_excel_{export_token}",
            )

        with export_col2:
            st.download_button(
                "📄 Download PDF Financial Statements",
                data=pdf_bytes,
                file_name=f"{safe_company}_Financial_Statements.pdf",
                mime="application/pdf",
                key=f"download_pdf_{export_token}",
            )


        # =====================================================
        # FOOTER
        # =====================================================

        st.divider()

        st.caption(
            "Accountra • Schedule III-style financial statement "
            "presentation • Review all classifications and required "
            "disclosures before using statements for statutory filing."
        )

# =====================================================
# CREATOR CREDIT
# =====================================================

st.markdown(
    f'<div style="text-align:center;font-size:.78rem;opacity:.52;margin-top:2rem;margin-bottom:.8rem;">Built by {CREATOR_NAME} • Accountra</div>',
    unsafe_allow_html=True,
)

# =====================================================
# FEEDBACK & BUG REPORT
# =====================================================

st.divider()
st.markdown("## 💬 Feedback & Bug Report")
st.caption(
    "Help us improve Accountra. Tell us what worked, what didn't, "
    "or what you would like to see next."
)

feedback_type = st.radio(
    "What would you like to send?",
    ["💬 Feedback", "🐛 Report a Bug"],
    horizontal=True,
    key="feedback_type",
)

with st.form("feedback_form"):
    feedback_message = st.text_area(
        "Your message",
        placeholder=(
            "Tell us what happened, what you expected, or what you would like us to improve..."
        ),
        height=120,
        key="feedback_message",
    )

    feedback_email = st.text_input(
        "Email (optional)",
        placeholder="you@example.com",
        key="feedback_email",
    )

    submitted = st.form_submit_button(
        "Send Feedback",
        type="primary",
        use_container_width=True,
    )

    if submitted:
        if feedback_message.strip():
            st.session_state["feedback_submitted"] = True
            st.session_state["feedback_last_type"] = feedback_type
            st.session_state["feedback_last_message"] = feedback_message.strip()
            st.success(
                "Thanks! Your feedback was recorded for this session. "
                "We'll connect a persistent feedback inbox before public launch."
            )
        else:
            st.warning("Please enter a message before submitting.")


# =========================================================
