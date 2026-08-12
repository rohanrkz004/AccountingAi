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
LOGO_PATH = BASE_DIR / "assets" / "accountra_mark.png"
UI_ICON_PATH = BASE_DIR / "assets" / "accountra_mark.png"
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
                "comparative_loaded", "export_excel_bytes", "export_pdf_bytes",
                "export_signature"
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

# =========================================================
# ACCOUNTRA V5 — PRODUCT WORKSPACE UI
# =========================================================

st.set_page_config(
    page_title="Accountra — AI Financial Workspace",
    page_icon=str(BASE_DIR / "assets" / "accountra_favicon.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# State
# -----------------------------

def init_ui_state():
    defaults = {
        "app_page": "home",
        "company_name": "ABC Private Limited",
        "cin": "",
        "business_nature": "",
        "reporting_date": date.today(),
        "materiality_threshold": 20.0,
        "reset_nonce": 0,
        "file_token": None,
        "prepared": False,
        "results": None,
        "comparative_results": None,
        "comparative_loaded": False,
        "generated_tb_confirmed": False,
        "input_mode": "trial_balance",
        "theme_mode": "system",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_ui_state()

# -----------------------------
# Premium visual system
# -----------------------------

st.markdown(r'''
<style>
:root {
  --acc-primary:#5b5ce2;
  --acc-primary-2:#7c6df2;
  --acc-cyan:#16a6b6;
  --acc-ink:#111827;
  --acc-muted:#667085;
  --acc-line:#e7eaf0;
  --acc-line-strong:#d8dce6;
  --acc-bg:#f6f7fb;
  --acc-surface:#ffffff;
  --acc-surface-2:#fbfcfe;
  --acc-success:#159570;
  --acc-warning:#c77a16;
  --acc-danger:#d84c5b;
  --acc-radius:18px;
  --acc-shadow:0 10px 35px rgba(15,23,42,.06);
  --acc-shadow-hover:0 18px 48px rgba(15,23,42,.10);
}

html, body, [class*="css"] {
  font-family: Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

[data-testid="stAppViewContainer"] { background:var(--acc-bg); }
.block-container { max-width:1480px !important; padding:1.2rem 2rem 4rem !important; }
header[data-testid="stHeader"] { background:transparent !important; }
footer, #MainMenu, [data-testid="stToolbar"], [data-testid="stDecoration"] { display:none !important; }

/* Sidebar */
section[data-testid="stSidebar"] {
  background:var(--acc-surface) !important;
  border-right:1px solid var(--acc-line) !important;
  min-width:245px !important;
  max-width:245px !important;
}
section[data-testid="stSidebar"] > div { padding:1.2rem .9rem 1rem !important; }
section[data-testid="stSidebar"] .block-container { padding:0 !important; }

.brand {
  display:flex; align-items:center; gap:.7rem; padding:.4rem .45rem 1.35rem;
}
.brand-mark {
  width:38px; height:38px; border-radius:12px; object-fit:cover;
  box-shadow:0 7px 18px rgba(91,92,226,.20);
}
.brand-name { font-weight:850; font-size:1.05rem; letter-spacing:-.03em; color:var(--acc-ink); }
.brand-sub { font-size:.67rem; color:var(--acc-muted); margin-top:.1rem; }

.nav-label { font-size:.68rem; font-weight:800; text-transform:uppercase; letter-spacing:.10em; color:#98a0b3; padding:.45rem .65rem .35rem; }
.nav-caption { color:var(--acc-muted); font-size:.76rem; padding:.7rem .65rem; line-height:1.45; }

section[data-testid="stSidebar"] .stButton > button {
  width:100%; border:0 !important; background:transparent !important; color:var(--acc-muted) !important;
  box-shadow:none !important; border-radius:11px !important; text-align:left !important;
  padding:.62rem .7rem !important; min-height:2.45rem !important; font-weight:700 !important;
  transition:background .16s ease, color .16s ease, transform .16s ease;
}
section[data-testid="stSidebar"] .stButton > button:hover {
  background:#f1f2ff !important; color:var(--acc-primary) !important; transform:translateX(2px);
}
.nav-active {
  background:linear-gradient(90deg,#eeeeff,#f7f7ff); color:var(--acc-primary); border-left:3px solid var(--acc-primary);
  border-radius:10px; padding:.62rem .7rem; font-weight:800; margin:.08rem 0;
}

/* Generic */
button, input, textarea, select { font-family:inherit !important; }
.stButton > button, .stFormSubmitButton > button, [data-testid="stDownloadButton"] button {
  border-radius:11px !important; min-height:2.55rem !important; font-weight:750 !important;
  border:1px solid var(--acc-line-strong) !important; transition:all .18s ease !important;
}
.stButton > button:hover, .stFormSubmitButton > button:hover, [data-testid="stDownloadButton"] button:hover {
  transform:translateY(-1px); box-shadow:0 9px 24px rgba(15,23,42,.08) !important; border-color:#bfc4ff !important;
}
div[data-baseweb="input"] > div, div[data-baseweb="select"] > div, div[data-baseweb="textarea"] > div { border-radius:11px !important; border-color:var(--acc-line-strong) !important; }

/* App header */
.topbar {
  display:flex; align-items:center; justify-content:space-between; gap:1rem; margin-bottom:1.1rem;
  padding:.35rem 0 .8rem; border-bottom:1px solid var(--acc-line);
}
.topbar-title { font-size:.78rem; color:var(--acc-muted); font-weight:700; }
.topbar-title strong { color:var(--acc-ink); }

.page-eyebrow { color:var(--acc-primary); font-size:.69rem; font-weight:850; text-transform:uppercase; letter-spacing:.11em; margin-bottom:.45rem; }
.page-title { color:var(--acc-ink); font-size:clamp(1.8rem,3vw,2.55rem); line-height:1.05; font-weight:900; letter-spacing:-.055em; margin:0; }
.page-subtitle { color:var(--acc-muted); font-size:.92rem; line-height:1.55; margin-top:.45rem; }

.hero {
  position:relative; overflow:hidden; padding:2rem 2.1rem; border:1px solid #e0e1ff; border-radius:24px;
  background:radial-gradient(circle at 90% 10%,rgba(22,166,182,.15),transparent 25%),
             radial-gradient(circle at 8% 95%,rgba(91,92,226,.12),transparent 30%),
             linear-gradient(135deg,#ffffff,#f4f4ff);
  box-shadow:var(--acc-shadow); margin-bottom:1rem;
}
.hero::after { content:""; position:absolute; width:240px;height:240px;right:-130px;top:-120px;border:1px solid #dedfff;border-radius:50%; }
.hero-content { position:relative; z-index:1; }

.card {
  background:var(--acc-surface); border:1px solid var(--acc-line); border-radius:var(--acc-radius); padding:1.15rem 1.2rem;
  box-shadow:var(--acc-shadow); transition:transform .2s ease, box-shadow .2s ease, border-color .2s ease; height:100%;
}
.card:hover { transform:translateY(-2px); box-shadow:var(--acc-shadow-hover); border-color:#d8d9ff; }
.card-title { color:var(--acc-ink); font-size:.96rem; font-weight:850; letter-spacing:-.02em; }
.card-caption { color:var(--acc-muted); font-size:.75rem; margin-top:.22rem; }

.kpi-grid { display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:.75rem; margin:1rem 0; }
.kpi { background:var(--acc-surface); border:1px solid var(--acc-line); border-radius:16px; padding:1rem; box-shadow:var(--acc-shadow); transition:all .2s ease; }
.kpi:hover { transform:translateY(-2px); box-shadow:var(--acc-shadow-hover); }
.kpi-label { color:var(--acc-muted); font-size:.68rem; font-weight:800; text-transform:uppercase; letter-spacing:.075em; }
.kpi-value { color:var(--acc-ink); font-size:1.42rem; font-weight:900; letter-spacing:-.045em; margin-top:.35rem; }
.kpi-meta { color:var(--acc-muted); font-size:.7rem; margin-top:.25rem; }

.status-pill { display:inline-flex; align-items:center; gap:.35rem; border-radius:999px; padding:.34rem .62rem; font-size:.68rem; font-weight:850; }
.status-good { background:#e8f8f1; color:#087b5e; }
.status-warn { background:#fff3df; color:#9a5c0b; }
.status-bad { background:#ffe9ed; color:#b53648; }

.insight { padding:.75rem .8rem; border:1px solid var(--acc-line); border-radius:12px; background:var(--acc-surface-2); margin:.45rem 0; }
.insight-title { font-size:.8rem; font-weight:850; color:var(--acc-ink); }
.insight-copy { font-size:.72rem; color:var(--acc-muted); line-height:1.45; margin-top:.15rem; }
.insight-good { border-left:3px solid var(--acc-success); }
.insight-warn { border-left:3px solid var(--acc-warning); }
.insight-danger { border-left:3px solid var(--acc-danger); }

.quick-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:.55rem; }
.quick { border:1px solid var(--acc-line); border-radius:12px; padding:.75rem; background:var(--acc-surface-2); font-weight:750; color:var(--acc-ink); transition:all .16s ease; }
.quick:hover { border-color:#cfd0ff; background:#f8f8ff; transform:translateY(-1px); }

.section-head { display:flex; align-items:end; justify-content:space-between; gap:1rem; margin:1.45rem 0 .7rem; }
.section-title { font-size:1.02rem; font-weight:850; color:var(--acc-ink); letter-spacing:-.025em; }
.section-caption { color:var(--acc-muted); font-size:.75rem; }

/* Custom data tables */
.table-shell { border:1px solid var(--acc-line); border-radius:16px; overflow:auto; background:var(--acc-surface); box-shadow:var(--acc-shadow); }
.acc-table { width:100%; min-width:880px; border-collapse:collapse; }
.acc-table th { position:sticky; top:0; z-index:1; background:#f8f9fc; color:#6b7280; font-size:.67rem; text-transform:uppercase; letter-spacing:.07em; padding:.75rem .8rem; text-align:left; border-bottom:1px solid var(--acc-line); white-space:nowrap; }
.acc-table td { color:var(--acc-ink); font-size:.76rem; padding:.68rem .8rem; border-bottom:1px solid #f0f1f4; vertical-align:middle; }
.acc-table tr:hover td { background:#fafaff; }
.acc-table .num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
.acc-table .center { text-align:center; }
.tag { display:inline-flex; padding:.25rem .48rem; border-radius:999px; background:#f1f2ff; color:#4d4fc6; font-size:.64rem; font-weight:800; }
.tag-good { background:#e8f8f1; color:#087b5e; }
.tag-warn { background:#fff3df; color:#9a5c0b; }
.tag-danger { background:#ffe9ed; color:#b53648; }
.conf-track { width:80px; height:5px; background:#eceef3; border-radius:99px; overflow:hidden; display:inline-block; vertical-align:middle; margin-right:.35rem; }
.conf-fill { height:100%; border-radius:99px; background:linear-gradient(90deg,var(--acc-primary),var(--acc-cyan)); }

/* Statement */
.statement-card { border:1px solid var(--acc-line); border-radius:18px; overflow:auto; background:var(--acc-surface); box-shadow:var(--acc-shadow); }
.statement-head { padding:1rem 1.1rem; border-bottom:1px solid var(--acc-line); background:linear-gradient(135deg,#f8f8ff,#f7fbfc); }
.statement-title { font-weight:900; color:var(--acc-ink); font-size:1rem; }
.statement-subtitle { color:var(--acc-muted); font-size:.72rem; margin-top:.2rem; }
.statement-table { width:100%; min-width:760px; border-collapse:collapse; }
.statement-table th { padding:.7rem .85rem; color:#707789; font-size:.65rem; text-transform:uppercase; letter-spacing:.07em; border-bottom:1px solid var(--acc-line); text-align:left; }
.statement-table td { padding:.66rem .85rem; color:var(--acc-ink); font-size:.78rem; border-bottom:1px solid #f0f1f4; }
.statement-table .amount { text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }
.statement-table .note { width:70px; text-align:center; color:var(--acc-muted); }
.statement-table .section td { font-weight:900; background:#f5f5ff; border-top:1px solid #dedfff; }
.statement-table .subsection td { font-weight:800; }
.statement-table .indent td:first-child { padding-left:1.7rem; }
.statement-table .total td { font-weight:850; border-top:1px solid var(--acc-line-strong); }
.statement-table .grand-total td { font-weight:950; border-top:2px solid #bfc0ff; background:#f7f7ff; }

/* Validation */
.validation-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:.65rem; }
.check { border:1px solid var(--acc-line); border-radius:14px; padding:.85rem; background:var(--acc-surface); display:flex; align-items:center; justify-content:space-between; gap:1rem; }
.check-name { font-weight:800; color:var(--acc-ink); font-size:.78rem; }
.check-detail { color:var(--acc-muted); font-size:.68rem; margin-top:.16rem; }

/* Reports */
.report-card { min-height:150px; display:flex; flex-direction:column; justify-content:space-between; }
.report-icon { width:36px;height:36px;border-radius:11px;background:#f0f0ff;color:var(--acc-primary);display:grid;place-items:center;font-weight:900; }

/* Landing */
.landing { max-width:1180px; margin:2.5rem auto; }
.landing-title { font-size:clamp(2.7rem,6vw,5rem); font-weight:950; line-height:.98; letter-spacing:-.065em; color:var(--acc-ink); }
.landing-title span { background:linear-gradient(100deg,#5354df,#13a6b5); -webkit-background-clip:text; background-clip:text; color:transparent; }
.landing-copy { max-width:700px; color:var(--acc-muted); font-size:1.03rem; line-height:1.7; margin-top:1rem; }

/* Hide native app chrome */
[data-testid="stFileUploaderDropzone"] { border-radius:14px !important; border:1.5px dashed #c8caf5 !important; background:#fafaff !important; }
[data-testid="stDataFrame"] { border-radius:14px !important; }

@media (max-width:1050px) {
  .kpi-grid { grid-template-columns:repeat(3,minmax(0,1fr)); }
}
@media (max-width:760px) {
  .block-container { padding:.8rem .75rem 3rem !important; }
  .kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .validation-grid { grid-template-columns:1fr; }
  .hero { padding:1.35rem; }
  .page-title { font-size:1.85rem; }
}
@media (max-width:480px) {
  .kpi-grid { grid-template-columns:1fr; }
  .quick-grid { grid-template-columns:1fr; }
  .kpi-value { font-size:1.3rem; }
}
@media (prefers-color-scheme: dark) {
  :root {
    --acc-ink:#f3f5fa; --acc-muted:#aab2c1; --acc-line:#2b3140; --acc-line-strong:#3a4252;
    --acc-bg:#0c1018; --acc-surface:#121824; --acc-surface-2:#171e2b;
    --acc-shadow:0 10px 35px rgba(0,0,0,.25); --acc-shadow-hover:0 18px 48px rgba(0,0,0,.35);
  }
  section[data-testid="stSidebar"] { background:#0f141e !important; }
  .nav-active { background:#1b2033; }
  section[data-testid="stSidebar"] .stButton > button:hover { background:#1a2030 !important; }
  .hero { background:radial-gradient(circle at 90% 10%,rgba(22,166,182,.10),transparent 25%),linear-gradient(135deg,#121827,#151b2a); border-color:#2f3650; }
  .kpi,.card,.table-shell,.statement-card,.check,.insight,.quick { background:var(--acc-surface); }
  .acc-table th { background:#171d29; }
  .acc-table td,.statement-table td { border-color:#242b37; }
  .acc-table tr:hover td { background:#171d29; }
  .statement-head { background:linear-gradient(135deg,#171b2a,#121b21); }
  .statement-table .section td { background:#1a1d33; }
  .statement-table .grand-total td { background:#191c32; }
  [data-testid="stFileUploaderDropzone"] { background:#121824 !important; border-color:#414872 !important; }
}
</style>
''', unsafe_allow_html=True)

# -----------------------------
# Small render helpers
# -----------------------------

def html_escape(value):
    import html
    return html.escape(str(value if value is not None else ""))


def money(value):
    return indian_currency(float(value or 0))


def set_page(page):
    st.session_state["app_page"] = page


def nav_button(key, label):
    active = st.session_state.get("app_page", "home") == key
    if active:
        st.markdown(f'<div class="nav-active">{html_escape(label)}</div>', unsafe_allow_html=True)
    else:
        if st.button(label, key=f"nav_{key}", use_container_width=True):
            st.session_state["app_page"] = key
            st.rerun()


def page_header(eyebrow, title, subtitle=""):
    st.markdown(f'<div class="page-eyebrow">{html_escape(eyebrow)}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-title">{html_escape(title)}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="page-subtitle">{html_escape(subtitle)}</div>', unsafe_allow_html=True)


def metric_card(label, value, meta=""):
    return f'''<div class="kpi"><div class="kpi-label">{html_escape(label)}</div><div class="kpi-value">{html_escape(value)}</div><div class="kpi-meta">{html_escape(meta)}</div></div>'''


def insight_card(title, copy, tone="good"):
    return f'''<div class="insight insight-{tone}"><div class="insight-title">{html_escape(title)}</div><div class="insight-copy">{html_escape(copy)}</div></div>'''


def render_nav_sidebar():
    with st.sidebar:
        st.markdown(f'''<div class="brand"><img class="brand-mark" src="assets/accountra_mark.png"><div><div class="brand-name">Accountra</div><div class="brand-sub">AI financial workspace</div></div></div>''', unsafe_allow_html=True)
        st.markdown('<div class="nav-label">Workspace</div>', unsafe_allow_html=True)
        nav_button("home", "Overview")
        nav_button("trial_balance", "Trial Balance")
        nav_button("ai_review", "AI Review")
        nav_button("statements", "Financial Statements")
        nav_button("validation", "Validation")
        nav_button("reports", "Reports & Export")
        st.markdown('<div class="nav-label" style="margin-top:.8rem">Manage</div>', unsafe_allow_html=True)
        nav_button("settings", "Report Settings")
        st.markdown('<div class="nav-caption">Accountra turns a Trial Balance into a reviewed, validated financial reporting workflow.</div>', unsafe_allow_html=True)
        if st.session_state.get("prepared"):
            st.markdown('<div style="margin:.6rem .55rem"><span class="status-pill status-good">● Workspace ready</span></div>', unsafe_allow_html=True)
        if st.button("Reset workspace", key="sidebar_reset", use_container_width=True):
            clear_accounting_session()
            st.session_state["app_page"] = "home"
            st.rerun()


def render_topbar():
    company = st.session_state.get("company_name") or "Your company"
    fy = financial_year_label(st.session_state.get("reporting_date", date.today()))
    st.markdown(f'''<div class="topbar"><div class="topbar-title"><strong>Accountra</strong> <span>·</span> {html_escape(company)}</div><div class="topbar-title">{html_escape(fy)} <span>·</span> AI-assisted</div></div>''', unsafe_allow_html=True)


def financial_year_label(reporting_date):
    d = reporting_date or date.today()
    return f"FY {d.year-1}-{str(d.year)[-2:]}" if d.month <= 3 else f"FY {d.year}-{str(d.year+1)[-2:]}"

# -----------------------------
# Input + processing helpers
# -----------------------------

def normalize_tb(uploaded_file):
    if uploaded_file is None:
        return None, None
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as exc:
        return None, f"Could not read the uploaded Trial Balance: {exc}"
    df.columns = (df.columns.astype(str).str.strip().str.lower().str.replace("₹", "", regex=False).str.replace("(", "", regex=False).str.replace(")", "", regex=False).str.strip())
    df = df.rename(columns={
        "account":"Account", "account name":"Account", "ledger":"Account", "ledger account":"Account", "particulars":"Account",
        "debit":"Debit", "debits":"Debit", "dr":"Debit", "credit":"Credit", "credits":"Credit", "cr":"Credit",
    })
    if not {"Account","Debit","Credit"}.issubset(df.columns):
        return None, f"Required columns are Account, Debit and Credit. Detected: {df.columns.tolist()}"
    total_names = {"total","grand total","trial balance total","subtotal","total trial balance"}
    df = df[~df["Account"].astype(str).str.strip().str.lower().isin(total_names)].copy()
    df["Debit"] = clean_number_series(df["Debit"])
    df["Credit"] = clean_number_series(df["Credit"])
    return df, None


def prepare_classifications(df, comparative_df=None):
    results = []
    total_rows = max(len(df), 1)
    progress = st.progress(0, text="Classifying accounts…")
    for index, row in df.iterrows():
        account = str(row["Account"]).strip()
        debit = float(row["Debit"] or 0)
        credit = float(row["Credit"] or 0)
        result = classify_account(account, debit, credit)
        if result is None:
            try:
                result = classify_account_ai(account, debit, credit)
            except Exception:
                result = None
        if result is None:
            result = make_result("Unknown", "NEEDS_REVIEW", "NEEDS_REVIEW", "Unable to classify the account.", ambiguous=True, confidence=0, missing_information="Manual review required.")
        nature = result.get("nature", "Unknown")
        classification = result.get("classification", "NEEDS_REVIEW")
        statement = result.get("statement", "NEEDS_REVIEW")
        ambiguous = bool(result.get("ambiguous", True))
        confidence = float(result.get("confidence", 0) or 0)
        reason = result.get("reason", "") or ""
        missing_information = result.get("missing_information")
        if classification not in APPROVED_HEADS:
            classification, statement, ambiguous, confidence = "NEEDS_REVIEW", "NEEDS_REVIEW", True, min(confidence, .50)
            reason, missing_information = "AI returned an unapproved classification.", "Manual classification required."
        elif classification in APPROVED_EXPENSE_HEADS | APPROVED_INCOME_HEADS:
            statement = "Profit & Loss"
        else:
            statement = "Balance Sheet"
        results.append({"Account":account,"Debit":debit,"Credit":credit,"Nature":nature,"Classification":classification,"Statement":statement,"Ambiguous":ambiguous,"Confidence":confidence,"Reason":reason,"Missing Information":missing_information})
        progress.progress((index+1)/total_rows, text=f"Classifying {index+1:,} of {len(df):,} accounts")
    progress.empty()
    results_df = pd.DataFrame(results)
    if comparative_df is not None and len(comparative_df):
        comp = classify_comparative_tb(comparative_df)
        st.session_state["comparative_results"] = comp.to_dict("records")
        st.session_state["comparative_loaded"] = True
    else:
        st.session_state.pop("comparative_results", None)
        st.session_state["comparative_loaded"] = False
    return results_df


def apply_overrides(results_df):
    df = results_df.copy()
    for key, override in list(st.session_state.items()):
        if not key.startswith("override_"):
            continue
        account_name = key.replace("override_", "", 1)
        mask = df["Account"] == account_name
        df.loc[mask, "Classification"] = override
        df.loc[mask, "Statement"] = "Profit & Loss" if override in (APPROVED_INCOME_HEADS | APPROVED_EXPENSE_HEADS) else "Balance Sheet"
        df.loc[mask, "Ambiguous"] = False
        df.loc[mask, "Confidence"] = 1.0
    return df


def compute_financials(results_df):
    pnl_df = results_df[results_df["Statement"] == "Profit & Loss"].copy()
    revenue_rows = pnl_df[pnl_df["Classification"].isin(APPROVED_INCOME_HEADS)]
    revenue_summary = revenue_rows.groupby("Classification")[["Debit","Credit"]].sum() if len(revenue_rows) else pd.DataFrame(columns=["Debit","Credit"])
    if len(revenue_summary): revenue_summary["Net"] = revenue_summary["Credit"] - revenue_summary["Debit"]
    revenue_ops = float(revenue_summary.loc["Revenue from Operations","Net"]) if "Revenue from Operations" in revenue_summary.index else 0.0
    other_income = float(revenue_summary.loc["Other Income","Net"]) if "Other Income" in revenue_summary.index else 0.0
    total_revenue = float(revenue_summary["Net"].sum()) if len(revenue_summary) else 0.0
    pre_tax_heads = APPROVED_EXPENSE_HEADS - {"Tax Expense"}
    expense_rows = pnl_df[pnl_df["Classification"].isin(pre_tax_heads)]
    expense_summary = expense_rows.groupby("Classification")[["Debit","Credit"]].sum() if len(expense_rows) else pd.DataFrame(columns=["Debit","Credit"])
    if len(expense_summary): expense_summary["Net"] = expense_summary["Debit"] - expense_summary["Credit"]
    total_expenses = float(expense_summary["Net"].sum()) if len(expense_summary) else 0.0
    tax_rows = pnl_df[pnl_df["Classification"] == "Tax Expense"]
    tax_summary = tax_rows[["Debit","Credit"]].sum() if len(tax_rows) else pd.Series({"Debit":0.0,"Credit":0.0})
    tax_expense = max(0.0, float(tax_summary["Debit"] - tax_summary["Credit"]))
    pbt = total_revenue - total_expenses
    profit = pbt - tax_expense

    asset_rows = results_df[results_df["Classification"].isin(APPROVED_ASSET_HEADS)]
    asset_summary = asset_rows.groupby("Classification")[["Debit","Credit"]].sum() if len(asset_rows) else pd.DataFrame(columns=["Debit","Credit"])
    if len(asset_summary): asset_summary["Net"] = asset_summary["Debit"] - asset_summary["Credit"]
    total_assets = float(asset_summary["Net"].sum()) if len(asset_summary) else 0.0
    liability_rows = results_df[results_df["Classification"].isin(APPROVED_LIABILITY_HEADS)]
    liability_summary = liability_rows.groupby("Classification")[["Debit","Credit"]].sum() if len(liability_rows) else pd.DataFrame(columns=["Debit","Credit"])
    if len(liability_summary): liability_summary["Net"] = liability_summary["Credit"] - liability_summary["Debit"]
    total_liabilities = float(liability_summary["Net"].sum()) if len(liability_summary) else 0.0
    equity_rows = results_df[results_df["Classification"].isin(APPROVED_EQUITY_HEADS)]
    equity_summary = equity_rows.groupby("Classification")[["Debit","Credit"]].sum() if len(equity_rows) else pd.DataFrame(columns=["Debit","Credit"])
    if len(equity_summary): equity_summary["Net"] = equity_summary["Credit"] - equity_summary["Debit"]
    total_equity = float(equity_summary["Net"].sum()) if len(equity_summary) else 0.0
    total_equity_and_liabilities = total_equity + profit + total_liabilities

    nca = [("Property, Plant and Equipment",["PPE"]),("Intangible Assets",["Intangible Assets"]),("Capital Work-in-Progress",["Capital Work-in-Progress"]),("Intangible Assets under Development",["Intangible Assets Under Development"]),("Investment Property",["Investment Property"]),("Non-current Investments",["Investments"]),("Other Non-current Assets",["Other Non-current Assets"])]
    ca = [("Inventories",["Inventories"]),("Trade Receivables",["Trade Receivables"]),("Cash and Cash Equivalents",["Cash & Cash Equivalents"]),("Other Current Assets",["Other Current Assets"])]
    ncl = [("Long-term Borrowings",["Non-current Borrowings"]),("Other Long-term Liabilities",["Other Non-current Liabilities"])]
    cl = [("Short-term Borrowings",["Current Borrowings"]),("Trade Payables",["Trade Payables"]),("Other Current Liabilities",["Other Current Liabilities"]),("Short-term Provisions",["Provisions"])]
    def group_amount(summary, classes):
        return sum(float(summary.loc[c,"Net"]) for c in classes if c in summary.index)
    share_capital = group_amount(equity_summary,["Share Capital"])
    other_equity = group_amount(equity_summary,["Other Equity","Capital Account"])
    shareholders_funds = share_capital + other_equity
    ncl_total = sum(group_amount(liability_summary,c) for _,c in ncl)
    cl_total = sum(group_amount(liability_summary,c) for _,c in cl)
    nca_total = sum(group_amount(asset_summary,c) for _,c in nca)
    ca_total = sum(group_amount(asset_summary,c) for _,c in ca)
    balance_difference = total_assets - total_equity_and_liabilities

    comparative_previous = {}
    movement_df = pd.DataFrame()
    if st.session_state.get("comparative_results"):
        comp = pd.DataFrame(st.session_state["comparative_results"])
        if len(comp):
            pg = comp.groupby("Classification")[["Debit","Credit"]].sum(); pg["Net"] = pg["Credit"] - pg["Debit"]
            comparative_previous = pg["Net"].to_dict()
            cg = results_df.groupby("Classification")[["Debit","Credit"]].sum(); cg["Net"] = cg["Credit"] - cg["Debit"]
            rows=[]
            for head in sorted(set(cg.index)|set(pg.index)):
                cur=float(cg.loc[head,"Net"]) if head in cg.index else 0.0; prev=float(pg.loc[head,"Net"]) if head in pg.index else 0.0
                change=cur-prev; pct=None if abs(prev)<.005 else change/abs(prev)*100
                rows.append({"Classification":head,"Current Period":cur,"Previous Period":prev,"Change":change,"Change %":pct,"Material Movement":bool(pct is not None and abs(pct)>=float(st.session_state.get("materiality_threshold",20)))})
            movement_df=pd.DataFrame(rows)

    pnl_order=["Cost of Materials Consumed","Purchases","Changes in Inventories","Employee Benefits Expense","Finance Costs","Depreciation & Amortisation","Other Expenses"]
    def prev_amount(c): return comparative_previous.get(c)
    pnl_rows=[
      {"label":"I. Revenue from Operations","kind":"section"},{"label":"Revenue from Operations","note":"1","current":revenue_ops,"previous":prev_amount("Revenue from Operations"),"kind":"indent"},
      {"label":"II. Other Income","kind":"section"},{"label":"Other Income","note":"2","current":other_income,"previous":prev_amount("Other Income"),"kind":"indent"},
      {"label":"III. Total Revenue","current":total_revenue,"previous":sum(v for k,v in comparative_previous.items() if k in APPROVED_INCOME_HEADS) if comparative_previous else None,"kind":"total"},
      {"label":"IV. Expenses","kind":"section"}]
    note_no=3
    for h in pnl_order:
        amt=float(expense_summary.loc[h,"Net"]) if h in expense_summary.index else 0.0
        if abs(amt)>.005:
            pnl_rows.append({"label":h,"note":str(note_no),"current":amt,"previous":prev_amount(h),"kind":"indent"}); note_no+=1
    pnl_rows += [{"label":"Total Expenses","current":total_expenses,"previous":sum(v for k,v in comparative_previous.items() if k in APPROVED_EXPENSE_HEADS and k!="Tax Expense") if comparative_previous else None,"kind":"total"},{"label":"Profit Before Tax","current":pbt,"kind":"subtotal"},{"label":"Tax Expense","note":str(note_no),"current":tax_expense,"previous":prev_amount("Tax Expense"),"kind":"indent"},{"label":"Profit for the Period","current":profit,"kind":"grand-total"}]

    bs_rows=[{"label":"I. EQUITY AND LIABILITIES","kind":"section"},{"label":"1. Shareholders' Funds","kind":"subsection"},{"label":"Share Capital","note":"1","current":share_capital,"previous":prev_amount("Share Capital"),"kind":"indent"},{"label":"Other Equity","note":"2","current":other_equity,"previous":prev_amount("Other Equity") if prev_amount("Other Equity") is not None else prev_amount("Capital Account"),"kind":"indent"},{"label":"2. Non-current Liabilities","kind":"subsection"}]
    note_no=3
    for label,classes in ncl:
        amt=group_amount(liability_summary,classes)
        if abs(amt)>.005: bs_rows.append({"label":label,"note":str(note_no),"current":amt,"previous":sum(prev_amount(c) or 0 for c in classes) if comparative_previous else None,"kind":"indent"}); note_no+=1
    bs_rows.append({"label":"3. Current Liabilities","kind":"subsection"})
    for label,classes in cl:
        amt=group_amount(liability_summary,classes)
        if abs(amt)>.005: bs_rows.append({"label":label,"note":str(note_no),"current":amt,"previous":sum(prev_amount(c) or 0 for c in classes) if comparative_previous else None,"kind":"indent"}); note_no+=1
    bs_rows += [{"label":"Total Equity and Liabilities","current":total_equity_and_liabilities,"kind":"grand-total"},{"label":"II. ASSETS","kind":"section"},{"label":"1. Non-current Assets","kind":"subsection"}]
    for label,classes in nca:
        amt=group_amount(asset_summary,classes)
        if abs(amt)>.005: bs_rows.append({"label":label,"note":str(note_no),"current":amt,"previous":sum(prev_amount(c) or 0 for c in classes) if comparative_previous else None,"kind":"indent"}); note_no+=1
    bs_rows.append({"label":"2. Current Assets","kind":"subsection"})
    for label,classes in ca:
        amt=group_amount(asset_summary,classes)
        if abs(amt)>.005: bs_rows.append({"label":label,"note":str(note_no),"current":amt,"previous":sum(prev_amount(c) or 0 for c in classes) if comparative_previous else None,"kind":"indent"}); note_no+=1
    bs_rows.append({"label":"Total Assets","current":total_assets,"kind":"grand-total"})

    notes=[]
    if "Investments" in results_df["Classification"].values: notes.append(("Investments","Presented under non-current investments by default; confirm current/non-current presentation when required facts are unavailable."))
    if "Provisions" in results_df["Classification"].values: notes.append(("Provisions","Presented under short-term provisions by default because the current classification engine does not capture a separate long-term provision head."))
    if "Capital Account" in results_df["Classification"].values: notes.append(("Capital Account","Presented within shareholders' funds / other equity for this Schedule III-style layout."))
    if not notes: notes.append(("Presentation","No additional Schedule III presentation assumptions were detected."))
    return locals()


def render_statement_table(title, subtitle, rows):
    parts=[]
    for r in rows:
        kind=r.get("kind","line"); label=html_escape(r.get("label","")); note=html_escape(r.get("note","")); cur=r.get("current"); prev=r.get("previous")
        cur_html=money(cur) if cur is not None else ""; prev_html=money(prev) if prev is not None else "—"
        parts.append(f'<tr class="{html_escape(kind)}"><td>{label}</td><td class="note">{note}</td><td class="amount">{html_escape(cur_html)}</td><td class="amount">{html_escape(prev_html)}</td></tr>')
    st.markdown(f'''<div class="statement-card"><div class="statement-head"><div class="statement-title">{html_escape(title)}</div><div class="statement-subtitle">{html_escape(subtitle)}</div></div><table class="statement-table"><thead><tr><th>Particulars</th><th class="note">Note</th><th class="amount">Current Period</th><th class="amount">Previous Period</th></tr></thead><tbody>{''.join(parts)}</tbody></table></div>''', unsafe_allow_html=True)


def render_tb_table(df):
    rows=[]
    for _,r in df.iterrows():
        amb=bool(r.get("Ambiguous",False)); conf=float(r.get("Confidence",0) or 0); pct=max(0,min(100,conf*100)); tag="tag-warn" if amb else "tag-good"; label="Review" if amb else "Ready"
        rows.append(f'''<tr><td>{html_escape(r['Account'])}</td><td class="num">{html_escape(money(r['Debit']))}</td><td class="num">{html_escape(money(r['Credit']))}</td><td><span class="tag">{html_escape(r['Nature'])}</span></td><td><span class="tag">{html_escape(r['Classification'])}</span></td><td class="center"><span class="tag {tag}">{label}</span></td><td class="center"><span class="conf-track"><span class="conf-fill" style="width:{pct:.0f}%"></span></span>{conf:.0%}</td></tr>''')
    st.markdown(f'''<div class="table-shell"><table class="acc-table"><thead><tr><th>Account</th><th class="num">Debit</th><th class="num">Credit</th><th>Nature</th><th>Classification</th><th class="center">Status</th><th class="center">Confidence</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>''', unsafe_allow_html=True)


def build_export_artifacts(results_df, fin):
    validation_rows = validation_check_rows(results_df, fin)
    pnl_export=[]
    for r in fin["pnl_rows"]:
        pnl_export.append((r["label"], r.get("current") if r.get("current") is not None else ""))
    bs_export=[]
    for r in fin["bs_rows"]:
        bs_export.append((r["label"], r.get("current") if r.get("current") is not None else ""))
    notes_rows=fin["notes"]
    excel=make_schedule3_excel(st.session_state["company_name"],st.session_state["cin"],st.session_state["reporting_date"],results_df,pnl_export,bs_export,validation_rows,notes_rows)
    pdf=make_schedule3_pdf(st.session_state["company_name"],st.session_state["cin"],st.session_state["reporting_date"],pnl_export,bs_export,validation_rows,notes_rows)
    return excel,pdf,validation_rows


def validation_check_rows(results_df, fin):
    # Recompute the uploaded TB totals from classified rows; this is the same accounting validation basis as the original app.
    debit=float(results_df["Debit"].sum()); credit=float(results_df["Credit"].sum()); tb_bal=abs(debit-credit)<.01
    invalid=int((~results_df["Classification"].isin(APPROVED_HEADS)).sum())
    numeric_ok=pd.api.types.is_numeric_dtype(results_df["Debit"]) and pd.api.types.is_numeric_dtype(results_df["Credit"])
    pnl_reconciles=abs(fin["total_revenue"]-fin["total_expenses"]-fin["tax_expense"]-fin["profit"])<.01
    bs_bal=abs(fin["balance_difference"])<.01
    review=int(results_df["Ambiguous"].fillna(True).sum())
    return [("Trial Balance balances","PASS" if tb_bal else "REVIEW",f"Debit {money(debit)} · Credit {money(credit)}"),("All account classifications are approved","PASS" if invalid==0 else "REVIEW",f"{invalid} unapproved classification(s)"),("Debit and Credit values are numeric","PASS" if numeric_ok else "REVIEW","Numeric validation"),("Profit & Loss reconciles","PASS" if pnl_reconciles else "REVIEW",f"Profit {money(fin['profit'])}"),("Balance Sheet Tally","PASS" if bs_bal else "REVIEW",f"Difference {money(fin['balance_difference'])}"),("No accounts require manual review","PASS" if review==0 else "REVIEW",f"{review} account(s) flagged")]

# -----------------------------
# Sidebar + topbar
# -----------------------------
render_nav_sidebar()
render_topbar()

# -----------------------------
# HOME / DASHBOARD
# -----------------------------

def render_home():
    prepared=bool(st.session_state.get("prepared") and st.session_state.get("results"))
    if not prepared:
        st.markdown('''<div class="hero"><div class="hero-content"><div class="page-eyebrow">AI FINANCIAL WORKSPACE</div><div class="landing-title">From Trial Balance to <span>decision-ready statements.</span></div><div class="landing-copy">Upload your Trial Balance, classify accounts with deterministic rules plus AI fallback, review exceptions, validate the numbers and generate professional financial statements — all in one workspace.</div></div></div>''', unsafe_allow_html=True)
        a,b,c=st.columns(3)
        for col,title,copy in [(a,"1 · Upload","Bring your Excel or CSV Trial Balance into a clean workspace."),(b,"2 · Review","See classifications, confidence and exceptions before anything is final."),(c,"3 · Report","Generate Schedule III-style statements, validation and exports.")]:
            with col: st.markdown(f'<div class="card"><div class="card-title">{title}</div><div class="card-caption">{copy}</div></div>',unsafe_allow_html=True)
        st.markdown('<div class="section-head"><div><div class="section-title">Start your workspace</div><div class="section-caption">Nothing is processed until you choose a file and prepare the statements.</div></div></div>',unsafe_allow_html=True)
        render_input_panel()
        return
    results_df=apply_overrides(pd.DataFrame(st.session_state["results"]))
    st.session_state["results"]=results_df.to_dict("records")
    fin=compute_financials(results_df)
    tb_debit=float(results_df["Debit"].sum()); tb_credit=float(results_df["Credit"].sum()); diff=tb_debit-tb_credit
    review=int(results_df["Ambiguous"].fillna(True).sum())
    st.markdown(f'<div class="hero"><div class="hero-content"><div class="page-eyebrow">ACCOUNTRA INTELLIGENCE · OVERVIEW</div><div class="page-title">Good afternoon, {html_escape(st.session_state["company_name"])}.</div><div class="page-subtitle">Your books are prepared for {html_escape(financial_year_label(st.session_state["reporting_date"]))}. Here is what deserves your attention.</div></div></div>',unsafe_allow_html=True)
    status="Balanced" if abs(diff)<.01 else "Needs attention"
    status_cls="status-good" if abs(diff)<.01 else "status-bad"
    st.markdown('<div class="kpi-grid">'+''.join([metric_card("Total Assets",money(fin["total_assets"]),"Balance Sheet"),metric_card("Total Liabilities",money(fin["total_liabilities"]),"Balance Sheet"),metric_card("Total Equity",money(fin["total_equity"]),"Before current profit"),metric_card("Net Profit",money(fin["profit"]),"Profit for the period"),metric_card("Trial Balance",status,f"{review} review item(s)")])+"</div>",unsafe_allow_html=True)
    left,right=st.columns([1.55,1])
    with left:
        st.markdown('<div class="card"><div class="card-title">Financial overview</div><div class="card-caption">Current-period composition from the validated classifications.</div>',unsafe_allow_html=True)
        values=[("Revenue",fin["total_revenue"]),("Expenses",fin["total_expenses"]),("Assets",fin["total_assets"]),("Liabilities",fin["total_liabilities"]),("Equity",fin["total_equity"])]
        maxv=max([abs(v) for _,v in values] or [1])
        bars=''.join([f'<div style="display:grid;grid-template-columns:110px 1fr 105px;gap:.7rem;align-items:center;margin:.75rem 0"><div style="font-size:.72rem;color:var(--acc-muted);font-weight:750">{html_escape(k)}</div><div style="height:9px;background:#eceef5;border-radius:99px;overflow:hidden"><div style="height:100%;width:{abs(v)/maxv*100:.1f}%;background:linear-gradient(90deg,#5b5ce2,#16a6b6);border-radius:99px"></div></div><div style="text-align:right;font-size:.72rem;font-weight:800;color:var(--acc-ink)">{html_escape(money(v))}</div></div>' for k,v in values])
        st.markdown(bars+'</div>',unsafe_allow_html=True)
    with right:
        st.markdown('<div class="card"><div class="card-title">Accountra Intelligence</div><div class="card-caption">What deserves your attention right now.</div>',unsafe_allow_html=True)
        if abs(diff)<.01: st.markdown(insight_card("Trial Balance is balanced","Debit and credit totals agree. The core input check passed.","good"),unsafe_allow_html=True)
        else: st.markdown(insight_card("Trial Balance needs review",f"The difference is {money(diff)}. Review the uploaded ledger before relying on statements.","danger"),unsafe_allow_html=True)
        if review: st.markdown(insight_card(f"{review} account(s) need review","Open AI Review to inspect confidence, reasons and manual overrides.","warn"),unsafe_allow_html=True)
        else: st.markdown(insight_card("Classification review is clear","No accounts are currently flagged for manual review.","good"),unsafe_allow_html=True)
        if abs(fin["balance_difference"])<.01: st.markdown(insight_card("Balance Sheet tallies","Assets equal equity and liabilities after the current-period profit bridge.","good"),unsafe_allow_html=True)
        else: st.markdown(insight_card("Balance Sheet does not tally",f"Difference: {money(fin['balance_difference'])}. Review classifications and equity treatment.","danger"),unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)
    st.markdown('<div class="section-head"><div><div class="section-title">Quick actions</div><div class="section-caption">Jump directly into the part of the workflow you need.</div></div></div>',unsafe_allow_html=True)
    qa=st.columns(4)
    for col,key,title,sub in [(qa[0],"trial_balance","Trial Balance","Search and inspect accounts"),(qa[1],"ai_review","AI Review","Resolve exceptions"),(qa[2],"statements","Statements","Review P&L and Balance Sheet"),(qa[3],"reports","Reports","Export working papers")]:
        with col:
            if st.button(title,key=f"qa_{key}",use_container_width=True): st.session_state["app_page"]=key; st.rerun()
            st.caption(sub)

# -----------------------------
# Input panel
# -----------------------------

def render_input_panel():
    left,right=st.columns([1.5,1])
    with left:
        st.markdown('<div class="card"><div class="card-title">Upload your Trial Balance</div><div class="card-caption">Excel, XLS or CSV · Account, Debit and Credit columns.</div>',unsafe_allow_html=True)
        if "reset_nonce" not in st.session_state: st.session_state["reset_nonce"]=0
        uploaded=st.file_uploader("Choose Trial Balance",type=["xlsx","xls","csv"],key=f"trial_balance_upload_v5_{st.session_state['reset_nonce']}",label_visibility="collapsed")
        st.markdown('</div>',unsafe_allow_html=True)
    with right:
        st.markdown('<div class="card"><div class="card-title">Report context</div><div class="card-caption">These details flow into the financial statements and exports.</div>',unsafe_allow_html=True)
        st.text_input("Company / Entity Name",key="company_name")
        st.text_input("CIN / Registration No. (optional)",key="cin")
        st.date_input("Reporting Date",key="reporting_date")
        st.markdown('</div>',unsafe_allow_html=True)
    if uploaded:
        df,error=normalize_tb(uploaded)
        if error: st.error(error); return
        if uploaded.size>15*1024*1024: st.error("File is larger than 15 MB."); return
        token=f"tb_{uploaded.name}_{uploaded.size}"
        if st.session_state.get("file_token")!=token:
            st.session_state["file_token"]=token; st.session_state["prepared"]=False; st.session_state.pop("results",None); st.session_state.pop("comparative_results",None)
        st.session_state["source_df"]=df.to_dict("records")
        st.success(f"{uploaded.name} loaded · {len(df):,} accounts")
        c1,c2,c3=st.columns(3)
        c1.metric("Accounts",f"{len(df):,}"); c2.metric("Debit",money(df["Debit"].sum())); c3.metric("Credit",money(df["Credit"].sum()))
        diff=float(df["Debit"].sum()-df["Credit"].sum())
        if abs(diff)<.01: st.success("Trial Balance is balanced.")
        else: st.error(f"Trial Balance is not balanced · Difference {money(diff)}")
        with st.expander("Preview uploaded accounts",expanded=False): st.dataframe(df,use_container_width=True,hide_index=True)
        comparative=st.file_uploader("Previous-year Trial Balance (optional)",type=["xlsx","xls","csv"],key=f"comparative_trial_balance_upload_v5_{st.session_state['reset_nonce']}")
        comp_df=None
        if comparative:
            comp_df,_=normalize_tb(comparative)
            if comp_df is not None: st.session_state["comparative_source_df"]=comp_df.to_dict("records")
        can_prepare=abs(diff)<.01
        if st.button("Prepare financial workspace",type="primary",use_container_width=True,key="prepare_v5"):
            if not can_prepare: st.error("Cannot prepare financial statements until the Trial Balance balances.")
            else:
                with st.spinner("Preparing your accounting workspace…"):
                    res=prepare_classifications(df,comp_df)
                st.session_state["results"]=res.to_dict("records"); st.session_state["prepared"]=True; st.session_state["app_page"]="home"; st.rerun()

# -----------------------------
# Trial Balance page
# -----------------------------

def render_trial_balance():
    page_header("WORKSPACE","Trial Balance","Inspect the source accounts, balances and classifications in one high-density workspace.")
    if not st.session_state.get("prepared"):
        render_input_panel(); return
    df=apply_overrides(pd.DataFrame(st.session_state["results"]))
    st.session_state["results"]=df.to_dict("records")
    q=st.text_input("Search accounts",placeholder="Search account name…",key="tb_search")
    f1,f2,f3=st.columns([1,1,1.2])
    with f1: nature=st.selectbox("Nature",["All"]+sorted(df["Nature"].dropna().unique().tolist()),key="tb_nature")
    with f2: cls=st.selectbox("Classification",["All"]+sorted(df["Classification"].dropna().unique().tolist()),key="tb_class")
    with f3: sort=st.selectbox("Sort by",["Account","Debit","Credit","Confidence"],key="tb_sort")
    view=df.copy()
    if q: view=view[view["Account"].astype(str).str.contains(q,case=False,na=False)]
    if nature!="All": view=view[view["Nature"]==nature]
    if cls!="All": view=view[view["Classification"]==cls]
    view=view.sort_values(sort,ascending=(sort=="Account"),na_position="last")
    st.caption(f"Showing {len(view):,} of {len(df):,} accounts")
    render_tb_table(view)

# -----------------------------
# AI Review page
# -----------------------------

def render_ai_review():
    page_header("ACCOUNTRA INTELLIGENCE","AI Review","A focused exception queue for ambiguous, low-confidence or incomplete classifications.")
    if not st.session_state.get("prepared"): st.info("Prepare a Trial Balance first."); return
    df=apply_overrides(pd.DataFrame(st.session_state["results"]))
    st.session_state["results"]=df.to_dict("records")
    review=df[df["Ambiguous"].fillna(True) | (pd.to_numeric(df["Confidence"],errors="coerce").fillna(0)<.80)].copy()
    c1,c2,c3=st.columns(3)
    c1.metric("Accounts needing review",f"{len(review):,}"); c2.metric("High confidence",f"{int((pd.to_numeric(df['Confidence'],errors='coerce')>=.80).sum()):,}"); c3.metric("Missing information",f"{int(df['Missing Information'].fillna('').astype(str).str.strip().ne('').sum()):,}")
    if not len(review):
        st.markdown(insight_card("All clear","No accounts are currently below the review threshold.","good"),unsafe_allow_html=True); return
    for idx,(_,row) in enumerate(review.iterrows()):
        conf=float(row.get("Confidence",0) or 0); tone="danger" if conf<.5 else "warn"
        st.markdown(f'''<div class="card" style="margin:.7rem 0"><div style="display:flex;justify-content:space-between;gap:1rem"><div><div class="card-title">{html_escape(row['Account'])}</div><div class="card-caption">{html_escape(row['Nature'])} · {html_escape(row['Classification'])}</div></div><span class="tag tag-warn">{conf:.0%} confidence</span></div><div style="margin-top:.7rem;font-size:.78rem;color:var(--acc-muted)"><strong style="color:var(--acc-ink)">Why:</strong> {html_escape(row['Reason'])}</div><div style="margin-top:.35rem;font-size:.75rem;color:var(--acc-muted)"><strong style="color:var(--acc-ink)">Missing information:</strong> {html_escape(row['Missing Information'] or 'None specified')}</div></div>''',unsafe_allow_html=True)
        nature=row["Nature"]
        if nature=="Liability": options=["Current Borrowings","Non-current Borrowings"] if row["Classification"]=="Borrowings" else ["Trade Payables","Provisions","Other Current Liabilities","Other Non-current Liabilities"]
        elif nature=="Asset": options=sorted(APPROVED_ASSET_HEADS)
        elif nature=="Equity": options=sorted(APPROVED_EQUITY_HEADS)
        elif nature=="Income": options=sorted(APPROVED_INCOME_HEADS)
        elif nature=="Expense": options=sorted(APPROVED_EXPENSE_HEADS)
        else: options=[]
        if options:
            c1,c2=st.columns([2,1])
            with c1: newc=st.selectbox("Change classification",options,index=options.index(row["Classification"]) if row["Classification"] in options else 0,key=f"review_class_{idx}_{row['Account']}")
            with c2:
                if st.button("Apply",key=f"review_apply_{idx}",use_container_width=True):
                    st.session_state[f"override_{row['Account']}"]=newc; st.rerun()

# -----------------------------
# Statements page
# -----------------------------

def render_statements():
    page_header("REPORTING","Financial Statements","Schedule III-style Profit & Loss and Balance Sheet presentation from the validated classifications.")
    if not st.session_state.get("prepared"): st.info("Prepare a Trial Balance first."); return
    df=apply_overrides(pd.DataFrame(st.session_state["results"])); st.session_state["results"]=df.to_dict("records"); fin=compute_financials(df)
    a,b=st.columns(2)
    with a: st.markdown(metric_card("Profit Before Tax",money(fin["pbt"]),"Current period"),unsafe_allow_html=True)
    with b: st.markdown(metric_card("Balance Sheet Difference",money(fin["balance_difference"]),"Should be ₹0"),unsafe_allow_html=True)
    st.markdown('<div class="section-head"><div><div class="section-title">Statement of Profit & Loss</div><div class="section-caption">Current period with previous period where supplied.</div></div></div>',unsafe_allow_html=True)
    render_statement_table("STATEMENT OF PROFIT & LOSS",f"For the period ended {st.session_state['reporting_date'].strftime('%d %B %Y')}",fin["pnl_rows"])
    st.markdown('<div class="section-head"><div><div class="section-title">Balance Sheet</div><div class="section-caption">Equity and liabilities compared with assets.</div></div></div>',unsafe_allow_html=True)
    render_statement_table("BALANCE SHEET",f"As at {st.session_state['reporting_date'].strftime('%d %B %Y')}",fin["bs_rows"])
    if abs(fin["balance_difference"])<.01: st.success("Balance Sheet tallies.")
    else: st.error(f"Balance Sheet does not tally · Difference {money(fin['balance_difference'])}")
    st.markdown('<div class="section-head"><div><div class="section-title">Notes to Accounts</div><div class="section-caption">Presentation assumptions and account composition.</div></div></div>',unsafe_allow_html=True)
    for note,detail in fin["notes"]: st.markdown(insight_card(note,detail,"warn"),unsafe_allow_html=True)
    if len(fin["movement_df"]):
        st.markdown('<div class="section-head"><div><div class="section-title">Comparative movement</div><div class="section-caption">Material movement threshold applied from Report Settings.</div></div></div>',unsafe_allow_html=True)
        mv=fin["movement_df"].copy(); mv["Current Period"]=mv["Current Period"].map(money); mv["Previous Period"]=mv["Previous Period"].map(money); mv["Change"]=mv["Change"].map(money); mv["Change %"]=mv["Change %"].map(lambda x:"—" if pd.isna(x) else f"{x:+.1f}%"); st.dataframe(mv,use_container_width=True,hide_index=True)

# -----------------------------
# Validation page
# -----------------------------

def render_validation():
    page_header("CONTROL CENTER","Validation","A single place to verify the Trial Balance, classification quality, P&L reconciliation and Balance Sheet tally.")
    if not st.session_state.get("prepared"): st.info("Prepare a Trial Balance first."); return
    df=apply_overrides(pd.DataFrame(st.session_state["results"])); st.session_state["results"]=df.to_dict("records"); fin=compute_financials(df); checks=validation_check_rows(df,fin)
    passed=sum(1 for _,s,_ in checks if s=="PASS")
    tone="good" if passed==len(checks) else "warn"
    st.markdown(f'<div class="hero"><div class="hero-content"><div class="page-eyebrow">VALIDATION SCORE</div><div class="page-title">{passed}/{len(checks)} checks passed</div><div class="page-subtitle">Core accounting controls are shown below. Review any item marked REVIEW before relying on the generated reports.</div></div></div>',unsafe_allow_html=True)
    for name,status,detail in checks:
        cls="status-good" if status=="PASS" else "status-warn"
        st.markdown(f'<div class="check"><div><div class="check-name">{html_escape(name)}</div><div class="check-detail">{html_escape(detail)}</div></div><span class="status-pill {cls}">{html_escape(status)}</span></div>',unsafe_allow_html=True)
    if fin["pbt"]<0 and fin["tax_expense"]>0: st.warning("PBT is negative while tax expense is present. Review whether the tax amount represents deferred tax or another applicable adjustment.")

# -----------------------------
# Reports page
# -----------------------------

def render_reports():
    page_header("DELIVERABLES","Reports & Export","Download the working paper and financial statements after reviewing the validation center.")
    if not st.session_state.get("prepared"): st.info("Prepare a Trial Balance first."); return
    df=apply_overrides(pd.DataFrame(st.session_state["results"])); st.session_state["results"]=df.to_dict("records"); fin=compute_financials(df)
    excel,pdf,checks=build_export_artifacts(df,fin)
    safe=re.sub(r"[^A-Za-z0-9_-]+","_",st.session_state["company_name"].strip() or "Accountra").strip("_")
    a,b=st.columns(2)
    with a:
        st.markdown('<div class="card report-card"><div><div class="report-icon">X</div><div class="card-title" style="margin-top:.7rem">Excel working paper</div><div class="card-caption">Classified Trial Balance, P&L, Balance Sheet, validation and notes.</div></div></div>',unsafe_allow_html=True)
        st.download_button("Download Excel",data=excel,file_name=f"{safe}_Schedule_III_Working_Paper.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="report_excel_v5")
    with b:
        st.markdown('<div class="card report-card"><div><div class="report-icon">PDF</div><div class="card-title" style="margin-top:.7rem">PDF financial statements</div><div class="card-caption">Presentation-ready P&L, Balance Sheet, validation and notes.</div></div></div>',unsafe_allow_html=True)
        st.download_button("Download PDF",data=pdf,file_name=f"{safe}_Financial_Statements.pdf",mime="application/pdf",use_container_width=True,key="report_pdf_v5")
    st.markdown('<div class="section-head"><div><div class="section-title">Export readiness</div><div class="section-caption">Reports reflect the current validated session state.</div></div></div>',unsafe_allow_html=True)
    if all(s=="PASS" for _,s,_ in checks): st.success("All validation checks passed. Reports are ready for review.")
    else: st.warning("Reports are available, but one or more validation checks require attention.")

# -----------------------------
# Settings page
# -----------------------------

def render_settings():
    page_header("MANAGE","Report Settings","Control the entity details and comparative review threshold used throughout the workspace.")
    st.text_input("Company / Entity Name",key="company_name")
    st.text_input("CIN / Registration No. (optional)",key="cin")
    st.date_input("Reporting Date",key="reporting_date")
    st.number_input("Comparative movement review threshold (%)",min_value=1.0,max_value=100.0,step=5.0,key="materiality_threshold")
    st.caption(f"Financial year: {financial_year_label(st.session_state['reporting_date'])}")
    st.markdown('<div class="section-head"><div><div class="section-title">Session controls</div><div class="section-caption">Reset removes uploaded and generated accounting state while keeping these report settings.</div></div></div>',unsafe_allow_html=True)
    if st.button("Reset accounting workspace",type="secondary",use_container_width=False,key="settings_reset"):
        clear_accounting_session(); st.session_state["app_page"]="home"; st.rerun()

# -----------------------------
# Route — exactly one workspace
# -----------------------------

page=st.session_state.get("app_page","home")
if page=="home": render_home()
elif page=="trial_balance": render_trial_balance()
elif page=="ai_review": render_ai_review()
elif page=="statements": render_statements()
elif page=="validation": render_validation()
elif page=="reports": render_reports()
elif page=="settings": render_settings()
else:
    st.session_state["app_page"]="home"; st.rerun()

st.markdown('<div style="text-align:center;color:var(--acc-muted);font-size:.68rem;margin-top:3rem">Accountra · AI-assisted accounting workflow · Review classifications and statutory disclosures before filing.</div>',unsafe_allow_html=True)
