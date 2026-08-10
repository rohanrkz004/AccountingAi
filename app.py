import json
import os
import re
from io import BytesIO
from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st
from openai import OpenAI

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)


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
    cover["A1"] = company_name
    cover["A1"].font = Font(size=20, bold=True, color="FFFFFF")
    cover["A1"].fill = PatternFill("solid", fgColor="17365D")
    cover.merge_cells("A1:D2")
    cover["A4"] = "Accounting AI — Schedule III-style Financial Statements"
    cover["A4"].font = Font(size=14, bold=True)
    cover["A5"] = f"Reporting date: {reporting_date.strftime('%d %B %Y')}"
    cover["A6"] = f"CIN: {cin}" if cin else "CIN: Not provided"
    cover["A8"] = "Important"
    cover["A8"].font = Font(bold=True)
    cover["A9"] = (
        "This is an AI-assisted accounting preparation report. "
        "Review classifications, notes and statutory disclosures before filing."
    )
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

    story = [
        Paragraph(company_name, title_style),
        Paragraph(
            f"Schedule III-style Financial Statements<br/>"
            f"As at / for the period ended {reporting_date.strftime('%d %B %Y')}"
            + (f"<br/>CIN: {cin}" if cin else ""),
            subtitle_style,
        ),
    ]

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


def _get_secret(name, default=None):
    """Read a Streamlit secret, falling back to an environment variable."""
    try:
        value = st.secrets.get(name)
        if value:
            return value
    except Exception:
        pass

    return os.getenv(name, default)


def classify_account_ai(account, debit, credit):
    """
    Cloud-safe AI fallback.

    Deterministic rules run first. This function is only called for
    accounts that the rulebook could not classify. If no API key is
    configured, the app safely falls back to manual review.
    """
    api_key = _get_secret("OPENAI_API_KEY")
    model = _get_secret("OPENAI_MODEL", "gpt-5")

    if not api_key:
        return None

    prompt = f"""
You are an accounting classification assistant for Indian companies.

Classify this ledger account for financial statement preparation.

Account: {account}
Debit: ₹{debit}
Credit: ₹{credit}

Approved Balance Sheet Assets:
- PPE
- Capital Work-in-Progress
- Intangible Assets
- Intangible Assets Under Development
- Investment Property
- Investments
- Inventories
- Trade Receivables
- Cash & Cash Equivalents
- Other Current Assets
- Other Non-current Assets

Approved Equity:
- Share Capital
- Other Equity
- Capital Account

Approved Liabilities:
- Borrowings
- Current Borrowings
- Non-current Borrowings
- Trade Payables
- Provisions
- Other Current Liabilities
- Other Non-current Liabilities

Approved Profit & Loss:
- Revenue from Operations
- Other Income
- Cost of Materials Consumed
- Purchases
- Changes in Inventories
- Employee Benefits Expense
- Finance Costs
- Depreciation & Amortisation
- Other Expenses
- Tax Expense

Rules:
- Never invent facts.
- If current/non-current information is missing for a borrowing,
  use Borrowings and set ambiguous=true.
- If the account is a payable/receivable, classify it according to
  what is owed/receivable.
- Do not infer maturity merely from the word "loan".
- Return only JSON.
- Confidence must be between 0 and 1.
- Use only one of the approved classifications above.

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

    try:
        client = OpenAI(api_key=api_key)

        response = client.responses.create(
            model=model,
            input=prompt,
            text={"format": {"type": "json_object"}},
        )

        return json.loads(response.output_text)

    except Exception as exc:
        print(f"AI classification error for {account}: {exc}")
        return None



# =========================================================
# STREAMLIT APP — SCHEDULE III PRESENTATION
# =========================================================

st.set_page_config(
    page_title="AccountingAI",
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

    .validation-pass {
        padding: .7rem .9rem;
        border-radius: 10px;
        border: 1px solid rgba(34,197,94,.25);
        margin: .35rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-hero">
        <h1>📊 Accounting AI</h1>
        <p>Trial Balance → AI classification → Schedule III-style financial statements → validation</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# COMPARATIVE / ANALYTICS HELPERS
# =========================================================

def classify_uploaded_tb(tb_df, allow_ai=True):
    """Classify a Trial Balance for current or comparative analysis."""
    work = tb_df.copy()

    work["Account"] = work["Account"].astype(str).str.strip()
    work["Debit"] = clean_number_series(work["Debit"])
    work["Credit"] = clean_number_series(work["Credit"])

    rows = []

    for _, row in work.iterrows():
        account = str(row["Account"]).strip()
        debit = float(row["Debit"] or 0)
        credit = float(row["Credit"] or 0)

        result = classify_account(account, debit, credit)

        if result is None and allow_ai:
            try:
                result = classify_account_ai(
                    account,
                    debit,
                    credit,
                )
            except Exception:
                result = None

        if result is None:
            result = make_result(
                "Unknown",
                "NEEDS_REVIEW",
                "NEEDS_REVIEW",
                "Unable to classify this comparative account.",
                ambiguous=True,
                confidence=0,
                missing_information="Manual review required.",
            )

        nature = result.get("nature", "Unknown")
        classification = result.get(
            "classification",
            "NEEDS_REVIEW",
        )
        statement = result.get(
            "statement",
            "NEEDS_REVIEW",
        )

        if classification not in APPROVED_HEADS:
            classification = "NEEDS_REVIEW"
            statement = "NEEDS_REVIEW"
            nature = "Unknown"

        elif classification in (
            APPROVED_INCOME_HEADS
            | APPROVED_EXPENSE_HEADS
        ):
            statement = "Profit & Loss"
        else:
            statement = "Balance Sheet"

        rows.append({
            "Account": account,
            "Debit": debit,
            "Credit": credit,
            "Nature": nature,
            "Classification": classification,
            "Statement": statement,
            "Ambiguous": bool(result.get("ambiguous", True)),
            "Confidence": float(result.get("confidence", 0) or 0),
            "Reason": result.get("reason", "") or "",
            "Missing Information": result.get(
                "missing_information"
            ) or "",
        })

    return pd.DataFrame(rows)


def summarize_by_classification(results):
    """Return signed current/prior values by classification."""
    if results is None or results.empty:
        return pd.DataFrame(
            columns=["Classification", "Amount"]
        )

    rows = []

    for _, row in results.iterrows():
        classification = row["Classification"]
        debit = float(row["Debit"] or 0)
        credit = float(row["Credit"] or 0)

        if classification in APPROVED_ASSET_HEADS:
            amount = debit - credit
        elif classification in (
            APPROVED_LIABILITY_HEADS
            | APPROVED_EQUITY_HEADS
            | APPROVED_INCOME_HEADS
        ):
            amount = credit - debit
        elif classification in APPROVED_EXPENSE_HEADS:
            amount = debit - credit
        else:
            amount = 0.0

        rows.append((classification, amount))

    summary = (
        pd.DataFrame(rows, columns=["Classification", "Amount"])
        .groupby("Classification", as_index=False)["Amount"]
        .sum()
    )

    return summary


def indian_percentage(value):
    try:
        return f"{float(value):.1f}%"
    except Exception:
        return "—"



# =========================================================
# SIDEBAR — REPORTING INFORMATION
# =========================================================

with st.sidebar:
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

    reporting_date = st.date_input(
        "Reporting Date",
        value=date.today(),
        key="reporting_date",
    )

    business_nature = st.selectbox(
        "Nature of Business",
        [
            "Trading",
            "Manufacturing",
            "Services",
            "Mixed / Other",
        ],
        key="business_nature",
    )

    materiality_threshold = st.number_input(
        "Movement review threshold (%)",
        min_value=1.0,
        max_value=100.0,
        value=20.0,
        step=5.0,
        key="materiality_threshold",
        help="Used for comparative movement flags.",
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

    ai_configured = bool(_get_secret("OPENAI_API_KEY"))

    if ai_configured:
        st.success("🤖 AI fallback: Connected")
    else:
        st.info(
            "🤖 AI fallback: Not configured — deterministic rules "
            "and manual review remain available."
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

# =========================================================
# STEP 1 — UPLOAD
# =========================================================

st.markdown("## 1️⃣ Upload Trial Balance")
st.caption(
    "Upload an Excel or CSV Trial Balance containing Account, Debit and Credit columns."
)

uploaded_file = st.file_uploader(
    "Choose a file",
    type=["xlsx", "xls", "csv"],
    help="Supported formats: .xlsx, .xls and .csv",
    key="trial_balance_upload",
)

if uploaded_file:
    # A newly uploaded current-year TB starts a fresh comparison context.
    st.session_state["has_comparative"] = False
    st.session_state.pop("comparative_results", None)

    st.success(f"Uploaded: **{uploaded_file.name}**")

    st.markdown("### 📊 Optional Comparative Trial Balance")
    comparative_file = st.file_uploader(
        "Upload previous-year Trial Balance (optional)",
        type=["xlsx", "xls", "csv"],
        help="Optional: used to show year-on-year movement analysis.",
        key="comparative_trial_balance_upload",
    )

    file_name = uploaded_file.name.lower()

    if uploaded_file.size and uploaded_file.size > 15 * 1024 * 1024:
        st.error(
            "File is larger than 15 MB. Please upload a smaller Trial Balance."
        )
        st.stop()

    try:
        if file_name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as upload_error:
        st.error(
            "Could not read the uploaded file. "
            "Please check that it is a valid Excel/CSV Trial Balance."
        )
        st.exception(upload_error)
        st.stop()

    # -----------------------------------------------------
    # NORMALIZE COLUMNS
    # -----------------------------------------------------

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace("₹", "", regex=False)
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )

    df = df.rename(
        columns={
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
        }
    )

    required_columns = {"Account", "Debit", "Credit"}

    if not required_columns.issubset(df.columns):
        st.error(
            "The file must contain Account, Debit and Credit columns. "
            f"Detected: {df.columns.tolist()}"
        )
        st.stop()

    # -----------------------------------------------------
    # REMOVE TOTAL / SUBTOTAL ROWS
    # -----------------------------------------------------

    total_row_names = {
        "total",
        "grand total",
        "trial balance total",
        "subtotal",
        "total trial balance",
    }

    df = df[
        ~df["Account"]
        .astype(str)
        .str.strip()
        .str.lower()
        .isin(total_row_names)
    ].copy()

    # -----------------------------------------------------
    # NUMERIC CLEANING — BEFORE ANY CALCULATION
    # -----------------------------------------------------

    df["Debit"] = clean_number_series(df["Debit"])
    df["Credit"] = clean_number_series(df["Credit"])

    # -----------------------------------------------------
    # OPTIONAL COMPARATIVE FILE
    # -----------------------------------------------------

    comparative_df = None

    if comparative_file is not None:

        if (
            comparative_file.size
            and comparative_file.size > 15 * 1024 * 1024
        ):
            st.error(
                "Comparative file is larger than 15 MB."
            )
            st.stop()

        try:
            comparative_name = comparative_file.name.lower()

            if comparative_name.endswith(".csv"):
                comparative_df = pd.read_csv(comparative_file)
            else:
                comparative_df = pd.read_excel(comparative_file)

            comparative_df.columns = (
                comparative_df.columns.astype(str)
                .str.strip()
                .str.lower()
                .str.replace("₹", "", regex=False)
                .str.replace("(", "", regex=False)
                .str.replace(")", "", regex=False)
                .str.strip()
            )

            comparative_df = comparative_df.rename(
                columns={
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
                }
            )

            if not {"Account", "Debit", "Credit"}.issubset(
                comparative_df.columns
            ):
                st.warning(
                    "Comparative file was uploaded but does not contain "
                    "Account, Debit and Credit columns. "
                    "Comparative analysis will be skipped."
                )
                comparative_df = None
            else:
                comparative_df = comparative_df[
                    ~comparative_df["Account"]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    .isin({
                        "total",
                        "grand total",
                        "trial balance total",
                        "subtotal",
                        "total trial balance",
                    })
                ].copy()

                comparative_df["Debit"] = clean_number_series(
                    comparative_df["Debit"]
                )
                comparative_df["Credit"] = clean_number_series(
                    comparative_df["Credit"]
                )

                st.success(
                    f"Comparative TB loaded: **{comparative_file.name}**"
                )

        except Exception as comparative_error:
            st.warning(
                "Could not read the comparative Trial Balance. "
                "The current-year workflow will continue."
            )
            comparative_df = None

    # -----------------------------------------------------
    # PREVIEW
    # -----------------------------------------------------

    st.markdown("### 📄 Trial Balance Preview")

    preview_col1, preview_col2, preview_col3 = st.columns(3)

    with preview_col1:
        st.metric("Accounts", f"{len(df):,}")

    with preview_col2:
        st.metric("Total Debit", indian_currency(df["Debit"].sum()))

    with preview_col3:
        st.metric("Total Credit", indian_currency(df["Credit"].sum()))

    st.dataframe(
        df,
        width="stretch",
        hide_index=True,
    )

    # -----------------------------------------------------
    # TRIAL BALANCE VALIDATION
    # -----------------------------------------------------

    total_debit = float(df["Debit"].sum())
    total_credit = float(df["Credit"].sum())
    difference = total_debit - total_credit

    st.markdown("### ⚖️ Trial Balance Validation")

    if abs(difference) < 0.01:
        st.success(
            f"**Trial Balance is balanced ✅**  "
            f"Debit: {indian_currency(total_debit)}  |  "
            f"Credit: {indian_currency(total_credit)}"
        )
    else:
        st.error(
            f"**Trial Balance is NOT balanced ❌**  "
            f"Debit: {indian_currency(total_debit)}  |  "
            f"Credit: {indian_currency(total_credit)}  |  "
            f"Difference: {indian_currency(abs(difference))}"
        )

    # =====================================================
    # STEP 2 — PREPARE FINANCIAL STATEMENTS
    # =====================================================

    st.divider()
    st.markdown("## 2️⃣ Prepare Financial Statements")

    if st.button(
        "📄 Prepare Financial Statements",
        key="prepare_fs_button_v2",
        type="primary",
        use_container_width=True,
    ):

        if abs(difference) >= 0.01:
            st.error(
                "Cannot prepare financial statements because "
                "the Trial Balance is not balanced."
            )
        else:
            st.session_state["prepared"] = True
            st.session_state.pop("results", None)

            # Fresh preparation = clear old overrides.
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

                result = classify_account(
                    account,
                    debit,
                    credit,
                )

                if result is None:
                    st.write(
                        f"🤖 AI review required: **{account}**"
                    )
                    try:
                        result = classify_account_ai(
                            account,
                            debit,
                            credit,
                        )
                    except Exception as ai_error:
                        st.warning(
                            f"AI review failed for **{account}**. "
                            "The account has been marked for manual review."
                        )
                        result = None

                if result is None:
                    result = make_result(
                        "Unknown",
                        "NEEDS_REVIEW",
                        "NEEDS_REVIEW",
                        "Unable to classify the account.",
                        ambiguous=True,
                        confidence=0,
                        missing_information="Manual review required.",
                    )

                nature = result.get("nature", "Unknown")
                classification = result.get(
                    "classification",
                    "NEEDS_REVIEW",
                )
                statement = result.get(
                    "statement",
                    "NEEDS_REVIEW",
                )
                ambiguous = bool(
                    result.get("ambiguous", True)
                )
                confidence = float(
                    result.get("confidence", 0) or 0
                )
                reason = result.get("reason", "") or ""
                missing_information = result.get(
                    "missing_information"
                )

                # -------------------------------------------------
                # HARD VALIDATION OF AI OUTPUT
                # -------------------------------------------------

                if classification not in APPROVED_HEADS:

                    classification = "NEEDS_REVIEW"
                    statement = "NEEDS_REVIEW"
                    ambiguous = True
                    confidence = min(
                        confidence,
                        0.50,
                    )
                    reason = (
                        "AI returned an unapproved classification."
                    )
                    missing_information = (
                        "Manual classification required."
                    )

                elif (
                    classification
                    in APPROVED_EXPENSE_HEADS
                    | APPROVED_INCOME_HEADS
                ):

                    statement = "Profit & Loss"

                else:

                    statement = "Balance Sheet"

                results.append(
                    {
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
                    }
                )

                progress.progress(
                    (index + 1) / total_rows
                )

            results_df = pd.DataFrame(results)

            results_df["Debit"] = clean_number_series(
                results_df["Debit"]
            )
            results_df["Credit"] = clean_number_series(
                results_df["Credit"]
            )

            st.session_state["results"] = (
                results_df.to_dict("records")
            )

            # Prepare optional comparative results once, so they are
            # available to the movement-analysis dashboard.
            if comparative_df is not None:
                comparative_results = classify_uploaded_tb(
                    comparative_df,
                    allow_ai=True,
                )
                st.session_state["comparative_results"] = (
                    comparative_results.to_dict("records")
                )
                st.session_state["has_comparative"] = True
            else:
                st.session_state.pop("comparative_results", None)
                st.session_state["has_comparative"] = False

            st.success(
                "AI analysis completed successfully! ✅"
            )


# =========================================================
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
    # COMPARATIVE ANALYSIS
    # =====================================================

    if (
        st.session_state.get("has_comparative", False)
        and "comparative_results" in st.session_state
    ):

        comparative_results_df = pd.DataFrame(
            st.session_state["comparative_results"]
        )

        current_summary = summarize_by_classification(
            results_df
        ).rename(columns={"Amount": "Current"})

        prior_summary = summarize_by_classification(
            comparative_results_df
        ).rename(columns={"Amount": "Previous"})

        movement_df = pd.merge(
            current_summary,
            prior_summary,
            on="Classification",
            how="outer",
        ).fillna(0)

        movement_df["Change"] = (
            movement_df["Current"]
            - movement_df["Previous"]
        )

        def pct_change(row):
            previous = float(row["Previous"])
            change = float(row["Change"])

            if abs(previous) < 0.01:
                return None

            return (change / abs(previous)) * 100

        movement_df["Change %"] = movement_df.apply(
            pct_change,
            axis=1,
        )

        movement_df["Material Movement"] = movement_df[
            "Change %"
        ].apply(
            lambda x:
                bool(
                    x is not None
                    and abs(x) >= materiality_threshold
                )
        )

        st.divider()
        st.markdown("## 📈 Comparative Financial Analysis")

        movement_count = int(
            movement_df["Material Movement"].sum()
        )

        metric_cols = st.columns(3)

        with metric_cols[0]:
            st.metric(
                "Current-Year Heads",
                len(current_summary),
            )

        with metric_cols[1]:
            st.metric(
                "Comparative-Year Heads",
                len(prior_summary),
            )

        with metric_cols[2]:
            st.metric(
                "Significant Movements",
                movement_count,
            )

        display_movement = movement_df.copy()

        display_movement["Current"] = display_movement[
            "Current"
        ].map(indian_currency)

        display_movement["Previous"] = display_movement[
            "Previous"
        ].map(indian_currency)

        display_movement["Change"] = display_movement[
            "Change"
        ].map(indian_currency)

        display_movement["Change %"] = display_movement[
            "Change %"
        ].apply(
            lambda x:
                "—"
                if x is None
                else indian_percentage(x)
        )

        display_movement = display_movement[
            [
                "Classification",
                "Previous",
                "Current",
                "Change",
                "Change %",
                "Material Movement",
            ]
        ]

        st.dataframe(
            display_movement.sort_values(
                "Material Movement",
                ascending=False,
            ),
            width="stretch",
            hide_index=True,
        )

        significant_df = movement_df[
            movement_df["Material Movement"]
        ].copy()

        if len(significant_df):
            st.markdown("### ⚠️ Significant Movements")

            for _, movement in significant_df.iterrows():

                pct = movement["Change %"]

                direction = (
                    "increased"
                    if movement["Change"] > 0
                    else "decreased"
                )

                st.write(
                    f"**{movement['Classification']}** "
                    f"{direction} by "
                    f"**{indian_currency(abs(movement['Change']))}** "
                    f"({indian_percentage(abs(pct))})."
                )
        else:
            st.success(
                "No movements crossed the selected review threshold."
            )


    # =====================================================
    # SCHEDULE III — STATEMENT OF PROFIT & LOSS
    # =====================================================

    st.divider()
    st.markdown("## 5️⃣ Statement of Profit & Loss")

    st.markdown(
        f"""
        <div class="section-card">
            <div class="fs-title">
                {company_name}
            </div>
            <div class="fs-subtitle">
                Statement of Profit and Loss for the period ended
                {reporting_date.strftime("%d %B %Y")}
            </div>
            <div class="schedule-note">
                Schedule III-style presentation. The mapping is based on
                the classifications generated by the accounting engine.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pnl_df = results_df[
        results_df["Statement"]
        == "Profit & Loss"
    ].copy()

    revenue_rows = pnl_df[
        pnl_df["Classification"].isin(
            APPROVED_INCOME_HEADS
        )
    ]

    revenue_summary = (
        revenue_rows
        .groupby("Classification")[
            ["Debit", "Credit"]
        ]
        .sum()
        if len(revenue_rows)
        else pd.DataFrame(
            columns=["Debit", "Credit"]
        )
    )

    if len(revenue_summary):

        revenue_summary["Net"] = (
            revenue_summary["Credit"]
            - revenue_summary["Debit"]
        )

    total_revenue = (
        float(revenue_summary["Net"].sum())
        if len(revenue_summary)
        else 0.0
    )

    # Tax is presented separately below Profit Before Tax.
    pre_tax_expense_heads = (
        APPROVED_EXPENSE_HEADS
        - {"Tax Expense"}
    )

    expense_rows = pnl_df[
        pnl_df["Classification"].isin(
            pre_tax_expense_heads
        )
    ]

    expense_summary = (
        expense_rows
        .groupby("Classification")[
            ["Debit", "Credit"]
        ]
        .sum()
        if len(expense_rows)
        else pd.DataFrame(
            columns=["Debit", "Credit"]
        )
    )

    if len(expense_summary):

        expense_summary["Net"] = (
            expense_summary["Debit"]
            - expense_summary["Credit"]
        )

    total_expenses = (
        float(expense_summary["Net"].sum())
        if len(expense_summary)
        else 0.0
    )

    tax_rows = pnl_df[
        pnl_df["Classification"]
        == "Tax Expense"
    ]

    tax_summary = (
        tax_rows[["Debit", "Credit"]].sum()
        if len(tax_rows)
        else pd.Series(
            {
                "Debit": 0.0,
                "Credit": 0.0,
            }
        )
    )

    tax_expense = max(
        0.0,
        float(
            tax_summary["Debit"]
            - tax_summary["Credit"]
        ),
    )

    profit_before_tax = (
        total_revenue
        - total_expenses
    )

    profit = (
        profit_before_tax
        - tax_expense
    )

    # -----------------------------------------------------
    # DISPLAY — VERTICAL, TOP TO BOTTOM
    # -----------------------------------------------------

    st.markdown(
        "### I. Revenue from Operations & Other Income"
    )

    revenue_ops = float(
        revenue_summary.loc[
            "Revenue from Operations",
            "Net",
        ]
    ) if "Revenue from Operations" in revenue_summary.index else 0.0

    other_income = float(
        revenue_summary.loc[
            "Other Income",
            "Net",
        ]
    ) if "Other Income" in revenue_summary.index else 0.0

    st.write(
        f"**1. Revenue from Operations** — "
        f"{indian_currency(revenue_ops)}"
    )

    st.write(
        f"**2. Other Income** — "
        f"{indian_currency(other_income)}"
    )

    st.markdown(
        f"**Total Revenue** — "
        f"{indian_currency(total_revenue)}"
    )

    st.divider()

    st.markdown("### II. Expenses")

    pnl_order = [
        "Cost of Materials Consumed",
        "Purchases",
        "Changes in Inventories",
        "Employee Benefits Expense",
        "Finance Costs",
        "Depreciation & Amortisation",
        "Other Expenses",
    ]

    for head in pnl_order:
        amount = (
            float(
                expense_summary.loc[
                    head,
                    "Net",
                ]
            )
            if head in expense_summary.index
            else 0.0
        )

        if abs(amount) > 0.005:
            st.write(
                f"**{head}** — "
                f"{indian_currency(amount)}"
            )

    st.markdown(
        f"**Total Expenses** — "
        f"{indian_currency(total_expenses)}"
    )

    st.divider()

    st.write(
        f"**Profit Before Tax** — "
        f"{indian_currency(profit_before_tax)}"
    )

    st.write(
        f"**Tax Expense** — "
        f"{indian_currency(tax_expense)}"
    )

    st.markdown(
        f"### **Profit for the Period — "
        f"{indian_currency(profit)}**"
    )

    st.divider()

    # =====================================================
    # SCHEDULE III — BALANCE SHEET
    # =====================================================

    st.divider()
    st.markdown("## 6️⃣ Balance Sheet")

    st.markdown(
        f"""
        <div class="section-card">
            <div class="fs-title">
                {company_name}
            </div>
            <div class="fs-subtitle">
                Balance Sheet as at
                {reporting_date.strftime("%d %B %Y")}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -----------------------------------------------------
    # BALANCE SHEET SUMMARIES
    # -----------------------------------------------------

    asset_rows = results_df[
        results_df["Classification"].isin(
            APPROVED_ASSET_HEADS
        )
    ]

    asset_summary = (
        asset_rows
        .groupby("Classification")[
            ["Debit", "Credit"]
        ]
        .sum()
        if len(asset_rows)
        else pd.DataFrame(
            columns=["Debit", "Credit"]
        )
    )

    if len(asset_summary):
        asset_summary["Net"] = (
            asset_summary["Debit"]
            - asset_summary["Credit"]
        )

    liability_rows = results_df[
        results_df["Classification"].isin(
            APPROVED_LIABILITY_HEADS
        )
    ]

    liability_summary = (
        liability_rows
        .groupby("Classification")[
            ["Debit", "Credit"]
        ]
        .sum()
        if len(liability_rows)
        else pd.DataFrame(
            columns=["Debit", "Credit"]
        )
    )

    if len(liability_summary):
        liability_summary["Net"] = (
            liability_summary["Credit"]
            - liability_summary["Debit"]
        )

    equity_rows = results_df[
        results_df["Classification"].isin(
            APPROVED_EQUITY_HEADS
        )
    ]

    equity_summary = (
        equity_rows
        .groupby("Classification")[
            ["Debit", "Credit"]
        ]
        .sum()
        if len(equity_rows)
        else pd.DataFrame(
            columns=["Debit", "Credit"]
        )
    )

    if len(equity_summary):
        equity_summary["Net"] = (
            equity_summary["Credit"]
            - equity_summary["Debit"]
        )

    total_assets = (
        float(asset_summary["Net"].sum())
        if len(asset_summary)
        else 0.0
    )

    total_liabilities = (
        float(liability_summary["Net"].sum())
        if len(liability_summary)
        else 0.0
    )

    total_equity = (
        float(equity_summary["Net"].sum())
        if len(equity_summary)
        else 0.0
    )

    total_equity_and_liabilities = (
        total_equity
        + profit
        + total_liabilities
    )

    # -----------------------------------------------------
    # SCHEDULE III DISPLAY GROUPS
    # -----------------------------------------------------

    non_current_asset_groups = [
        (
            "Property, Plant and Equipment",
            ["PPE"],
        ),
        (
            "Intangible Assets",
            ["Intangible Assets"],
        ),
        (
            "Capital Work-in-Progress",
            ["Capital Work-in-Progress"],
        ),
        (
            "Intangible Assets under Development",
            ["Intangible Assets Under Development"],
        ),
        (
            "Investment Property",
            ["Investment Property"],
        ),
        (
            "Non-current Investments",
            ["Investments"],
        ),
        (
            "Other Non-current Assets",
            ["Other Non-current Assets"],
        ),
    ]

    current_asset_groups = [
        (
            "Inventories",
            ["Inventories"],
        ),
        (
            "Trade Receivables",
            ["Trade Receivables"],
        ),
        (
            "Cash and Cash Equivalents",
            ["Cash & Cash Equivalents"],
        ),
        (
            "Other Current Assets",
            ["Other Current Assets"],
        ),
    ]

    non_current_liability_groups = [
        (
            "Long-term Borrowings",
            ["Non-current Borrowings"],
        ),
        (
            "Other Long-term Liabilities",
            ["Other Non-current Liabilities"],
        ),
    ]

    current_liability_groups = [
        (
            "Short-term Borrowings",
            ["Current Borrowings"],
        ),
        (
            "Trade Payables",
            ["Trade Payables"],
        ),
        (
            "Other Current Liabilities",
            ["Other Current Liabilities"],
        ),
        (
            "Short-term Provisions",
            ["Provisions"],
        ),
    ]

    def group_amount(summary, classifications):
        total = 0.0
        for classification in classifications:
            if classification in summary.index:
                total += float(
                    summary.loc[
                        classification,
                        "Net",
                    ]
                )
        return total

    # -----------------------------------------------------
    # DISPLAY
    # -----------------------------------------------------

    # -----------------------------------------------------
    # DISPLAY — VERTICAL, TOP TO BOTTOM
    # -----------------------------------------------------

    st.markdown("### I. Equity and Liabilities")

    st.markdown("**1. Shareholders' Funds**")

    share_capital = group_amount(
        equity_summary,
        ["Share Capital"],
    )

    other_equity = group_amount(
        equity_summary,
        ["Other Equity", "Capital Account"],
    )

    st.write(
        f"Share Capital — "
        f"{indian_currency(share_capital)}"
    )

    st.write(
        f"Reserves and Surplus / Other Equity — "
        f"{indian_currency(other_equity)}"
    )

    shareholders_funds = share_capital + other_equity

    st.markdown(
        f"**Total Shareholders' Funds — "
        f"{indian_currency(shareholders_funds)}**"
    )

    st.divider()

    st.markdown("**2. Non-current Liabilities**")

    non_current_liability_total = 0.0

    for label, classifications in non_current_liability_groups:
        amount = group_amount(
            liability_summary,
            classifications,
        )

        if abs(amount) > 0.005:
            st.write(
                f"{label} — "
                f"{indian_currency(amount)}"
            )

        non_current_liability_total += amount

    st.markdown(
        f"**Total Non-current Liabilities — "
        f"{indian_currency(non_current_liability_total)}**"
    )

    st.divider()

    st.markdown("**3. Current Liabilities**")

    current_liability_total = 0.0

    for label, classifications in current_liability_groups:
        amount = group_amount(
            liability_summary,
            classifications,
        )

        if abs(amount) > 0.005:
            st.write(
                f"{label} — "
                f"{indian_currency(amount)}"
            )

        current_liability_total += amount

    st.markdown(
        f"**Total Current Liabilities — "
        f"{indian_currency(current_liability_total)}**"
    )

    st.divider()

    st.markdown(
        f"### Total Equity and Liabilities — "
        f"{indian_currency(total_equity_and_liabilities)}"
    )

    st.divider()

    st.markdown("### II. Assets")

    st.markdown("**1. Non-current Assets**")

    non_current_asset_total = 0.0

    for label, classifications in non_current_asset_groups:
        amount = group_amount(
            asset_summary,
            classifications,
        )

        if abs(amount) > 0.005:
            st.write(
                f"{label} — "
                f"{indian_currency(amount)}"
            )

        non_current_asset_total += amount

    st.markdown(
        f"**Total Non-current Assets — "
        f"{indian_currency(non_current_asset_total)}**"
    )

    st.divider()

    st.markdown("**2. Current Assets**")

    current_asset_total = 0.0

    for label, classifications in current_asset_groups:
        amount = group_amount(
            asset_summary,
            classifications,
        )

        if abs(amount) > 0.005:
            st.write(
                f"{label} — "
                f"{indian_currency(amount)}"
            )

        current_asset_total += amount

    st.markdown(
        f"**Total Current Assets — "
        f"{indian_currency(current_asset_total)}**"
    )

    st.divider()

    st.markdown(
        f"### Total Assets — "
        f"{indian_currency(total_assets)}"
    )

    # -----------------------------------------------------
    # BALANCE SHEET CHECK
    # -----------------------------------------------------

    balance_difference = (
        total_assets
        - total_equity_and_liabilities
    )

    if abs(balance_difference) < 0.01:

        st.success(
            f"**Balance Sheet Tally ✅**  "
            f"Difference: {indian_currency(0)}"
        )

    else:

        st.error(
            f"**Balance Sheet Tally Failed ❌**  "
            f"Difference: "
            f"{indian_currency(abs(balance_difference))}"
        )

    # =====================================================
    # PRESENTATION NOTES
    # =====================================================

    st.divider()
    st.markdown("## 📝 Presentation Notes")

    presentation_notes = []

    if "Investments" in results_df["Classification"].values:

        presentation_notes.append(
            "Investments are displayed under Non-current "
            "Investments by default. Current/non-current "
            "classification should be confirmed where the "
            "Trial Balance does not provide the required facts."
        )

    if "Provisions" in results_df["Classification"].values:

        presentation_notes.append(
            "Provisions are displayed under Short-term "
            "Provisions by default because the current "
            "classification engine does not yet capture "
            "a separate long-term provision head."
        )

    if "Capital Account" in results_df["Classification"].values:

        presentation_notes.append(
            "Capital Account is presented within "
            "Shareholders' Funds / Other Equity for the "
            "Schedule III-style layout."
        )

    if presentation_notes:

        for note in presentation_notes:
            st.info(f"ℹ️ {note}")

    else:

        st.success(
            "No additional Schedule III presentation assumptions "
            "were detected."
        )


    # =====================================================
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
    # DETAILED NOTES / ACCOUNT BREAKDOWNS
    # =====================================================

    st.divider()
    st.markdown("## 📚 Detailed Account Notes")

    note_groups = [
        (
            "Property, Plant and Equipment",
            {"PPE", "Capital Work-in-Progress"},
        ),
        (
            "Intangible Assets",
            {
                "Intangible Assets",
                "Intangible Assets Under Development",
            },
        ),
        (
            "Investments",
            {"Investments"},
        ),
        (
            "Inventories",
            {"Inventories"},
        ),
        (
            "Trade Receivables",
            {"Trade Receivables"},
        ),
        (
            "Cash and Cash Equivalents",
            {"Cash & Cash Equivalents"},
        ),
        (
            "Borrowings",
            {
                "Borrowings",
                "Current Borrowings",
                "Non-current Borrowings",
            },
        ),
        (
            "Trade Payables",
            {"Trade Payables"},
        ),
        (
            "Equity",
            {
                "Share Capital",
                "Other Equity",
                "Capital Account",
            },
        ),
    ]

    for note_title, classifications in note_groups:

        note_df = results_df[
            results_df["Classification"].isin(
                classifications
            )
        ].copy()

        if note_df.empty:
            continue

        with st.expander(note_title):

            note_display = note_df[
                [
                    "Account",
                    "Debit",
                    "Credit",
                    "Classification",
                    "Confidence",
                ]
            ].copy()

            note_display["Debit"] = note_display[
                "Debit"
            ].map(indian_currency)

            note_display["Credit"] = note_display[
                "Credit"
            ].map(indian_currency)

            note_display["Confidence"] = note_display[
                "Confidence"
            ].apply(
                lambda x: f"{float(x) * 100:.0f}%"
            )

            st.dataframe(
                note_display,
                width="stretch",
                hide_index=True,
            )


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

    high_confidence_review_count = int(
        (
            pd.to_numeric(
                results_df["Confidence"],
                errors="coerce",
            ).fillna(0) < 0.80
        ).sum()
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
            "No unapproved classifications",
            invalid_classification_count == 0,
        ),
        (
            "No high-confidence review blockers",
            high_confidence_review_count == 0
            or review_count > 0,
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

    with export_col1:
        st.download_button(
            "📊 Download Excel Working Paper",
            data=excel_bytes,
            file_name=f"{safe_company}_Schedule_III_Working_Paper.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_excel_final",
        )

    with export_col2:
        st.download_button(
            "📄 Download PDF Financial Statements",
            data=pdf_bytes,
            file_name=f"{safe_company}_Financial_Statements.pdf",
            mime="application/pdf",
            key="download_pdf_final",
        )



    # =====================================================
    # PRODUCTION READINESS PANEL
    # =====================================================

    st.divider()
    st.markdown("## 🚀 Product Readiness")

    readiness = {
        "Trial Balance validation": tb_balanced,
        "Classification engine": len(results_df) > 0,
        "Balance Sheet Tally": bs_balanced,
        "P&L reconciliation": pnl_reconciles,
        "Manual review resolved": review_count == 0,
        "Numeric data": numeric_data_ok,
        "No unapproved classifications": invalid_classification_count == 0,
        "Company/reporting details": bool(company_name.strip()),
    }

    ready_count = sum(bool(value) for value in readiness.values())
    ready_total = len(readiness)

    st.progress(
        ready_count / max(ready_total, 1),
        text=f"Preparation readiness: {ready_count}/{ready_total}",
    )

    readiness_df = pd.DataFrame(
        [
            {
                "Area": area,
                "Status": "READY" if ready else "REVIEW",
            }
            for area, ready in readiness.items()
        ]
    )

    st.dataframe(
        readiness_df,
        width="stretch",
        hide_index=True,
    )

    st.caption(
        "Production deployment still requires server-side security, "
        "authentication/access control, privacy policy, logging, "
        "monitoring and real-world accounting validation."
    )


    # =====================================================
    # FOOTER
    # =====================================================

    st.divider()

    st.caption(
        "Accounting AI v1 • Schedule III-style financial statement "
        "presentation • Review all classifications and required "
        "disclosures before using statements for statutory filing."
    )
