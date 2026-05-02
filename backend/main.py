from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
import io
from datetime import datetime

app = FastAPI(title="FX Financial Report Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Column detection ──────────────────────────────────────────────────────────

# Maps normalized column keywords → canonical field name
COLUMN_ALIASES = {
    "account": ["account", "code", "acct", "account number", "account code"],
    "description": ["description", "account name", "name", "desc", "title", "detail"],
    "amount": ["amount", "balance", "total", "value", "amt", "debit", "credit"],
}

def find_columns(df: pd.DataFrame) -> dict:
    """Flexibly map dataframe columns to account/description/amount."""
    mapping = {}
    for col in df.columns:
        normalized = str(col).strip().lower()
        for canonical, aliases in COLUMN_ALIASES.items():
            if canonical not in mapping:
                for alias in aliases:
                    if alias in normalized:
                        mapping[canonical] = col
                        break
    return mapping


def pick_financial_sheet(xl: pd.ExcelFile) -> str:
    """Return the most likely financial sheet, skipping Filters/metadata sheets."""
    skip = {"filters", "filter", "settings", "config", "metadata", "lookup"}
    for name in xl.sheet_names:
        if name.strip().lower() not in skip:
            return name
    return xl.sheet_names[0]


def parse_df(contents: bytes) -> tuple[pd.DataFrame, str]:
    """
    Parse uploaded Excel, auto-detect columns, skip section headers/totals rows.
    Returns (df_work, sheet_name).
    """
    try:
        xl = pd.ExcelFile(io.BytesIO(contents))
    except Exception as e:
        raise HTTPException(400, f"Could not open Excel file: {e}")

    sheet = pick_financial_sheet(xl)

    # Try reading with header on row 0; if columns not found, try row 1 or 2
    df = None
    col_map = {}
    for header_row in range(0, 4):
        try:
            candidate = pd.read_excel(io.BytesIO(contents), sheet_name=sheet, header=header_row)
            candidate.columns = [str(c).strip() for c in candidate.columns]
            mapping = find_columns(candidate)
            if len(mapping) >= 2:          # need at least description + amount
                df = candidate
                col_map = mapping
                break
        except Exception:
            continue

    if df is None or not col_map:
        raise HTTPException(
            400,
            f"Could not detect required columns (account/description + amount/balance). "
            f"Sheet '{sheet}' columns found: {list(pd.read_excel(io.BytesIO(contents), sheet_name=sheet, nrows=3).columns)}"
        )

    # Ensure all three canonical fields exist (account may be optional)
    if "description" not in col_map:
        raise HTTPException(400, "Could not find a description/account name column.")
    if "amount" not in col_map:
        raise HTTPException(400, "Could not find an amount/balance column.")

    # Rename to canonical names
    rename = {v: k for k, v in col_map.items()}
    df = df.rename(columns=rename)

    # If no account/code column, add a blank one
    if "account" not in df.columns:
        df.insert(0, "account", "")

    # Keep only our three columns (plus any extras we don't need)
    df = df[["account", "description", "amount"]].copy()

    # Convert amount to numeric; rows where it's NaN after conversion are headers/labels
    df["amount_raw"] = pd.to_numeric(df["amount"], errors="coerce")

    # Mark row type:
    #   "header"  → description has text, amount is blank/NaN (section headings, Total rows)
    #   "data"    → has a numeric amount
    df["row_type"] = df["amount_raw"].apply(lambda x: "data" if pd.notna(x) else "header")

    # Clean up
    df["account"] = df["account"].fillna("").astype(str).str.strip()
    df["description"] = df["description"].fillna("").astype(str).str.strip()
    df["amount_raw"] = df["amount_raw"].fillna(0)

    # Drop completely empty rows
    df = df[~((df["account"] == "") & (df["description"] == "") & (df["amount_raw"] == 0))]
    df = df.reset_index(drop=True)

    return df, sheet


def apply_rate(df: pd.DataFrame, rate: float) -> pd.DataFrame:
    df = df.copy()
    df["amount_original"]  = df["amount_raw"]
    df["amount_converted"] = df.apply(
        lambda r: r["amount_raw"] * rate if r["row_type"] == "data" else r["amount_raw"],
        axis=1
    )
    return df


# ── Excel builder ─────────────────────────────────────────────────────────────

def build_excel(df: pd.DataFrame, rate: float, from_cur: str, to_cur: str, sheet_name: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Converted Report"

    thin  = Side(style="thin",   color="CCCCCC")
    thick = Side(style="medium", color="888888")
    data_border   = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
    header_border = Border(left=thick, right=thick, top=thick, bottom=thick)

    # ── Banner rows ──
    ws.merge_cells("A1:E1")
    ws["A1"] = f"Financial Report — {sheet_name}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=14, color="E8D5B7")
    ws["A1"].fill      = PatternFill("solid", start_color="0F3460")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:E2")
    ws["A2"] = (f"Exchange Rate: 1 {from_cur} = {rate} {to_cur}   |   "
                f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    ws["A2"].font      = Font(name="Arial", size=10, color="A8C5DA")
    ws["A2"].fill      = PatternFill("solid", start_color="16213E")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.append([])  # blank spacer row 3

    # ── Column headers ──
    headers = ["Code", "Account Name", f"Original ({from_cur})", f"Converted ({to_cur})", "Rate"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font      = Font(name="Arial", bold=True, color="E8D5B7", size=11)
        c.fill      = PatternFill("solid", start_color="1A1A2E")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = header_border
    ws.row_dimensions[4].height = 22

    # ── Data rows ──
    is_header_fill  = PatternFill("solid", start_color="2D2B55")
    alt_fill        = PatternFill("solid", start_color="F5F3FF")
    data_row_count  = 0

    for i, row in df.iterrows():
        r = 5 + i
        is_hdr = row["row_type"] == "header"

        if is_hdr:
            # Section header / subtotal row — span description, bold, indigo tint
            ws.merge_cells(f"B{r}:C{r}")
            c_desc = ws.cell(r, 2, str(row["description"]))
            c_desc.font      = Font(name="Arial", bold=True, size=10, color="C4B5FD")
            c_desc.fill      = is_header_fill
            c_desc.alignment = Alignment(vertical="center")
            c_desc.border    = data_border
            ws.cell(r, 1).fill   = is_header_fill
            ws.cell(r, 1).border = data_border
            # Show the amount for subtotal rows (non-zero headers)
            if row["amount_raw"] != 0:
                for ci, key in [(4, "amount_original"), (5, "amount_converted")]:
                    c = ws.cell(r, ci, float(row[key]))
                    c.number_format = "#,##0.00"
                    c.alignment     = Alignment(horizontal="right", vertical="center")
                    c.fill          = is_header_fill
                    c.border        = data_border
                    c.font          = Font(name="Arial", bold=True, size=10, color="C4B5FD")
            else:
                ws.cell(r, 4).fill   = is_header_fill
                ws.cell(r, 4).border = data_border
                ws.cell(r, 5).fill   = is_header_fill
                ws.cell(r, 5).border = data_border
            ws.cell(r, 3).fill   = is_header_fill
            ws.cell(r, 3).border = data_border
        else:
            data_row_count += 1
            row_fill = alt_fill if data_row_count % 2 == 0 else None
            # Code
            c1 = ws.cell(r, 1, str(row["account"]))
            c1.font      = Font(name="Arial", size=10, color="6B7280")
            c1.alignment = Alignment(vertical="center")
            c1.border    = data_border
            if row_fill: c1.fill = row_fill
            # Description
            c2 = ws.cell(r, 2, str(row["description"]))
            c2.font      = Font(name="Arial", size=10)
            c2.alignment = Alignment(vertical="center", indent=1)
            c2.border    = data_border
            if row_fill: c2.fill = row_fill
            # Amounts
            for ci, key in [(3, "amount_original"), (4, "amount_converted")]:
                c = ws.cell(r, ci, float(row[key]))
                c.number_format = "#,##0.00"
                c.alignment     = Alignment(horizontal="right", vertical="center")
                c.border        = data_border
                c.font          = Font(name="Arial", size=10)
                if row_fill: c.fill = row_fill
            # Rate
            c5 = ws.cell(r, 5, rate)
            c5.number_format = "0.0000"
            c5.alignment     = Alignment(horizontal="center", vertical="center")
            c5.border        = data_border
            c5.font          = Font(name="Arial", size=10, color="9CA3AF")
            if row_fill: c5.fill = row_fill

    # ── Column widths ──
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 12

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── PDF builder ───────────────────────────────────────────────────────────────

def build_pdf(df: pd.DataFrame, rate: float, from_cur: str, to_cur: str, sheet_name: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.6*inch, rightMargin=0.6*inch,
                            topMargin=0.7*inch, bottomMargin=0.7*inch)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("T2", parent=styles["Title"], fontSize=16,
                                 textColor=colors.HexColor("#0F3460"), spaceAfter=4,
                                 fontName="Helvetica-Bold")
    sub_style   = ParagraphStyle("S2", parent=styles["Normal"], fontSize=9,
                                 textColor=colors.HexColor("#374151"), spaceAfter=14)

    story = [
        Paragraph(f"Financial Report — {sheet_name}", title_style),
        Paragraph(
            f"Rate: 1 {from_cur} = {rate} {to_cur} &nbsp;|&nbsp; "
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            sub_style),
    ]

    col_widths = [0.85*inch, 3.0*inch, 1.15*inch, 1.15*inch, 0.65*inch]
    table_data = [["Code", "Account Name", f"Orig.\n({from_cur})", f"Conv.\n({to_cur})", "Rate"]]

    for _, row in df.iterrows():
        is_hdr = row["row_type"] == "header"
        orig = f"{float(row['amount_original']):,.2f}" if row["amount_raw"] != 0 else ""
        conv = f"{float(row['amount_converted']):,.2f}" if row["amount_raw"] != 0 else ""
        table_data.append([
            "" if is_hdr else str(row["account"]),
            str(row["description"]),
            orig, conv,
            "" if is_hdr else f"{rate:.4f}",
        ])

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    # Build per-row styles
    cmd = [
        # Header row
        ("BACKGROUND",    (0, 0), (-1,  0), colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",     (0, 0), (-1,  0), colors.HexColor("#E8D5B7")),
        ("FONTNAME",      (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (0, 0), (-1,  0), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("ALIGN",         (2, 1), ( 4, -1), "RIGHT"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUND", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F3FF")]),
    ]
    # Style section-header rows
    for ri, row in enumerate(df.itertuples(), start=1):
        if row.row_type == "header":
            cmd += [
                ("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#2D2B55")),
                ("TEXTCOLOR",  (0, ri), (-1, ri), colors.HexColor("#C4B5FD")),
                ("FONTNAME",   (0, ri), (-1, ri), "Helvetica-Bold"),
                ("SPAN",       (0, ri), (1,  ri)),
            ]

    tbl.setStyle(TableStyle(cmd))
    story.append(tbl)
    doc.build(story)
    return buf.getvalue()


# ── Shared request parser ─────────────────────────────────────────────────────

async def read_request(file: UploadFile, exchange_rate: float):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx / .xls files are supported.")
    if exchange_rate <= 0:
        raise HTTPException(400, "Exchange rate must be a positive number.")
    contents = await file.read()
    df, sheet = parse_df(contents)
    df = apply_rate(df, exchange_rate)
    return df, exchange_rate, sheet, contents


# ── Routes ────────────────────────────────────────────────────────────────────

@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    exchange_rate: float = Form(...),
    from_currency: str = Form("USD"),
    to_currency: str = Form("EUR"),
):
    df, rate, sheet, _ = await read_request(file, exchange_rate)
    data_rows = df[df["row_type"] == "data"]
    return {
        "rows": len(data_rows),
        "total_original":  float(data_rows["amount_original"].sum()),
        "total_converted": float(data_rows["amount_converted"].sum()),
        "from_currency": from_currency,
        "to_currency":   to_currency,
        "exchange_rate": rate,
        "sheet_name":    sheet,
        "preview": df[["account","description","amount_original","amount_converted","row_type"]]
                     .head(15).to_dict("records"),
    }


@app.post("/download/xlsx")
async def download_xlsx(
    file: UploadFile = File(...),
    exchange_rate: float = Form(...),
    from_currency: str = Form("USD"),
    to_currency: str = Form("EUR"),
):
    df, rate, sheet, _ = await read_request(file, exchange_rate)
    data = build_excel(df, rate, from_currency, to_currency, sheet)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=converted_report.xlsx"},
    )


@app.post("/download/pdf")
async def download_pdf(
    file: UploadFile = File(...),
    exchange_rate: float = Form(...),
    from_currency: str = Form("USD"),
    to_currency: str = Form("EUR"),
):
    df, rate, sheet, _ = await read_request(file, exchange_rate)
    data = build_pdf(df, rate, from_currency, to_currency, sheet)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=converted_report.pdf"},
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
