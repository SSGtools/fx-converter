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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
import io
from datetime import datetime

app = FastAPI(title="FX Financial Report Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # tighten to your frontend domain in production
    allow_methods=["*"],
    allow_headers=["*"],
)

REQUIRED_COLUMNS = {"account", "description", "amount"}


def find_columns(df: pd.DataFrame) -> dict:
    mapping = {}
    for col in df.columns:
        normalized = col.strip().lower()
        for req in REQUIRED_COLUMNS:
            if req in normalized and req not in mapping:
                mapping[req] = col
    return mapping


def parse_df(contents: bytes) -> pd.DataFrame:
    try:
        df = pd.read_excel(io.BytesIO(contents))
    except Exception as e:
        raise HTTPException(400, f"Could not parse Excel file: {e}")
    col_map = find_columns(df)
    missing = REQUIRED_COLUMNS - set(col_map.keys())
    if missing:
        raise HTTPException(400, f"Missing columns: {', '.join(missing)}. Found: {list(df.columns)}")
    df = df.rename(columns={v: k for k, v in col_map.items()})
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    df["amount_original"] = df["amount"]
    df["amount_converted"] = df["amount"]
    return df


async def read_request(file: UploadFile, exchange_rate: float):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx / .xls files are supported.")
    if exchange_rate <= 0:
        raise HTTPException(400, "Exchange rate must be a positive number.")
    contents = await file.read()
    df = parse_df(contents)
    df["amount_converted"] = df["amount_original"] * exchange_rate
    return df, exchange_rate


def build_excel(df: pd.DataFrame, rate: float, from_cur: str, to_cur: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Converted Report"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title banner
    ws.merge_cells("A1:E1")
    ws["A1"] = "Financial Report — Currency Conversion"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="E8D5B7")
    ws["A1"].fill = PatternFill("solid", start_color="0F3460")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Subtitle
    ws.merge_cells("A2:E2")
    ws["A2"] = f"Rate: 1 {from_cur} = {rate} {to_cur}   |   Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = Font(name="Arial", size=10, color="A8C5DA")
    ws["A2"].fill = PatternFill("solid", start_color="16213E")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.append([])  # spacer row 3

    # Column headers
    headers = ["Account", "Description", f"Original ({from_cur})", f"Converted ({to_cur})", "Rate Applied"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = Font(name="Arial", bold=True, color="E8D5B7", size=11)
        c.fill = PatternFill("solid", start_color="1A1A2E")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws.row_dimensions[4].height = 22

    # Data rows
    for i, (_, row) in enumerate(df.iterrows()):
        r = 5 + i
        ws.cell(r, 1, str(row["account"])).border = border
        ws.cell(r, 2, str(row["description"])).border = border
        for ci, key in [(3, "amount_original"), (4, "amount_converted")]:
            c = ws.cell(r, ci, float(row[key]))
            c.number_format = "#,##0.00"
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = border
            c.font = Font(name="Arial", size=10)
        rate_cell = ws.cell(r, 5, rate)
        rate_cell.number_format = "0.0000"
        rate_cell.alignment = Alignment(horizontal="center", vertical="center")
        rate_cell.border = border
        rate_cell.font = Font(name="Arial", size=10)
        if r % 2 == 0:
            for ci in range(1, 6):
                ws.cell(r, ci).fill = PatternFill("solid", start_color="F0F4F8")

    # Totals row
    tr = 5 + len(df)
    for ci in range(1, 6):
        c = ws.cell(tr, ci)
        c.fill = PatternFill("solid", start_color="0F3460")
        c.font = Font(name="Arial", bold=True, size=10, color="E8D5B7")
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(tr, 1).value = "TOTAL"
    for ci, col in [(3, "C"), (4, "D")]:
        ws.cell(tr, ci).value = f"=SUM({col}5:{col}{tr-1})"
        ws.cell(tr, ci).number_format = "#,##0.00"
        ws.cell(tr, ci).alignment = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_pdf(df: pd.DataFrame, rate: float, from_cur: str, to_cur: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Financial Report — Currency Conversion", ParagraphStyle(
            "Title2", parent=styles["Title"], fontSize=18,
            textColor=colors.HexColor("#0F3460"), spaceAfter=4, fontName="Helvetica-Bold")),
        Paragraph(
            f"Rate: 1 {from_cur} = {rate} {to_cur} &nbsp;|&nbsp; "
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            ParagraphStyle("Sub", parent=styles["Normal"], fontSize=10,
                           textColor=colors.HexColor("#16213E"), spaceAfter=16)),
    ]
    rows = [["Account", "Description", f"Original\n({from_cur})", f"Converted\n({to_cur})", "Rate"]]
    for _, row in df.iterrows():
        rows.append([
            str(row.get("account", "")), str(row.get("description", "")),
            f"{float(row['amount_original']):,.2f}",
            f"{float(row['amount_converted']):,.2f}",
            f"{rate:.4f}",
        ])
    rows.append(["TOTAL", "", f"{df['amount_original'].sum():,.2f}",
                 f"{df['amount_converted'].sum():,.2f}", ""])

    tbl = Table(rows, colWidths=[1.1*inch, 2.5*inch, 1.3*inch, 1.3*inch, 0.8*inch], repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1,  0), colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",     (0, 0), (-1,  0), colors.HexColor("#E8D5B7")),
        ("FONTNAME",      (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 9),
        ("ALIGN",         (0, 0), (-1,  0), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUND", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F0F4F8")]),
        ("FONTNAME",      (0, 1), (-1, -2), "Helvetica"),
        ("ALIGN",         (2, 1), ( 4, -1), "RIGHT"),
        ("BACKGROUND",    (0,-1), (-1, -1), colors.HexColor("#0F3460")),
        ("TEXTCOLOR",     (0,-1), (-1, -1), colors.HexColor("#E8D5B7")),
        ("FONTNAME",      (0,-1), (-1, -1), "Helvetica-Bold"),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(tbl)
    doc.build(story)
    return buf.getvalue()


# ── Routes ──────────────────────────────────────────────────────────────────

@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    exchange_rate: float = Form(...),
    from_currency: str = Form("USD"),
    to_currency: str = Form("EUR"),
):
    df, rate = await read_request(file, exchange_rate)
    return {
        "rows": len(df),
        "total_original": float(df["amount_original"].sum()),
        "total_converted": float(df["amount_converted"].sum()),
        "from_currency": from_currency,
        "to_currency": to_currency,
        "exchange_rate": rate,
        "preview": df[["account", "description", "amount_original", "amount_converted"]]
                     .head(10).to_dict("records"),
    }


@app.post("/download/xlsx")
async def download_xlsx(
    file: UploadFile = File(...),
    exchange_rate: float = Form(...),
    from_currency: str = Form("USD"),
    to_currency: str = Form("EUR"),
):
    df, rate = await read_request(file, exchange_rate)
    data = build_excel(df, rate, from_currency, to_currency)
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
    df, rate = await read_request(file, exchange_rate)
    data = build_pdf(df, rate, from_currency, to_currency)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=converted_report.pdf"},
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
