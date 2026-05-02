from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response
import httpx
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
import io, json, string
from datetime import datetime

# ── App + forced CORS middleware ──────────────────────────────────────────────

app = FastAPI(title="FX Financial Report Converter")

class ForceCORSMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        if request.method == "OPTIONS":
            response = Response()
        else:
            response = await call_next(request)
        response.headers["Access-Control-Allow-Origin"]  = "*"
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "*"
        return response

app.add_middleware(ForceCORSMiddleware)

SKIP_SHEETS = {"filters","filter","settings","config","metadata","lookup"}

# ── Live FX rate proxy (avoids browser CORS on frankfurter.app) ───────────────

@app.get("/fx-rate")
async def fx_rate(from_currency: str, to_currency: str):
    try:
        async with httpx.AsyncClient(timeout=8) as client:
            res = await client.get(
                "https://api.frankfurter.app/latest",
                params={"from": from_currency, "to": to_currency}
            )
            return res.json()
    except Exception as e:
        raise HTTPException(502, f"Could not fetch live rate: {e}")


# ── Structure detection ───────────────────────────────────────────────────────

def pick_sheet(xl):
    for name in xl.sheet_names:
        if name.strip().lower() not in SKIP_SHEETS:
            return name
    return xl.sheet_names[0]


def detect_structure(contents: bytes):
    xl    = pd.ExcelFile(io.BytesIO(contents))
    sheet = pick_sheet(xl)
    df_raw = pd.read_excel(io.BytesIO(contents), sheet_name=sheet, header=None)

    # Find the header row (contains account/code/name AND balance/amount)
    header_row_idx = None
    for ri in range(min(8, len(df_raw))):
        vals = [str(v).lower().strip() for v in df_raw.iloc[ri] if pd.notna(v)]
        has_label  = any(k in v for v in vals for k in ["code","account","name","description"])
        has_amount = any(k in v for v in vals for k in ["balance","amount","total","value"])
        if has_label and has_amount:
            header_row_idx = ri
            break

    if header_row_idx is None:
        raise HTTPException(400, f"Could not find column headers in sheet '{sheet}'.")

    # Read with the detected header row
    df_hdr = pd.read_excel(io.BytesIO(contents), sheet_name=sheet, header=header_row_idx)
    df_hdr.columns = [str(c).strip() for c in df_hdr.columns]

    code_col = next((c for c in df_hdr.columns if "code" in c.lower()), None)
    name_col = next((c for c in df_hdr.columns
                     if any(k in c.lower() for k in ["account name","account","description","name"])
                     and c != code_col), None)
    amt_raw  = [c for c in df_hdr.columns
                if c not in [code_col, name_col]
                and any(k in c.lower() for k in ["balance","amount","total","value"])]

    if not name_col:
        raise HTTPException(400, f"No account name column found. Columns: {list(df_hdr.columns)}")
    if not amt_raw:
        raise HTTPException(400, f"No amount/balance column found. Columns: {list(df_hdr.columns)}")

    n_amt = len(amt_raw)
    all_col_names    = list(df_hdr.columns)
    amt_col_indices  = [all_col_names.index(c) for c in amt_raw]

    # Scan ALL pre-header rows to find branch/column names at the amount column positions
    branch_names = []
    for ri in range(header_row_idx):
        row_vals   = list(df_raw.iloc[ri])
        candidates = []
        for idx in amt_col_indices:
            if idx < len(row_vals) and pd.notna(row_vals[idx]):
                s = str(row_vals[idx]).strip()
                if s and s.lower() != "nan" and not (s.isdigit() and len(s) == 4):
                    candidates.append(s)
        if len(candidates) == n_amt:
            branch_names = candidates
            break
        elif len(candidates) > len(branch_names):
            branch_names = candidates  # keep best partial match so far

    # Fall back to generic names if nothing found
    if not branch_names:
        branch_names = ["Balance"] if n_amt == 1 else [f"Column {i+1}" for i in range(n_amt)]

    # Rename amount columns to branch names
    rename_map = dict(zip(amt_raw, branch_names[:n_amt]))
    df = df_hdr.rename(columns=rename_map)

    # Standardise code / account_name
    if code_col:
        df = df.rename(columns={code_col: "code"})
    else:
        df.insert(0, "code", "")
    df = df.rename(columns={name_col: "account_name"})

    keep = ["code","account_name"] + branch_names
    df   = df[[c for c in keep if c in df.columns]].copy()

    df["code"]         = df["code"].fillna("").astype(str).str.strip().str.replace(r"\.0$","",regex=True)
    df["account_name"] = df["account_name"].fillna("").astype(str).str.strip()
    for col in branch_names:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Row type: data if any amount column has a value, else header/section
    df["row_type"] = df[branch_names].notna().any(axis=1).map({True:"data",False:"header"})

    # Drop completely blank rows
    df = df[~(
        (df["code"] == "") &
        (df["account_name"] == "") &
        df[branch_names].isna().all(axis=1)
    )].reset_index(drop=True)

    return df, sheet, branch_names, len(branch_names) > 1


def apply_rate(df, branch_cols, selected_cols, rate):
    df = df.copy()
    for col in branch_cols:
        df[f"{col}__orig"] = df[col]
        df[f"{col}__conv"] = df[col] * rate if col in selected_cols else df[col]
    return df


# ── Excel builder ─────────────────────────────────────────────────────────────

def build_excel(df, branch_cols, selected_cols, rate, from_cur, to_cur, sheet_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "Converted Report"

    thin = Side(style="thin", color="CCCCCC")
    b    = Border(left=thin, right=thin, top=thin, bottom=thin)

    col_headers = ["Code", "Account Name"]
    col_keys    = []
    for bc in branch_cols:
        col_headers.append(f"{bc}\n({from_cur})")
        if bc in selected_cols:
            col_headers.append(f"{bc}\n({to_cur}) ✓")
        col_keys.append((f"{bc}__orig", f"{bc}__conv" if bc in selected_cols else None))

    n_cols = len(col_headers)

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.cell(1,1).value     = f"Financial Report — {sheet_name}"
    ws.cell(1,1).font      = Font(name="Arial", bold=True, size=14, color="E8D5B7")
    ws.cell(1,1).fill      = PatternFill("solid", start_color="0F3460")
    ws.cell(1,1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Subtitle
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ws.cell(2,1).value = (
        f"Rate: 1 {from_cur} = {rate} {to_cur}   |   "
        f"Converted: {', '.join(selected_cols) or 'none'}   |   "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    ws.cell(2,1).font      = Font(name="Arial", size=9, color="A8C5DA")
    ws.cell(2,1).fill      = PatternFill("solid", start_color="16213E")
    ws.cell(2,1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18
    ws.append([])  # spacer row 3

    # Column headers row 4
    for ci, h in enumerate(col_headers, 1):
        c = ws.cell(4, ci, h)
        c.font      = Font(name="Arial", bold=True, size=10, color="E8D5B7")
        c.fill      = PatternFill("solid", start_color="1A1A2E")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = b
    ws.row_dimensions[4].height = 32

    data_count = 0
    for _, row in df.iterrows():
        r      = ws.max_row + 1
        is_hdr = row["row_type"] == "header"
        if not is_hdr:
            data_count += 1
        alt = data_count % 2 == 0

        row_fill = (PatternFill("solid", start_color="2D2B55") if is_hdr else
                    PatternFill("solid", start_color="F5F3FF") if alt else None)

        # Code cell
        c1 = ws.cell(r, 1, "" if is_hdr else str(row["code"]))
        c1.font      = Font(name="Arial", size=9, color="C4B5FD" if is_hdr else "9CA3AF")
        c1.border    = b
        c1.alignment = Alignment(vertical="center")
        if row_fill: c1.fill = row_fill

        # Account name cell
        c2 = ws.cell(r, 2, str(row["account_name"]))
        c2.font      = Font(name="Arial", bold=is_hdr, size=10,
                            color="C4B5FD" if is_hdr else "1F2937")
        c2.border    = b
        c2.alignment = Alignment(vertical="center", indent=0 if is_hdr else 1)
        if row_fill: c2.fill = row_fill

        # Amount columns
        ci = 3
        for orig_key, conv_key in col_keys:
            v = row.get(orig_key)
            c = ws.cell(r, ci, float(v) if pd.notna(v) else None)
            c.number_format = "#,##0.00"
            c.alignment     = Alignment(horizontal="right", vertical="center")
            c.border        = b
            c.font          = Font(name="Arial", bold=is_hdr, size=9,
                                   color="C4B5FD" if is_hdr else "374151")
            if row_fill: c.fill = row_fill
            ci += 1

            if conv_key:
                cv = row.get(conv_key)
                cc = ws.cell(r, ci, float(cv) if pd.notna(cv) else None)
                cc.number_format = "#,##0.00"
                cc.alignment     = Alignment(horizontal="right", vertical="center")
                cc.border        = b
                cc.font          = Font(name="Arial", bold=is_hdr, size=9,
                                        color="86EFAC" if is_hdr else "065F46")
                cc.fill          = row_fill if row_fill else PatternFill("solid", start_color="F0FDF4")
                ci += 1

    # Column widths
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 42
    alpha = list(string.ascii_uppercase)
    for i in range(n_cols - 2):
        ws.column_dimensions[alpha[2 + i]].width = 17

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── PDF builder ───────────────────────────────────────────────────────────────

def build_pdf(df, branch_cols, selected_cols, rate, from_cur, to_cur, sheet_name):
    buf  = io.BytesIO()
    page = landscape(letter) if len(branch_cols) > 2 else letter
    doc  = SimpleDocTemplate(buf, pagesize=page,
                              leftMargin=0.5*inch, rightMargin=0.5*inch,
                              topMargin=0.6*inch,  bottomMargin=0.6*inch)
    styles = getSampleStyleSheet()
    story  = [
        Paragraph(f"Financial Report — {sheet_name}", ParagraphStyle(
            "T", parent=styles["Title"], fontSize=15,
            textColor=colors.HexColor("#0F3460"), spaceAfter=3, fontName="Helvetica-Bold")),
        Paragraph(
            f"Rate: 1 {from_cur} = {rate} {to_cur} &nbsp;|&nbsp; "
            f"Converted: {', '.join(selected_cols) or 'none'} &nbsp;|&nbsp; "
            f"{datetime.now().strftime('%Y-%m-%d %H:%M')}",
            ParagraphStyle("S", parent=styles["Normal"], fontSize=8,
                           textColor=colors.HexColor("#374151"), spaceAfter=12)),
    ]

    avail_w = page[0] - inch
    name_w  = 2.2*inch
    code_w  = 0.7*inch
    n_amt   = len(branch_cols) + len(selected_cols)
    amt_w   = max(0.65*inch, min(1.1*inch, (avail_w - name_w - code_w) / max(n_amt,1)))

    col_headers = ["Code", "Account Name"]
    col_widths  = [code_w, name_w]
    for bc in branch_cols:
        col_headers.append(f"{bc[:18]}\n({from_cur})")
        col_widths.append(amt_w)
        if bc in selected_cols:
            col_headers.append(f"{bc[:18]}\n({to_cur})✓")
            col_widths.append(amt_w)

    rows = [col_headers]
    for _, row in df.iterrows():
        is_hdr = row["row_type"] == "header"
        tr = ["" if is_hdr else str(row["code"]), str(row["account_name"])]
        for bc in branch_cols:
            v = row.get(f"{bc}__orig")
            tr.append(f"{float(v):,.2f}" if pd.notna(v) else "")
            if bc in selected_cols:
                cv = row.get(f"{bc}__conv")
                tr.append(f"{float(cv):,.2f}" if pd.notna(cv) else "")
        rows.append(tr)

    tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    cmd = [
        ("BACKGROUND",    (0, 0), (-1,  0), colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",     (0, 0), (-1,  0), colors.HexColor("#E8D5B7")),
        ("FONTNAME",      (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 7),
        ("ALIGN",         (0, 0), (-1,  0), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("ALIGN",         (2, 1), (-1, -1), "RIGHT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#E5E7EB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUND", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F3FF")]),
    ]
    for ri, row in enumerate(df.itertuples(), start=1):
        if row.row_type == "header":
            cmd += [
                ("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#2D2B55")),
                ("TEXTCOLOR",  (0, ri), (-1, ri), colors.HexColor("#C4B5FD")),
                ("FONTNAME",   (0, ri), (-1, ri), "Helvetica-Bold"),
            ]
    tbl.setStyle(TableStyle(cmd))
    story.append(tbl)
    doc.build(story)
    return buf.getvalue()


# ── Shared request helper ─────────────────────────────────────────────────────

async def read_request(file, exchange_rate, selected_cols_json):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx / .xls files supported.")
    if exchange_rate <= 0:
        raise HTTPException(400, "Exchange rate must be positive.")
    contents = await file.read()
    df, sheet, branch_cols, is_multi = detect_structure(contents)
    sel = json.loads(selected_cols_json) if selected_cols_json else branch_cols
    sel = [c for c in sel if c in branch_cols] or branch_cols
    df  = apply_rate(df, branch_cols, sel, exchange_rate)
    return df, sheet, branch_cols, sel, exchange_rate


# ── Routes ────────────────────────────────────────────────────────────────────

@app.post("/detect")
async def detect(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx / .xls files supported.")
    contents = await file.read()
    df, sheet, branch_cols, is_multi = detect_structure(contents)
    data_rows = df[df["row_type"] == "data"]
    return {
        "sheet_name":  sheet,
        "branch_cols": branch_cols,
        "is_multi":    is_multi,
        "row_count":   len(data_rows),
    }


@app.post("/preview")
async def preview(
    file:          UploadFile = File(...),
    exchange_rate: float      = Form(...),
    from_currency: str        = Form("USD"),
    to_currency:   str        = Form("EUR"),
    selected_cols: str        = Form("[]"),
):
    df, sheet, branch_cols, sel, rate = await read_request(file, exchange_rate, selected_cols)
    data_rows = df[df["row_type"] == "data"]
    totals = {}
    for bc in branch_cols:
        totals[f"{bc}__orig"] = float(data_rows[f"{bc}__orig"].sum(skipna=True))
        if bc in sel:
            totals[f"{bc}__conv"] = float(data_rows[f"{bc}__conv"].sum(skipna=True))
    preview_keys = (
        ["code","account_name","row_type"] +
        [k for bc in branch_cols
         for k in ([f"{bc}__orig", f"{bc}__conv"] if bc in sel else [f"{bc}__orig"])]
    )
    return {
        "sheet_name":    sheet,
        "branch_cols":   branch_cols,
        "selected_cols": sel,
        "is_multi":      len(branch_cols) > 1,
        "rows":          len(data_rows),
        "totals":        totals,
        "from_currency": from_currency,
        "to_currency":   to_currency,
        "exchange_rate": rate,
        "preview":       df[preview_keys].head(15).to_dict("records"),
    }


@app.post("/download/xlsx")
async def download_xlsx(
    file:          UploadFile = File(...),
    exchange_rate: float      = Form(...),
    from_currency: str        = Form("USD"),
    to_currency:   str        = Form("EUR"),
    selected_cols: str        = Form("[]"),
):
    df, sheet, branch_cols, sel, rate = await read_request(file, exchange_rate, selected_cols)
    data = build_excel(df, branch_cols, sel, rate, from_currency, to_currency, sheet)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=converted_report.xlsx"},
    )


@app.post("/download/pdf")
async def download_pdf(
    file:          UploadFile = File(...),
    exchange_rate: float      = Form(...),
    from_currency: str        = Form("USD"),
    to_currency:   str        = Form("EUR"),
    selected_cols: str        = Form("[]"),
):
    df, sheet, branch_cols, sel, rate = await read_request(file, exchange_rate, selected_cols)
    data = build_pdf(df, branch_cols, sel, rate, from_currency, to_currency, sheet)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=converted_report.pdf"},
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
