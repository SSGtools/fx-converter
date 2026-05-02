# FX Financial Report Converter

Upload an Excel financial report, apply an exchange rate, and download
converted outputs as a styled Excel file and PDF — no temp files, no storage.

## Architecture

```
POST /preview       → JSON summary + first 10 rows
POST /download/xlsx → streams .xlsx directly to browser
POST /download/pdf  → streams .pdf directly to browser
GET  /health        → {"status":"ok"}
```

Every request is stateless. Files are processed in memory and streamed back
immediately. Nothing is written to disk on the server.

---

## Run locally

### Backend

```bash
cd backend
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
# API docs → http://localhost:8000/docs
```

### Frontend (new terminal)

```bash
cd frontend
npm install
npm run dev
# → http://localhost:3000
```

---

## Deploy to Railway (recommended)

Railway auto-detects Dockerfiles and gives you two free services.

### 1 · Push to GitHub

```bash
git init && git add . && git commit -m "init"
gh repo create fx-converter --public --push
```

### 2 · Deploy the backend

1. Go to railway.app → New Project → Deploy from GitHub repo
2. Select your repo, point to the **`backend/`** folder
3. Railway picks up `Dockerfile` automatically
4. Once deployed, copy the public URL (e.g. `https://fx-api.up.railway.app`)

### 3 · Deploy the frontend

1. In the same Railway project → New Service → GitHub repo again
2. Point to the **`frontend/`** folder
3. Add an environment variable:
   ```
   VITE_API_URL = https://fx-api.up.railway.app
   ```
4. Deploy — Railway runs `npm install && npm run build` then serves `dist/`

### 4 · Tighten CORS (optional but recommended)

In `backend/main.py`, replace `allow_origins=["*"]` with your frontend URL:

```python
allow_origins=["https://fx-app.up.railway.app"],
```

Redeploy the backend.

---

## Deploy to Render

Same idea — create two Web Services from the same repo,
set `Root Directory` to `backend/` or `frontend/` for each.
Add `VITE_API_URL` as an environment variable on the frontend service.

---

## Excel input format

Your file needs columns containing the words Account, Description, and Amount
(case-insensitive, partial match — "Account Code", "Desc.", "Amt" won't work
but "Account Number", "Description", "Amount USD" will).

| Account | Description             | Amount     |
|---------|-------------------------|------------|
| 1001    | Cash and Equivalents    | 150000.00  |
| 2001    | Accounts Payable        | -45000.00  |
| 3001    | Revenue - Product Sales | 320000.75  |

A sample file is included: `sample_financial_report.xlsx`
