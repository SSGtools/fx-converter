import { useState, useRef, useCallback } from "react";

// In production set this to your Railway backend URL, e.g. https://fx-api.up.railway.app
const API = import.meta.env.VITE_API_URL ?? "http://localhost:8000";

const CURRENCIES = ["USD","EUR","GBP","JPY","CAD","AUD","CHF","CNY","HKD","SGD",
                    "INR","BRL","MXN","KRW","AED","SAR","SEK","NOK","DKK","NZD"];

const fmt = (n, d = 2) =>
  new Intl.NumberFormat("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }).format(n);

// Build a FormData from current state — used by preview AND both downloads
function buildForm(file, rate, fromCur, toCur) {
  const fd = new FormData();
  fd.append("file", file);
  fd.append("exchange_rate", rate);
  fd.append("from_currency", fromCur);
  fd.append("to_currency", toCur);
  return fd;
}

// POST to an endpoint and trigger a browser download from the binary response
async function downloadFile(endpoint, form, filename) {
  const res = await fetch(`${API}${endpoint}`, { method: "POST", body: form });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Download failed" }));
    throw new Error(err.detail ?? "Download failed");
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{
      background: "var(--color-background-secondary)",
      borderRadius: "var(--border-radius-lg)",
      padding: "1rem 1.25rem",
      border: accent ? "1.5px solid var(--color-border-info)" : ".5px solid var(--color-border-tertiary)",
    }}>
      <p style={{ fontSize: 11, color: "var(--color-text-secondary)", margin: "0 0 4px",
                  textTransform: "uppercase", letterSpacing: ".06em" }}>{label}</p>
      <p style={{ fontSize: 22, fontWeight: 500, margin: 0,
                  color: accent ? "var(--color-text-info)" : "var(--color-text-primary)" }}>{value}</p>
      {sub && <p style={{ fontSize: 11, color: "var(--color-text-tertiary)", margin: "4px 0 0" }}>{sub}</p>}
    </div>
  );
}

function DropZone({ onFile, file }) {
  const [dragging, setDragging] = useState(false);
  const ref = useRef();
  const onDrop = useCallback(e => {
    e.preventDefault(); setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f && (f.name.endsWith(".xlsx") || f.name.endsWith(".xls"))) onFile(f);
  }, [onFile]);
  return (
    <div onClick={() => ref.current.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={onDrop}
      style={{
        border: `2px dashed ${dragging ? "var(--color-border-info)" : "var(--color-border-secondary)"}`,
        borderRadius: "var(--border-radius-lg)", padding: "2.5rem 1rem",
        textAlign: "center", cursor: "pointer",
        background: dragging ? "var(--color-background-info)" : "var(--color-background-secondary)",
        transition: "all .15s",
      }}>
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
        onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
      <div style={{ fontSize: 30, marginBottom: 8 }}>📊</div>
      {file
        ? <p style={{ margin: 0, fontWeight: 500, color: "var(--color-text-primary)" }}>{file.name}</p>
        : <>
            <p style={{ margin: "0 0 4px", fontWeight: 500, color: "var(--color-text-primary)" }}>
              Drop your Excel file here
            </p>
            <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-secondary)" }}>
              .xlsx or .xls · needs Account, Description, Amount columns
            </p>
          </>
      }
    </div>
  );
}

function PreviewTable({ rows, fromCur, toCur }) {
  if (!rows?.length) return null;
  const cols = [
    { key: "account",          label: "Account",              right: false },
    { key: "description",      label: "Description",          right: false },
    { key: "amount_original",  label: `Original (${fromCur})`,  right: true  },
    { key: "amount_converted", label: `Converted (${toCur})`, right: true  },
  ];
  return (
    <div style={{ overflowX: "auto", borderRadius: "var(--border-radius-md)",
                  border: ".5px solid var(--color-border-tertiary)" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, tableLayout: "fixed" }}>
        <thead>
          <tr style={{ background: "var(--color-background-secondary)" }}>
            {cols.map(c => (
              <th key={c.key} style={{ padding: "8px 12px", textAlign: c.right ? "right" : "left",
                fontWeight: 500, fontSize: 12, color: "var(--color-text-secondary)",
                borderBottom: ".5px solid var(--color-border-tertiary)", whiteSpace: "nowrap" }}>
                {c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ background: i % 2 === 0
              ? "var(--color-background-primary)" : "var(--color-background-secondary)" }}>
              {cols.map(c => {
                const v = row[c.key];
                const isAmt = c.right;
                const neg = isAmt && v < 0;
                return (
                  <td key={c.key} style={{
                    padding: "7px 12px", textAlign: c.right ? "right" : "left",
                    borderBottom: ".5px solid var(--color-border-tertiary)",
                    color: c.key === "amount_converted"
                      ? (neg ? "var(--color-text-danger)" : "var(--color-text-info)")
                      : neg ? "var(--color-text-danger)" : "var(--color-text-primary)",
                    fontWeight: c.key === "amount_converted" ? 500 : 400,
                    overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                  }}>
                    {isAmt ? fmt(v) : String(v ?? "—")}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const sel = { padding: "8px 10px", borderRadius: "var(--border-radius-md)",
  border: ".5px solid var(--color-border-secondary)", background: "var(--color-background-primary)",
  color: "var(--color-text-primary)", fontSize: 14, cursor: "pointer" };

const btn = (primary) => ({
  padding: "9px 18px", borderRadius: "var(--border-radius-md)", fontSize: 14, fontWeight: 500,
  cursor: "pointer", border: primary ? "none" : ".5px solid var(--color-border-secondary)",
  background: primary ? "var(--color-text-primary)" : "var(--color-background-secondary)",
  color: primary ? "var(--color-background-primary)" : "var(--color-text-primary)",
  fontFamily: "inherit", transition: "opacity .15s",
});

export default function App() {
  const [file, setFile]           = useState(null);
  const [rate, setRate]           = useState("1.08");
  const [fromCur, setFromCur]     = useState("USD");
  const [toCur, setToCur]         = useState("EUR");
  const [loading, setLoading]     = useState(false);
  const [dlLoading, setDlLoading] = useState("");   // "xlsx" | "pdf" | ""
  const [error, setError]         = useState("");
  const [result, setResult]       = useState(null);

  const rateVal = parseFloat(rate);
  const rateHint = rateVal > 0
    ? `1 ${fromCur} = ${rateVal} ${toCur}  ·  1 ${toCur} = ${(1/rateVal).toFixed(4)} ${fromCur}`
    : "Enter a positive rate";

  const handlePreview = async () => {
    if (!file)            { setError("Please upload an Excel file."); return; }
    if (!(rateVal > 0))   { setError("Enter a valid positive exchange rate."); return; }
    setLoading(true); setError(""); setResult(null);
    try {
      const res  = await fetch(`${API}/preview`, { method: "POST", body: buildForm(file, rateVal, fromCur, toCur) });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail ?? "Preview failed.");
      setResult(data);
    } catch (e) {
      setError(e.message.includes("fetch")
        ? `Cannot reach backend at ${API}. Is the server running?`
        : e.message);
    } finally { setLoading(false); }
  };

  const handleDownload = async (format) => {
    if (!file || !(rateVal > 0)) return;
    setDlLoading(format); setError("");
    try {
      await downloadFile(
        `/download/${format}`,
        buildForm(file, rateVal, fromCur, toCur),
        `converted_report.${format}`,
      );
    } catch (e) { setError(e.message); }
    finally { setDlLoading(""); }
  };

  return (
    <div style={{ maxWidth: 720, margin: "0 auto", padding: "2rem 1rem", fontFamily: "var(--font-sans)" }}>

      {/* Header */}
      <div style={{ marginBottom: "2rem" }}>
        <p style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary)",
                    textTransform: "uppercase", letterSpacing: ".08em", margin: "0 0 6px" }}>
          Financial Tools
        </p>
        <h1 style={{ fontSize: 26, fontWeight: 500, margin: "0 0 6px", color: "var(--color-text-primary)" }}>
          FX Report Converter
        </h1>
        <p style={{ fontSize: 14, color: "var(--color-text-secondary)", margin: 0 }}>
          Upload an Excel financial report, apply an exchange rate, preview the result, then download Excel or PDF.
        </p>
      </div>

      {/* Upload */}
      <section style={{ marginBottom: "1.5rem" }}>
        <p style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary)",
                    textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 10px" }}>
          1 · Upload report
        </p>
        <DropZone onFile={f => { setFile(f); setResult(null); setError(""); }} file={file} />
        {file && (
          <button onClick={() => { setFile(null); setResult(null); }}
            style={{ marginTop: 8, background: "none", border: "none", cursor: "pointer",
                     fontSize: 12, color: "var(--color-text-secondary)", padding: 0 }}>
            Remove file ×
          </button>
        )}
      </section>

      {/* Rate */}
      <section style={{ marginBottom: "1.5rem" }}>
        <p style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary)",
                    textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 10px" }}>
          2 · Exchange rate
        </p>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 28px 1fr 1fr", gap: 10, alignItems: "end" }}>
          <div>
            <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>
              From
            </label>
            <select value={fromCur} onChange={e => setFromCur(e.target.value)} style={sel}>
              {CURRENCIES.map(c => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div style={{ fontSize: 18, color: "var(--color-text-tertiary)", paddingBottom: 9, textAlign: "center" }}>→</div>
          <div>
            <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>
              To
            </label>
            <select value={toCur} onChange={e => setToCur(e.target.value)} style={sel}>
              {CURRENCIES.map(c => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>
              Rate (1 {fromCur} = ? {toCur})
            </label>
            <input type="number" min="0.0001" step="0.0001" value={rate}
              onChange={e => setRate(e.target.value)}
              style={{ ...sel, width: "100%" }} placeholder="e.g. 1.08" />
          </div>
        </div>
        <p style={{ fontSize: 12, color: "var(--color-text-tertiary)", margin: "8px 0 0" }}>{rateHint}</p>
      </section>

      {/* Error */}
      {error && (
        <div style={{ background: "var(--color-background-danger)", border: ".5px solid var(--color-border-danger)",
          borderRadius: "var(--border-radius-md)", padding: "10px 14px", marginBottom: "1rem",
          fontSize: 13, color: "var(--color-text-danger)" }}>
          {error}
        </div>
      )}

      {/* Preview button */}
      <button onClick={handlePreview} disabled={loading}
        style={{ ...btn(true), width: "100%", justifyContent: "center",
                 padding: 11, marginBottom: "2rem", opacity: loading ? .6 : 1 }}>
        {loading ? "Loading preview…" : "Preview Conversion →"}
      </button>

      {/* Results */}
      {result && (
        <section>
          <p style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary)",
                      textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 12px" }}>
            3 · Results
          </p>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))",
                        gap: 10, marginBottom: "1.5rem" }}>
            <StatCard label="Rows converted"           value={result.rows} />
            <StatCard label={`Total (${result.from_currency})`} value={fmt(result.total_original)} />
            <StatCard label={`Total (${result.to_currency})`}   value={fmt(result.total_converted)} accent />
            <StatCard label="Rate applied" value={result.exchange_rate}
                      sub={`1 ${result.from_currency} → ${result.to_currency}`} />
          </div>

          <p style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary)",
                      textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 10px" }}>
            Preview — first {result.preview.length} rows
          </p>
          <PreviewTable rows={result.preview} fromCur={result.from_currency} toCur={result.to_currency} />

          {/* Download buttons */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: "1.5rem" }}>
            <button onClick={() => handleDownload("xlsx")} disabled={!!dlLoading}
              style={{ ...btn(false), width: "100%", justifyContent: "center", opacity: dlLoading ? .6 : 1 }}>
              {dlLoading === "xlsx" ? "Generating…" : "⬇ Download Excel (.xlsx)"}
            </button>
            <button onClick={() => handleDownload("pdf")} disabled={!!dlLoading}
              style={{ ...btn(true), width: "100%", justifyContent: "center", opacity: dlLoading ? .6 : 1 }}>
              {dlLoading === "pdf" ? "Generating…" : "⬇ Download PDF Report"}
            </button>
          </div>
        </section>
      )}
    </div>
  );
}
