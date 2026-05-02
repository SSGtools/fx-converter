import { useState, useRef, useCallback, useEffect } from "react";

const API = import.meta.env.VITE_API_URL ?? "http://localhost:8000";

// frankfurter.app is free, no API key needed
const FX_API = "https://api.frankfurter.app";

const CURRENCIES = [
  "AED","AUD","BGN","BRL","CAD","CHF","CNY","CZK","DKK","EUR",
  "GBP","HKD","HUF","IDR","ILS","INR","ISK","JPY","KRW","MXN",
  "MYR","NOK","NZD","PHP","PLN","RON","SEK","SGD","THB","TRY",
  "USD","ZAR"
];

const CURRENCY_NAMES = {
  AED:"UAE Dirham", AUD:"Australian Dollar", BGN:"Bulgarian Lev",
  BRL:"Brazilian Real", CAD:"Canadian Dollar", CHF:"Swiss Franc",
  CNY:"Chinese Yuan", CZK:"Czech Koruna", DKK:"Danish Krone",
  EUR:"Euro", GBP:"British Pound", HKD:"Hong Kong Dollar",
  HUF:"Hungarian Forint", IDR:"Indonesian Rupiah", ILS:"Israeli Shekel",
  INR:"Indian Rupee", ISK:"Icelandic Króna", JPY:"Japanese Yen",
  KRW:"South Korean Won", MXN:"Mexican Peso", MYR:"Malaysian Ringgit",
  NOK:"Norwegian Krone", NZD:"New Zealand Dollar", PHP:"Philippine Peso",
  PLN:"Polish Zloty", RON:"Romanian Leu", SEK:"Swedish Krona",
  SGD:"Singapore Dollar", THB:"Thai Baht", TRY:"Turkish Lira",
  USD:"US Dollar", ZAR:"South African Rand"
};

const fmt = (n, d = 2) =>
  new Intl.NumberFormat("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }).format(n);

function buildForm(file, rate, fromCur, toCur) {
  const fd = new FormData();
  fd.append("file", file);
  fd.append("exchange_rate", rate);
  fd.append("from_currency", fromCur);
  fd.append("to_currency", toCur);
  return fd;
}

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

// ── Styles ────────────────────────────────────────────────────────────────────

const styles = {
  page: {
    minHeight: "100vh",
    background: "linear-gradient(135deg, #0f0c29, #302b63, #24243e)",
    padding: "2rem 1rem",
    fontFamily: "'Inter', 'Segoe UI', system-ui, sans-serif",
    color: "#f0f0f0",
  },
  card: {
    background: "rgba(255,255,255,0.06)",
    backdropFilter: "blur(12px)",
    borderRadius: 16,
    border: "1px solid rgba(255,255,255,0.12)",
    padding: "1.5rem",
    marginBottom: "1.5rem",
  },
  sectionHead: {
    fontSize: 11, fontWeight: 600, letterSpacing: ".1em",
    textTransform: "uppercase", color: "#a78bfa", marginBottom: 12,
  },
  input: {
    width: "100%", padding: "10px 12px",
    borderRadius: 10, border: "1px solid rgba(255,255,255,0.15)",
    background: "rgba(255,255,255,0.08)", color: "#f0f0f0",
    fontSize: 14, fontFamily: "inherit", outline: "none",
    transition: "border-color .2s",
  },
  select: {
    padding: "10px 12px", borderRadius: 10,
    border: "1px solid rgba(255,255,255,0.15)",
    background: "rgba(30,20,60,0.9)", color: "#f0f0f0",
    fontSize: 14, fontFamily: "inherit", cursor: "pointer", outline: "none",
    width: "100%",
  },
  btnPrimary: {
    width: "100%", padding: "13px", borderRadius: 12, border: "none",
    background: "linear-gradient(135deg, #7c3aed, #4f46e5)",
    color: "#fff", fontSize: 15, fontWeight: 600,
    cursor: "pointer", fontFamily: "inherit",
    boxShadow: "0 4px 20px rgba(124,58,237,0.4)",
    transition: "opacity .2s, transform .1s",
  },
  btnSecondary: {
    flex: 1, padding: "11px 16px", borderRadius: 12,
    border: "1px solid rgba(255,255,255,0.2)",
    background: "rgba(255,255,255,0.07)", color: "#f0f0f0",
    fontSize: 14, fontWeight: 500, cursor: "pointer", fontFamily: "inherit",
    transition: "background .2s",
  },
  btnAccent: {
    flex: 1, padding: "11px 16px", borderRadius: 12, border: "none",
    background: "linear-gradient(135deg, #7c3aed, #4f46e5)",
    color: "#fff", fontSize: 14, fontWeight: 600,
    cursor: "pointer", fontFamily: "inherit",
    boxShadow: "0 2px 12px rgba(124,58,237,0.35)",
    transition: "opacity .2s",
  },
  error: {
    background: "rgba(239,68,68,0.15)", border: "1px solid rgba(239,68,68,0.4)",
    borderRadius: 10, padding: "10px 14px", fontSize: 13, color: "#fca5a5",
    marginBottom: "1rem",
  },
  success: {
    background: "rgba(34,197,94,0.12)", border: "1px solid rgba(34,197,94,0.3)",
    borderRadius: 10, padding: "10px 14px", fontSize: 13, color: "#86efac",
    marginBottom: "1rem",
  },
};

// ── Sub-components ─────────────────────────────────────────────────────────────

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{
      background: accent ? "rgba(124,58,237,0.2)" : "rgba(255,255,255,0.06)",
      borderRadius: 12, padding: "1rem 1.25rem",
      border: accent ? "1px solid rgba(124,58,237,0.5)" : "1px solid rgba(255,255,255,0.1)",
    }}>
      <p style={{ fontSize: 11, color: "#a78bfa", margin: "0 0 4px",
                  textTransform: "uppercase", letterSpacing: ".06em" }}>{label}</p>
      <p style={{ fontSize: 20, fontWeight: 600, margin: 0,
                  color: accent ? "#c4b5fd" : "#f0f0f0" }}>{value}</p>
      {sub && <p style={{ fontSize: 11, color: "#9ca3af", margin: "4px 0 0" }}>{sub}</p>}
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
    <div
      onClick={() => ref.current.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={onDrop}
      style={{
        border: `2px dashed ${dragging ? "#7c3aed" : "rgba(255,255,255,0.2)"}`,
        borderRadius: 14, padding: "2.5rem 1rem", textAlign: "center", cursor: "pointer",
        background: dragging ? "rgba(124,58,237,0.1)" : "rgba(255,255,255,0.03)",
        transition: "all .15s",
      }}>
      <input ref={ref} type="file" accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
      <div style={{ fontSize: 36, marginBottom: 10 }}>📊</div>
      {file
        ? <p style={{ margin: 0, fontWeight: 600, color: "#c4b5fd", fontSize: 15 }}>{file.name}</p>
        : <>
            <p style={{ margin: "0 0 6px", fontWeight: 600, fontSize: 15 }}>
              Drop your Excel file here
            </p>
            <p style={{ margin: 0, fontSize: 13, color: "#9ca3af" }}>
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
    { key: "amount_original",  label: `Original (${fromCur})`,  right: true },
    { key: "amount_converted", label: `Converted (${toCur})`, right: true },
  ];
  return (
    <div style={{ overflowX: "auto", borderRadius: 12,
                  border: "1px solid rgba(255,255,255,0.1)" }}>
      <table style={{ width: "100%", borderCollapse: "collapse",
                      fontSize: 13, tableLayout: "fixed" }}>
        <thead>
          <tr style={{ background: "rgba(124,58,237,0.2)" }}>
            {cols.map(c => (
              <th key={c.key} style={{
                padding: "9px 12px", textAlign: c.right ? "right" : "left",
                fontWeight: 600, fontSize: 11, color: "#c4b5fd",
                textTransform: "uppercase", letterSpacing: ".06em",
                borderBottom: "1px solid rgba(255,255,255,0.1)",
              }}>{c.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{
              background: i % 2 === 0 ? "rgba(255,255,255,0.02)" : "rgba(255,255,255,0.05)"
            }}>
              {cols.map(c => {
                const v = row[c.key];
                const neg = c.right && v < 0;
                return (
                  <td key={c.key} style={{
                    padding: "8px 12px", textAlign: c.right ? "right" : "left",
                    borderBottom: "1px solid rgba(255,255,255,0.06)",
                    color: c.key === "amount_converted"
                      ? (neg ? "#f87171" : "#a78bfa")
                      : neg ? "#f87171" : "#e5e7eb",
                    fontWeight: c.key === "amount_converted" ? 600 : 400,
                    overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                  }}>
                    {c.right ? fmt(v) : String(v ?? "—")}
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

// ── Main App ──────────────────────────────────────────────────────────────────

export default function App() {
  const [file, setFile]           = useState(null);
  const [rate, setRate]           = useState("");
  const [fromCur, setFromCur]     = useState("USD");
  const [toCur, setToCur]         = useState("EUR");
  const [loading, setLoading]     = useState(false);
  const [dlLoading, setDlLoading] = useState("");
  const [rateLoading, setRateLoading] = useState(false);
  const [error, setError]         = useState("");
  const [rateSource, setRateSource] = useState(""); // "live" | "manual"
  const [rateDate, setRateDate]   = useState("");
  const [result, setResult]       = useState(null);

  // Fetch live rate whenever currencies change
  useEffect(() => {
    if (fromCur === toCur) { setRate("1"); setRateSource("manual"); return; }
    setRateLoading(true);
    setRateSource("");
    fetch(`${FX_API}/latest?from=${fromCur}&to=${toCur}`)
      .then(r => r.json())
      .then(data => {
        const r = data.rates?.[toCur];
        if (r) {
          setRate(String(r));
          setRateSource("live");
          setRateDate(data.date ?? "");
        }
      })
      .catch(() => setRateSource("manual"))
      .finally(() => setRateLoading(false));
  }, [fromCur, toCur]);

  const rateVal = parseFloat(rate);

  function buildForm(file, rate, fromCur, toCur) {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("exchange_rate", rate);
    fd.append("from_currency", fromCur);
    fd.append("to_currency", toCur);
    return fd;
  }

  const handlePreview = async () => {
    if (!file)          { setError("Please upload an Excel file."); return; }
    if (!(rateVal > 0)) { setError("Enter a valid positive exchange rate."); return; }
    setLoading(true); setError(""); setResult(null);
    try {
      const res  = await fetch(`${API}/preview`, { method: "POST", body: buildForm(file, rateVal, fromCur, toCur) });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail ?? "Preview failed.");
      setResult(data);
    } catch (e) {
      setError(e.message.includes("fetch")
        ? `Cannot reach backend at ${API}.`
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
    <div style={styles.page}>
      <div style={{ maxWidth: 740, margin: "0 auto" }}>

        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: "2.5rem" }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>💱</div>
          <h1 style={{ fontSize: 30, fontWeight: 700, margin: "0 0 8px",
                       background: "linear-gradient(135deg, #c4b5fd, #818cf8)",
                       WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
            FX Report Converter
          </h1>
          <p style={{ fontSize: 15, color: "#9ca3af", margin: 0 }}>
            Upload a financial report · apply live exchange rates · export to Excel &amp; PDF
          </p>
        </div>

        {/* Step 1 — Upload */}
        <div style={styles.card}>
          <p style={styles.sectionHead}>① Upload Report</p>
          <DropZone onFile={f => { setFile(f); setResult(null); setError(""); }} file={file} />
          {file && (
            <button onClick={() => { setFile(null); setResult(null); }}
              style={{ marginTop: 8, background: "none", border: "none", cursor: "pointer",
                       fontSize: 12, color: "#9ca3af", padding: 0 }}>
              Remove ×
            </button>
          )}
        </div>

        {/* Step 2 — Currency & Rate */}
        <div style={styles.card}>
          <p style={styles.sectionHead}>② Exchange Rate</p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 24px 1fr", gap: 12, marginBottom: 14, alignItems: "end" }}>
            <div>
              <label style={{ fontSize: 12, color: "#9ca3af", display: "block", marginBottom: 6 }}>From</label>
              <select value={fromCur} onChange={e => setFromCur(e.target.value)} style={styles.select}>
                {CURRENCIES.map(c => <option key={c} value={c}>{c} — {CURRENCY_NAMES[c]}</option>)}
              </select>
            </div>
            <div style={{ fontSize: 18, color: "#7c3aed", textAlign: "center", paddingBottom: 10 }}>⇄</div>
            <div>
              <label style={{ fontSize: 12, color: "#9ca3af", display: "block", marginBottom: 6 }}>To</label>
              <select value={toCur} onChange={e => setToCur(e.target.value)} style={styles.select}>
                {CURRENCIES.map(c => <option key={c} value={c}>{c} — {CURRENCY_NAMES[c]}</option>)}
              </select>
            </div>
          </div>

          <div>
            <label style={{ fontSize: 12, color: "#9ca3af", display: "block", marginBottom: 6 }}>
              Rate &nbsp;
              {rateLoading && <span style={{ color: "#a78bfa" }}>fetching live rate…</span>}
              {!rateLoading && rateSource === "live" && (
                <span style={{ color: "#86efac" }}>
                  ✓ Live rate · {rateDate}
                  <button onClick={() => setRateSource("manual")}
                    style={{ marginLeft: 8, background: "none", border: "none",
                             color: "#9ca3af", cursor: "pointer", fontSize: 11 }}>
                    edit manually
                  </button>
                </span>
              )}
              {!rateLoading && rateSource === "manual" && (
                <span style={{ color: "#fbbf24" }}>manual</span>
              )}
            </label>
            <input
              type="number" min="0.0001" step="0.0001"
              value={rate} onChange={e => { setRate(e.target.value); setRateSource("manual"); }}
              style={styles.input} placeholder="e.g. 1.08"
              readOnly={rateLoading}
            />
            {rateVal > 0 && (
              <p style={{ fontSize: 12, color: "#6b7280", margin: "6px 0 0" }}>
                1 {fromCur} = {fmt(rateVal, 4)} {toCur}
                &nbsp;·&nbsp;
                1 {toCur} = {fmt(1/rateVal, 4)} {fromCur}
              </p>
            )}
          </div>
        </div>

        {/* Error */}
        {error && <div style={styles.error}>{error}</div>}

        {/* Convert */}
        <button onClick={handlePreview} disabled={loading || rateLoading}
          style={{ ...styles.btnPrimary, opacity: (loading || rateLoading) ? .6 : 1,
                   marginBottom: "1.5rem" }}>
          {loading ? "Processing…" : "Preview Conversion →"}
        </button>

        {/* Results */}
        {result && (
          <div style={styles.card}>
            <p style={styles.sectionHead}>③ Results</p>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))",
                          gap: 10, marginBottom: "1.5rem" }}>
              <StatCard label="Rows"                              value={result.rows} />
              <StatCard label={`Total ${result.from_currency}`}   value={fmt(result.total_original)} />
              <StatCard label={`Total ${result.to_currency}`}     value={fmt(result.total_converted)} accent />
              <StatCard label="Rate" value={fmt(result.exchange_rate, 4)}
                        sub={`${result.from_currency} → ${result.to_currency}`} />
            </div>

            <p style={{ ...styles.sectionHead, marginBottom: 10 }}>
              Preview — first {result.preview.length} rows
            </p>
            <PreviewTable rows={result.preview}
                          fromCur={result.from_currency} toCur={result.to_currency} />

            <div style={{ display: "flex", gap: 12, marginTop: "1.5rem" }}>
              <button onClick={() => handleDownload("xlsx")} disabled={!!dlLoading}
                style={{ ...styles.btnSecondary, opacity: dlLoading ? .6 : 1 }}>
                {dlLoading === "xlsx" ? "Generating…" : "⬇ Excel (.xlsx)"}
              </button>
              <button onClick={() => handleDownload("pdf")} disabled={!!dlLoading}
                style={{ ...styles.btnAccent, opacity: dlLoading ? .6 : 1 }}>
                {dlLoading === "pdf" ? "Generating…" : "⬇ PDF Report"}
              </button>
            </div>
          </div>
        )}

        <p style={{ textAlign: "center", fontSize: 12, color: "#4b5563", marginTop: "1.5rem" }}>
          Live rates via frankfurter.app · No data stored on server
        </p>
      </div>
    </div>
  );
}
