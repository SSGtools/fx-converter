import { useState, useRef, useCallback, useEffect } from "react";

const API    = import.meta.env.VITE_API_URL ?? "http://localhost:8000";

const CURRENCIES = [
  "AED","AUD","BGN","BRL","CAD","CHF","CNY","CZK","DKK","EUR",
  "GBP","HKD","HUF","IDR","ILS","INR","ISK","JPY","KRW","MXN",
  "MYR","NOK","NZD","PHP","PLN","RON","SEK","SGD","THB","TRY",
  "USD","XCD","ZAR"
];
const CURRENCY_NAMES = {
  AED:"UAE Dirham",AUD:"Australian Dollar",BGN:"Bulgarian Lev",BRL:"Brazilian Real",
  CAD:"Canadian Dollar",CHF:"Swiss Franc",CNY:"Chinese Yuan",CZK:"Czech Koruna",
  DKK:"Danish Krone",EUR:"Euro",GBP:"British Pound",HKD:"Hong Kong Dollar",
  HUF:"Hungarian Forint",IDR:"Indonesian Rupiah",ILS:"Israeli Shekel",INR:"Indian Rupee",
  ISK:"Icelandic Króna",JPY:"Japanese Yen",KRW:"South Korean Won",MXN:"Mexican Peso",
  MYR:"Malaysian Ringgit",NOK:"Norwegian Krone",NZD:"New Zealand Dollar",PHP:"Philippine Peso",
  PLN:"Polish Zloty",RON:"Romanian Leu",SEK:"Swedish Krona",SGD:"Singapore Dollar",
  THB:"Thai Baht",TRY:"Turkish Lira",USD:"US Dollar",XCD:"East Caribbean Dollar",
  ZAR:"South African Rand"
};

const fmt = (n, d=2) => n==null?"—":
  new Intl.NumberFormat("en-US",{minimumFractionDigits:d,maximumFractionDigits:d}).format(n);

function buildForm(files, rate, fromCur, toCur, selectedCols) {
  const fd = new FormData();
  files.forEach(f => fd.append("files", f.file));
  fd.append("exchange_rate", rate);
  fd.append("from_currency", fromCur);
  fd.append("to_currency",   toCur);
  fd.append("selected_cols", JSON.stringify(selectedCols));
  return fd;
}

async function downloadFile(endpoint, form, filename) {
  const res = await fetch(`${API}${endpoint}`, { method:"POST", body:form });
  if (!res.ok) { const e = await res.json().catch(()=>({})); throw new Error(e.detail??"Download failed"); }
  const blob = await res.blob();
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

const S = {
  page:  { minHeight:"100vh", background:"linear-gradient(135deg,#0f0c29,#302b63,#24243e)",
           padding:"2rem 1rem", fontFamily:"'Inter','Segoe UI',system-ui,sans-serif", color:"#f0f0f0" },
  card:  { background:"rgba(255,255,255,0.06)", backdropFilter:"blur(12px)", borderRadius:16,
           border:"1px solid rgba(255,255,255,0.12)", padding:"1.5rem", marginBottom:"1.5rem" },
  sHead: { fontSize:11, fontWeight:600, letterSpacing:".1em", textTransform:"uppercase",
           color:"#a78bfa", marginBottom:12 },
  input: { width:"100%", padding:"10px 12px", borderRadius:10,
           border:"1px solid rgba(255,255,255,0.15)", background:"rgba(255,255,255,0.08)",
           color:"#f0f0f0", fontSize:14, fontFamily:"inherit", outline:"none" },
  select:{ padding:"10px 12px", borderRadius:10, border:"1px solid rgba(255,255,255,0.15)",
           background:"rgba(30,20,60,0.9)", color:"#f0f0f0", fontSize:14,
           fontFamily:"inherit", cursor:"pointer", outline:"none", width:"100%" },
  btnPri:{ width:"100%", padding:"13px", borderRadius:12, border:"none",
           background:"linear-gradient(135deg,#7c3aed,#4f46e5)", color:"#fff",
           fontSize:15, fontWeight:600, cursor:"pointer", fontFamily:"inherit",
           boxShadow:"0 4px 20px rgba(124,58,237,0.4)" },
  btnSec:{ flex:1, padding:"11px 16px", borderRadius:12,
           border:"1px solid rgba(255,255,255,0.2)", background:"rgba(255,255,255,0.07)",
           color:"#f0f0f0", fontSize:14, fontWeight:500, cursor:"pointer", fontFamily:"inherit" },
  btnAcc:{ flex:1, padding:"11px 16px", borderRadius:12, border:"none",
           background:"linear-gradient(135deg,#7c3aed,#4f46e5)", color:"#fff",
           fontSize:14, fontWeight:600, cursor:"pointer", fontFamily:"inherit" },
  error: { background:"rgba(239,68,68,0.15)", border:"1px solid rgba(239,68,68,0.4)",
           borderRadius:10, padding:"10px 14px", fontSize:13, color:"#fca5a5", marginBottom:"1rem" },
};

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{ background:accent?"rgba(124,58,237,0.2)":"rgba(255,255,255,0.06)",
                  borderRadius:12, padding:"1rem 1.25rem",
                  border:accent?"1px solid rgba(124,58,237,0.5)":"1px solid rgba(255,255,255,0.1)" }}>
      <p style={{fontSize:11,color:"#a78bfa",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:".06em"}}>{label}</p>
      <p style={{fontSize:18,fontWeight:600,margin:0,color:accent?"#c4b5fd":"#f0f0f0"}}>{value}</p>
      {sub&&<p style={{fontSize:11,color:"#9ca3af",margin:"4px 0 0"}}>{sub}</p>}
    </div>
  );
}

// ── Drop Zone (multi-file) ────────────────────────────────────────────────────
function DropZone({ onFiles }) {
  const [drag,setDrag] = useState(false);
  const ref = useRef();

  const process = useCallback(fileList => {
    const valid = Array.from(fileList).filter(f =>
      f.name.endsWith(".xlsx") || f.name.endsWith(".xls"));
    if (valid.length) onFiles(valid);
  }, [onFiles]);

  const onDrop = useCallback(e => {
    e.preventDefault(); setDrag(false);
    process(e.dataTransfer.files);
  }, [process]);

  return (
    <div onClick={()=>ref.current.click()}
         onDragOver={e=>{e.preventDefault();setDrag(true);}}
         onDragLeave={()=>setDrag(false)} onDrop={onDrop}
         style={{ border:`2px dashed ${drag?"#7c3aed":"rgba(255,255,255,0.2)"}`,
                  borderRadius:14, padding:"2.5rem 1rem", textAlign:"center", cursor:"pointer",
                  background:drag?"rgba(124,58,237,0.1)":"rgba(255,255,255,0.03)", transition:"all .15s" }}>
      <input ref={ref} type="file" accept=".xlsx,.xls" multiple style={{display:"none"}}
             onChange={e=>process(e.target.files)}/>
      <div style={{fontSize:36,marginBottom:10}}>📂</div>
      <p style={{margin:"0 0 6px",fontWeight:600,fontSize:15}}>Drop Excel reports here</p>
      <p style={{margin:0,fontSize:13,color:"#9ca3af"}}>
        Multiple files supported · .xlsx or .xls · P&amp;L, Balance Sheet, multi-branch
      </p>
    </div>
  );
}

// ── File Queue Item ───────────────────────────────────────────────────────────
function FileItem({ item, onRemove }) {
  const statusColor = {
    pending:    "#9ca3af",
    detecting:  "#a78bfa",
    ready:      "#86efac",
    error:      "#f87171",
  }[item.status] ?? "#9ca3af";

  const statusLabel = {
    pending:   "Pending…",
    detecting: "Detecting…",
    ready:     `${item.detected?.sheet_name} · ${item.detected?.row_count} rows`,
    error:     item.error ?? "Error",
  }[item.status];

  return (
    <div style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 14px",
                  borderRadius:10, background:"rgba(255,255,255,0.04)",
                  border:"1px solid rgba(255,255,255,0.08)", marginBottom:8 }}>
      <div style={{fontSize:20}}>📊</div>
      <div style={{flex:1,minWidth:0}}>
        <p style={{margin:0,fontWeight:500,fontSize:13,color:"#f0f0f0",
                   overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
          {item.file.name}
        </p>
        <p style={{margin:0,fontSize:11,color:statusColor}}>{statusLabel}</p>
      </div>
      {item.status==="ready"&&item.detected?.is_multi&&(
        <span style={{fontSize:10,padding:"2px 8px",borderRadius:999,
                      background:"rgba(124,58,237,0.25)",color:"#c4b5fd",
                      border:"1px solid rgba(124,58,237,0.3)",whiteSpace:"nowrap"}}>
          {item.detected.branch_cols.length} branches
        </span>
      )}
      <button onClick={()=>onRemove(item.id)}
              style={{background:"none",border:"none",cursor:"pointer",
                      color:"#6b7280",fontSize:16,padding:"0 4px",lineHeight:1}}>×</button>
    </div>
  );
}

// ── Summary Table (post-convert) ──────────────────────────────────────────────
function SummaryTable({ reports, fromCur, toCur }) {
  if (!reports?.length) return null;
  let grandO = 0, grandC = 0;
  return (
    <div style={{overflowX:"auto",borderRadius:12,border:"1px solid rgba(255,255,255,0.1)",marginBottom:"1rem"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead>
          <tr style={{background:"rgba(124,58,237,0.25)"}}>
            {["Report","Sheet",`Original (${fromCur})`,`Converted (${toCur})`, "Difference"].map((h,i)=>(
              <th key={i} style={{padding:"9px 12px",textAlign:i>=2?"right":"left",
                fontWeight:600,fontSize:10,color:"#c4b5fd",textTransform:"uppercase",
                letterSpacing:".06em",borderBottom:"1px solid rgba(255,255,255,0.1)",
                whiteSpace:"nowrap"}}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {reports.map((r,i)=>{
            const o=r.total_orig??0, c=r.total_conv??0, diff=c-o;
            grandO+=o; grandC+=c;
            return (
              <tr key={i} style={{background:i%2===0?"rgba(255,255,255,0.02)":"rgba(255,255,255,0.05)"}}>
                <td style={{padding:"8px 12px",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:"#e5e7eb",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:200}}>
                  {r.filename}</td>
                <td style={{padding:"8px 12px",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:"#9ca3af",fontSize:12}}>{r.sheet_name}</td>
                <td style={{padding:"8px 12px",textAlign:"right",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:"#e5e7eb"}}>{fmt(o)}</td>
                <td style={{padding:"8px 12px",textAlign:"right",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:"#a78bfa",fontWeight:600}}>{fmt(c)}</td>
                <td style={{padding:"8px 12px",textAlign:"right",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:diff>=0?"#86efac":"#f87171",fontWeight:500}}>{fmt(diff)}</td>
              </tr>
            );
          })}
          <tr style={{background:"rgba(15,52,96,0.6)"}}>
            <td colSpan={2} style={{padding:"9px 12px",fontWeight:700,color:"#e8d5b7",fontSize:13}}>
              GRAND TOTAL</td>
            <td style={{padding:"9px 12px",textAlign:"right",fontWeight:700,color:"#e8d5b7"}}>{fmt(grandO)}</td>
            <td style={{padding:"9px 12px",textAlign:"right",fontWeight:700,color:"#c4b5fd"}}>{fmt(grandC)}</td>
            <td style={{padding:"9px 12px",textAlign:"right",fontWeight:700,
              color:(grandC-grandO)>=0?"#86efac":"#f87171"}}>{fmt(grandC-grandO)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [queue,setQueue]           = useState([]);  // [{id,file,status,detected,error}]
  const [rate,setRate]             = useState("");
  const [fromCur,setFromCur]       = useState("USD");
  const [toCur,setToCur]           = useState("EUR");
  const [rateLoading,setRateLoading] = useState(false);
  const [rateSource,setRateSource] = useState("");
  const [rateDate,setRateDate]     = useState("");
  const [loading,setLoading]       = useState(false);
  const [dlLoading,setDlLoading]   = useState("");
  const [error,setError]           = useState("");
  const [results,setResults]       = useState(null); // [{filename,sheet_name,total_orig,total_conv}]

  // Live rate fetch
  useEffect(()=>{
    if(!fromCur||!toCur||fromCur===toCur){setRate("1");setRateSource("manual");return;}
    setRateLoading(true); setRateSource("");
    const tryFetch = async () => {
      try {
        const r = await fetch(`https://api.frankfurter.dev/v2/rates?base=${fromCur}&quotes=${toCur}`);
        if(r.ok){ const d=await r.json(); const v=d.rates?.[toCur]; if(v) return {rate:v,date:d.date??""}; }
      } catch(_){}
      try {
        const r = await fetch(`https://open.er-api.com/v6/latest/${fromCur}`);
        if(r.ok){ const d=await r.json(); const v=d.rates?.[toCur]; if(v) return {rate:v,date:d.time_last_update_utc?.slice(0,10)??""}; }
      } catch(_){}
      try {
        const r = await fetch(`https://api.frankfurter.app/latest?from=${fromCur}&to=${toCur}`);
        if(r.ok){ const d=await r.json(); const v=d.rates?.[toCur]; if(v) return {rate:v,date:d.date??""}; }
      } catch(_){}
      return null;
    };
    tryFetch().then(res=>{
      if(res){ setRate(res.rate.toFixed(6).replace(/\.?0+$/,"")); setRateSource("live"); setRateDate(res.date); }
      else { setRateSource("manual"); }
    }).finally(()=>setRateLoading(false));
  },[fromCur,toCur]);

  const rateVal = parseFloat(rate);

  // Add files to queue and auto-detect
  const handleFiles = useCallback(async (newFiles) => {
    const items = newFiles.map(f => ({
      id: `${f.name}-${Date.now()}-${Math.random()}`,
      file: f, status: "detecting", detected: null, error: null
    }));
    setQueue(q => [...q, ...items]);
    setResults(null); setError("");

    // Detect each file
    for (const item of items) {
      const fd = new FormData(); fd.append("file", item.file);
      try {
        const res  = await fetch(`${API}/detect`, { method:"POST", body:fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail ?? "Detection failed");
        setQueue(q => q.map(qi => qi.id===item.id
          ? {...qi, status:"ready", detected:data} : qi));
      } catch(e) {
        setQueue(q => q.map(qi => qi.id===item.id
          ? {...qi, status:"error", error:e.message} : qi));
      }
    }
  }, []);

  const removeFile = useCallback(id => {
    setQueue(q => q.filter(qi => qi.id !== id));
    setResults(null);
  }, []);

  const readyFiles = queue.filter(q => q.status === "ready");

  const handleConvert = async (format) => {
    if (!readyFiles.length)  { setError("Add at least one Excel file."); return; }
    if (!(rateVal > 0))      { setError("Enter a valid exchange rate."); return; }
    setDlLoading(format); setError("");

    // For multi-file we need all selected cols — use all branch cols per file
    const form = new FormData();
    readyFiles.forEach(item => form.append("files", item.file));
    form.append("exchange_rate", rateVal);
    form.append("from_currency", fromCur);
    form.append("to_currency",   toCur);
    form.append("selected_cols", JSON.stringify([])); // backend defaults to all

    const filename = readyFiles.length === 1
      ? `converted_report.${format}`
      : `financial_reports_package.${format}`;

    try {
      await downloadFile(`/download/${format}`, form, filename);
      // Build summary from detected data for display
      setResults(readyFiles.map(item => ({
        filename: item.file.name.replace(/\.xlsx?$/i,""),
        sheet_name: item.detected?.sheet_name ?? "",
        total_orig: null, // not available without full preview
        total_conv: null,
      })));
    } catch(e) {
      setError(e.message);
    } finally { setDlLoading(""); }
  };

  return (
    <div style={S.page}>
      <div style={{maxWidth:820,margin:"0 auto"}}>

        {/* Header */}
        <div style={{textAlign:"center",marginBottom:"2.5rem"}}>
          <div style={{fontSize:40,marginBottom:12}}>💱</div>
          <h1 style={{fontSize:30,fontWeight:700,margin:"0 0 8px",
            background:"linear-gradient(135deg,#c4b5fd,#818cf8)",
            WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>
            FX Report Converter
          </h1>
          <p style={{fontSize:15,color:"#9ca3af",margin:0}}>
            Upload multiple financial reports · apply live exchange rates · export a professional package
          </p>
        </div>

        {/* Step 1 — Upload */}
        <div style={S.card}>
          <p style={S.sHead}>① Upload Reports</p>
          <DropZone onFiles={handleFiles}/>

          {queue.length > 0 && (
            <div style={{marginTop:16}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <p style={{...S.sHead,margin:0}}>
                  {queue.length} file{queue.length!==1?"s":""} queued · {readyFiles.length} ready
                </p>
                <button onClick={()=>{setQueue([]);setResults(null);}}
                  style={{background:"none",border:"none",cursor:"pointer",
                          fontSize:12,color:"#9ca3af",padding:0}}>
                  clear all
                </button>
              </div>
              {queue.map(item => (
                <FileItem key={item.id} item={item} onRemove={removeFile}/>
              ))}
            </div>
          )}
        </div>

        {/* Step 2 — Rate */}
        <div style={S.card}>
          <p style={S.sHead}>② Exchange Rate</p>
          <div style={{display:"grid",gridTemplateColumns:"1fr 24px 1fr",gap:12,marginBottom:14,alignItems:"end"}}>
            <div>
              <label style={{fontSize:12,color:"#9ca3af",display:"block",marginBottom:6}}>From</label>
              <select value={fromCur} onChange={e=>setFromCur(e.target.value)} style={S.select}>
                {CURRENCIES.map(c=><option key={c} value={c}>{c} — {CURRENCY_NAMES[c]}</option>)}
              </select>
            </div>
            <div style={{fontSize:18,color:"#7c3aed",textAlign:"center",paddingBottom:10}}>⇄</div>
            <div>
              <label style={{fontSize:12,color:"#9ca3af",display:"block",marginBottom:6}}>To</label>
              <select value={toCur} onChange={e=>setToCur(e.target.value)} style={S.select}>
                {CURRENCIES.map(c=><option key={c} value={c}>{c} — {CURRENCY_NAMES[c]}</option>)}
              </select>
            </div>
          </div>
          <div>
            <label style={{fontSize:12,color:"#9ca3af",display:"block",marginBottom:6}}>
              Rate &nbsp;
              {rateLoading&&<span style={{color:"#a78bfa"}}>fetching live rate…</span>}
              {!rateLoading&&rateSource==="live"&&(
                <span style={{color:"#86efac"}}>✓ Live · {rateDate}
                  <button onClick={()=>setRateSource("manual")}
                    style={{marginLeft:8,background:"none",border:"none",
                            color:"#9ca3af",cursor:"pointer",fontSize:11}}>
                    edit manually
                  </button>
                </span>
              )}
              {!rateLoading&&rateSource==="manual"&&rate&&<span style={{color:"#fbbf24"}}>manual</span>}
              {!rateLoading&&rateSource==="manual"&&!rate&&(
                <span style={{color:"#f87171"}}>live rate unavailable — enter manually</span>
              )}
            </label>
            <input type="number" min="0.0001" step="0.0001" value={rate}
              onChange={e=>{setRate(e.target.value);setRateSource("manual");}}
              style={S.input} placeholder="e.g. 1.08" readOnly={rateLoading}/>
            {rateVal>0&&(
              <p style={{fontSize:12,color:"#6b7280",margin:"6px 0 0"}}>
                1 {fromCur} = {fmt(rateVal,4)} {toCur} · 1 {toCur} = {fmt(1/rateVal,4)} {fromCur}
              </p>
            )}
          </div>
        </div>

        {error&&<div style={S.error}>{error}</div>}

        {/* Download buttons */}
        {readyFiles.length > 0 && (
          <div style={{display:"flex",gap:12,marginBottom:"1.5rem"}}>
            <button onClick={()=>handleConvert("xlsx")} disabled={!!dlLoading||rateLoading||!(rateVal>0)}
              style={{...S.btnSec,opacity:(dlLoading||rateLoading||!(rateVal>0))?.6:1}}>
              {dlLoading==="xlsx"
                ? "Generating…"
                : readyFiles.length>1
                  ? `⬇ Excel Package (${readyFiles.length} reports, 1 file)`
                  : "⬇ Excel Report"}
            </button>
            <button onClick={()=>handleConvert("pdf")} disabled={!!dlLoading||rateLoading||!(rateVal>0)}
              style={{...S.btnAcc,opacity:(dlLoading||rateLoading||!(rateVal>0))?.6:1}}>
              {dlLoading==="pdf"
                ? "Generating…"
                : readyFiles.length>1
                  ? `⬇ PDF Package (${readyFiles.length} reports)`
                  : "⬇ PDF Report"}
            </button>
          </div>
        )}

        {/* What's included info box */}
        {readyFiles.length > 1 && (
          <div style={{...S.card,background:"rgba(124,58,237,0.08)",
                       border:"1px solid rgba(124,58,237,0.2)",marginTop:"-0.5rem"}}>
            <p style={{...S.sHead,marginBottom:8}}>What's included in the package</p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,fontSize:13,color:"#c4b5fd"}}>
              <div>
                <p style={{margin:"0 0 4px",fontWeight:600}}>📊 Excel Package</p>
                <p style={{margin:0,color:"#9ca3af",fontSize:12}}>
                  Summary tab with grand totals · One tab per report · Professional styling
                </p>
              </div>
              <div>
                <p style={{margin:"0 0 4px",fontWeight:600}}>📄 PDF Package</p>
                <p style={{margin:0,color:"#9ca3af",fontSize:12}}>
                  Cover page with summary table · One section per report · Page breaks between reports
                </p>
              </div>
            </div>
          </div>
        )}

        <p style={{textAlign:"center",fontSize:12,color:"#4b5563",marginTop:"1rem"}}>
          Live rates via frankfurter.dev · Supports P&amp;L, Balance Sheet, multi-branch · No data stored
        </p>
      </div>
    </div>
  );
}
