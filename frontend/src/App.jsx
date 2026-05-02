import { useState, useRef, useCallback, useEffect } from "react";

const API = import.meta.env.VITE_API_URL ?? "http://localhost:8000";
const FX_API = "https://api.frankfurter.app";

const CURRENCIES = [
  "AED","AUD","BGN","BRL","CAD","CHF","CNY","CZK","DKK","EUR",
  "GBP","HKD","HUF","IDR","ILS","INR","ISK","JPY","KRW","MXN",
  "MYR","NOK","NZD","PHP","PLN","RON","SEK","SGD","THB","TRY",
  "USD","ZAR"
];

const CURRENCY_NAMES = {
  AED:"UAE Dirham",AUD:"Australian Dollar",BGN:"Bulgarian Lev",BRL:"Brazilian Real",
  CAD:"Canadian Dollar",CHF:"Swiss Franc",CNY:"Chinese Yuan",CZK:"Czech Koruna",
  DKK:"Danish Krone",EUR:"Euro",GBP:"British Pound",HKD:"Hong Kong Dollar",
  HUF:"Hungarian Forint",IDR:"Indonesian Rupiah",ILS:"Israeli Shekel",INR:"Indian Rupee",
  ISK:"Icelandic Króna",JPY:"Japanese Yen",KRW:"South Korean Won",MXN:"Mexican Peso",
  MYR:"Malaysian Ringgit",NOK:"Norwegian Krone",NZD:"New Zealand Dollar",PHP:"Philippine Peso",
  PLN:"Polish Zloty",RON:"Romanian Leu",SEK:"Swedish Krona",SGD:"Singapore Dollar",
  THB:"Thai Baht",TRY:"Turkish Lira",USD:"US Dollar",ZAR:"South African Rand"
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

const S = {
  page: {
    minHeight:"100vh",
    background:"linear-gradient(135deg,#0f0c29,#302b63,#24243e)",
    padding:"2rem 1rem",
    fontFamily:"'Inter','Segoe UI',system-ui,sans-serif",
    color:"#f0f0f0",
  },
  card: {
    background:"rgba(255,255,255,0.06)",
    backdropFilter:"blur(12px)",
    borderRadius:16,
    border:"1px solid rgba(255,255,255,0.12)",
    padding:"1.5rem",
    marginBottom:"1.5rem",
  },
  sHead: {
    fontSize:11,fontWeight:600,letterSpacing:".1em",
    textTransform:"uppercase",color:"#a78bfa",marginBottom:12,
  },
  input: {
    width:"100%",padding:"10px 12px",borderRadius:10,
    border:"1px solid rgba(255,255,255,0.15)",
    background:"rgba(255,255,255,0.08)",color:"#f0f0f0",
    fontSize:14,fontFamily:"inherit",outline:"none",
  },
  select: {
    padding:"10px 12px",borderRadius:10,
    border:"1px solid rgba(255,255,255,0.15)",
    background:"rgba(30,20,60,0.9)",color:"#f0f0f0",
    fontSize:14,fontFamily:"inherit",cursor:"pointer",outline:"none",width:"100%",
  },
  btnPrimary: {
    width:"100%",padding:"13px",borderRadius:12,border:"none",
    background:"linear-gradient(135deg,#7c3aed,#4f46e5)",
    color:"#fff",fontSize:15,fontWeight:600,cursor:"pointer",
    fontFamily:"inherit",boxShadow:"0 4px 20px rgba(124,58,237,0.4)",
  },
  btnSec: {
    flex:1,padding:"11px 16px",borderRadius:12,
    border:"1px solid rgba(255,255,255,0.2)",
    background:"rgba(255,255,255,0.07)",color:"#f0f0f0",
    fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"inherit",
  },
  btnAcc: {
    flex:1,padding:"11px 16px",borderRadius:12,border:"none",
    background:"linear-gradient(135deg,#7c3aed,#4f46e5)",
    color:"#fff",fontSize:14,fontWeight:600,cursor:"pointer",
    fontFamily:"inherit",boxShadow:"0 2px 12px rgba(124,58,237,0.35)",
  },
  error: {
    background:"rgba(239,68,68,0.15)",border:"1px solid rgba(239,68,68,0.4)",
    borderRadius:10,padding:"10px 14px",fontSize:13,color:"#fca5a5",marginBottom:"1rem",
  },
};

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{
      background: accent ? "rgba(124,58,237,0.2)" : "rgba(255,255,255,0.06)",
      borderRadius:12, padding:"1rem 1.25rem",
      border: accent ? "1px solid rgba(124,58,237,0.5)" : "1px solid rgba(255,255,255,0.1)",
    }}>
      <p style={{fontSize:11,color:"#a78bfa",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:".06em"}}>{label}</p>
      <p style={{fontSize:20,fontWeight:600,margin:0,color:accent?"#c4b5fd":"#f0f0f0"}}>{value}</p>
      {sub && <p style={{fontSize:11,color:"#9ca3af",margin:"4px 0 0"}}>{sub}</p>}
    </div>
  );
}

function DropZone({ onFile, file }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  const onDrop = useCallback(e => {
    e.preventDefault(); setDrag(false);
    const f = e.dataTransfer.files[0];
    if (f && (f.name.endsWith(".xlsx")||f.name.endsWith(".xls"))) onFile(f);
  }, [onFile]);
  return (
    <div onClick={() => ref.current.click()}
      onDragOver={e=>{e.preventDefault();setDrag(true);}}
      onDragLeave={()=>setDrag(false)} onDrop={onDrop}
      style={{
        border:`2px dashed ${drag?"#7c3aed":"rgba(255,255,255,0.2)"}`,
        borderRadius:14,padding:"2.5rem 1rem",textAlign:"center",cursor:"pointer",
        background:drag?"rgba(124,58,237,0.1)":"rgba(255,255,255,0.03)",transition:"all .15s",
      }}>
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}}
        onChange={e=>{if(e.target.files[0])onFile(e.target.files[0]);}}/>
      <div style={{fontSize:36,marginBottom:10}}>📊</div>
      {file
        ? <p style={{margin:0,fontWeight:600,color:"#c4b5fd",fontSize:15}}>{file.name}</p>
        : <>
            <p style={{margin:"0 0 6px",fontWeight:600,fontSize:15}}>Drop your Excel report here</p>
            <p style={{margin:0,fontSize:13,color:"#9ca3af"}}>
              .xlsx or .xls · Profit &amp; Loss, Balance Sheet, and similar reports supported
            </p>
          </>
      }
    </div>
  );
}

function PreviewTable({ rows, fromCur, toCur }) {
  if (!rows?.length) return null;
  return (
    <div style={{overflowX:"auto",borderRadius:12,border:"1px solid rgba(255,255,255,0.1)"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,tableLayout:"fixed"}}>
        <thead>
          <tr style={{background:"rgba(124,58,237,0.25)"}}>
            {["Code","Account Name",`Original (${fromCur})`,`Converted (${toCur})`].map((h,i)=>(
              <th key={i} style={{
                padding:"9px 12px",textAlign:i>=2?"right":"left",
                fontWeight:600,fontSize:11,color:"#c4b5fd",
                textTransform:"uppercase",letterSpacing:".06em",
                borderBottom:"1px solid rgba(255,255,255,0.1)",
              }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => {
            const isHdr = row.row_type === "header";
            return (
              <tr key={i} style={{
                background: isHdr
                  ? "rgba(45,43,85,0.8)"
                  : i%2===0?"rgba(255,255,255,0.02)":"rgba(255,255,255,0.05)"
              }}>
                <td style={{padding:"7px 12px",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:"#6b7280",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  {isHdr ? "" : String(row.account||"")}
                </td>
                <td style={{padding:"7px 12px",borderBottom:"1px solid rgba(255,255,255,0.06)",
                  color:isHdr?"#c4b5fd":"#e5e7eb",fontWeight:isHdr?600:400,
                  overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                  paddingLeft:isHdr?"12px":"20px"}}>
                  {String(row.description||"")}
                </td>
                {[row.amount_original, row.amount_converted].map((v,vi)=>(
                  <td key={vi} style={{
                    padding:"7px 12px",textAlign:"right",
                    borderBottom:"1px solid rgba(255,255,255,0.06)",
                    color: isHdr?"#c4b5fd": vi===1 ? (v<0?"#f87171":"#a78bfa") : (v<0?"#f87171":"#e5e7eb"),
                    fontWeight:isHdr||vi===1?600:400,
                  }}>
                    {row.amount_original===0&&isHdr?"":fmt(v)}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export default function App() {
  const [file,setFile]           = useState(null);
  const [rate,setRate]           = useState("");
  const [fromCur,setFromCur]     = useState("USD");
  const [toCur,setToCur]         = useState("EUR");
  const [loading,setLoading]     = useState(false);
  const [dlLoading,setDlLoading] = useState("");
  const [rateLoading,setRateLoading] = useState(false);
  const [error,setError]         = useState("");
  const [rateSource,setRateSource] = useState("");
  const [rateDate,setRateDate]   = useState("");
  const [result,setResult]       = useState(null);

  useEffect(()=>{
    if(fromCur===toCur){setRate("1");setRateSource("manual");return;}
    setRateLoading(true); setRateSource("");
    fetch(`${FX_API}/latest?from=${fromCur}&to=${toCur}`)
      .then(r=>r.json())
      .then(data=>{
        const r=data.rates?.[toCur];
        if(r){setRate(String(r));setRateSource("live");setRateDate(data.date??"");}
      })
      .catch(()=>setRateSource("manual"))
      .finally(()=>setRateLoading(false));
  },[fromCur,toCur]);

  const rateVal = parseFloat(rate);

  const handlePreview = async () => {
    if(!file){setError("Please upload an Excel file.");return;}
    if(!(rateVal>0)){setError("Enter a valid positive exchange rate.");return;}
    setLoading(true);setError("");setResult(null);
    try{
      const res  = await fetch(`${API}/preview`,{method:"POST",body:buildForm(file,rateVal,fromCur,toCur)});
      const data = await res.json();
      if(!res.ok) throw new Error(data.detail??"Preview failed.");
      setResult(data);
    }catch(e){
      setError(e.message.includes("fetch")?`Cannot reach backend at ${API}.`:e.message);
    }finally{setLoading(false);}
  };

  const handleDownload = async (format) => {
    if(!file||!(rateVal>0))return;
    setDlLoading(format);setError("");
    try{
      await downloadFile(`/download/${format}`,buildForm(file,rateVal,fromCur,toCur),`converted_report.${format}`);
    }catch(e){setError(e.message);}
    finally{setDlLoading("");}
  };

  return (
    <div style={S.page}>
      <div style={{maxWidth:760,margin:"0 auto"}}>

        {/* Header */}
        <div style={{textAlign:"center",marginBottom:"2.5rem"}}>
          <div style={{fontSize:40,marginBottom:12}}>💱</div>
          <h1 style={{fontSize:30,fontWeight:700,margin:"0 0 8px",
            background:"linear-gradient(135deg,#c4b5fd,#818cf8)",
            WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>
            FX Report Converter
          </h1>
          <p style={{fontSize:15,color:"#9ca3af",margin:0}}>
            Upload a financial report · apply live exchange rates · export to Excel &amp; PDF
          </p>
        </div>

        {/* Upload */}
        <div style={S.card}>
          <p style={S.sHead}>① Upload Report</p>
          <DropZone onFile={f=>{setFile(f);setResult(null);setError("");}} file={file}/>
          {file&&(
            <button onClick={()=>{setFile(null);setResult(null);}}
              style={{marginTop:8,background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#9ca3af",padding:0}}>
              Remove ×
            </button>
          )}
        </div>

        {/* Rate */}
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
              {rateLoading && <span style={{color:"#a78bfa"}}>fetching live rate…</span>}
              {!rateLoading&&rateSource==="live"&&(
                <span style={{color:"#86efac"}}>
                  ✓ Live rate · {rateDate}
                  <button onClick={()=>setRateSource("manual")}
                    style={{marginLeft:8,background:"none",border:"none",color:"#9ca3af",cursor:"pointer",fontSize:11}}>
                    edit manually
                  </button>
                </span>
              )}
              {!rateLoading&&rateSource==="manual"&&<span style={{color:"#fbbf24"}}>manual</span>}
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

        <button onClick={handlePreview} disabled={loading||rateLoading}
          style={{...S.btnPrimary,opacity:(loading||rateLoading)?.6:1,marginBottom:"1.5rem"}}>
          {loading?"Processing…":"Preview Conversion →"}
        </button>

        {result&&(
          <div style={S.card}>
            <p style={S.sHead}>③ Results — {result.sheet_name}</p>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10,marginBottom:"1.5rem"}}>
              <StatCard label="Data rows"                         value={result.rows}/>
              <StatCard label={`Total ${result.from_currency}`}   value={fmt(result.total_original)}/>
              <StatCard label={`Total ${result.to_currency}`}     value={fmt(result.total_converted)} accent/>
              <StatCard label="Rate" value={fmt(result.exchange_rate,4)} sub={`${result.from_currency} → ${result.to_currency}`}/>
            </div>
            <p style={{...S.sHead,marginBottom:10}}>Preview — first {result.preview.length} rows</p>
            <PreviewTable rows={result.preview} fromCur={result.from_currency} toCur={result.to_currency}/>
            <div style={{display:"flex",gap:12,marginTop:"1.5rem"}}>
              <button onClick={()=>handleDownload("xlsx")} disabled={!!dlLoading}
                style={{...S.btnSec,opacity:dlLoading?.6:1}}>
                {dlLoading==="xlsx"?"Generating…":"⬇ Excel (.xlsx)"}
              </button>
              <button onClick={()=>handleDownload("pdf")} disabled={!!dlLoading}
                style={{...S.btnAcc,opacity:dlLoading?.6:1}}>
                {dlLoading==="pdf"?"Generating…":"⬇ PDF Report"}
              </button>
            </div>
          </div>
        )}

        <p style={{textAlign:"center",fontSize:12,color:"#4b5563",marginTop:"1.5rem"}}>
          Live rates via frankfurter.app · Supports P&amp;L, Balance Sheet &amp; more · No data stored
        </p>
      </div>
    </div>
  );
}
