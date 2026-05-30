import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

const API = import.meta.env.VITE_API_BASE || "/api";

// =====================================================================
// CSS — Bloomberg-style dense data terminal
// =====================================================================
const css = `
:root{
  --bg:#0a0e14;--bg2:#0f1419;--bg3:#141b22;--bg4:#1c2530;
  --border:#1f2937;--border2:#374151;
  --text:#e6e9ed;--dim:#7a8a9a;--dimmer:#4b5563;
  --green:#52ffae;--green-dim:#0d8050;--green-glow:rgba(82,255,174,0.15);
  --red:#ff5e6e;--red-dim:#7f1d1d;--red-bg:rgba(255,94,110,0.1);
  --yellow:#ffd166;--yellow-bg:rgba(255,209,102,0.1);
  --blue:#5e9eff;--purple:#b794f6;--orange:#ff9a3c;
  --mono:'Space Mono',ui-monospace,Menlo,monospace;
}
.light{
  --bg:#f7f8fa;--bg2:#ffffff;--bg3:#eef1f5;--bg4:#e2e7ee;
  --border:#d1d8e0;--border2:#a8b3bf;
  --text:#0f172a;--dim:#475569;--dimmer:#94a3b8;
  --green:#10b981;--green-dim:#a7f3d0;--green-glow:rgba(16,185,129,0.12);
  --red:#ef4444;--red-bg:rgba(239,68,68,0.08);
  --yellow:#f59e0b;--yellow-bg:rgba(245,158,11,0.1);
}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--text);font-family:var(--mono);font-size:13px;line-height:1.4;}
.app{max-width:1480px;margin:0 auto;padding:14px 18px;}

/* ===== TICKER STRIP (top) ===== */
.ticker-strip{display:flex;gap:24px;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:8px 14px;margin-bottom:12px;overflow:hidden;white-space:nowrap;}
.ticker-item{font-size:11px;letter-spacing:0.5px;}
.ticker-sym{color:var(--dim);font-weight:700;}
.ticker-price{margin-left:6px;}
.ticker-chg{margin-left:6px;font-size:10px;}
.up{color:var(--green);}
.down{color:var(--red);}

/* ===== TOPBAR ===== */
.topbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;padding:0 4px;}
.brand{display:flex;align-items:center;gap:12px;}
.brand-name{font-size:11px;letter-spacing:4px;color:var(--green);font-weight:700;text-transform:uppercase;}
.brand-tag{font-size:10px;color:var(--dim);letter-spacing:2px;}
.market-status{font-size:9px;letter-spacing:2px;padding:3px 8px;border-radius:3px;background:var(--bg3);text-transform:uppercase;font-weight:700;}
.market-status.open{background:var(--green-glow);color:var(--green);}
.market-status.closed{background:var(--bg3);color:var(--dim);}
.topbar-right{display:flex;gap:6px;}
.btn-icon{background:var(--bg2);border:1px solid var(--border);color:var(--text);padding:7px 12px;font-size:10px;font-family:var(--mono);letter-spacing:1.5px;cursor:pointer;border-radius:4px;text-transform:uppercase;font-weight:700;}
.btn-icon:hover{border-color:var(--green);color:var(--green);}

/* ===== TABS ===== */
.tabs{display:flex;gap:0;border-bottom:1px solid var(--border);margin-bottom:14px;}
.tab{padding:9px 16px;background:transparent;border:none;border-bottom:2px solid transparent;color:var(--dim);font-family:var(--mono);font-size:11px;letter-spacing:2px;font-weight:700;text-transform:uppercase;cursor:pointer;}
.tab.active{color:var(--green);border-bottom-color:var(--green);}
.tab:hover{color:var(--text);}
.tab-count{font-size:9px;background:var(--bg3);padding:1px 6px;border-radius:8px;margin-left:6px;color:var(--dim);}

/* ===== METHOD BANNER ===== */
.method-banner{background:linear-gradient(135deg,rgba(183,148,246,0.08),rgba(82,255,174,0.04));border:1px solid var(--purple);border-radius:5px;padding:8px 14px;margin-bottom:14px;font-size:11px;letter-spacing:1.5px;color:var(--purple);font-weight:700;text-transform:uppercase;}

/* ===== ANALYZE LAYOUT (Bloomberg grid) ===== */
.analyze-grid{display:grid;grid-template-columns:340px 1fr;gap:14px;}
@media(max-width:1024px){.analyze-grid{grid-template-columns:1fr;}}

/* Left column — input */
.panel{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:14px;}
.panel-title{font-size:10px;letter-spacing:2px;color:var(--dim);text-transform:uppercase;font-weight:700;margin-bottom:10px;border-bottom:1px solid var(--border);padding-bottom:7px;}
.input-row{position:relative;}
.symbol-input{width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:13px 14px;font-size:18px;font-family:var(--mono);letter-spacing:3px;font-weight:700;outline:none;border-radius:5px;text-transform:uppercase;}
.symbol-input:focus{border-color:var(--green);}

/* Autocomplete dropdown */
.ac-list{position:absolute;top:100%;left:0;right:0;margin-top:2px;background:var(--bg2);border:1px solid var(--green);border-radius:5px;max-height:300px;overflow-y:auto;z-index:50;}
.ac-item{padding:9px 12px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:10px;border-bottom:1px solid var(--border);}
.ac-item:hover,.ac-item.focused{background:var(--bg3);}
.ac-sym{font-weight:700;font-size:12px;letter-spacing:1.5px;}
.ac-name{font-size:10px;color:var(--dim);flex:1;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.ac-exch{font-size:9px;background:var(--bg4);padding:2px 5px;border-radius:3px;color:var(--dim);letter-spacing:1px;}

.tf-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-top:12px;}
.tf-btn{padding:8px 0;background:var(--bg);border:1px solid var(--border);color:var(--dim);font-family:var(--mono);font-size:10px;letter-spacing:1.5px;cursor:pointer;border-radius:4px;font-weight:700;}
.tf-btn.active{background:var(--green);color:var(--bg);border-color:var(--green);}
.tf-btn:hover:not(.active){border-color:var(--green-dim);color:var(--text);}

.analyze-btn{width:100%;background:var(--green);color:var(--bg);border:none;padding:14px;font-size:12px;font-weight:700;font-family:var(--mono);letter-spacing:3px;text-transform:uppercase;cursor:pointer;margin-top:14px;border-radius:5px;}
.analyze-btn:disabled{opacity:0.4;cursor:not-allowed;}
.analyze-btn:hover:not(:disabled){background:#3fe89a;}

/* ===== RIGHT COLUMN — RESULTS ===== */
.result-row{display:grid;grid-template-columns:1.2fr 1fr;gap:14px;}
@media(max-width:900px){.result-row{grid-template-columns:1fr;}}

.price-panel{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:18px;}
.price-row{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:14px;}
.price-symbol{font-size:18px;font-weight:700;letter-spacing:3px;}
.price-exchange{font-size:9px;color:var(--dim);letter-spacing:2px;background:var(--bg4);padding:3px 7px;border-radius:3px;margin-left:8px;}
.price-current{font-size:32px;font-weight:700;letter-spacing:1px;}
.price-change{font-size:13px;letter-spacing:1px;margin-top:4px;}

.price-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px;padding-top:14px;border-top:1px solid var(--border);}
.price-cell{}
.price-cell-label{font-size:9px;color:var(--dim);letter-spacing:1.5px;text-transform:uppercase;margin-bottom:3px;}
.price-cell-val{font-size:14px;font-weight:700;}

/* Verdict panel */
.verdict-panel{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:18px;}
.verdict-badge{display:inline-block;padding:5px 12px;font-size:11px;font-weight:700;letter-spacing:2px;border-radius:3px;text-transform:uppercase;margin-bottom:14px;}
.verdict-tradable{background:var(--green-glow);color:var(--green);border:1px solid var(--green);}
.verdict-wait{background:var(--yellow-bg);color:var(--yellow);border:1px solid var(--yellow);}
.verdict-setup{font-size:15px;font-weight:700;margin-bottom:10px;}
.verdict-reason{font-size:11px;color:var(--dim);line-height:1.6;}
.verdict-conf{font-size:10px;letter-spacing:1.5px;text-transform:uppercase;margin-top:8px;color:var(--dim);}
.verdict-conf-high{color:var(--green);}
.verdict-conf-medium{color:var(--yellow);}

/* Trade plan */
.plan-panel{background:var(--bg2);border:1px solid var(--green);border-radius:6px;padding:18px;margin-top:14px;}
.plan-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;}
.plan-cell-label{font-size:9px;color:var(--dim);letter-spacing:2px;text-transform:uppercase;margin-bottom:3px;}
.plan-cell-val{font-size:18px;font-weight:700;}
.plan-cell-entry{color:var(--green);}
.plan-cell-sl{color:var(--red);}
.plan-cell-tg{color:var(--blue);}
.plan-rr{font-size:13px;color:var(--purple);}

/* Methodology box */
.method-box{background:var(--bg2);border:1px solid var(--purple);border-radius:6px;padding:14px;margin-top:14px;}
.method-box-title{font-size:10px;letter-spacing:2px;color:var(--purple);text-transform:uppercase;font-weight:700;margin-bottom:8px;}
.method-box-text{font-size:11px;line-height:1.6;color:var(--text);}

/* ===== RULES CHECKLIST TABLE ===== */
.rules-panel{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:14px;margin-top:14px;}
.rules-title{font-size:10px;letter-spacing:2px;color:var(--dim);text-transform:uppercase;font-weight:700;margin-bottom:10px;}
.rules-summary{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;font-size:11px;}
.rules-score{font-size:18px;font-weight:700;color:var(--green);}
.rules-table{width:100%;border-collapse:collapse;font-size:11px;}
.rules-table th{text-align:left;color:var(--dim);font-size:9px;letter-spacing:1.5px;text-transform:uppercase;padding:6px 8px;border-bottom:1px solid var(--border);}
.rules-table td{padding:7px 8px;border-bottom:1px solid var(--border);}
.rule-check-yes{color:var(--green);font-weight:700;font-size:14px;}
.rule-check-no{color:var(--red);font-weight:700;font-size:14px;}
.rule-section-row{background:var(--bg3);font-size:9px;letter-spacing:2px;color:var(--purple);text-transform:uppercase;font-weight:700;}

/* ===== ERROR BOX ===== */
.err{background:var(--red-bg);border:1px solid var(--red);color:var(--red);padding:11px 14px;border-radius:5px;font-size:11px;margin-top:12px;}

/* ===== SCANNER ===== */
.scanner-grid{display:grid;gap:10px;}
.scanner-card{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:14px;display:grid;grid-template-columns:140px 1fr 160px 120px;gap:14px;align-items:center;}
@media(max-width:768px){.scanner-card{grid-template-columns:1fr;}}
.scanner-sym{font-size:14px;font-weight:700;letter-spacing:2px;}
.scanner-setup{font-size:11px;color:var(--text);}
.scanner-meta{font-size:10px;color:var(--dim);}
.scanner-score{font-size:18px;font-weight:700;color:var(--green);text-align:center;}
.scanner-actions{display:flex;gap:6px;}

/* ===== JOURNAL & PORTFOLIO ===== */
.table{width:100%;border-collapse:collapse;font-size:11px;}
.table th{text-align:left;color:var(--dim);font-size:9px;letter-spacing:1.5px;text-transform:uppercase;padding:9px 10px;background:var(--bg3);border-bottom:1px solid var(--border);}
.table td{padding:10px;border-bottom:1px solid var(--border);}
.table tr:hover td{background:var(--bg3);}

.empty{text-align:center;color:var(--dim);padding:60px 20px;font-size:12px;}
.empty-big{font-size:36px;color:var(--dimmer);margin-bottom:14px;}

/* Sub-buttons */
.mini-btn{background:var(--bg3);border:1px solid var(--border);color:var(--text);padding:6px 11px;font-size:10px;font-family:var(--mono);letter-spacing:1px;cursor:pointer;border-radius:3px;text-transform:uppercase;font-weight:700;}
.mini-btn:hover{border-color:var(--green);color:var(--green);}
.mini-btn-danger:hover{border-color:var(--red);color:var(--red);}
.mini-btn-primary{background:var(--green);color:var(--bg);border-color:var(--green);}
.mini-btn-primary:hover{background:#3fe89a;color:var(--bg);}

/* alert badge */
.alert-row{display:flex;justify-content:space-between;align-items:center;padding:10px 14px;background:var(--bg2);border:1px solid var(--border);border-radius:5px;margin-bottom:6px;font-size:11px;}
.alert-active{border-left:3px solid var(--green);}
.alert-triggered{border-left:3px solid var(--red);}

/* Scan trigger button */
.scan-trigger{display:flex;gap:10px;align-items:center;margin-bottom:14px;}
`;

// =====================================================================
// HELPERS
// =====================================================================
async function apiCall(path, opts = {}) {
  const r = await fetch(API + path, {
    ...opts,
    headers: { "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!r.ok) {
    const err = await r.text();
    throw new Error(`API ${r.status}: ${err.slice(0, 200)}`);
  }
  return r.json();
}

function fmt(n, d = 2) {
  if (n == null || isNaN(n)) return "—";
  return Number(n).toFixed(d);
}
function fmtPct(n) {
  if (n == null || isNaN(n)) return "—";
  const sign = n >= 0 ? "+" : "";
  return `${sign}${Number(n).toFixed(2)}%`;
}

// =====================================================================
// MAIN APP
// =====================================================================
export default function App() {
  const [theme, setTheme] = useState(() => localStorage.getItem("td:theme") || "dark");
  const [tab, setTab] = useState("analyze");
  const [marketStatus, setMarketStatus] = useState("");

  // Ticker strip data (Nifty + Sensex + popular stocks)
  const [tickers, setTickers] = useState([]);

  // Analyze tab
  const [sym, setSym] = useState("");
  const [tf, setTf] = useState("1D");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [err, setErr] = useState("");
  const abortCtrl = useRef(null);

  // Autocomplete
  const [acItems, setAcItems] = useState([]);
  const [acShow, setAcShow] = useState(false);
  const [acFocused, setAcFocused] = useState(-1);
  const acTimer = useRef(null);

  // Journal / Portfolio / Alerts
  const [journal, setJournal] = useState([]);
  const [portfolio, setPortfolio] = useState({ cash: 100000, positions: [], totalValue: 100000, pnl: 0 });
  const [alerts, setAlerts] = useState([]);

  // Scanner
  const [scanResults, setScanResults] = useState(null);
  const [scanRunning, setScanRunning] = useState(false);

  // ============ EFFECTS ============
  useEffect(() => {
    document.body.className = theme;
    localStorage.setItem("td:theme", theme);
  }, [theme]);

  useEffect(() => {
    loadData();
    loadLastScan();
    loadTickers();
    const tickInt = setInterval(loadTickers, 60_000); // refresh every 60s
    return () => clearInterval(tickInt);
  }, []);

  async function loadData() {
    try { setJournal(await apiCall("/journal")); } catch (e) { console.warn(e); }
    try { setPortfolio(await apiCall("/portfolio")); } catch (e) { console.warn(e); }
    try { setAlerts(await apiCall("/alerts")); } catch (e) { console.warn(e); }
  }

  async function loadTickers() {
    const symbols = ["NIFTY", "SENSEX", "BANKNIFTY", "RELIANCE", "TCS", "HDFCBANK", "INFY", "ICICIBANK"];
    const out = [];
    for (const s of symbols) {
      try {
        const d = await fetch(`${API}/price/${s}`).then(r => r.json());
        if (d?.current_price) out.push({ symbol: s, price: d.current_price, chg: d.change_pct || 0 });
        if (s === "NIFTY") setMarketStatus(d.market_status || "");
      } catch {}
    }
    setTickers(out);
  }

  async function loadLastScan() {
    try {
      const d = await apiCall("/scan/last");
      setScanResults(d);
    } catch (e) { /* nothing yet */ }
  }

  // ============ AUTOCOMPLETE ============
  function onSymChange(v) {
    setSym(v.toUpperCase());
    setAcFocused(-1);
    if (acTimer.current) clearTimeout(acTimer.current);
    if (!v.trim()) { setAcItems([]); setAcShow(false); return; }
    acTimer.current = setTimeout(async () => {
      try {
        const items = await apiCall(`/search?q=${encodeURIComponent(v)}`);
        setAcItems(items);
        setAcShow(items.length > 0);
      } catch (e) { console.warn(e); }
    }, 200);
  }

  function pickAc(it) {
    setSym(it.symbol);
    setAcItems([]); setAcShow(false); setAcFocused(-1);
  }

  function onSymKey(e) {
    if (!acShow) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setAcFocused(i => Math.min(i + 1, acItems.length - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setAcFocused(i => Math.max(i - 1, 0)); }
    else if (e.key === "Enter" && acFocused >= 0) { e.preventDefault(); pickAc(acItems[acFocused]); }
    else if (e.key === "Escape") { setAcShow(false); }
  }

  // ============ ANALYZE ============
  async function analyze() {
    if (!sym.trim()) return;
    setErr(""); setResult(null); setLoading(true); setAcShow(false);
    if (abortCtrl.current) abortCtrl.current.abort();
    abortCtrl.current = new AbortController();
    try {
      const d = await fetch(API + "/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal: abortCtrl.current.signal,
        body: JSON.stringify({ symbol: sym.trim().toUpperCase(), tf }),
      }).then(r => r.ok ? r.json() : r.text().then(t => { throw new Error(t); }));
      setResult(d);
    } catch (e) {
      if (e.name !== "AbortError") setErr(e.message);
    } finally { setLoading(false); }
  }

  async function saveToJournal() {
    if (!result) return;
    try {
      await apiCall("/journal", { method: "POST", body: JSON.stringify(result) });
      setJournal(await apiCall("/journal"));
      alert("Saved to journal");
    } catch (e) { alert("Save failed: " + e.message); }
  }

  async function setAlert(price, condition) {
    if (!result) return;
    try {
      await apiCall("/alerts", {
        method: "POST",
        body: JSON.stringify({ symbol: result.symbol, price, condition, setup: result.setup }),
      });
      setAlerts(await apiCall("/alerts"));
      alert(`Alert set: ${result.symbol} ${condition} ₹${price}`);
    } catch (e) { alert("Alert failed: " + e.message); }
  }

  // ============ SCANNER ============
  async function runScan() {
    if (!confirm("Run a full Nifty 500 scan? Takes ~3 minutes.")) return;
    setScanRunning(true);
    try {
      const d = await apiCall("/scan/run", { method: "POST", body: JSON.stringify({ notify: true }) });
      setScanResults(d);
      alert(`Scan done. Found ${d.findings.length} setups. Telegram sent.`);
    } catch (e) { alert("Scan failed: " + e.message); }
    finally { setScanRunning(false); }
  }

  // ============ RENDER ============
  return (
    <>
      <style>{css}</style>
      <div className="app">
        {/* TICKER STRIP */}
        {tickers.length > 0 && (
          <div className="ticker-strip">
            {tickers.map(t => (
              <div className="ticker-item" key={t.symbol}>
                <span className="ticker-sym">{t.symbol}</span>
                <span className="ticker-price">₹{fmt(t.price)}</span>
                <span className={`ticker-chg ${t.chg >= 0 ? "up" : "down"}`}>{fmtPct(t.chg)}</span>
              </div>
            ))}
          </div>
        )}

        {/* TOPBAR */}
        <div className="topbar">
          <div className="brand">
            <span className="brand-name">▸ Trading Desk</span>
            <span className="brand-tag">PAATHSHAALA METHOD</span>
            <span className={`market-status ${marketStatus === "live" ? "open" : "closed"}`}>
              ● {marketStatus === "live" ? "Live" : "Closed"}
            </span>
          </div>
          <div className="topbar-right">
            <button className="btn-icon" onClick={() => setTheme(theme === "dark" ? "light" : "dark")}>
              {theme === "dark" ? "☀ Light" : "🌙 Dark"}
            </button>
          </div>
        </div>

        {/* TABS */}
        <div className="tabs">
          <button className={`tab ${tab === "analyze" ? "active" : ""}`} onClick={() => setTab("analyze")}>Analyze</button>
          <button className={`tab ${tab === "scanner" ? "active" : ""}`} onClick={() => setTab("scanner")}>
            Scanner {scanResults?.findings?.length > 0 && <span className="tab-count">{scanResults.findings.length}</span>}
          </button>
          <button className={`tab ${tab === "alerts" ? "active" : ""}`} onClick={() => setTab("alerts")}>
            Alerts {alerts.length > 0 && <span className="tab-count">{alerts.length}</span>}
          </button>
          <button className={`tab ${tab === "journal" ? "active" : ""}`} onClick={() => setTab("journal")}>
            Journal {journal.length > 0 && <span className="tab-count">{journal.length}</span>}
          </button>
          <button className={`tab ${tab === "portfolio" ? "active" : ""}`} onClick={() => setTab("portfolio")}>Portfolio</button>
          <button className={`tab ${tab === "method" ? "active" : ""}`} onClick={() => setTab("method")}>Method</button>
        </div>

        {/* ========= ANALYZE TAB ========= */}
        {tab === "analyze" && (
          <div className="analyze-grid">
            {/* LEFT — Input */}
            <div>
              <div className="method-banner">✓ Paathshaala Method Active (CA Rahul Ranka)</div>

              <div className="panel">
                <div className="panel-title">Symbol</div>
                <div className="input-row">
                  <input
                    className="symbol-input"
                    value={sym}
                    onChange={e => onSymChange(e.target.value)}
                    onKeyDown={onSymKey}
                    onFocus={() => acItems.length > 0 && setAcShow(true)}
                    onBlur={() => setTimeout(() => setAcShow(false), 200)}
                    placeholder="RELIANCE"
                  />
                  {acShow && acItems.length > 0 && (
                    <div className="ac-list">
                      {acItems.map((it, i) => (
                        <div key={i} className={`ac-item ${i === acFocused ? "focused" : ""}`} onMouseDown={() => pickAc(it)}>
                          <span className="ac-sym">{it.symbol}</span>
                          <span className="ac-name">{it.name}</span>
                          <span className="ac-exch">{it.exchange}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <div className="panel-title" style={{ marginTop: 16 }}>Timeframe</div>
                <div className="tf-grid">
                  {["1m", "5m", "15m", "30m", "1h", "4h", "1D", "1W"].map(t => (
                    <button key={t} className={`tf-btn ${tf === t ? "active" : ""}`} onClick={() => setTf(t)}>{t}</button>
                  ))}
                </div>

                <button className="analyze-btn" onClick={analyze} disabled={loading}>
                  {loading ? "Analyzing…" : "Analyze ▸"}
                </button>

                {err && <div className="err">⚠ {err}</div>}
              </div>
            </div>

            {/* RIGHT — Results */}
            <div>
              {!result && !loading && (
                <div className="panel">
                  <div className="empty">
                    <div className="empty-big">▸</div>
                    Enter a symbol on the left and click Analyze<br />
                    to run Paathshaala methodology check.
                  </div>
                </div>
              )}

              {loading && (
                <div className="panel">
                  <div className="empty">
                    <div className="empty-big">⟳</div>
                    Running Paathshaala analysis…<br />
                    <span style={{ fontSize: 10, color: "var(--dim)" }}>Fetching live data, computing rules, calling Gemini</span>
                  </div>
                </div>
              )}

              {result && (
                <>
                  <div className="result-row">
                    {/* Price card */}
                    <div className="price-panel">
                      <div className="price-row">
                        <div>
                          <span className="price-symbol">{result.symbol}</span>
                          <span className="price-exchange">{result.exchange || "NSE"}</span>
                        </div>
                      </div>
                      <div className="price-current">₹{fmt(result.current_price)}</div>
                      <div className={`price-change ${(result.change_pct || 0) >= 0 ? "up" : "down"}`}>
                        {fmtPct(result.change_pct)} ({result.change >= 0 ? "+" : ""}₹{fmt(result.change)})
                      </div>
                      <div className="price-grid">
                        <div className="price-cell"><div className="price-cell-label">Open</div><div className="price-cell-val">₹{fmt(result.day_open)}</div></div>
                        <div className="price-cell"><div className="price-cell-label">High</div><div className="price-cell-val">₹{fmt(result.day_high)}</div></div>
                        <div className="price-cell"><div className="price-cell-label">Low</div><div className="price-cell-val">₹{fmt(result.day_low)}</div></div>
                        <div className="price-cell"><div className="price-cell-label">Prev Close</div><div className="price-cell-val">₹{fmt(result.prev_close)}</div></div>
                      </div>
                    </div>

                    {/* Verdict card */}
                    <div className="verdict-panel">
                      <span className={`verdict-badge ${result.tradable ? "verdict-tradable" : "verdict-wait"}`}>
                        {result.tradable ? "▸ Tradable" : "⏸ Wait"}
                      </span>
                      <div className="verdict-setup">{result.setup || "—"} <span style={{ color: "var(--dim)", fontSize: 11 }}>({result.trend})</span></div>
                      <div className="verdict-reason">{result.reason}</div>
                      <div className={`verdict-conf verdict-conf-${(result.confidence || "low").toLowerCase()}`}>
                        Confidence: {result.confidence}
                      </div>
                    </div>
                  </div>

                  {/* Trade plan */}
                  {result.tradable && result.entry && (
                    <div className="plan-panel">
                      <div className="plan-grid">
                        <div><div className="plan-cell-label">Entry</div><div className="plan-cell-val plan-cell-entry">₹{fmt(result.entry)}</div></div>
                        <div><div className="plan-cell-label">Stop Loss</div><div className="plan-cell-val plan-cell-sl">₹{fmt(result.sl)}</div></div>
                        <div><div className="plan-cell-label">Target</div><div className="plan-cell-val plan-cell-tg">₹{fmt(result.target)}</div></div>
                        <div><div className="plan-cell-label">R:R</div><div className="plan-cell-val plan-rr">{result.risk_reward}</div></div>
                      </div>
                      <div style={{ marginTop: 14, display: "flex", gap: 8 }}>
                        <button className="mini-btn mini-btn-primary" onClick={saveToJournal}>💾 Save Trade</button>
                        <button className="mini-btn" onClick={() => setAlert(result.entry, "below")}>🔔 Alert at Entry</button>
                      </div>
                    </div>
                  )}

                  {!result.tradable && result.alert_price && (
                    <div className="plan-panel" style={{ borderColor: "var(--yellow)" }}>
                      <div className="plan-grid" style={{ gridTemplateColumns: "1fr 1fr" }}>
                        <div><div className="plan-cell-label">Watch Level</div><div className="plan-cell-val" style={{ color: "var(--yellow)" }}>₹{fmt(result.alert_price)}</div></div>
                        <div><div className="plan-cell-label">Wait For</div><div className="plan-cell-val" style={{ fontSize: 11, fontWeight: 400, color: "var(--text)" }}>{result.when_to_enter}</div></div>
                      </div>
                      <div style={{ marginTop: 14 }}>
                        <button className="mini-btn mini-btn-primary" onClick={() => setAlert(result.alert_price, result.alert_condition || "above")}>
                          🔔 Set Watch Alert
                        </button>
                      </div>
                    </div>
                  )}

                  {/* Methodology details */}
                  {result.methodology_applied && (
                    <div className="method-box">
                      <div className="method-box-title">▸ Methodology Applied</div>
                      <div className="method-box-text">{result.methodology_applied}</div>
                    </div>
                  )}

                  {/* RULES CHECKLIST TABLE */}
                  {result.rules_checklist && result.rules_checklist.length > 0 && (
                    <div className="rules-panel">
                      <div className="rules-summary">
                        <div className="rules-title">Paathshaala Rules Checklist</div>
                        <div className="rules-score">{result.rule_score || 0}%</div>
                      </div>
                      <table className="rules-table">
                        <thead><tr><th style={{ width: 30 }}></th><th>Rule</th><th>Detail</th></tr></thead>
                        <tbody>
                          {(() => {
                            const sections = {};
                            result.rules_checklist.forEach(r => {
                              if (!sections[r.section]) sections[r.section] = [];
                              sections[r.section].push(r);
                            });
                            const out = [];
                            for (const [sec, rules] of Object.entries(sections)) {
                              out.push(<tr key={sec + "_h"} className="rule-section-row"><td colSpan={3}>{sec}</td></tr>);
                              rules.forEach((r, idx) => {
                                out.push(
                                  <tr key={sec + "_" + idx}>
                                    <td className={r.passed ? "rule-check-yes" : "rule-check-no"}>{r.passed ? "✓" : "✗"}</td>
                                    <td>{r.rule}</td>
                                    <td style={{ color: "var(--dim)", fontSize: 10 }}>{r.detail}</td>
                                  </tr>
                                );
                              });
                            }
                            return out;
                          })()}
                        </tbody>
                      </table>
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        )}

        {/* ========= SCANNER TAB ========= */}
        {tab === "scanner" && (
          <div>
            <div className="scan-trigger">
              <button className="btn-icon" onClick={runScan} disabled={scanRunning} style={{ background: "var(--green)", color: "var(--bg)", borderColor: "var(--green)" }}>
                {scanRunning ? "Scanning…" : "▸ Run Scan Now"}
              </button>
              <span style={{ fontSize: 11, color: "var(--dim)" }}>
                Auto-runs daily at 9:30 AM IST. Last run: {scanResults?.runAt ? new Date(scanResults.runAt).toLocaleString("en-IN") : "Never"}
              </span>
            </div>

            {!scanResults?.findings?.length && (
              <div className="empty">
                <div className="empty-big">⊟</div>
                No scan results yet. Click "Run Scan Now" to check Nifty 500 for setups.<br />
                <span style={{ fontSize: 10 }}>The scan runs automatically at 9:30 AM IST every trading day.</span>
              </div>
            )}

            <div className="scanner-grid">
              {(scanResults?.findings || []).map((f, i) => (
                <div className="scanner-card" key={i}>
                  <div>
                    <div className="scanner-sym">{f.symbol}</div>
                    <div className="scanner-meta">₹{fmt(f.price)}</div>
                  </div>
                  <div>
                    <div className="scanner-setup">{f.setup}</div>
                    <div className="scanner-meta">{f.reason}</div>
                  </div>
                  <div className="scanner-meta">
                    RSI: {fmt(f.rsi, 1)}<br/>
                    Support: ₹{fmt(f.support)}
                  </div>
                  <div>
                    <div className="scanner-score">{f.score}%</div>
                    <button className="mini-btn" style={{ width: "100%", marginTop: 6 }} onClick={() => { setSym(f.symbol); setTf("1D"); setTab("analyze"); setTimeout(analyze, 100); }}>
                      Analyze ▸
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ========= ALERTS TAB ========= */}
        {tab === "alerts" && (
          <div>
            {alerts.length === 0 ? (
              <div className="empty">
                <div className="empty-big">🔔</div>
                No active alerts. Analyze a stock and click "Alert at Entry" or "Set Watch Alert" to create one.
              </div>
            ) : (
              alerts.map(a => (
                <div key={a.id} className={`alert-row ${a.triggered ? "alert-triggered" : "alert-active"}`}>
                  <div>
                    <strong style={{ letterSpacing: 1.5 }}>{a.symbol}</strong>
                    <span style={{ color: "var(--dim)", marginLeft: 10 }}>{a.setup || "—"}</span>
                  </div>
                  <div>
                    <span style={{ color: "var(--yellow)" }}>{a.condition} ₹{fmt(a.price)}</span>
                    {a.triggered && <span style={{ color: "var(--red)", marginLeft: 10, fontSize: 10, letterSpacing: 2 }}>TRIGGERED</span>}
                  </div>
                  <button className="mini-btn mini-btn-danger" onClick={async () => {
                    await apiCall(`/alerts/${a.id}`, { method: "DELETE" });
                    setAlerts(await apiCall("/alerts"));
                  }}>Delete</button>
                </div>
              ))
            )}
          </div>
        )}

        {/* ========= JOURNAL TAB ========= */}
        {tab === "journal" && (
          <div className="panel">
            {journal.length === 0 ? (
              <div className="empty"><div className="empty-big">📔</div>Journal empty. Save analyses from the Analyze tab.</div>
            ) : (
              <>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 14 }}>
                  <div className="panel-title">Journal ({journal.length})</div>
                  <button className="mini-btn" onClick={() => {
                    const ws = XLSX.utils.json_to_sheet(journal);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Journal");
                    XLSX.writeFile(wb, `tradingdesk-journal-${new Date().toISOString().slice(0,10)}.xlsx`);
                  }}>⬇ Export Excel</button>
                </div>
                <table className="table">
                  <thead><tr>
                    <th>Date</th><th>Symbol</th><th>TF</th><th>Setup</th>
                    <th>Entry</th><th>SL</th><th>Target</th><th>R:R</th><th>Conf</th>
                  </tr></thead>
                  <tbody>
                    {journal.map(j => (
                      <tr key={j.id}>
                        <td>{new Date(j.created_at || j.createdAt).toLocaleDateString("en-IN")}</td>
                        <td><strong>{j.symbol}</strong></td>
                        <td>{j.tf}</td>
                        <td>{j.setup}</td>
                        <td>₹{fmt(j.entry)}</td>
                        <td>₹{fmt(j.sl)}</td>
                        <td>₹{fmt(j.target)}</td>
                        <td>{j.risk_reward}</td>
                        <td>{j.confidence}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>
        )}

        {/* ========= PORTFOLIO TAB ========= */}
        {tab === "portfolio" && (
          <div className="panel">
            <div className="panel-title">Portfolio (Mock ₹1L)</div>
            <div className="price-grid" style={{ borderTop: "none", paddingTop: 0 }}>
              <div className="price-cell"><div className="price-cell-label">Cash</div><div className="price-cell-val">₹{fmt(portfolio.cash, 0)}</div></div>
              <div className="price-cell"><div className="price-cell-label">Total Value</div><div className="price-cell-val">₹{fmt(portfolio.totalValue, 0)}</div></div>
              <div className="price-cell"><div className="price-cell-label">P&L</div><div className={`price-cell-val ${portfolio.pnl >= 0 ? "up" : "down"}`}>₹{fmt(portfolio.pnl, 0)}</div></div>
              <div className="price-cell"><div className="price-cell-label">Positions</div><div className="price-cell-val">{portfolio.positions.length}</div></div>
            </div>
            {portfolio.positions.length > 0 && (
              <table className="table" style={{ marginTop: 18 }}>
                <thead><tr><th>Symbol</th><th>Qty</th><th>Entry</th><th>LTP</th><th>P&L</th></tr></thead>
                <tbody>
                  {portfolio.positions.map(p => (
                    <tr key={p.symbol}>
                      <td><strong>{p.symbol}</strong></td><td>{p.qty}</td>
                      <td>₹{fmt(p.entry)}</td><td>₹{fmt(p.ltp)}</td>
                      <td className={p.pnl >= 0 ? "up" : "down"}>₹{fmt(p.pnl, 0)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* ========= METHOD TAB ========= */}
        {tab === "method" && (
          <div className="panel" style={{ padding: 26 }}>
            <div className="panel-title">Paathshaala Methodology (CA Rahul Ranka)</div>
            <div style={{ fontSize: 12, lineHeight: 1.8, color: "var(--text)" }}>
              <p><strong style={{ color: "var(--green)" }}>Top 3 Setups</strong></p>
              <ol>
                <li><strong>Reversal from Support</strong> — stock falls 5%+, finds tested support (2+ touches), shows bullish reversal candle (engulfing/hammer/morning star) with volume confirmation.</li>
                <li><strong>Breakout (Jam ke kaam)</strong> — price breaks above 20-bar high in a rising-SMA regime; both SMA20 and SMA50 sloping up.</li>
                <li><strong>Pullback to Moving Average</strong> — in uptrend, price pulls back to SMA20 or SMA50 and bounces.</li>
              </ol>
              <p><strong style={{ color: "var(--green)" }}>Final Reversal Checklist</strong></p>
              <ul>
                <li>Proper fall before reversal attempt (≥5%)</li>
                <li>Strong support level identified (2+ touches)</li>
                <li>Bullish reversal candle present</li>
                <li>Volume above 20-bar average</li>
                <li>RSI reversing up from oversold (&lt;45 zone)</li>
                <li>RSI not overbought (≤82)</li>
                <li>R:R ≥ 1:2</li>
              </ul>
              <p><strong style={{ color: "var(--green)" }}>Regimes</strong></p>
              <ul>
                <li><strong>Jam ke kaam</strong> (trending up) — both SMAs rising → momentum/breakout trades OK</li>
                <li><strong>Tham ke kaam</strong> (range/flat) — only reversal trades from support</li>
                <li><strong>Falling</strong> — both SMAs falling → avoid longs entirely</li>
              </ul>
            </div>
          </div>
        )}
      </div>
    </>
  );
}
