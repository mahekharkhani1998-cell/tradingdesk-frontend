import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import Login from "./Login.jsx";

const API = import.meta.env.VITE_API_BASE || "/api";
const TFS = ["1m","3m","5m","15m","30m","1h","4h","1D","1W"];
const TABS = ["Analyze","Scanner","Journal","Portfolio","Method"];

const NSE_STOCKS = [
  ["RELIANCE","Reliance"],["TCS","TCS"],["HDFCBANK","HDFC Bank"],["BHARTIARTL","Bharti Airtel"],["ICICIBANK","ICICI Bank"],
  ["INFY","Infosys"],["SBIN","SBI"],["HINDUNILVR","HUL"],["ITC","ITC"],["LICI","LIC"],
  ["BAJFINANCE","Bajaj Finance"],["LT","L&T"],["HCLTECH","HCL Tech"],["KOTAKBANK","Kotak Bank"],["MARUTI","Maruti Suzuki"],
  ["SUNPHARMA","Sun Pharma"],["AXISBANK","Axis Bank"],["M&M","Mahindra"],["ULTRACEMCO","UltraTech"],["ASIANPAINT","Asian Paints"],
  ["NTPC","NTPC"],["BAJAJFINSV","Bajaj Finserv"],["ADANIPORTS","Adani Ports"],["ADANIENT","Adani Enterprises"],["TITAN","Titan"],
  ["ONGC","ONGC"],["NESTLEIND","Nestle"],["JSWSTEEL","JSW Steel"],["POWERGRID","Power Grid"],["TATAMOTORS","Tata Motors"],
  ["COALINDIA","Coal India"],["SBILIFE","SBI Life"],["GRASIM","Grasim"],["BAJAJ-AUTO","Bajaj Auto"],["WIPRO","Wipro"],
  ["HDFCLIFE","HDFC Life"],["TATASTEEL","Tata Steel"],["BEL","Bharat Electronics"],["TECHM","Tech Mahindra"],["EICHERMOT","Eicher"],
  ["DRREDDY","Dr Reddy's"],["HINDALCO","Hindalco"],["CIPLA","Cipla"],["BPCL","BPCL"],["HEROMOTOCO","Hero MotoCorp"],
  ["BRITANNIA","Britannia"],["APOLLOHOSP","Apollo Hospitals"],["INDUSINDBK","IndusInd"],["SHRIRAMFIN","Shriram Finance"],
  ["TATACONSUM","Tata Consumer"],["TRENT","Trent"],["DMART","DMart"],["DLF","DLF"],["GAIL","GAIL"],["IRCTC","IRCTC"],
  ["VEDL","Vedanta"],["ZOMATO","Zomato"],["VBL","Varun Beverages"],["TATAPOWER","Tata Power"],["NHPC","NHPC"],
  ["NMDC","NMDC"],["PFC","PFC"],["RECLTD","REC"],["IOC","IOC"],["INDIGO","IndiGo"],["HAL","HAL"],["BOSCHLTD","Bosch"],
  ["HAVELLS","Havells"],["ICICIGI","ICICI Lombard"],["ICICIPRULI","ICICI Prudential"],["SBICARD","SBI Cards"],
  ["DABUR","Dabur"],["COLPAL","Colgate"],["GODREJCP","Godrej Consumer"],["MARICO","Marico"],["PIDILITIND","Pidilite"],
  ["AUBANK","AU SFB"],["FEDERALBNK","Federal Bank"],["BANDHANBNK","Bandhan"],["IDFCFIRSTB","IDFC First"],["YESBANK","Yes Bank"],
  ["PNB","PNB"],["BANKBARODA","BoB"],["CANBK","Canara Bank"],["JINDALSTEL","Jindal Steel"],["SAIL","SAIL"],
  ["AUROPHARMA","Aurobindo"],["LUPIN","Lupin"],["BIOCON","Biocon"],["DIVISLAB","Divi's Labs"],["ALKEM","Alkem"],
  ["TORNTPHARM","Torrent Pharma"],["TVSMOTOR","TVS Motor"],["ASHOKLEY","Ashok Leyland"],["MOTHERSON","Motherson"],
  ["LTIM","LTIMindtree"],["PERSISTENT","Persistent"],["COFORGE","Coforge"],["MPHASIS","Mphasis"],["KPITTECH","KPIT"],
  ["DELHIVERY","Delhivery"],["NYKAA","Nykaa"],["PAYTM","Paytm"],["POLYCAB","Polycab"],["DIXON","Dixon Tech"],
  ["UPL","UPL"],["MUTHOOTFIN","Muthoot Finance"],["CHOLAFIN","Cholamandalam"],["LICHSGFIN","LIC Housing"],
  ["BHEL","BHEL"],["SIEMENS","Siemens"],["ABB","ABB India"],["CGPOWER","CG Power"],["IRFC","IRFC"],["RVNL","RVNL"],
  ["SHREECEM","Shree Cement"],["AMBUJACEM","Ambuja"],["ACC","ACC"],["JKCEMENT","JK Cement"],["DALBHARAT","Dalmia Bharat"],
  ["INDHOTEL","Indian Hotels"],["OBEROIRLTY","Oberoi Realty"],["GODREJPROP","Godrej Properties"],["LODHA","Macrotech"],
  ["TATAELXSI","Tata Elxsi"],["TATACOMM","Tata Comm"],["DATAMATICS","Datamatics"],["NEWGEN","Newgen Software"],
  ["KEC","KEC International"],["VOLTAS","Voltas"],["WHIRLPOOL","Whirlpool"],["BLUESTARCO","Blue Star"],
  ["NIFTY","Nifty 50"],["BANKNIFTY","Bank Nifty"],["FINNIFTY","Fin Nifty"],["SENSEX","Sensex"]
];

const SECTORS = {
  "Nifty Top 10": ["RELIANCE","TCS","HDFCBANK","BHARTIARTL","ICICIBANK","INFY","SBIN","HINDUNILVR","ITC","LICI"],
  "Banking": ["HDFCBANK","ICICIBANK","SBIN","KOTAKBANK","AXISBANK","INDUSINDBK","BANKBARODA","FEDERALBNK","PNB","AUBANK"],
  "IT": ["TCS","INFY","HCLTECH","WIPRO","TECHM","LTIM","COFORGE","PERSISTENT","MPHASIS","KPITTECH"],
  "Auto": ["MARUTI","M&M","TATAMOTORS","BAJAJ-AUTO","HEROMOTOCO","EICHERMOT","TVSMOTOR","ASHOKLEY","MOTHERSON","BOSCHLTD"],
  "Pharma": ["SUNPHARMA","DRREDDY","CIPLA","DIVISLAB","APOLLOHOSP","TORNTPHARM","LUPIN","AUROPHARMA","BIOCON","ALKEM"],
  "FMCG": ["HINDUNILVR","ITC","NESTLEIND","BRITANNIA","DABUR","MARICO","TATACONSUM","COLPAL","GODREJCP","VBL"],
  "Metals": ["JSWSTEEL","TATASTEEL","HINDALCO","JINDALSTEL","VEDL","NMDC","SAIL"]
};

const PAATHSHAALA_INFO = `═══════════════════════════════════
PAATHSHAALA METHODOLOGY (CA Rahul Ranka)
═══════════════════════════════════

▸ TOP 3 RANKED SETUPS
1. Reversal Candle from Support + RSI reversing from oversold 40 + High volume
2. Unexpected Move
3. Breakout on Daily TF + all SMAs (20/50/200) rising

▸ FINAL REVERSAL BUYING CHECKLIST
• Proper Fall — stock visibly declined
• 3rd touch of support (clear U or V swings)
• Bullish Engulfing / Hammer / Morning Star
• Volume higher than previous
• Daily: Nifty 100, Weekly: Nifty 200, Monthly: Nifty 500
• Nifty 500 must be above 20 & 50 SMA both

▸ SMA REGIME
• Both 20 & 50 rising → "Jam ke kaam karna hai"
• Either flat/falling → "Tham ke kaam karna hai"
• Both falling → Avoid OR max 30% quantity

▸ RSI (Paathshaala)
• > 65 = strong bullish
• < 40 = bearish
• > 82-85 = book 30-50%, no fresh entries
• Buy zone: reversing UP from oversold (~40)

▸ TARGETS & SL
• Target: nearest UNBROKEN swing high
• SL: most recent swing low (never shift down)
• Min R:R = 1:2

▸ TRAILING EXIT (SMA based)
• Closes below 20 SMA + bearish candle → exit 50%
• Closes below 50 SMA + bearish candle → exit 100%`;

const css = `
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body.dark{
  --bg:#080a0f;--bg2:#0d1017;--bg3:#121620;--border:#1c2030;--border2:#252b3d;
  --text:#d8ddf0;--dim:#4a5068;--dim2:#6b7590;
  --green:#3dffa0;--red:#ff4d6a;--blue:#38c8ff;--yellow:#ffcc44;--purple:#b794ff;
  --green-bg:#0a1f15;--red-bg:#1f0a0d;--blue-bg:#091520;--yellow-bg:#1a1400;--purple-bg:#15101f;
}
body.light{
  --bg:#fafaf7;--bg2:#ffffff;--bg3:#f3f3ee;--border:#e3e3dc;--border2:#cfcfc4;
  --text:#1a1c24;--dim:#9aa0a8;--dim2:#6a6e78;
  --green:#0d8f4f;--red:#d63846;--blue:#1e7fb8;--yellow:#b48400;--purple:#7a52d4;
  --green-bg:#e8f5ed;--red-bg:#fae5e7;--blue-bg:#e5f1f9;--yellow-bg:#fbf3dc;--purple-bg:#f0eafb;
}
body{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;}
.app{min-height:100vh;max-width:720px;margin:0 auto;padding:22px 18px 40px;}
.topbar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:22px;}
.brand-block{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap;}
.brand{font-family:'Space Mono',monospace;font-size:11px;letter-spacing:3px;color:var(--green);text-transform:uppercase;font-weight:700;}
.theme-toggle{background:var(--bg2);border:1px solid var(--border);color:var(--text);padding:7px 12px;font-size:11px;font-family:'Space Mono',monospace;border-radius:6px;cursor:pointer;letter-spacing:0.5px;}
.tabs{display:flex;gap:2px;background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:3px;margin-bottom:20px;overflow-x:auto;}
.tab{flex:1;min-width:80px;padding:8px 10px;font-size:11px;font-weight:500;color:var(--dim2);background:none;border:none;cursor:pointer;border-radius:6px;letter-spacing:0.5px;font-family:'DM Sans',sans-serif;white-space:nowrap;}
.tab.on{background:var(--bg3);color:var(--text);border:1px solid var(--border2);font-weight:600;}
.method-badge{background:var(--purple-bg);border:1px solid var(--purple);border-radius:6px;padding:8px 12px;margin-bottom:14px;font-size:11px;color:var(--purple);font-family:'Space Mono',monospace;letter-spacing:0.5px;font-weight:700;}
.sym-wrap{position:relative;margin-bottom:10px;}
.sym-in{width:100%;background:var(--bg2);border:1px solid var(--border);color:var(--text);padding:14px 18px;font-size:18px;font-family:'Space Mono',monospace;font-weight:700;letter-spacing:3px;text-transform:uppercase;outline:none;border-radius:8px;}
.sym-in:focus{border-color:var(--green);}
.sym-in::placeholder{color:var(--dim);font-size:12px;letter-spacing:2px;font-weight:400;}
.sym-dropdown{position:absolute;top:calc(100% + 4px);left:0;right:0;background:var(--bg2);border:1px solid var(--border2);border-radius:8px;max-height:280px;overflow-y:auto;z-index:50;box-shadow:0 10px 30px rgba(0,0,0,0.3);}
.sym-opt{padding:10px 16px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:10px;border-bottom:1px solid var(--border);}
.sym-opt:last-child{border-bottom:none;}
.sym-opt:hover,.sym-opt.kb-active{background:var(--bg3);}
.sym-opt-sym{font-family:'Space Mono',monospace;font-size:13px;font-weight:700;color:var(--green);letter-spacing:1px;}
.sym-opt-name{font-size:11px;color:var(--dim2);text-align:right;}
.tf-label{font-size:9px;letter-spacing:3px;color:var(--dim);font-family:'Space Mono',monospace;text-transform:uppercase;margin-bottom:7px;}
.tf-row{display:flex;gap:5px;flex-wrap:wrap;margin-bottom:14px;}
.tf-btn{padding:5px 12px;font-size:11px;font-family:'Space Mono',monospace;background:var(--bg2);border:1px solid var(--border);color:var(--dim2);border-radius:5px;cursor:pointer;font-weight:700;}
.tf-btn.on{background:var(--green-bg);border-color:var(--green);color:var(--green);}
.go-btn{width:100%;background:var(--green);color:var(--bg);border:none;padding:14px;font-size:13px;font-weight:700;font-family:'Space Mono',monospace;letter-spacing:2px;text-transform:uppercase;cursor:pointer;border-radius:8px;}
.go-btn:disabled{opacity:0.4;cursor:not-allowed;}
.cancel-btn{width:100%;background:var(--red-bg);color:var(--red);border:1px solid var(--red);padding:10px;font-size:11px;font-weight:700;font-family:'Space Mono',monospace;letter-spacing:1px;cursor:pointer;border-radius:8px;margin-top:8px;}
.rcard{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:20px;margin-top:20px;}
.rcard.go{border-color:var(--green);background:var(--green-bg);}
.rcard.wait{border-color:var(--yellow);background:var(--yellow-bg);}
.rtop{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;gap:12px;}
.rsym{font-size:24px;font-weight:700;font-family:'Space Mono',monospace;letter-spacing:1px;}
.rmeta{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;align-items:center;}
.pill{font-size:10px;font-family:'Space Mono',monospace;padding:3px 8px;border-radius:4px;font-weight:700;letter-spacing:0.5px;}
.pill-go{background:var(--green-bg);color:var(--green);border:1px solid var(--green);}
.pill-wait{background:var(--yellow-bg);color:var(--yellow);border:1px solid var(--yellow);}
.pill-bull{color:var(--green);}.pill-bear{color:var(--red);}.pill-side{color:var(--yellow);}
.pill-ch{color:var(--green);}.pill-cm{color:var(--yellow);}.pill-cl{color:var(--red);}
.dot{color:var(--dim);}
.cmp-block{text-align:right;flex-shrink:0;}
.cmp-lbl{font-size:9px;letter-spacing:2px;color:var(--dim);font-family:'Space Mono',monospace;margin-bottom:3px;}
.cmp-val{font-size:21px;font-weight:700;font-family:'Space Mono',monospace;}
.reson{background:var(--bg);border-left:3px solid var(--border2);padding:12px 14px;border-radius:0 6px 6px 0;font-size:13px;line-height:1.75;color:var(--dim2);margin-bottom:14px;}
.basis-block{font-size:11px;font-family:'Space Mono',monospace;color:var(--dim2);background:var(--purple-bg);border:1px solid var(--purple);padding:9px 12px;border-radius:6px;margin-bottom:14px;line-height:1.7;}
.basis-lbl{color:var(--purple);font-weight:700;letter-spacing:1px;}
.pgrid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-bottom:12px;}
.pcell{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:11px 13px;}
.pcell-lbl{font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--dim);font-family:'Space Mono',monospace;margin-bottom:4px;}
.pcell-val{font-size:16px;font-weight:700;font-family:'Space Mono',monospace;}
.ve{color:var(--blue);}.vsl{color:var(--red);}.vt{color:var(--green);}
.bot-strip{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:6px;}
.rr-pill{font-size:11px;font-family:'Space Mono',monospace;padding:4px 10px;background:var(--green-bg);border:1px solid var(--green);color:var(--green);border-radius:4px;font-weight:700;}
.t2-note{font-size:12px;font-family:'Space Mono',monospace;color:var(--dim2);}
.t2-v{color:var(--green);font-weight:700;}
.wait-box{background:var(--bg);border-left:3px solid var(--yellow);border-radius:0 6px 6px 0;padding:12px 14px;font-size:13px;color:var(--dim2);line-height:1.75;margin-bottom:14px;}
.wait-lbl{font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--yellow);font-family:'Space Mono',monospace;font-weight:700;display:block;margin-bottom:4px;}
.al-row{display:flex;justify-content:space-between;align-items:center;background:var(--bg);border:1px solid var(--yellow);border-radius:8px;padding:12px 14px;margin-bottom:10px;gap:10px;}
.al-price{font-size:18px;font-weight:700;font-family:'Space Mono',monospace;color:var(--yellow);}
.al-lbl2{font-size:9px;letter-spacing:2px;color:var(--dim);font-family:'Space Mono',monospace;margin-bottom:3px;}
.act-row{display:flex;gap:8px;margin-top:14px;flex-wrap:wrap;}
.btn{padding:9px 16px;font-size:11px;font-family:'Space Mono',monospace;border-radius:6px;cursor:pointer;border:1px solid;letter-spacing:0.5px;font-weight:700;}
.btn-green{background:var(--green-bg);border-color:var(--green);color:var(--green);}
.btn-blue{background:var(--blue-bg);border-color:var(--blue);color:var(--blue);}
.btn-dim{background:var(--bg2);border-color:var(--border);color:var(--dim2);}
.btn-red{background:var(--red-bg);border-color:var(--red);color:var(--red);}
.btn-yellow{background:var(--yellow-bg);border-color:var(--yellow);color:var(--yellow);}
.btn:disabled{opacity:0.4;cursor:not-allowed;}
.sec-head{font-size:9px;letter-spacing:3px;text-transform:uppercase;color:var(--dim);font-family:'Space Mono',monospace;display:flex;align-items:center;gap:10px;margin:22px 0 10px;font-weight:700;}
.sec-head::after{content:'';flex:1;height:1px;background:var(--border);}
.al-item{background:var(--bg2);border:1px solid var(--border);border-left:3px solid var(--yellow);border-radius:8px;padding:11px 14px;margin-bottom:7px;display:flex;justify-content:space-between;align-items:center;gap:8px;}
.al-item.hit{border-left-color:var(--green);}
.al-sym-name{font-weight:700;font-family:'Space Mono',monospace;font-size:13px;}
.al-detail{font-size:10px;font-family:'Space Mono',monospace;color:var(--dim);margin-top:2px;}
.rm-btn{background:none;border:none;color:var(--dim);cursor:pointer;font-size:18px;padding:2px 6px;line-height:1;}
.stat-row{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;}
.stat-card{flex:1;min-width:80px;background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;text-align:center;}
.stat-n{font-size:22px;font-weight:700;font-family:'Space Mono',monospace;}
.stat-l{font-size:9px;letter-spacing:2px;color:var(--dim);font-family:'Space Mono',monospace;margin-top:2px;}
.filter-row{display:flex;gap:5px;margin-bottom:14px;flex-wrap:wrap;}
.f-btn{padding:5px 12px;font-size:10px;font-family:'Space Mono',monospace;background:var(--bg2);border:1px solid var(--border);color:var(--dim2);border-radius:5px;cursor:pointer;font-weight:700;}
.f-btn.on{background:var(--bg3);border-color:var(--border2);color:var(--text);}
.jrow,.trow{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:14px 16px;margin-bottom:8px;}
.jrow-top,.trow-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;gap:10px;}
.jsym,.tsym{font-size:15px;font-weight:700;font-family:'Space Mono',monospace;}
.jmeta,.tmeta{font-size:10px;font-family:'Space Mono',monospace;color:var(--dim);margin-top:2px;}
.jlevels{display:flex;gap:10px;flex-wrap:wrap;font-size:11px;font-family:'Space Mono',monospace;margin-bottom:8px;}
.jl-e{color:var(--blue);}.jl-sl{color:var(--red);}.jl-t{color:var(--green);}.jl-al{color:var(--yellow);}
.status-sel{font-size:10px;font-family:'Space Mono',monospace;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:4px 8px;border-radius:5px;cursor:pointer;outline:none;}
.jreason{font-size:11px;color:var(--dim2);line-height:1.6;border-top:1px solid var(--border);padding-top:8px;margin-top:4px;}
.port-stat{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;margin-bottom:16px;}
.ps-card{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;}
.ps-val{font-size:16px;font-weight:700;font-family:'Space Mono',monospace;}
.ps-lbl{font-size:9px;letter-spacing:1.5px;color:var(--dim);font-family:'Space Mono',monospace;margin-top:2px;}
.p-green{color:var(--green);}.p-red{color:var(--red);}
.tgrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:6px;margin-bottom:8px;}
.tc{background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:7px 9px;}
.tc-lbl{font-size:8px;letter-spacing:1.5px;color:var(--dim);font-family:'Space Mono',monospace;margin-bottom:3px;}
.tc-val{font-size:12px;font-weight:700;font-family:'Space Mono',monospace;}
.pnl-strip{display:flex;gap:8px;align-items:center;font-size:11px;font-family:'Space Mono',monospace;padding-top:8px;border-top:1px solid var(--border);flex-wrap:wrap;}
.method-box{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:16px;}
.method-text{font-size:11px;font-family:'Space Mono',monospace;color:var(--dim2);line-height:1.9;white-space:pre-wrap;}
.method-note{font-size:11px;color:var(--dim2);line-height:1.7;padding:10px 12px;background:var(--blue-bg);border:1px solid var(--blue);border-radius:6px;margin-bottom:14px;}
.err-box{margin-top:14px;padding:11px 14px;background:var(--red-bg);border:1px solid var(--red);border-radius:8px;color:var(--red);font-size:12px;font-family:'Space Mono',monospace;}
.warn-box{margin-top:10px;padding:9px 12px;background:var(--yellow-bg);border:1px solid var(--yellow);border-radius:6px;color:var(--yellow);font-size:11px;font-family:'Space Mono',monospace;line-height:1.6;}
.ldg-sym{font-family:'Space Mono',monospace;font-size:30px;font-weight:700;color:var(--green);animation:blink 1s ease-in-out infinite;}
@keyframes blink{0%,100%{opacity:1;}50%{opacity:0.2;}}
.ldg-msg{font-size:10px;letter-spacing:3px;text-transform:uppercase;color:var(--dim);font-family:'Space Mono',monospace;margin-top:12px;}
.ldg-time{font-size:10px;font-family:'Space Mono',monospace;color:var(--dim2);margin-top:6px;}
.empty{text-align:center;padding:40px 20px;color:var(--dim);font-size:12px;font-family:'Space Mono',monospace;line-height:2;}
.export-row{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;}
.badge-hit{font-size:9px;font-family:'Space Mono',monospace;padding:2px 7px;border-radius:3px;font-weight:700;letter-spacing:0.5px;}
.badge-open{background:var(--blue-bg);color:var(--blue);border:1px solid var(--blue);}
.badge-target{background:var(--green-bg);color:var(--green);border:1px solid var(--green);}
.badge-sl{background:var(--red-bg);color:var(--red);border:1px solid var(--red);}
.badge-closed{background:var(--bg3);color:var(--dim2);border:1px solid var(--border);}
.added-pill{font-size:10px;font-family:'Space Mono',monospace;padding:6px 12px;border-radius:5px;background:var(--green-bg);color:var(--green);border:1px solid var(--green);font-weight:700;letter-spacing:0.5px;}
.market-status{display:inline-flex;align-items:center;gap:6px;font-size:10px;font-family:'Space Mono',monospace;padding:4px 10px;border-radius:5px;letter-spacing:0.5px;font-weight:700;}
.ms-open{background:var(--green-bg);color:var(--green);border:1px solid var(--green);}
.ms-closed{background:var(--bg3);color:var(--dim2);border:1px solid var(--border);}
.ms-dot{width:6px;height:6px;border-radius:50%;background:currentColor;animation:pulse 1.5s ease-in-out infinite;}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:0.3;}}
.scan-grid{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px;}
.scan-pick{padding:6px 12px;font-size:11px;font-family:'Space Mono',monospace;background:var(--bg2);border:1px solid var(--border);color:var(--text);border-radius:5px;cursor:pointer;font-weight:700;}
.scan-pick.on{background:var(--purple-bg);border-color:var(--purple);color:var(--purple);}
.scan-progress{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;margin-bottom:12px;font-size:11px;font-family:'Space Mono',monospace;color:var(--dim2);}
.scan-bar{height:4px;background:var(--bg3);border-radius:2px;margin-top:6px;overflow:hidden;}
.scan-bar-fill{height:100%;background:var(--green);transition:width 0.3s;}
.scan-result{background:var(--bg2);border:1px solid var(--green);border-left:3px solid var(--green);border-radius:8px;padding:12px 14px;margin-bottom:8px;cursor:pointer;}
.scan-result-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;}
.scan-result-sym{font-size:14px;font-weight:700;font-family:'Space Mono',monospace;}
.scan-result-cmp{font-size:13px;font-weight:700;font-family:'Space Mono',monospace;color:var(--green);}
.scan-result-meta{font-size:10px;font-family:'Space Mono',monospace;color:var(--dim2);}
`;

function fmt(n){return n==null?"—":"₹"+Number(n).toLocaleString("en-IN",{maximumFractionDigits:2,minimumFractionDigits:2});}
function fmtShort(n){if(n==null)return"—";if(Math.abs(n)>=100000)return"₹"+(n/100000).toFixed(2)+"L";return fmt(n);}
function statusLabel(s){return s==="open"?"Open":s==="target_hit"?"✓ Target":s==="sl_hit"?"✕ SL":"Closed";}
function statusCls(s){return s==="open"?"badge-open":s==="target_hit"?"badge-target":s==="sl_hit"?"badge-sl":"badge-closed";}
function isMarketOpen(){
  const ist=new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));
  const day=ist.getDay();if(day===0||day===6)return false;
  const m=ist.getHours()*60+ist.getMinutes();
  return m>=(9*60+15)&&m<=(15*60+30);
}
function nextMarketOpenStr(){
  const ist=new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));
  const day=ist.getDay();const m=ist.getHours()*60+ist.getMinutes();
  if(day>=1&&day<=5&&m<(9*60+15))return"today 9:15 AM IST";
  if(day===5&&m>(15*60+30))return"Monday 9:15 AM IST";
  if(day===6||day===0)return"Monday 9:15 AM IST";
  return"tomorrow 9:15 AM IST";
}
const sleep=ms=>new Promise(r=>setTimeout(r,ms));

async function apiCall(path, opts = {}) {
  const token = localStorage.getItem("td:token");
  const r = await fetch(API + path, {
    ...opts,
    headers: {
      "Content-Type": "application/json",
      ...(token ? { "Authorization": "Bearer " + token } : {}),
      ...(opts.headers || {}),
    },
  });
  if (r.status === 401) {
    localStorage.removeItem("td:token");
    window.location.reload();
    throw new Error("Session expired");
  }
  if (!r.ok) {
    const err = await r.text();
    throw new Error(`API ${r.status}: ${err.slice(0,200)}`);
  }
  return r.json();
}

export default function App(){
  const [theme,setTheme]=useState(()=>localStorage.getItem("td:theme")||"dark");
  const [token,setToken]=useState(()=>localStorage.getItem("td:token"));
  const [tab,setTab]=useState(0);
  const [sym,setSym]=useState("");
  const [tf,setTf]=useState("15m");
  const [showSuggest,setShowSuggest]=useState(false);
  const [kbIdx,setKbIdx]=useState(-1);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [error,setError]=useState("");
  const [warn,setWarn]=useState("");
  const [alerts,setAlerts]=useState([]);
  const [journal,setJournal]=useState([]);
  const [portfolio,setPortfolio]=useState([]);
  const [ldgMsg,setLdgMsg]=useState("");
  const [ldgSec,setLdgSec]=useState(0);
  const [refreshing,setRefreshing]=useState(false);
  const [jFilter,setJFilter]=useState("all");
  const [savedToJournal,setSavedToJournal]=useState(false);
  const [addedToPort,setAddedToPort]=useState(false);
  const [marketOpen,setMarketOpen]=useState(isMarketOpen());
  const [scanSector,setScanSector]=useState("Banking");
  const [scanTf,setScanTf]=useState("1D");
  const [scanning,setScanning]=useState(false);
  const [scanProgress,setScanProgress]=useState({done:0,total:0,current:""});
  const [scanResults,setScanResults]=useState([]);
  const ldgTimer=useRef(null);
  const ldgSecTimer=useRef(null);
  const abortCtrl=useRef(null);
  const scanAbort=useRef(false);
  const LM=["Fetching live data from Yahoo…","Computing indicators…","Applying Paathshaala rules…","Building setup…"];

  useEffect(()=>{document.body.className=theme;localStorage.setItem("td:theme",theme);},[theme]);
  useEffect(()=>{const i=setInterval(()=>setMarketOpen(isMarketOpen()),60000);return()=>clearInterval(i);},[]);

  useEffect(()=>{loadData();},[]);

  async function loadData(){
    try{const j=await apiCall("/journal");setJournal(j);}catch(e){console.warn(e);}
    try{const p=await apiCall("/portfolio");setPortfolio(p);}catch(e){console.warn(e);}
    try{const a=await apiCall("/alerts");setAlerts(a);}catch(e){console.warn(e);}
  }

  const suggestions=useMemo(()=>{
    if(!sym||sym.length<1)return[];
    const q=sym.toUpperCase();
    return NSE_STOCKS.filter(([s,n])=>s.startsWith(q)||s.includes(q)||n.toUpperCase().includes(q)).slice(0,8);
  },[sym]);

  function pickSuggestion(s){setSym(s);setShowSuggest(false);setKbIdx(-1);}

  function cancelLoad(){
    if(abortCtrl.current){abortCtrl.current.abort();abortCtrl.current=null;}
    clearInterval(ldgTimer.current);clearInterval(ldgSecTimer.current);
    setLoading(false);setLdgSec(0);setError("Cancelled.");
  }

  async function analyze(){
    if(!sym.trim())return;
    setLoading(true);setError("");setWarn("");setResult(null);setSavedToJournal(false);setAddedToPort(false);setShowSuggest(false);
    let mi=0;setLdgMsg(LM[0]);setLdgSec(0);
    ldgTimer.current=setInterval(()=>{mi=(mi+1)%LM.length;setLdgMsg(LM[mi]);},1500);
    ldgSecTimer.current=setInterval(()=>setLdgSec(s=>s+1),1000);
    abortCtrl.current=new AbortController();
    const timeoutId=setTimeout(()=>{if(abortCtrl.current)abortCtrl.current.abort();},90000);
    try{
      const d=await fetch(API+"/analyze",{
        method:"POST",
        headers:{
          "Content-Type":"application/json",
          "Authorization":"Bearer "+(localStorage.getItem("td:token")||""),
        },
        signal:abortCtrl.current.signal,
        body:JSON.stringify({symbol:sym.trim().toUpperCase(),tf})
      }).then(r=>{
        if(r.status===401){localStorage.removeItem("td:token");window.location.reload();throw new Error("Session expired");}
        return r.ok?r.json():r.text().then(t=>{throw new Error(t);});
      });
      setResult(d);
    }catch(e){
      if(e.name==="AbortError")setError("Request timed out after 90 seconds. Try again.");
      else setError(e.message||"Analysis failed.");
    }finally{
      clearTimeout(timeoutId);
      clearInterval(ldgTimer.current);clearInterval(ldgSecTimer.current);
      setLoading(false);setLdgSec(0);abortCtrl.current=null;
    }
  }

  async function saveToJournal(){
    if(!result)return;
    try{
      const saved=await apiCall("/journal",{method:"POST",body:JSON.stringify({...result,tf,status:"open"})});
      setJournal([saved,...journal]);setSavedToJournal(true);
    }catch(e){setError(e.message);}
  }

  async function addToPortfolio(){
    if(!result?.entry)return;
    const qty=Math.floor(100000/result.entry);if(qty<1)return;
    try{
      const t={symbol:result.symbol,tf,entry:result.entry,sl:result.sl,target:result.target,target2:result.target2,
        qty,investment:qty*result.entry,current_price:result.current_price,
        pnl:(result.current_price-result.entry)*qty,
        pnl_pct:((result.current_price-result.entry)/result.entry)*100,status:"open"};
      const saved=await apiCall("/portfolio",{method:"POST",body:JSON.stringify(t)});
      setPortfolio([saved,...portfolio]);setAddedToPort(true);
    }catch(e){setError(e.message);}
  }

  async function refreshPrices(){
    if(!isMarketOpen())setWarn("Market closed — showing last-known prices. Live updates resume "+nextMarketOpenStr()+".");
    setRefreshing(true);
    try{const updated=await apiCall("/portfolio/refresh",{method:"POST"});await loadData();}
    catch(e){setError(e.message);}
    finally{setRefreshing(false);}
  }

  async function addAlert(){
    if(!result?.alert_price)return;
    const cond=result.alert_condition||"below";
    if(cond==="below"&&result.current_price<=result.alert_price){
      setWarn("Alert level ₹"+result.alert_price+" already at/below CMP. Pick a level genuinely below.");return;}
    if(cond==="above"&&result.current_price>=result.alert_price){
      setWarn("Alert level ₹"+result.alert_price+" already at/above CMP. Pick a level genuinely above.");return;}
    try{
      const saved=await apiCall("/alerts",{method:"POST",body:JSON.stringify({
        symbol:result.symbol,tf,price:result.alert_price,condition:cond,initial_price:result.current_price})});
      setAlerts([...alerts,saved]);
    }catch(e){setError(e.message);}
  }

  async function removeAlert(id){
    try{await apiCall("/alerts/"+id,{method:"DELETE"});setAlerts(alerts.filter(a=>a.id!==id));}
    catch(e){setError(e.message);}
  }

  async function updateJournalStatus(id,status){
    try{const u=await apiCall("/journal/"+id,{method:"PATCH",body:JSON.stringify({status})});
      setJournal(journal.map(x=>x.id===id?u:x));}catch(e){setError(e.message);}
  }

  async function updatePortfolioStatus(id,status){
    try{const u=await apiCall("/portfolio/"+id,{method:"PATCH",body:JSON.stringify({status})});
      setPortfolio(portfolio.map(x=>x.id===id?u:x));}catch(e){setError(e.message);}
  }

  async function runScan(){
    const list=SECTORS[scanSector]||[];if(!list.length)return;
    setScanning(true);setScanResults([]);scanAbort.current=false;
    setScanProgress({done:0,total:list.length,current:""});
    const results=[];
    for(let i=0;i<list.length;i++){
      if(scanAbort.current)break;
      const s=list[i];
      setScanProgress({done:i,total:list.length,current:s});
      try{
        const d=await fetch(API+"/analyze",{
          method:"POST",
          headers:{
            "Content-Type":"application/json",
            "Authorization":"Bearer "+(localStorage.getItem("td:token")||""),
          },
          body:JSON.stringify({symbol:s,tf:scanTf})
        }).then(r=>{
          if(r.status===401){localStorage.removeItem("td:token");window.location.reload();throw new Error("Session expired");}
          return r.json();
        });
        if(d?.tradable)results.push({...d,tf:scanTf});
        setScanResults([...results]);
      }catch(e){console.warn(s,e.message);}
      await sleep(4500); // free Gemini = 15 req/min, stay safe
    }
    setScanProgress({done:list.length,total:list.length,current:""});
    setScanning(false);
  }

  function stopScan(){scanAbort.current=true;}
  function loadScanResult(r){setResult(r);setSym(r.symbol);setTf(r.tf);setTab(0);}

  function exportXLSX(){
    const jR=journal.map(e=>({Date:new Date(e.created_at).toLocaleDateString("en-IN"),Symbol:e.symbol,TF:e.tf,
      Tradable:e.tradable?"Yes":"No",Trend:e.trend,Setup:e.setup||"",Confidence:e.confidence,
      CMP:e.current_price,Entry:e.entry||"",SL:e.sl||"",Target:e.target||"","Target 2":e.target2||"",
      "R:R":e.risk_reward||"",Status:statusLabel(e.status),"Paathshaala Rules":e.methodology_applied||"",Analysis:e.reason||""}));
    const pR=portfolio.map(t=>({Date:new Date(t.created_at).toLocaleDateString("en-IN"),Symbol:t.symbol,TF:t.tf,
      Entry:t.entry,SL:t.sl||"",Target:t.target||"",Qty:t.qty,Investment:Number(t.investment).toFixed(2),
      "Current Price":t.current_price||"","P&L (₹)":t.pnl?Number(t.pnl).toFixed(2):"",
      "P&L %":t.pnl_pct?Number(t.pnl_pct).toFixed(2):"",Status:statusLabel(t.status)}));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(jR),"Journal");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(pR),"Portfolio");
    XLSX.writeFile(wb,"trading_journal.xlsx");
  }

  const openTrades=portfolio.filter(t=>t.status==="open");
  const totalPnL=portfolio.filter(t=>t.current_price!=null).reduce((s,t)=>s+Number(t.pnl||0),0);
  const winners=portfolio.filter(t=>t.status==="target_hit").length;
  const losers=portfolio.filter(t=>t.status==="sl_hit").length;
  const filteredJ=jFilter==="all"?journal:journal.filter(e=>e.status===jFilter);
  const trendCls=t=>t==="Bullish"?"pill-bull":t==="Bearish"?"pill-bear":"pill-side";
  const confCls=c=>c==="High"?"pill-ch":c==="Medium"?"pill-cm":"pill-cl";

  if (!token) {
    return <Login onLoggedIn={(t) => { setToken(t); window.location.reload(); }} />;
  }

  return(<>
    <style>{css}</style>
    <div className="app">
      <div className="topbar">
        <div className="brand-block">
          <span className="brand">▸ Trading Desk</span>
          <span className={`market-status ${marketOpen?"ms-open":"ms-closed"}`}>
            <span className="ms-dot"></span>{marketOpen?"OPEN":"CLOSED"}
          </span>
        </div>
        <button className="theme-toggle" onClick={()=>setTheme(theme==="dark"?"light":"dark")}>
          {theme==="dark"?"☀ LIGHT":"🌙 DARK"}
        </button>
        <button className="theme-toggle" style={{marginLeft:6}} onClick={()=>{
          if(confirm("Log out?")){localStorage.removeItem("td:token");window.location.reload();}
        }}>🔒 LOGOUT</button>
      </div>
      <div className="tabs">{TABS.map((t,i)=>(<button key={t} className={`tab${tab===i?" on":""}`} onClick={()=>setTab(i)}>{t}</button>))}</div>
      <div className="method-badge">✓ PAATHSHAALA METHOD ACTIVE (CA Rahul Ranka)</div>

      {tab===0&&(<div>
        <div className="sym-wrap">
          <input className="sym-in" placeholder="Type stock — RELIANCE, INFY..."
            value={sym}
            onChange={e=>{setSym(e.target.value.toUpperCase());setShowSuggest(true);setKbIdx(-1);}}
            onFocus={()=>setShowSuggest(true)}
            onBlur={()=>setTimeout(()=>setShowSuggest(false),200)}
            onKeyDown={e=>{
              if(e.key==="ArrowDown"&&suggestions.length){e.preventDefault();setKbIdx(i=>Math.min(i+1,suggestions.length-1));}
              else if(e.key==="ArrowUp"&&suggestions.length){e.preventDefault();setKbIdx(i=>Math.max(i-1,0));}
              else if(e.key==="Enter"){if(kbIdx>=0&&suggestions[kbIdx])pickSuggestion(suggestions[kbIdx][0]);else analyze();}
              else if(e.key==="Escape"){setShowSuggest(false);setKbIdx(-1);}
            }}/>
          {showSuggest&&suggestions.length>0&&(<div className="sym-dropdown">
            {suggestions.map(([s,n],i)=>(
              <div key={s} className={`sym-opt ${i===kbIdx?"kb-active":""}`} onMouseDown={()=>pickSuggestion(s)}>
                <span className="sym-opt-sym">{s}</span><span className="sym-opt-name">{n}</span>
              </div>))}
          </div>)}
        </div>
        <div className="tf-label">Timeframe</div>
        <div className="tf-row">{TFS.map(t=>(<button key={t} className={`tf-btn${tf===t?" on":""}`} onClick={()=>setTf(t)}>{t}</button>))}</div>
        <button className="go-btn" disabled={loading||!sym.trim()} onClick={analyze}>{loading?ldgMsg:"Analyze Now →"}</button>
        {loading&&<button className="cancel-btn" onClick={cancelLoad}>✕ Cancel ({ldgSec}s)</button>}
        {error&&<div className="err-box">⚠ {error}</div>}
        {warn&&<div className="warn-box">⚠ {warn}</div>}
        {loading&&(<div style={{textAlign:"center",padding:"40px 0"}}>
          <div className="ldg-sym">{sym}</div><div className="ldg-msg">{ldgMsg}</div>
          <div className="ldg-time">{ldgSec}s elapsed</div>
        </div>)}

        {result&&!loading&&(<div className={`rcard ${result.tradable?"go":"wait"}`}>
          <div className="rtop">
            <div>
              <div className="rsym">{result.symbol}</div>
              <div className="rmeta">
                <span className={`pill ${trendCls(result.trend)}`} style={{fontSize:11}}>{result.trend}</span>
                <span className="dot">·</span>
                <span style={{fontSize:11,color:"var(--dim2)"}}>{result.setup||"—"}</span>
                <span className="dot">·</span>
                <span className={`pill ${confCls(result.confidence)}`} style={{fontSize:11}}>{result.confidence}</span>
              </div>
            </div>
            <div className="cmp-block">
              <div style={{marginBottom:8}}>
                {result.tradable?<span className="pill pill-go">✓ TRADABLE</span>:<span className="pill pill-wait">⏳ WAIT</span>}
              </div>
              <div className="cmp-lbl">LIVE PRICE</div>
              <div className="cmp-val">{fmt(result.current_price)}</div>
            </div>
          </div>
          <div className="reson">{result.reason}</div>
          {result.methodology_applied&&(<div className="basis-block">
            <span className="basis-lbl">PAATHSHAALA RULES: </span>{result.methodology_applied}
          </div>)}
          {result.tradable?(<>
            <div className="pgrid">
              <div className="pcell"><div className="pcell-lbl">Entry</div><div className="pcell-val ve">{fmt(result.entry)}</div></div>
              <div className="pcell"><div className="pcell-lbl">Stop Loss</div><div className="pcell-val vsl">{fmt(result.sl)}</div></div>
              <div className="pcell"><div className="pcell-lbl">Target 1</div><div className="pcell-val vt">{fmt(result.target)}</div></div>
            </div>
            <div className="bot-strip">
              {result.target2&&<span className="t2-note">T2: <span className="t2-v">{fmt(result.target2)}</span></span>}
              {result.risk_reward&&<span className="rr-pill">R:R {result.risk_reward}</span>}
            </div>
            <div className="act-row">
              {savedToJournal?<span className="added-pill">✓ Saved to Journal</span>:<button className="btn btn-blue" onClick={saveToJournal}>+ Save to Journal</button>}
              {addedToPort?<span className="added-pill">✓ Added ₹1L Position</span>:<button className="btn btn-green" onClick={addToPortfolio}>+ Add ₹1L Mock Position</button>}
            </div>
          </>):(<>
            <div className="wait-box">
              <span className="wait-lbl">When to Enter</span>{result.when_to_enter||"Wait for setup conditions."}
            </div>
            {result.alert_price&&(<div className="al-row">
              <div>
                <div className="al-lbl2">WATCH LEVEL — {result.alert_condition==="above"?"BREAK ABOVE":"DROP BELOW"}</div>
                <div className="al-price">{fmt(result.alert_price)}</div>
              </div>
              <button className="btn btn-yellow" onClick={addAlert}>🔔 Set Alert</button>
            </div>)}
            <div className="act-row">
              {savedToJournal?<span className="added-pill">✓ Saved to Journal</span>:<button className="btn btn-blue" onClick={saveToJournal}>+ Save to Journal</button>}
            </div>
          </>)}
        </div>)}

        {alerts.length>0&&(<>
          <div className="sec-head">Active Alerts ({alerts.filter(a=>!a.triggered).length})</div>
          {!marketOpen&&<div className="warn-box" style={{marginTop:0,marginBottom:10}}>Market closed — alerts paused. Resumes {nextMarketOpenStr()}.</div>}
          {alerts.map(a=>(<div key={a.id} className={`al-item ${a.triggered?"hit":""}`}>
            <div>
              <div className="al-sym-name">{a.triggered?"✅":"🔔"} {a.symbol}
                <span className={`badge-hit ${a.triggered?"badge-target":"badge-open"}`} style={{marginLeft:8}}>
                  {a.triggered?"Triggered":"Watching"}
                </span>
              </div>
              <div className="al-detail">
                {a.condition==="above"?"Break above":"Drop below"} {fmt(Number(a.price))} · {a.tf}
                {a.triggered&&a.triggered_at?` · hit ${new Date(a.triggered_at).toLocaleString("en-IN")}`:" · server polls every 5 min"}
              </div>
            </div>
            <button className="rm-btn" onClick={()=>removeAlert(a.id)}>×</button>
          </div>))}
        </>)}
      </div>)}

      {tab===1&&(<div>
        <div className="warn-box" style={{marginTop:0,marginBottom:14}}>
          Free Gemini = 15 req/min, 1500/day. Scanner uses 4.5s gap between stocks → 10 stocks ≈ 45s.
        </div>
        <div className="tf-label">Sector / Watchlist</div>
        <div className="scan-grid">
          {Object.keys(SECTORS).map(s=>(
            <button key={s} className={`scan-pick ${scanSector===s?"on":""}`} onClick={()=>!scanning&&setScanSector(s)} disabled={scanning}>
              {s} ({SECTORS[s].length})
            </button>))}
        </div>
        <div className="tf-label">Timeframe (Paathshaala: Daily for Nifty 100)</div>
        <div className="tf-row">
          {["15m","1h","1D","1W"].map(t=>(<button key={t} className={`tf-btn${scanTf===t?" on":""}`} onClick={()=>!scanning&&setScanTf(t)} disabled={scanning}>{t}</button>))}
        </div>
        {!scanning?<button className="go-btn" onClick={runScan}>▸ Scan {scanSector} on {scanTf}</button>:<button className="cancel-btn" onClick={stopScan}>✕ Stop Scan</button>}
        {(scanning||scanProgress.done>0)&&(<div className="scan-progress" style={{marginTop:14}}>
          <div>Progress: {scanProgress.done}/{scanProgress.total} {scanProgress.current&&"· checking "+scanProgress.current}</div>
          <div className="scan-bar"><div className="scan-bar-fill" style={{width:`${(scanProgress.done/Math.max(scanProgress.total,1))*100}%`}}></div></div>
        </div>)}
        {scanResults.length>0&&(<>
          <div className="sec-head">Tradable Setups ({scanResults.length})</div>
          {scanResults.map(r=>(<div key={r.symbol} className="scan-result" onClick={()=>loadScanResult(r)}>
            <div className="scan-result-top">
              <span className="scan-result-sym">{r.symbol}</span>
              <span className="scan-result-cmp">{fmt(r.current_price)}</span>
            </div>
            <div className="scan-result-meta">{r.setup} · {r.confidence} · E:{fmt(r.entry)} SL:{fmt(r.sl)} T:{fmt(r.target)} · {r.risk_reward}</div>
          </div>))}
        </>)}
        {!scanning&&scanProgress.done>0&&scanResults.length===0&&(<div className="empty">No Paathshaala setups in {scanSector} on {scanTf}.</div>)}
      </div>)}

      {tab===2&&(<div>
        <div className="stat-row">
          <div className="stat-card"><div className="stat-n" style={{color:"var(--text)"}}>{journal.length}</div><div className="stat-l">Total</div></div>
          <div className="stat-card"><div className="stat-n" style={{color:"var(--blue)"}}>{journal.filter(e=>e.status==="open").length}</div><div className="stat-l">Open</div></div>
          <div className="stat-card"><div className="stat-n" style={{color:"var(--green)"}}>{journal.filter(e=>e.status==="target_hit").length}</div><div className="stat-l">Targets</div></div>
          <div className="stat-card"><div className="stat-n" style={{color:"var(--red)"}}>{journal.filter(e=>e.status==="sl_hit").length}</div><div className="stat-l">SL Hits</div></div>
        </div>
        <div className="export-row"><button className="btn btn-dim" onClick={exportXLSX}>↓ Export .xlsx</button></div>
        <div className="filter-row">
          {[["all","All"],["open","Open"],["target_hit","Target"],["sl_hit","SL Hit"],["closed","Closed"]].map(([v,l])=>(
            <button key={v} className={`f-btn${jFilter===v?" on":""}`} onClick={()=>setJFilter(v)}>{l}</button>))}
        </div>
        {filteredJ.length===0?<div className="empty">No entries yet.<br/>Analyze a stock and click "Save to Journal".</div>:
          filteredJ.map(e=>(<div key={e.id} className="jrow">
            <div className="jrow-top">
              <div>
                <div className="jsym">{e.symbol} <span style={{fontWeight:400,fontSize:11,color:"var(--dim)"}}>{e.tf}</span></div>
                <div className="jmeta">{new Date(e.created_at).toLocaleString("en-IN")} · {e.trend} · {e.setup||"—"} · {e.confidence}</div>
              </div>
              <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
                <span className={`badge-hit ${statusCls(e.status)}`}>{statusLabel(e.status)}</span>
                <select className="status-sel" value={e.status} onChange={ev=>updateJournalStatus(e.id,ev.target.value)}>
                  <option value="open">Open</option><option value="target_hit">✓ Target Hit</option>
                  <option value="sl_hit">✕ SL Hit</option><option value="closed">Closed</option>
                </select>
              </div>
            </div>
            <div className="jlevels">
              <span>CMP: <span style={{color:"var(--text)",fontWeight:700}}>{fmt(e.current_price)}</span></span>
              {e.entry&&<span className="jl-e">E: {fmt(e.entry)}</span>}
              {e.sl&&<span className="jl-sl">SL: {fmt(e.sl)}</span>}
              {e.target&&<span className="jl-t">T: {fmt(e.target)}</span>}
              {e.alert_price&&!e.tradable&&<span className="jl-al">Watch: {fmt(e.alert_price)}</span>}
              {e.risk_reward&&<span style={{color:"var(--dim2)"}}>{e.risk_reward}</span>}
            </div>
            {e.methodology_applied&&(<div style={{fontSize:10,fontFamily:"'Space Mono',monospace",color:"var(--purple)",marginBottom:6,padding:"5px 8px",background:"var(--purple-bg)",border:"1px solid var(--purple)",borderRadius:5}}>
              <strong>RULES:</strong> {e.methodology_applied}
            </div>)}
            <div className="jreason">{e.reason}</div>
          </div>))}
      </div>)}

      {tab===3&&(<div>
        <div className="port-stat">
          <div className="ps-card"><div className="ps-val">{portfolio.length}</div><div className="ps-lbl">Trades</div></div>
          <div className="ps-card"><div className="ps-val">{openTrades.length}</div><div className="ps-lbl">Open</div></div>
          <div className="ps-card"><div className={`ps-val ${totalPnL>=0?"p-green":"p-red"}`}>{fmtShort(totalPnL)}</div><div className="ps-lbl">Total P&L</div></div>
          <div className="ps-card"><div className="ps-val" style={{color:"var(--dim2)"}}>{winners}W / {losers}L</div><div className="ps-lbl">W / L</div></div>
        </div>
        <div className="export-row">
          <button className="btn btn-dim" onClick={refreshPrices} disabled={refreshing||openTrades.length===0}>{refreshing?"Refreshing…":"↻ Refresh Prices"}</button>
          <button className="btn btn-dim" onClick={exportXLSX}>↓ Export .xlsx</button>
        </div>
        {!marketOpen&&portfolio.length>0&&<div className="warn-box" style={{marginTop:0,marginBottom:12}}>Market closed — refresh shows last prices.</div>}
        {portfolio.length===0?<div className="empty">No mock trades.<br/>Analyze a tradable stock and click "Add ₹1L Mock Position".</div>:
          portfolio.map(t=>{
            const pos=t.pnl!=null?Number(t.pnl)>=0:null;
            return(<div key={t.id} className="trow">
              <div className="trow-top">
                <div>
                  <div className="tsym">{t.symbol} <span style={{fontSize:10,color:"var(--dim)",fontWeight:400}}>{t.tf}</span></div>
                  <div className="tmeta">{new Date(t.created_at).toLocaleDateString("en-IN")} · Qty: {t.qty} · ₹{Number(t.investment).toFixed(0)} invested</div>
                </div>
                <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
                  <span className={`badge-hit ${statusCls(t.status)}`}>{statusLabel(t.status)}</span>
                  <select className="status-sel" value={t.status} onChange={ev=>updatePortfolioStatus(t.id,ev.target.value)}>
                    <option value="open">Open</option><option value="target_hit">✓ Target Hit</option>
                    <option value="sl_hit">✕ SL Hit</option><option value="closed">Closed</option>
                  </select>
                </div>
              </div>
              <div className="tgrid">
                <div className="tc"><div className="tc-lbl">Entry</div><div className="tc-val ve">{fmt(t.entry)}</div></div>
                <div className="tc"><div className="tc-lbl">SL</div><div className="tc-val vsl">{fmt(t.sl)}</div></div>
                <div className="tc"><div className="tc-lbl">Target</div><div className="tc-val vt">{fmt(t.target)}</div></div>
                <div className="tc"><div className="tc-lbl">Current</div><div className="tc-val">{fmt(t.current_price)}</div></div>
              </div>
              {t.pnl!=null&&(<div className="pnl-strip">
                <span className={pos?"p-green":"p-red"}>{pos?"▲":"▼"} P&L: {fmt(Number(t.pnl))}</span>
                <span className="dot">·</span>
                <span className={pos?"p-green":"p-red"}>{t.pnl_pct!=null?(Number(t.pnl_pct)>=0?"+":"")+Number(t.pnl_pct).toFixed(2)+"%":"—"}</span>
              </div>)}
            </div>);
          })}
      </div>)}

      {tab===4&&(<div>
        <div className="method-note">
          📋 <strong>Active methodology</strong> — applies to every analysis & scan. Currently using <strong>Paathshaala (CA Rahul Ranka)</strong>.
        </div>
        <div className="method-box">
          <div className="sec-head" style={{margin:"0 0 12px"}}>Rules in Use</div>
          <div className="method-text">{PAATHSHAALA_INFO}</div>
        </div>
      </div>)}
    </div>
  </>);
}
