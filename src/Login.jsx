import { useState } from "react";

const css = `
.login-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;}
.login-card{background:var(--bg2);border:1px solid var(--border);border-radius:14px;padding:32px 28px;width:100%;max-width:380px;}
.login-title{font-family:'Space Mono',monospace;font-size:11px;letter-spacing:3px;color:var(--green);text-transform:uppercase;font-weight:700;text-align:center;margin-bottom:8px;}
.login-sub{font-size:11px;color:var(--dim);text-align:center;margin-bottom:24px;font-family:'Space Mono',monospace;letter-spacing:1px;}
.login-input{width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:14px 18px;font-size:14px;font-family:'Space Mono',monospace;letter-spacing:2px;outline:none;border-radius:8px;margin-bottom:12px;}
.login-input:focus{border-color:var(--green);}
.login-btn{width:100%;background:var(--green);color:var(--bg);border:none;padding:13px;font-size:12px;font-weight:700;font-family:'Space Mono',monospace;letter-spacing:2px;text-transform:uppercase;cursor:pointer;border-radius:8px;}
.login-btn:disabled{opacity:0.4;cursor:not-allowed;}
.login-err{background:var(--red-bg);border:1px solid var(--red);color:var(--red);padding:9px 12px;border-radius:6px;font-size:11px;font-family:'Space Mono',monospace;margin-top:12px;text-align:center;}
`;

const API = import.meta.env.VITE_API_BASE || "/api";

export default function Login({ onLoggedIn }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit() {
    if (!pw) return;
    setErr(""); setLoading(true);
    try {
      const r = await fetch(API + "/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password: pw }),
      });
      if (!r.ok) {
        const t = await r.json().catch(() => ({ error: "Login failed" }));
        throw new Error(t.error || "Login failed");
      }
      const { token } = await r.json();
      localStorage.setItem("td:token", token);
      onLoggedIn(token);
    } catch (e) {
      setErr(e.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <>
      <style>{css}</style>
      <div className="login-wrap">
        <div className="login-card">
          <div className="login-title">▸ Trading Desk</div>
          <div className="login-sub">Enter password to continue</div>
          <input
            className="login-input"
            type="password"
            placeholder="••••••••"
            value={pw}
            onChange={e => setPw(e.target.value)}
            onKeyDown={e => e.key === "Enter" && submit()}
            autoFocus
          />
          <button className="login-btn" onClick={submit} disabled={loading || !pw}>
            {loading ? "Verifying…" : "Login"}
          </button>
          {err && <div className="login-err">⚠ {err}</div>}
        </div>
      </div>
    </>
  );
}
