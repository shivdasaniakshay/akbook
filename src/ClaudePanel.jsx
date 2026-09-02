import React, { useEffect, useState } from "react";
import { supabase } from "./sync.js";

// "Ask Claude" — describe a change; it becomes a GitHub issue that the Claude
// Code action picks up. Claude opens a pull request with a preview link; "Deploy"
// merges it and Vercel ships the new version to everyone.
const font = "'Avenir Next', 'Segoe UI', system-ui, sans-serif";

async function api(path, opts = {}) {
  const { data } = await supabase.auth.getSession();
  const r = await fetch(path, { ...opts, headers: { "Content-Type": "application/json", Authorization: "Bearer " + data?.session?.access_token, ...(opts.headers || {}) } });
  const body = await r.json().catch(() => null);
  if (!r.ok || body === null) throw new Error(body?.error || (r.ok ? "The API isn't deployed here (it runs on Vercel)." : r.statusText));
  return body;
}

export default function ClaudePanel({ onClose }) {
  const [text, setText] = useState("");
  const [items, setItems] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  const load = async () => { try { setItems(await api("/api/requests")); setErr(""); } catch (e) { setErr(e.message); setItems([]); } };
  useEffect(() => { load(); const t = setInterval(load, 20000); return () => clearInterval(t); }, []);

  const submit = async () => {
    if (!text.trim()) return;
    setBusy(true); setErr("");
    try { await api("/api/requests", { method: "POST", body: JSON.stringify({ text }) }); setText(""); await load(); }
    catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const deploy = async (pr) => {
    if (!window.confirm(`Deploy pull request #${pr} to the live site?`)) return;
    setBusy(true); setErr("");
    try { await api("/api/merge", { method: "POST", body: JSON.stringify({ pr }) }); await load(); }
    catch (e) { setErr(e.message); }
    setBusy(false);
  };

  const stage = (it) => {
    if (it.pr?.merged) return ["Deployed", "#6fbf73"];
    if (it.pr) return ["Ready to review", "#c9a44c"];
    if (it.state === "closed") return ["Closed", "#a89f8f"];
    return ["Claude is working…", "#8fb3e0"];
  };
  const link = { color: "#c9a44c", textDecoration: "none", fontWeight: 700 };

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 9500, background: "rgba(0,0,0,0.45)", fontFamily: font }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ position: "absolute", top: 0, right: 0, bottom: 0, width: "min(460px, 100vw)", background: "#1c1a17", color: "#efe8da", boxShadow: "-8px 0 30px rgba(0,0,0,0.4)", display: "flex", flexDirection: "column" }}>
        <div style={{ padding: "18px 22px 12px", borderBottom: "1px solid #302c27", display: "flex", alignItems: "center" }}>
          <div>
            <div style={{ fontFamily: "Georgia, serif", fontSize: 20 }}>Ask Claude</div>
            <div style={{ color: "#a89f8f", fontSize: 12 }}>Describe a change to the app. Claude builds it and posts a preview; you press Deploy.</div>
          </div>
          <button onClick={onClose} style={{ marginLeft: "auto", border: "none", background: "transparent", color: "#a89f8f", fontSize: 20, cursor: "pointer" }}>×</button>
        </div>
        <div style={{ padding: "14px 22px", borderBottom: "1px solid #302c27" }}>
          <textarea value={text} onChange={(e) => setText(e.target.value)} rows={5} placeholder={"e.g. Add a column to the Tabs ledger showing each player's running balance in the club's currency, and make the export include it."}
            style={{ width: "100%", boxSizing: "border-box", padding: 10, borderRadius: 7, border: "1px solid #46403a", background: "#26231f", color: "#efe8da", fontSize: 13.5, fontFamily: font, resize: "vertical" }} />
          <button onClick={submit} disabled={busy || !text.trim()} style={{ marginTop: 8, padding: "9px 16px", borderRadius: 7, border: "none", background: "#c9a44c", color: "#1c1a17", fontWeight: 700, fontSize: 13.5, cursor: "pointer", opacity: busy || !text.trim() ? 0.6 : 1 }}>
            {busy ? "Sending…" : "Send to Claude"}
          </button>
          {err && <div style={{ marginTop: 8, color: "#e07a6a", fontSize: 12.5 }}>{err}</div>}
        </div>
        <div style={{ flex: 1, overflow: "auto", padding: "10px 22px 30px" }}>
          <div style={{ color: "#a89f8f", fontSize: 11.5, textTransform: "uppercase", letterSpacing: 1, margin: "8px 0" }}>Requests</div>
          {items === null && <div style={{ color: "#a89f8f", fontSize: 13 }}>Loading…</div>}
          {items?.length === 0 && <div style={{ color: "#a89f8f", fontSize: 13 }}>No requests yet.</div>}
          {items?.map((it) => {
            const [label, color] = stage(it);
            return (
              <div key={it.number} style={{ padding: "12px 0", borderBottom: "1px solid #302c27" }}>
                <div style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
                  <span style={{ fontSize: 13.5, fontWeight: 700, flex: 1 }}>{it.title}</span>
                  <span style={{ fontSize: 11.5, color, whiteSpace: "nowrap" }}>{label}</span>
                </div>
                <div style={{ color: "#a89f8f", fontSize: 12, marginTop: 3 }}>{new Date(it.created_at).toLocaleString()} · <a href={it.html_url} target="_blank" rel="noreferrer" style={{ color: "#a89f8f" }}>#{it.number}</a></div>
                {it.pr && (
                  <div style={{ display: "flex", gap: 14, marginTop: 8, fontSize: 12.5, alignItems: "center" }}>
                    {it.pr.preview_url && <a href={it.pr.preview_url} target="_blank" rel="noreferrer" style={link}>Open preview</a>}
                    <a href={it.pr.html_url} target="_blank" rel="noreferrer" style={{ color: "#a89f8f" }}>View changes</a>
                    {!it.pr.merged && it.pr.state === "open" && (
                      <button onClick={() => deploy(it.pr.number)} disabled={busy} style={{ marginLeft: "auto", padding: "5px 12px", borderRadius: 6, border: "1px solid #c9a44c", background: "transparent", color: "#c9a44c", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>Deploy</button>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}
