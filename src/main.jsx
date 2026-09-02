import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import { configured, unlock, lock, savedPassphrase, onRemoteChange, onStatus, getStatus } from "./sync.js";
import ClaudePanel from "./ClaudePanel.jsx";

const font = "'Avenir Next', 'Segoe UI', system-ui, sans-serif";

// Passphrase screen. The same passphrase logs into Supabase and decrypts the data.
function Gate({ onUnlocked }) {
  const [pass, setPass] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [auto, setAuto] = useState(!!savedPassphrase());

  const go = async (p) => {
    setBusy(true); setErr("");
    try { await unlock(p); onUnlocked(); }
    catch (e) { setErr(e.message || String(e)); setAuto(false); }
    setBusy(false);
  };
  useEffect(() => { const p = savedPassphrase(); if (p) go(p); }, []);

  return (
    <div style={{ minHeight: "100vh", display: "grid", placeItems: "center", background: "#1c1a17", color: "#efe8da", fontFamily: font }}>
      <form onSubmit={(e) => { e.preventDefault(); go(pass); }} style={{ width: 340, padding: 32, background: "#26231f", borderRadius: 12, boxShadow: "0 12px 40px rgba(0,0,0,0.4)" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 26 }}>AK's Book</div>
        <div style={{ color: "#c9a44c", fontSize: 13, marginBottom: 22 }}>weekly accounting · shared</div>
        {auto ? <div style={{ color: "#a89f8f", fontSize: 13 }}>Unlocking…</div> : (
          <>
            <input autoFocus type="password" value={pass} onChange={(e) => setPass(e.target.value)} placeholder="Passphrase" disabled={busy}
              style={{ width: "100%", boxSizing: "border-box", padding: "11px 12px", borderRadius: 7, border: "1px solid #46403a", background: "#1c1a17", color: "#efe8da", fontSize: 15 }} />
            <button type="submit" disabled={busy || !pass} style={{ marginTop: 12, width: "100%", padding: "10px", borderRadius: 7, border: "none", background: "#c9a44c", color: "#1c1a17", fontWeight: 700, fontSize: 14, cursor: "pointer", opacity: busy || !pass ? 0.6 : 1 }}>
              {busy ? "Unlocking…" : "Unlock"}
            </button>
            {err && <div style={{ marginTop: 12, color: "#e07a6a", fontSize: 13 }}>{err}</div>}
            {!configured && <div style={{ marginTop: 12, color: "#a89f8f", fontSize: 12 }}>Supabase isn't configured yet — see SETUP.md.</div>}
          </>
        )}
      </form>
    </div>
  );
}

// Small floating dock: sync status, the Claude request panel, and Lock.
function Dock({ onLock, onOpenClaude }) {
  const [st, setSt] = useState({ s: getStatus(), msg: "" });
  useEffect(() => onStatus((s, msg) => setSt({ s, msg })), []);
  const color = { synced: "#6fbf73", saving: "#c9a44c", offline: "#a89f8f", error: "#e07a6a" }[st.s] || "#a89f8f";
  const label = { synced: "synced", saving: "saving…", offline: "offline", error: "save failed" }[st.s] || st.s;
  const btn = { border: "none", background: "transparent", color: "#efe8da", cursor: "pointer", fontSize: 12.5, fontWeight: 700, padding: "6px 10px" };
  return (
    <div title={st.msg} style={{ position: "fixed", right: 16, bottom: 16, zIndex: 9000, display: "flex", alignItems: "center", gap: 4, padding: "4px 6px 4px 12px", background: "rgba(28,26,23,0.95)", color: "#efe8da", borderRadius: 999, boxShadow: "0 6px 24px rgba(0,0,0,0.35)", fontFamily: font, fontSize: 12.5 }}>
      <span style={{ width: 8, height: 8, borderRadius: 4, background: color, marginRight: 6 }} />
      <span style={{ color: "#a89f8f" }}>{label}</span>
      <span style={{ width: 1, height: 16, background: "#46403a", margin: "0 6px" }} />
      <button style={btn} onClick={onOpenClaude}>Ask Claude</button>
      <button style={{ ...btn, color: "#a89f8f" }} onClick={onLock}>Lock</button>
    </div>
  );
}

function Root() {
  const [unlocked, setUnlocked] = useState(false);
  const [gen, setGen] = useState(0); // bump to remount App after a remote change
  const [claudeOpen, setClaudeOpen] = useState(false);
  useEffect(() => onRemoteChange(() => setGen((g) => g + 1)), []);
  if (!unlocked) return <Gate onUnlocked={() => setUnlocked(true)} />;
  return (
    <>
      <App key={gen} />
      <Dock onLock={async () => { await lock(); setUnlocked(false); }} onOpenClaude={() => setClaudeOpen(true)} />
      {claudeOpen && <ClaudePanel onClose={() => setClaudeOpen(false)} />}
    </>
  );
}

createRoot(document.getElementById("root")).render(<Root />);
