import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

// Persistence lives in sync.js: localStorage is the cache, Supabase is the shared (encrypted) truth.
import { store, blobPut as idbPut, blobGet as idbGet, blobDel as idbDel } from "./sync.js";

// ———————————————— Seed / snapshot ————————————————
// Everything the app knows lives under these localStorage keys.
const STATE_KEYS_STATIC = ["fishtank-config-v4", "fishtank-lastweek-v4", "agentclubs-v3", "tabs-v1", "allamerican-v1", "allamerican-lastweek-v1", "ownerclubs-v1", "archive-index-v1", "book-weeks-v1", "book-checklist-v1"];
// Owner clubs added later live under oc-cfg-* / oc-week-* keys — pick those up too.
const allStateKeys = () => {
  const all = [...STATE_KEYS_STATIC];
  try { for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); if (k && (k.startsWith("oc-cfg-") || k.startsWith("oc-week-")) && !all.includes(k)) all.push(k); } } catch (e) {}
  return all;
};

// To bake the CURRENT metadata into the app itself: header → "Export data",
// open the downloaded fishtank-seed-*.json, and paste its contents here in
// place of `null`. Anyone who runs this build then starts pre-loaded with
// that data. Alternatively (no code edit): drop the file next to index.html
// renamed to `seed.json` — the app fetches it on first run.
// A seed only fills keys the visitor doesn't already have; local edits win.
const SEED_STATE = null;

let seedApplied = false;
async function applySeedOnce() {
  if (seedApplied) return;
  seedApplied = true;
  let seed = SEED_STATE;
  if (!seed) {
    try {
      const r = await fetch("./seed.json", { cache: "no-store" });
      if (r.ok) seed = await r.json();
    } catch (e) {}
  }
  if (!seed || typeof seed !== "object") return;
  for (const k of new Set([...STATE_KEYS_STATIC, ...Object.keys(seed).filter((k) => k.startsWith("oc-cfg-") || k.startsWith("oc-week-"))])) {
    try {
      if (seed[k] != null && localStorage.getItem(k) == null)
        await store.set(k, typeof seed[k] === "string" ? seed[k] : JSON.stringify(seed[k]));
    } catch (e) {}
  }
}

function exportSnapshot() {
  const out = { app: "all-in-fish-tank", exportedAt: new Date().toISOString() };
  allStateKeys().forEach((k) => {
    const v = localStorage.getItem(k);
    if (v != null) { try { out[k] = JSON.parse(v); } catch (e) { out[k] = v; } }
  });
  const blob = new Blob([JSON.stringify(out, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `fishtank-seed-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}

async function importSnapshot(file) {
  const seed = JSON.parse(await file.text());
  let n = 0;
  const keys = [...new Set([...STATE_KEYS_STATIC, ...Object.keys(seed).filter((k) => k.startsWith("oc-cfg-") || k.startsWith("oc-week-"))])];
  for (const k of keys) {
    if (seed[k] != null) { await store.set(k, typeof seed[k] === "string" ? seed[k] : JSON.stringify(seed[k])); n++; }
  }
  if (!n) throw new Error("No recognized data in that file.");
}

// ————————————————————————————————————————————————————————————————
// All In Fish Tank — Weekly Settlement Engine (v4)
// Settlements · umbrella & SA reports · house-backed ledger
// (makeup nets = P&L + RB credit; action buys) · Ak-Jon recon ·
// light/dark theme. Config persists week to week.
// ————————————————————————————————————————————————————————————————

const PALETTES = {
  light: {
    "--paper": "#F7F2E8", "--card": "#FFFDF8", "--cream": "#F3EAD8",
    "--gold": "#B49A5E", "--goldAccent": "#8C7440", "--ink": "#2B241C", "--bar": "#332A22",
    "--green": "#1E7B34", "--red": "#C0392B", "--mute": "#8A7E6C", "--line": "#E4D9C2",
    "--rowAlt": "#FBF6EB", "--banner": "#FBF3D9", "--surface": "#FFFFFF", "--onGold": "#FFFFFF",
    "--barText": "#F2E7CC", "--barMute": "#C9BB9A", "--barSubtle": "#B4A585", "--barGold": "#F0C97A",
    "--barGreen": "#7CCB8B", "--barRed": "#F0958A",
    "--pillRedBg": "#F9E4E1", "--pillGreenBg": "#E3F1E5", "--pillBlueBg": "#E2ECF4", "--pillGoldBg": "#F1E8D3",
    "--pillBlueFg": "#31587A", "--chipOff": "#C9BFA9", "--errBg": "#F9E4E1",
  },
  dark: {
    "--paper": "#131009", "--card": "#1D180F", "--cream": "#262013", "--gold": "#D4B36A",
    "--goldAccent": "#E6C88A", "--ink": "#F3EAD6", "--bar": "#0B0906",
    "--green": "#6FD08C", "--red": "#EF8677", "--mute": "#A79A80", "--line": "#332B1C",
    "--rowAlt": "#231D10", "--banner": "#2E2614", "--surface": "#2A2313", "--onGold": "#171207",
    "--barText": "#F6ECD2", "--barMute": "#BCAD8C", "--barSubtle": "#A29476", "--barGold": "#F4CE7E",
    "--barGreen": "#84D695", "--barRed": "#F49C90",
    "--pillRedBg": "#43231C", "--pillGreenBg": "#1C3A24", "--pillBlueBg": "#1E2E3D", "--pillBlueFg": "#A6CBEA",
    "--pillGoldBg": "#3B3013", "--chipOff": "#5F5439", "--errBg": "#43231C",
  },
};
const C = {
  paper: "var(--paper)", card: "var(--card)", cream: "var(--cream)",
  gold: "var(--gold)", goldDark: "var(--goldAccent)", ink: "var(--ink)", bar: "var(--bar)",
  green: "var(--green)", red: "var(--red)", mute: "var(--mute)", line: "var(--line)",
  rowAlt: "var(--rowAlt)", banner: "var(--banner)", surface: "var(--surface)",
};

const DEFAULT_SA_DEALS = {
  "9416-1077": 80, "9956-9064": 75, "4590-9906": 75, "5207-4267": 82.5,
  "9812-1646": 80, "1788-2643": 75, "2340-6362": 65, "6340-0999": 75,
  "8775-1559": 75, "8389-0206": 80, "5532-5157": 80, "5349-7456": 100,
  "4243-3806": 80, "4617-6330": 90, "5805-5568": 80, "9818-6720": 75,
  "2981-4306": 65, "5349-5156": 90, "4091-6935": 80, "1581-1690": 75,
  "8474-9014": 80, "3011-0934": 70, "1652-5578": 90,
};
const DEFAULT_PLAYER_DEALS = {
  "2859-2602": 70, "1793-1073": 40, "3409-6651": 70,
  "8835-5140": 70, "7623-2110": 50, "7959-6530": 70,
};
const DEFAULT_NAMES = {
  "9416-1077": "SharpCheddar", "9956-9064": "FishSupport", "4590-9906": "CUMROCKET",
  "5207-4267": "H8Varience", "9812-1646": "Cashlover777", "1788-2643": "Animal3391",
  "2340-6362": "Humpback679", "6340-0999": "Leaderzay", "8775-1559": "Hannibal0",
  "8389-0206": "DiorSauvage", "5532-5157": "Rlawnsgud", "5349-7456": "punterx",
  "4243-3806": "shortbusbully88", "4617-6330": "BroadwaySupport", "5805-5568": "catdad777",
  "9818-6720": "wheatie DOG", "2981-4306": "WrongCalc", "5349-5156": "imnotpunting",
  "4091-6935": "GamadGadol", "1581-1690": "TheBettor", "8474-9014": "Ved_13",
  "3011-0934": "LuckilyLucky", "1652-5578": "Sus Moustache",
  "2859-2602": "krishdhawan", "1793-1073": "RobinhoodAP", "3409-6651": "syao12",
  "8835-5140": "Biggest Donk", "7623-2110": "soggy waffles", "7959-6530": "CorporalToenail",
};

const DEFAULT_CONFIG = {
  theme: "dark", themeV2: true,
  defaultTB: 80,
  saDeals: { ...DEFAULT_SA_DEALS },
  playerDeals: { ...DEFAULT_PLAYER_DEALS },
  // Action buys on regular players: house taxes pct% of wins / rebates pct% of losses.
  actionTax: {}, // memberId -> { pct, backer: 'split'|'ak'|'jon' }
  confirmedSAs: Object.fromEntries(Object.keys(DEFAULT_SA_DEALS).map((k) => [k, true])),
  confirmedPlayers: Object.fromEntries(Object.keys(DEFAULT_PLAYER_DEALS).map((k) => [k, true])),
  names: { ...DEFAULT_NAMES },
  ownAccounts: { ak: ["axe7777", "gimmezemoneys"], jon: ["Flashbrook123"] },
  // 'makeup': weekly net = P&L + RB credit. Above makeup → player paid their %
  //           of the excess, backer books the rest. Below → no cash, net accrues
  //           to makeup on the backer's book.
  // 'action': backer owns actionPct% of (P&L + RB); player settles the rest.
  backed: {
    dingleberry23: { name: "dingleberry23", deal: "makeup", rbNormal: 75, rbMakeup: 75, makeup: 0, playerProfitPct: 50, backer: "split" },
    wjewje12:      { name: "wjewje12", deal: "makeup", rbNormal: 65, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: "split" },
    niceblufflol:  { name: "niceblufflol", deal: "makeup", rbNormal: 65, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: "jon" },
    gigapuntwhale: { name: "gigapuntwhale", deal: "action", actionPct: 50, rbPct: 100, backer: "jon" },
    gruzzy:        { name: "gruzzy", deal: "action", actionPct: 75, rbPct: 100, backer: "jon" },
    jumpingguppy:  { name: "jumpingguppy", deal: "action", actionPct: 100, rbPct: 100, backer: "jon" },
  },
  umbrellas: [
    { id: "u-vivaan", name: "Vivaan Rastoghi", saIds: ["1788-2643", "4243-3806", "4617-6330", "9818-6720", "5349-5156", "4091-6935", "1581-1690", "1652-5578"] },
    { id: "u-penpaper", name: "Penpaper", saIds: ["9812-1646", "8389-0206", "5805-5568"] },
    { id: "u-snorlax", name: "Snorlax", saIds: ["4590-9906", "8775-1559", "5532-5157"] },
  ],
  assignments: {},
  fees: [
    { id: "acct", label: "Accountant", pct: 5, recipient: "external", paidBy: "split" },
    { id: "punterx", label: "punterx (open-sitting)", pct: 5, recipient: "external", paidBy: "split" },
    { id: "gruzzy", label: "gruzzy", pct: 1, recipient: "jon", paidBy: "split" },
  ],
  feeBase: "net",
  finalizedPeriods: {},
};

const fmt = (n, d = 2) =>
  (n < 0 ? "-" : "") + Math.abs(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtI = (n) => n.toLocaleString("en-US");
const money = (n) => (
  <span style={{ color: n > 0.005 ? C.green : n < -0.005 ? C.red : C.ink, fontWeight: 600 }}>{fmt(n)}</span>
);

// ———————————————— Parsing ————————————————
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const lower = wb.SheetNames.map((s) => s.toLowerCase());
  const clubIdx = lower.findIndex((s) => s.includes("club overview"));
  const unionIdx = lower.findIndex((s) => s.includes("union member statistics"));
  const sheetName = clubIdx >= 0 ? wb.SheetNames[clubIdx] : unionIdx >= 0 ? wb.SheetNames[unionIdx] : wb.SheetNames[0];
  const isUnion = clubIdx < 0 && unionIdx >= 0;
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });

  // Locate the header row. Club exports have "No." in col 0; union exports
  // prepend a "Club" column, putting "No." in col 1.
  let headerIdx = -1, off = 0, period = "";
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const c0 = String(rows[i]?.[0] ?? "");
    if (c0.startsWith("Period")) period = c0.replace("Period :", "").trim();
    if (c0 === "No.") { headerIdx = i; off = 0; break; }
    if (String(rows[i]?.[1] ?? "") === "No.") { headerIdx = i; off = 1; break; }
  }
  if (headerIdx < 0) {
    throw new Error(
      `Couldn't find the 'No.' header row on the "${sheetName}" sheet. ` +
      "Expected a weekly Club Overview export or a union export with a Union Member Statistics sheet."
    );
  }

  // Column map. Club export: fixed layout. Union export: same identity columns
  // shifted by 1, but P&L / fee / hands live in "Total" columns we find by
  // scanning the header band (section label row + the group row beneath it).
  let cHands = 10, cFee = 11, cPnl = 19, dataStart = headerIdx + 2;
  if (isUnion) {
    dataStart = headerIdx + 3; // 3-row header band (section / group / sub-label)
    const top = rows[headerIdx] || [], grp = rows[headerIdx + 1] || [];
    const sec = (label) => top.findIndex((v) => String(v ?? "").trim() === label);
    const pnlStart = sec("Player P&L"), feeStart = sec("Rake&Fee"), handsStart = sec("Played Hands");
    const totals = [];
    grp.forEach((v, j) => { if (String(v ?? "").trim() === "Total") totals.push(j); });
    const totalFor = (start) => {
      if (start < 0) return -1;
      // the Total belonging to a section is the first one at/after its start,
      // before the next section begins
      const cand = totals.filter((j) => j >= start);
      return cand.length ? cand[0] : -1;
    };
    // sections appear in order P&L → Rake&Fee → … → Played Hands, so consume
    // totals section by section
    const pick = (start, used) => {
      if (start < 0) return -1;
      const j = totals.find((t) => t >= start && !used.has(t));
      if (j != null) used.add(j);
      return j ?? -1;
    };
    const used = new Set();
    cPnl = pick(pnlStart, used);
    cFee = pick(feeStart, used);
    cHands = totals.length ? totals[totals.length - 1] : -1; // Played Hands Total is the last one
    if (cPnl < 0 || cFee < 0) {
      throw new Error("Union export found, but couldn't locate the P&L / Rake&Fee Total columns in the header.");
    }
  }

  const players = [];
  for (let i = dataStart; i < rows.length; i++) {
    const r = rows[i];
    // Skip blanks, sub-header rows, and summary rows (e.g. trailing TOTAL):
    // row number must be numeric and the member ID must be present.
    if (r?.[off] == null || !Number.isFinite(parseFloat(r[off]))) continue;
    const memberId = String(r[off + 7] ?? "").trim();
    const name = String(r[off + 8] ?? "").trim();
    if (!memberId || memberId === "-" || !name) continue;
    const num = (v) => { const x = parseFloat(v); return isNaN(x) ? 0 : x; };
    const p = {
      saId: String(r[off + 1] ?? "-").trim(), saName: String(r[off + 2] ?? "-").trim(),
      agentId: String(r[off + 3] ?? "-").trim(), agentName: String(r[off + 4] ?? "-").trim(),
      role: String(r[off + 6] ?? "").trim(), memberId, name,
      hands: cHands >= 0 ? num(r[cHands]) : 0, fee: num(r[cFee]), pnl: num(r[cPnl]),
    };
    if (p.hands > 0 || p.fee !== 0 || p.pnl !== 0) players.push(p);
  }
  if (!players.length) throw new Error("Parsed the sheet but found no active players (rows with hands, fee, or P&L).");

  // Bad beat jackpot: scan every sheet for a "bad beat" line item and take the
  // first number to the right of the label (summing if it appears per club).
  let jackpot = 0, jackpotFound = false;
  wb.SheetNames.forEach((sn) => {
    let rws;
    try { rws = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null }); } catch (e) { return; }
    rws.forEach((r) => {
      if (!Array.isArray(r)) return;
      const li = r.findIndex((v) => typeof v === "string" && /bad\s*beat/i.test(v));
      if (li < 0) return;
      for (let j = li + 1; j < r.length; j++) {
        const n = parseFloat(String(r[j] ?? "").replace(/,/g, ""));
        if (isFinite(n) && n !== 0) { jackpot += n; jackpotFound = true; break; }
      }
    });
  });
  return { players, period, jackpot, jackpotFound };
}

// ———————————————— Engine ————————————————
function buildModel(players, cfg, weekAdj, period) {
  // If this period was finalized, makeup math freezes at the balances it was
  // settled with (otherwise finalize would double-count the same week).
  const finEntry = period ? (cfg.finalizedPeriods || {})[period] : null;
  const snapMakeup = finEntry && typeof finEntry === "object" ? finEntry.snapshot || null : null;
  const ownMap = {};
  cfg.ownAccounts.ak.forEach((n) => (ownMap[n.trim().toLowerCase()] = "ak"));
  cfg.ownAccounts.jon.forEach((n) => (ownMap[n.trim().toLowerCase()] = "jon"));
  const backedMap = cfg.backed || {};

  const ownRows = [], backedRows = [], extRows = [];
  for (const p of players) {
    const lo = p.name.toLowerCase();
    if (ownMap[lo]) ownRows.push({ ...p, owner: ownMap[lo] });
    else if (backedMap[lo]) backedRows.push({ ...p, backedKey: lo });
    else extRows.push(p);
  }

  const splitTipback = (fee, adj) =>
    (Math.min(adj.amtA, fee) * adj.rateA) / 100 + (Math.max(0, fee - adj.amtA) * adj.rateB) / 100;

  // Owner accounts: 100% feeback.
  const own = ownRows.map((p) => ({ ...p, feeback: p.fee, position: p.pnl + p.fee }));

  // Normal external players (with optional action-buy tax/rebate on P&L)
  const ext = extRows.map((p) => {
    const tbPct = cfg.playerDeals[p.memberId] ?? (p.saId !== "-" ? cfg.saDeals[p.saId] : undefined) ?? cfg.defaultTB;
    const adj = weekAdj?.[`p:${p.memberId}`];
    const tipback = adj ? splitTipback(p.fee, adj) : (p.fee * tbPct) / 100;
    const at = (cfg.actionTax || {})[p.memberId];
    const net = p.pnl + tipback;
    const actionCut = at ? (net * at.pct) / 100 : 0; // pct% of net after rakeback: taxes wins, rebates losses
    return { ...p, tbPct, adjusted: !!adj, tipback, actionTaxPct: at ? at.pct : 0, actionBacker: at ? at.backer || "split" : null, actionCut, settlement: net - actionCut };
  });

  const saMap = new Map(); const individuals = [];
  for (const p of ext) {
    if (p.saId !== "-") {
      if (!saMap.has(p.saId)) saMap.set(p.saId, { key: `sa:${p.saId}`, type: "sa", id: p.saId, name: p.saName, members: [] });
      saMap.get(p.saId).members.push(p);
    } else {
      individuals.push({ key: `p:${p.memberId}`, type: "player", id: p.memberId, name: p.name, members: [p] });
    }
  }
  const agg = (e) => {
    const sum = (f) => e.members.reduce((a, m) => a + m[f], 0);
    return { ...e, hands: sum("hands"), pnl: sum("pnl"), fee: sum("fee"), tipback: sum("tipback"), settlement: sum("settlement") };
  };
  const saEntities = [...saMap.values()].map((e) => {
    let x = agg(e);
    const adj = weekAdj?.[e.key];
    if (adj) { const tb = splitTipback(x.fee, adj); x = { ...x, tipback: tb, settlement: x.pnl + tb, adjusted: true }; }
    return x;
  });
  const indEntities = individuals.map(agg).map((x) => ({ ...x, adjusted: x.members.some((m) => m.adjusted) }));

  // House-backed entities
  const backedEntities = backedRows.map((p) => {
    const b = backedMap[p.backedKey];
    const key = `b:${p.backedKey}`;
    const adj = weekAdj?.[key];
    if (b.deal === "action") {
      const rbTotal = adj ? splitTipback(p.fee, adj) : (p.fee * (b.rbPct ?? 100)) / 100;
      const gross = p.pnl + rbTotal;
      const backerShare = (gross * b.actionPct) / 100;
      const settlement = gross - backerShare;
      return {
        key, type: "backed", dealType: "action", id: p.memberId, name: p.name,
        members: [{ ...p, tbPct: b.rbPct ?? 100, tipback: rbTotal, settlement }],
        hands: p.hands, pnl: p.pnl, fee: p.fee, tipback: rbTotal, settlement,
        backer: b.backer, backerBook: backerShare, actionPct: b.actionPct, adjusted: !!adj,
      };
    }
    // Makeup deal: weekly net = P&L + RB credit.
    const entering = snapMakeup && snapMakeup[key.slice(2)] != null ? snapMakeup[key.slice(2)] : (b.makeup || 0);
    const inMakeup = entering > 0.005;
    const rb = inMakeup ? b.rbMakeup : b.rbNormal;
    const rbCredit = adj ? splitTipback(p.fee, adj) : (p.fee * rb) / 100;
    const net = p.pnl + rbCredit;
    const excess = Math.max(0, net - entering);
    const playerCash = (excess * (b.playerProfitPct ?? 50)) / 100;
    const makeupAfter = Math.max(0, entering - net);
    return {
      key, type: "backed", dealType: "makeup", id: p.memberId, name: p.name,
      members: [{ ...p, tbPct: rb, tipback: rbCredit, settlement: playerCash }],
      hands: p.hands, pnl: p.pnl, fee: p.fee, tipback: rbCredit, settlement: playerCash,
      backer: b.backer, backerBook: net - playerCash, net,
      rb, inMakeup, makeupBefore: entering, makeupAfter, adjusted: !!adj,
    };
  });

  // Umbrella merge
  const umbrellaOf = {}; (cfg.umbrellas || []).forEach((u) => u.saIds.forEach((id) => (umbrellaOf[id] = u)));
  const umbMap = new Map(); const looseSAs = [];
  for (const e of saEntities) {
    const u = umbrellaOf[e.id];
    if (u) {
      if (!umbMap.has(u.id)) umbMap.set(u.id, { key: `u:${u.id}`, type: "umbrella", id: u.id, name: u.name, subgroups: [], members: [] });
      const m = umbMap.get(u.id); m.subgroups.push(e); m.members.push(...e.members);
    } else looseSAs.push(e);
  }
  const umbEntities = [...umbMap.values()].map((e) => {
    const sum = (f) => e.subgroups.reduce((a, s) => a + s[f], 0);
    return { ...e, hands: sum("hands"), pnl: sum("pnl"), fee: sum("fee"), tipback: sum("tipback"), settlement: sum("settlement"), adjusted: e.subgroups.some((s) => s.adjusted) };
  });

  const entities = [...umbEntities, ...looseSAs, ...backedEntities, ...indEntities]
    .sort((a, b) => Math.abs(b.fee) - Math.abs(a.fee));

  // Club economics. For makeup players the RB credit leaves club profit and
  // moves onto the backer's book, so it counts as a tipback here.
  const clubRevenue = players.reduce((a, p) => a + p.fee, 0);
  const extTipbacks = ext.reduce((a, p) => a + p.tipback, 0);
  const backedRB = backedEntities.reduce((a, e) => a + e.tipback, 0);
  const ownFeeback = own.reduce((a, p) => a + p.feeback, 0);
  const tipbacksPaid = extTipbacks + backedRB + ownFeeback;
  const clubProfit = clubRevenue - tipbacksPaid;
  const feeRows = cfg.fees.map((f) => ({ ...f, amount: f.kind === "fixed" ? (f.amount || 0) : ((cfg.feeBase === "gross" ? clubRevenue : clubProfit) * f.pct) / 100 }));
  const totalFees = feeRows.reduce((a, f) => a + f.amount, 0);
  const netProfit = clubProfit - totalFees;

  const ownPosition = { ak: 0, jon: 0 };
  own.forEach((p) => (ownPosition[p.owner] += p.position));

  const backedBook = { ak: 0, jon: 0 };
  backedEntities.forEach((e) => {
    if (e.backer === "split") { backedBook.ak += e.backerBook / 2; backedBook.jon += e.backerBook / 2; }
    else backedBook[e.backer] += e.backerBook;
  });

  const taxBook = { ak: 0, jon: 0 };
  ext.forEach((p) => {
    if (!p.actionCut) return;
    if (p.actionBacker === "ak" || p.actionBacker === "jon") taxBook[p.actionBacker] += p.actionCut;
    else { taxBook.ak += p.actionCut / 2; taxBook.jon += p.actionCut / 2; }
  });

  const entitle = {
    ak: netProfit / 2 + ownPosition.ak + backedBook.ak + taxBook.ak,
    jon: netProfit / 2 + ownPosition.jon + backedBook.jon + taxBook.jon,
  };
  feeRows.forEach((f) => { if (f.recipient === "ak" || f.recipient === "jon") entitle[f.recipient] += f.amount; });

  const actual = { ak: 0, jon: 0 }; const unassigned = [];
  entities.forEach((e) => {
    const who = cfg.assignments[e.key];
    if (who === "ak" || who === "jon") actual[who] += -e.settlement;
    else unassigned.push(e);
  });
  feeRows.forEach((f) => {
    if (f.recipient === "external") {
      if (f.paidBy === "split") { actual.ak -= f.amount / 2; actual.jon -= f.amount / 2; }
      else actual[f.paidBy] -= f.amount;
    }
  });

  const akOwesJon = actual.ak - entitle.ak;
  const balanceOk = Math.abs(actual.ak + actual.jon - (entitle.ak + entitle.jon)) < 0.02 && unassigned.length === 0;

  const totals = {
    hands: entities.reduce((a, e) => a + e.hands, 0),
    pnl: entities.reduce((a, e) => a + e.pnl, 0),
    fee: entities.reduce((a, e) => a + e.fee, 0),
    tipback: entities.reduce((a, e) => a + e.tipback, 0),
    settlement: entities.reduce((a, e) => a + e.settlement, 0),
  };

  return { own, entities, backedEntities, saEntities, indEntities, umbEntities, looseSAs,
    clubRevenue, extTipbacks, backedRB, ownFeeback, tipbacksPaid, clubProfit, feeRows, totalFees, netProfit,
    ownPosition, backedBook, taxBook, entitle, actual, akOwesJon, unassigned, balanceOk, totals };
}

// ———————————————— Styled Excel export (ExcelJS) ————————————————
const n2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;
// Appends a player/agent/super-agent's export ID onto their name for every
// settlement Excel sheet — usernames can be reused or renamed, but the
// tracking system's own ID for that line never changes, so it's the
// foolproof way to tell two similarly-named players apart or re-match a
// player across weeks. No-op when there's no real ID to show (backed/manual
// entries, or an ID of "-").
const withId = (name, id) => (id && id !== "-" ? `${name} (${id})` : name);
const MONEY_FMT = "#,##0.00;(#,##0.00)";
const XLC = { gold: "FFB49A5E", bar: "FF332A22", cream: "FFF3EAD8", rowAlt: "FFFBF6EB", white: "FFFFFFFF", ink: "FF2B241C", green: "FF1E7B34", red: "FFC0392B", mute: "FF8A7E6C" };
const FBASE = { name: "Arial", size: 10, color: { argb: XLC.ink } };
const fillOf = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

function xTitle(ws, text) {
  const r = ws.addRow([text]);
  r.getCell(1).font = { ...FBASE, size: 13, bold: true };
  ws.addRow([]);
}
function xHeader(ws, labels, leftCols = 4) {
  const r = ws.addRow(labels);
  r.eachCell((c, i) => {
    c.font = { name: "Arial", size: 9, bold: true, color: { argb: XLC.white } };
    c.fill = fillOf(XLC.gold);
    c.alignment = { horizontal: i <= leftCols ? "left" : "right" };
  });
  return r.number;
}
function xMoney(cell, v, { bold = false, white = false, colorSign = true } = {}) {
  cell.value = n2(v);
  cell.numFmt = MONEY_FMT;
  const color = white ? XLC.white : colorSign && v > 0.005 ? XLC.green : colorSign && v < -0.005 ? XLC.red : XLC.ink;
  cell.font = { ...FBASE, bold, color: { argb: color } };
  cell.alignment = { horizontal: "right" };
}
function xNum(cell, v, fmt, { bold = false, white = false } = {}) {
  cell.value = v; cell.numFmt = fmt;
  cell.font = { ...FBASE, bold, color: { argb: white ? XLC.white : XLC.ink } };
  cell.alignment = { horizontal: "right" };
}
function xText(cell, v, { bold = false, white = false, mute = false } = {}) {
  cell.value = v;
  cell.font = { ...FBASE, bold, color: { argb: white ? XLC.white : mute ? XLC.mute : XLC.ink } };
  cell.alignment = { horizontal: "left" };
}
const DEAL_COLS = ["Player", "Device ID", "Super Agent", "Agent", "Hands", "Winnings", "Tips", "TB %", "Tipback", "Settlement"];
const DEAL_W = [20, 12, 17, 17, 9, 12, 12, 8, 12, 13];

function memberRow(ws, m, alt) {
  const r = ws.addRow([]);
  xText(r.getCell(1), m.name, { bold: true });
  xText(r.getCell(2), m.memberId, { mute: true });
  xText(r.getCell(3), withId(m.saName, m.saId), { mute: true });
  xText(r.getCell(4), withId(m.agentName, m.agentId), { mute: true });
  xNum(r.getCell(5), m.hands, "#,##0");
  xMoney(r.getCell(6), m.pnl, { colorSign: false });
  xMoney(r.getCell(7), m.fee, { colorSign: false });
  xNum(r.getCell(8), m.tbPct, '0.0"%"');
  xMoney(r.getCell(9), m.tipback, { colorSign: false });
  xMoney(r.getCell(10), m.settlement);
  if (alt) r.eachCell({ includeEmpty: true }, (c) => { if (!c.fill || !c.fill.fgColor) c.fill = fillOf(XLC.rowAlt); });
  return r;
}
function totalRow(ws, label, e, { dark = false } = {}) {
  const r = ws.addRow([]);
  xText(r.getCell(1), label, { bold: true, white: dark });
  xNum(r.getCell(5), e.hands, "#,##0", { bold: true, white: dark });
  xMoney(r.getCell(6), e.pnl, { bold: true, white: dark, colorSign: !dark });
  xMoney(r.getCell(7), e.fee, { bold: true, white: dark, colorSign: false });
  xNum(r.getCell(8), e.fee ? n2((e.tipback / e.fee) * 100) : 0, '0.0"%"', { bold: true, white: dark });
  xMoney(r.getCell(9), e.tipback, { bold: true, white: dark, colorSign: false });
  xMoney(r.getCell(10), e.settlement, { bold: true, white: dark, colorSign: !dark });
  r.eachCell({ includeEmpty: true }, (c) => { c.fill = fillOf(dark ? XLC.bar : XLC.cream); });
  for (let j = 1; j <= 10; j++) if (!r.getCell(j).value && r.getCell(j).value !== 0) r.getCell(j).fill = fillOf(dark ? XLC.bar : XLC.cream);
  return r;
}
function safeSheetName(base, wb) {
  let nm = String(base).replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 28) || "Sheet";
  let out = nm, i = 2;
  while (wb.worksheets.some((w) => w.name === out)) out = `${nm.slice(0, 25)} ${i++}`;
  return out;
}
function addDealSheet(wb, e, period) {
  const ws = wb.addWorksheet(safeSheetName(e.name, wb));
  DEAL_W.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `${e.name} — ${period || "this week"}`);
  const hr = xHeader(ws, DEAL_COLS);
  ws.views = [{ state: "frozen", ySplit: hr }];
  if (e.type === "umbrella") {
    e.subgroups.forEach((s) => {
      const sub = ws.addRow([]);
      xText(sub.getCell(1), `${withId(s.name, s.id)} — settlement ${fmt(s.settlement)}`, { bold: true });
      sub.eachCell({ includeEmpty: true }, (c) => (c.fill = fillOf(XLC.rowAlt)));
      for (let j = 1; j <= 10; j++) sub.getCell(j).fill = fillOf(XLC.rowAlt);
      [...s.members].sort((a, b) => b.fee - a.fee).forEach((m, i) => memberRow(ws, m, false));
    });
  } else {
    [...e.members].sort((a, b) => b.fee - a.fee).forEach((m, i) => memberRow(ws, m, i % 2 === 1));
  }
  totalRow(ws, "TOTAL", e);
  return ws;
}
// When _wbCapture is set, saveWb parks the workbook bytes there instead of
// downloading — used by the archive to keep a copy of the generated Excel.
let _wbCapture = null;
async function saveWb(wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  if (_wbCapture) { _wbCapture.push({ filename, buf }); return; }
  downloadBytes(buf, filename);
}
function downloadBytes(buf, filename, type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
  const url = URL.createObjectURL(new Blob([buf], { type }));
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ———————————————— Archive (workbooks live encrypted in Supabase Storage; see sync.js) ————————————————
const rawKey = (site, period) => `raw|${site}|${period}`;
const genKey = (site, period) => `gen|${site}|${period}`;
const ARCHIVE_INDEX_KEY = "archive-index-v1";
async function loadArchiveIndex() { try { const c = await store.get(ARCHIVE_INDEX_KEY); if (c?.value) return JSON.parse(c.value); } catch (e) {} return { weeks: [] }; }
async function saveArchiveIndex(idx) { try { await store.set(ARCHIVE_INDEX_KEY, JSON.stringify(idx)); } catch (e) {} }
// Store this week's finalized files: the raw upload (already parked under rawKey
// at upload time) plus the workbook(s) the app generates. Called when a new week
// replaces the old one, or from the "Archive this week" button.
async function archiveWeek(site, siteName, period, generate) {
  if (!period) return;
  _wbCapture = [];
  try { await generate(); } catch (e) {}
  const files = _wbCapture; _wbCapture = null;
  const idx = await loadArchiveIndex();
  const raw = await idbGet(rawKey(site, period)).catch(() => null);
  if (files.length) await idbPut(genKey(site, period), { site, siteName, period, files: files.map((f) => ({ name: f.filename, buf: f.buf })) });
  const rec = { site, siteName, period, archivedAt: new Date().toISOString(), hasRaw: !!raw, rawName: raw?.name || "", gen: files.map((f) => f.filename) };
  idx.weeks = [...idx.weeks.filter((w) => !(w.site === site && w.period === period)), rec];
  await saveArchiveIndex(idx);
  return rec;
}
async function deleteArchivedWeek(site, period) {
  await idbDel(rawKey(site, period)).catch(() => {});
  await idbDel(genKey(site, period)).catch(() => {});
  const idx = await loadArchiveIndex();
  idx.weeks = idx.weeks.filter((w) => !(w.site === site && w.period === period));
  await saveArchiveIndex(idx);
}

function ArchiveView({ clubs }) {
  const [idx, setIdx] = useState(null);
  const reload = async () => setIdx(await loadArchiveIndex());
  useEffect(() => { reload(); }, []);
  if (!idx) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;
  const dlRaw = async (w) => { const r = await idbGet(rawKey(w.site, w.period)); if (r?.buf) downloadBytes(r.buf, r.name || `${w.siteName}_${w.period}_raw.xlsx`); else window.alert("Raw file not stored for this week (uploaded before archiving existed)."); };
  const dlGen = async (w, i) => { const g = await idbGet(genKey(w.site, w.period)); const f = g?.files?.[i]; if (f) downloadBytes(f.buf, f.name); else window.alert("File not found."); };
  const dlAllGen = async (w) => { const g = await idbGet(genKey(w.site, w.period)); (g?.files || []).forEach((f) => downloadBytes(f.buf, f.name)); };
  const sites = [...new Set(idx.weeks.map((w) => w.site))];
  const siteLabel = (w) => (w.site === "fishtank" ? "Fish Tank" : (clubs.find((c) => c.id === w.site)?.name || w.siteName || w.site));
  const sortKey = (p) => (p || "").replace(/[^\d]/g, "");
  return (
    <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
      <div style={{ fontFamily: "Georgia, serif", fontSize: 19, marginBottom: 4 }}>Archive — finalized weeks</div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 16 }}>
        Each time a new weekly export is uploaded to Fish Tank or an owner club, the week it replaces is stored here: the <b>raw upload</b> and the app's <b>generated settlement workbook</b> (built with the deals in place at that moment). Use "Archive this week" on a site to store the current week without waiting for Monday.
      </div>
      {idx.weeks.length === 0 && <Card title="Nothing archived yet"><div style={{ color: C.mute, fontSize: 13 }}>Upload next week's export and this week's files will land here automatically.</div></Card>}
      {sites.map((site) => {
        const ws = idx.weeks.filter((w) => w.site === site).sort((a, b) => sortKey(b.period).localeCompare(sortKey(a.period)));
        return (
          <div key={site} style={{ marginBottom: 16 }}>
            <Card title={`${siteLabel(ws[0])} · ${ws.length} week${ws.length !== 1 ? "s" : ""}`}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr style={{ background: C.cream }}>
                  <th style={{ ...th, textAlign: "left" }}>Week</th><th style={{ ...th, textAlign: "left" }}>Archived</th>
                  <th style={{ ...th, textAlign: "left" }}>Raw upload</th><th style={{ ...th, textAlign: "left" }}>Generated files</th><th style={th}></th>
                </tr></thead>
                <tbody>
                  {ws.map((w, i) => (
                    <tr key={w.period} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                      <td style={{ ...tdL, fontWeight: 700 }}>{w.period}</td>
                      <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{(w.archivedAt || "").slice(0, 10)}</td>
                      <td style={tdL}>{w.hasRaw ? <Btn tone="ghost" small onClick={() => dlRaw(w)}>{w.rawName || "raw .xlsx"}</Btn> : <span style={{ color: C.mute, fontSize: 12 }}>not stored</span>}</td>
                      <td style={tdL}>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {(w.gen || []).map((n, j) => <Btn key={n + j} tone="gold" small onClick={() => dlGen(w, j)}>{n}</Btn>)}
                          {(w.gen || []).length > 1 && <Btn tone="ghost" small onClick={() => dlAllGen(w)}>all</Btn>}
                          {(w.gen || []).length === 0 && <span style={{ color: C.mute, fontSize: 12 }}>none</span>}
                        </div>
                      </td>
                      <td style={{ ...td, width: 40 }}><button onClick={async () => { if (window.confirm(`Delete archived week ${w.period}?`)) { await deleteArchivedWeek(w.site, w.period); reload(); } }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>
          </div>
        );
      })}
    </div>
  );
}
const periodSlug = (period) => (period || "week").replace(/[^\d]/g, "_").replace(/^_+|_+$/g, "") || "week";

async function downloadDealExcel(e, period) {
  const wb = new ExcelJS.Workbook();
  addDealSheet(wb, e, period);
  await saveWb(wb, `${e.name.replace(/[^\w]+/g, "_")}_${periodSlug(period)}.xlsx`);
}

async function downloadWorkbook(model, period, cfg) {
  const wb = new ExcelJS.Workbook();
  // — Settlements summary
  const ws = wb.addWorksheet("Settlements");
  [24, 14, 9, 13, 13, 9, 13, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `All In Fish Tank — Settlements ${period || ""}`.trim());
  const hr = xHeader(ws, ["Deal", "Type", "Hands", "Winnings", "Tips", "Avg TB %", "Tipback", "Settlement"], 2);
  ws.views = [{ state: "frozen", ySplit: hr }];
  model.entities.forEach((e, i) => {
    const r = ws.addRow([]);
    xText(r.getCell(1), e.type === "umbrella" ? e.name : withId(e.name, e.id), { bold: true });
    xText(r.getCell(2), e.type === "umbrella" ? "Umbrella" : e.type === "sa" ? "Super Agent" : e.type === "backed" ? (e.dealType === "action" ? "Action buy" : "Makeup deal") : "Player", { mute: true });
    xNum(r.getCell(3), e.hands, "#,##0");
    xMoney(r.getCell(4), e.pnl, { colorSign: false });
    xMoney(r.getCell(5), e.fee, { colorSign: false });
    xNum(r.getCell(6), e.fee ? n2((e.tipback / e.fee) * 100) : 0, '0.0"%"');
    xMoney(r.getCell(7), e.tipback, { colorSign: false });
    xMoney(r.getCell(8), e.settlement);
    if (i % 2 === 1) for (let j = 1; j <= 8; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const t = model.totals;
  const gr = ws.addRow([]);
  xText(gr.getCell(1), "GRAND TOTAL", { bold: true, white: true });
  xNum(gr.getCell(3), t.hands, "#,##0", { bold: true, white: true });
  xMoney(gr.getCell(4), t.pnl, { bold: true, white: true, colorSign: false });
  xMoney(gr.getCell(5), t.fee, { bold: true, white: true, colorSign: false });
  xNum(gr.getCell(6), t.fee ? n2((t.tipback / t.fee) * 100) : 0, '0.0"%"', { bold: true, white: true });
  xMoney(gr.getCell(7), t.tipback, { bold: true, white: true, colorSign: false });
  xMoney(gr.getCell(8), t.settlement, { bold: true, white: true, colorSign: false });
  for (let j = 1; j <= 8; j++) gr.getCell(j).fill = fillOf(XLC.bar);

  // — one sheet per umbrella / loose SA
  [...model.umbEntities, ...model.looseSAs].forEach((e) => addDealSheet(wb, e, period));

  // — Individuals
  if (model.indEntities.length) {
    const wi = wb.addWorksheet("Individuals");
    DEAL_W.forEach((w, i) => (wi.getColumn(i + 1).width = w));
    xTitle(wi, `Individual players (no super agent) — ${period || "this week"}`);
    xHeader(wi, DEAL_COLS);
    model.indEntities.forEach((e, i) => memberRow(wi, { ...e.members[0], saName: "-" }, i % 2 === 1));
  }

  // — House-backed
  const bMk = model.backedEntities.filter((e) => e.dealType === "makeup");
  const bAc = model.backedEntities.filter((e) => e.dealType === "action");
  if (bMk.length || bAc.length) {
    const wbk = wb.addWorksheet("House-Backed");
    [17, 15, 12, 8, 9, 12, 12, 12, 12, 13, 13, 13].forEach((w, i) => (wbk.getColumn(i + 1).width = w));
    xTitle(wbk, `House-backed players — ${period || "this week"}`);
    if (bMk.length) {
      const note = wbk.addRow(["Makeup deals — net = P&L + RB credit; settlement = player % of net above makeup; shortfall accrues to makeup."]);
      note.getCell(1).font = { ...FBASE, size: 9, color: { argb: XLC.mute } };
      xHeader(wbk, ["Player", "Backer", "Makeup in", "RB %", "Hands", "P&L", "Tips", "RB credit", "Net", "Settlement", "Backer book", "Makeup after"], 2);
      bMk.forEach((e, i) => {
        const r = wbk.addRow([]);
        xText(r.getCell(1), withId(e.name, e.id), { bold: true });
        xText(r.getCell(2), e.backer === "split" ? "Ak & Jon 50/50" : e.backer === "ak" ? "Ak" : "Jon", { mute: true });
        xMoney(r.getCell(3), e.makeupBefore, { colorSign: false });
        xNum(r.getCell(4), e.rb, '0.0"%"');
        xNum(r.getCell(5), e.hands, "#,##0");
        xMoney(r.getCell(6), e.pnl);
        xMoney(r.getCell(7), e.fee, { colorSign: false });
        xMoney(r.getCell(8), e.tipback, { colorSign: false });
        xMoney(r.getCell(9), e.net);
        xMoney(r.getCell(10), e.settlement, { colorSign: false });
        xMoney(r.getCell(11), e.backerBook);
        xMoney(r.getCell(12), e.makeupAfter, { colorSign: false });
        if (i % 2 === 1) for (let j = 1; j <= 12; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
      });
      wbk.addRow([]);
    }
    if (bAc.length) {
      const note = wbk.addRow(["Action buys — backer owns their % of (P&L + rakeback); player settles the rest."]);
      note.getCell(1).font = { ...FBASE, size: 9, color: { argb: XLC.mute } };
      xHeader(wbk, ["Player", "Backer", "Backer %", "RB %", "Hands", "P&L", "Tips", "Player settlement", "Backer book"], 2);
      bAc.forEach((e, i) => {
        const r = wbk.addRow([]);
        xText(r.getCell(1), withId(e.name, e.id), { bold: true });
        xText(r.getCell(2), e.backer === "ak" ? "Ak" : "Jon", { mute: true });
        xNum(r.getCell(3), e.actionPct, '0.0"%"');
        xNum(r.getCell(4), e.members[0].tbPct, '0.0"%"');
        xNum(r.getCell(5), e.hands, "#,##0");
        xMoney(r.getCell(6), e.pnl);
        xMoney(r.getCell(7), e.fee, { colorSign: false });
        xMoney(r.getCell(8), e.settlement);
        xMoney(r.getCell(9), e.backerBook);
        if (i % 2 === 1) for (let j = 1; j <= 9; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
      });
    }
  }

  // — Ak / Jon
  const wa = wb.addWorksheet("Ak-Jon");
  wa.getColumn(1).width = 44; wa.getColumn(2).width = 15;
  xTitle(wa, `Ak / Jon reconciliation — ${period || "this week"}`);
  const m = model;
  const line = (label, v, { bold = false, sign = true } = {}) => {
    const r = wa.addRow([]);
    xText(r.getCell(1), label, { bold, mute: !bold });
    xMoney(r.getCell(2), v, { bold, colorSign: sign });
  };
  line("Total tips collected (all accounts)", m.clubRevenue, { sign: false });
  line("Tipbacks to agents & players", -m.extTipbacks, { sign: false });
  line("Backed players' RB (cash + credits)", -m.backedRB, { sign: false });
  line("Owner accounts' 100% feeback", -m.ownFeeback, { sign: false });
  line("Club profit", m.clubProfit, { bold: true });
  m.feeRows.forEach((f) => line(`${f.label} — ${f.kind === "fixed" ? "fixed" : f.pct + "% of profit"}` + (f.recipient !== "external" ? ` → ${f.recipient === "ak" ? "Ak" : "Jon"}` : ""), -f.amount, { sign: false }));
  line("Net profit to split", m.netProfit, { bold: true });
  line("Each owner's half", m.netProfit / 2, { sign: false });
  wa.addRow([]);
  line("Ak — own accounts (P&L + 100% feeback)", m.ownPosition.ak);
  line("Ak — backed books", m.backedBook.ak);
  line("Ak — action-buy tax book", m.taxBook.ak);
  line("Ak — half of net profit", m.netProfit / 2, { sign: false });
  line("AK ENTITLEMENT", m.entitle.ak, { bold: true });
  wa.addRow([]);
  line("Jon — own accounts (P&L + 100% feeback)", m.ownPosition.jon);
  line("Jon — backed books", m.backedBook.jon);
  line("Jon — action-buy tax book", m.taxBook.jon);
  line("Jon — half of net profit", m.netProfit / 2, { sign: false });
  m.feeRows.filter((f) => f.recipient === "jon").forEach((f) => line(`Jon — ${f.label} fee`, f.amount, { sign: false }));
  line("JON ENTITLEMENT", m.entitle.jon, { bold: true });

  // ——— who settles what, per owner, plus the transfer ———
  const assignments = cfg?.assignments || {};
  [["ak", "Ak"], ["jon", "Jon"]].forEach(([w, W]) => {
    wa.addRow([]);
    const hr = wa.addRow([`${W} settles these deals`, "Settlement"]);
    hr.eachCell((c2, i) => { c2.font = { name: "Arial", size: 10, bold: true, color: { argb: XLC.white } }; c2.fill = fillOf(XLC.gold); c2.alignment = { horizontal: i === 1 ? "left" : "right" }; });
    m.entities.filter((e) => assignments[e.key] === w).forEach((e, i) => {
      const r = wa.addRow([]);
      xText(r.getCell(1), `${e.type === "umbrella" ? e.name : withId(e.name, e.id)} — ${e.settlement > 0.005 ? `${W} pays them` : e.settlement < -0.005 ? `they pay ${W}` : "even"}`, {});
      xMoney(r.getCell(2), e.settlement);
      if (i % 2 === 1) for (let j = 1; j <= 2; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
    });
    m.feeRows.filter((f) => f.recipient === "external").forEach((f) => {
      const share = f.paidBy === "split" ? f.amount / 2 : f.paidBy === w ? f.amount : 0;
      if (share > 0.005) { const r = wa.addRow([]); xText(r.getCell(1), `${f.label} (paid out)`, { mute: true }); xMoney(r.getCell(2), -share); }
    });
    const s1 = wa.addRow([]); xText(s1.getCell(1), `${W} — actual cash after settling`, { bold: true }); xMoney(s1.getCell(2), m.actual[w], { bold: true });
    for (let j = 1; j <= 2; j++) s1.getCell(j).fill = fillOf(XLC.cream);
    const s2 = wa.addRow([]); xText(s2.getCell(1), `${W} — entitlement`, {}); xMoney(s2.getCell(2), m.entitle[w]);
  });
  wa.addRow([]);
  const tr = wa.addRow([]);
  const owe = m.akOwesJon;
  xText(tr.getCell(1), Math.abs(owe) < 0.005 ? "PERFECTLY EVEN — NO TRANSFER" : owe > 0 ? "AK PAYS JON" : "JON PAYS AK", { bold: true, white: true });
  xMoney(tr.getCell(2), Math.abs(owe), { bold: true, white: true, colorSign: false });
  for (let j = 1; j <= 2; j++) tr.getCell(j).fill = fillOf(XLC.bar);

  await saveWb(wb, `FishTank_Settlements_${periodSlug(period)}.xlsx`);
}

// ———————————————— Export (clipboard — downloads are blocked in this sandbox) ————————————————
const toTSV = (header, rows) => [header, ...rows].map((r) => r.map((v) => String(v ?? "")).join("\t")).join("\n");

function ExportModal({ data, onClose }) {
  const [copied, setCopied] = useState(false);
  const taRef = useRef(null);
  if (!data) return null;
  const copy = async () => {
    try { await navigator.clipboard.writeText(data.text); setCopied(true); }
    catch (e) {
      taRef.current?.select();
      try { document.execCommand("copy"); setCopied(true); } catch (e2) {}
    }
    setTimeout(() => setCopied(false), 1800);
  };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.card, borderRadius: 12, padding: "20px 22px", width: "min(760px, 94vw)", maxHeight: "84vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px rgba(0,0,0,0.35)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
          <div style={{ fontFamily: "Georgia, serif", fontSize: 17, color: C.ink }}>{data.title}</div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <Btn tone="gold" small onClick={copy}>{copied ? "✓ Copied" : "Copy to clipboard"}</Btn>
            <Btn tone="ghost" small onClick={onClose}>Close</Btn>
          </div>
        </div>
        <div style={{ color: C.mute, fontSize: 12, marginBottom: 8 }}>
          Tab-separated — paste straight into Excel or Google Sheets and it lands in columns. (File downloads are blocked inside this app's sandbox.)
        </div>
        <textarea ref={taRef} readOnly value={data.text} onFocus={(e) => e.target.select()}
          style={{ ...inputS, width: "100%", boxSizing: "border-box", flex: 1, minHeight: 260, fontFamily: "ui-monospace, Menlo, monospace", fontSize: 11.5, whiteSpace: "pre", resize: "vertical" }} />
      </div>
    </div>
  );
}

// ———————————————— UI atoms ————————————————
const th = { fontSize: 10.5, letterSpacing: "0.09em", textTransform: "uppercase", color: C.goldDark, fontWeight: 700, padding: "8px 10px", textAlign: "right", whiteSpace: "nowrap" };
const td = { padding: "7px 10px", textAlign: "right", fontVariantNumeric: "tabular-nums", fontSize: 13.5, whiteSpace: "nowrap", color: C.ink };
const tdL = { ...td, textAlign: "left" };
const inputS = { padding: "5px 8px", border: `1px solid ${C.line}`, borderRadius: 5, fontSize: 13, color: C.ink, background: C.surface };

function Pill({ children, tone = "gold" }) {
  const bg = { red: "var(--pillRedBg)", green: "var(--pillGreenBg)", blue: "var(--pillBlueBg)", gold: "var(--pillGoldBg)" }[tone];
  const fg = { red: C.red, green: C.green, blue: "var(--pillBlueFg)", gold: C.goldDark }[tone];
  return <span style={{ background: bg, color: fg, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700, letterSpacing: "0.04em", whiteSpace: "nowrap" }}>{children}</span>;
}
const typePill = (e) =>
  e.type === "umbrella" ? <Pill tone="blue">umbrella · {e.subgroups.length} SAs</Pill>
  : e.type === "sa" ? <Pill tone="gold">SA · {e.members.length}</Pill>
  // All American-only entity types: an agent with no super agent above it,
  // the union roles that sit outside the SA/agent/player hierarchy, and a
  // DL umbrella folding several lines that are all settled by one person.
  : e.type === "agent" ? <Pill tone="gold">Agent · {e.members.length}</Pill>
  : e.type === "manager" ? <Pill tone="blue">Manager</Pill>
  : e.type === "master" ? <Pill tone="blue">Master</Pill>
  : e.type === "dlUmbrella" ? <Pill tone="blue">DL umbrella · {e.subgroups.length}</Pill>
  : e.type === "backed" ? <Pill tone="red">{e.dealType === "action" ? "action buy" : "house-backed"}</Pill>
  : <Pill tone="green">player</Pill>;

function PctInput({ value, onChange, width = 64, max = 100 }) {
  const [v, setV] = useState(value ?? "");
  useEffect(() => setV(value ?? ""), [value]);
  return (
    <span style={{ whiteSpace: "nowrap" }}>
      <input value={v} onChange={(e) => setV(e.target.value)}
        onBlur={() => { const n = parseFloat(v); onChange(isNaN(n) ? null : Math.max(0, Math.min(max, n))); }}
        onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
        style={{ ...inputS, width, textAlign: "right", fontVariantNumeric: "tabular-nums", border: `1px solid ${C.gold}` }} />
      <span style={{ marginLeft: 4, color: C.mute, fontSize: 12 }}>%</span>
    </span>
  );
}
function NumInput({ value, onChange, width = 90, disabled }) {
  const [v, setV] = useState(value ?? "");
  useEffect(() => setV(value ?? ""), [value]);
  return (
    <input value={v} onChange={(e) => setV(e.target.value)} disabled={disabled}
      onBlur={() => { const n = parseFloat(String(v).replace(/,/g, "")); onChange(isNaN(n) ? 0 : n); }}
      onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
      style={{ ...inputS, width, textAlign: "right", fontVariantNumeric: "tabular-nums", border: `1px solid ${C.gold}`, opacity: disabled ? 0.45 : 1, cursor: disabled ? "not-allowed" : "text" }} />
  );
}
function Btn({ children, onClick, tone = "dark", small, disabled }) {
  const s = { dark: { background: "var(--bar)", color: "var(--barText)" }, gold: { background: C.gold, color: "var(--onGold)" }, ghost: { background: "transparent", color: C.goldDark, border: `1px solid ${C.gold}` } }[tone];
  return (
    <button onClick={onClick} disabled={disabled} style={{ ...s, border: s.border || "none", borderRadius: 6, padding: small ? "5px 12px" : "9px 18px", fontSize: small ? 12 : 13.5, fontWeight: 700, letterSpacing: "0.03em", cursor: disabled ? "default" : "pointer", opacity: disabled ? 0.45 : 1 }}>
      {children}
    </button>
  );
}
const Card = ({ title, children, right }) => (
  <div style={{ background: C.card, borderRadius: 10, padding: "16px 20px", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
    <div style={{ display: "flex", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
      <div style={{ fontFamily: "Georgia, serif", fontSize: 16, color: C.ink }}>{title}</div>
      {right && <div style={{ marginLeft: "auto" }}>{right}</div>}
    </div>
    {children}
  </div>
);

// ———————————————— Main ————————————————
export default function App() {
  const [cfg, setCfg] = useState(DEFAULT_CONFIG);
  const [players, setPlayers] = useState(null);
  const [period, setPeriod] = useState("");
  const [weekAdj, setWeekAdj] = useState({});
  const [tab, setTab] = useState("settle");
  const [err, setErr] = useState("");
  const [saveNote, setSaveNote] = useState("");
  const [expanded, setExpanded] = useState({});
  const [loaded, setLoaded] = useState(false);
  const [clubs, setClubs] = useState([]);
  const fileRef = useRef(null);
  const seedRef = useRef(null);
  const saveClubs = async (next) => { setClubs(next); await saveOwnerClubs(next); };
  const addClub = async () => {
    const name = window.prompt("Name of the new owner club (e.g. Bazaar):");
    if (!name || !name.trim()) return;
    const c = newOwnerClub(name.trim());
    await saveClubs([...clubs, c]);
    up({ mode: "oc:" + c.id });
  };
  const deleteClub = async (id) => {
    const c = clubs.find((x) => x.id === id);
    if (!c) return;
    try { await store.del(c.cfgKey); await store.del(c.weekKey); } catch (e) {}
    await saveClubs(clubs.filter((x) => x.id !== id));
    up({ mode: "tabs" });
  };
  const onSeedFile = async (f) => {
    try { await importSnapshot(f); window.location.reload(); }
    catch (e) { setErr("Couldn't import data file: " + (e.message || e)); }
  };

  useEffect(() => {
    (async () => {
      await applySeedOnce();
      try { setClubs(await loadOwnerClubs()); } catch (e) {}
      try {
        const c = await store.get("fishtank-config-v4");
        if (c?.value) {
          const s = JSON.parse(c.value);
          if (!s.themeV2) { s.theme = "dark"; s.themeV2 = true; }
          setCfg({ ...DEFAULT_CONFIG, ...s,
            ownAccounts: s.ownAccounts || DEFAULT_CONFIG.ownAccounts,
            fees: s.fees || DEFAULT_CONFIG.fees,
            backed: s.backed || DEFAULT_CONFIG.backed,
            umbrellas: s.umbrellas || DEFAULT_CONFIG.umbrellas });
        }
      } catch (e) {}
      try {
        const d = await store.get("fishtank-lastweek-v4");
        if (d?.value) { const s = JSON.parse(d.value); setPlayers(s.players); setPeriod(s.period); setWeekAdj(s.weekAdj || {}); }
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  const persistCfg = useCallback(async (next) => {
    setCfg(next);
    try { await store.set("fishtank-config-v4", JSON.stringify(next)); setSaveNote(""); }
    catch (e) { setSaveNote("Change couldn't be saved — it still applies this session."); }
  }, []);
  const up = (patch) => persistCfg({ ...cfg, ...patch });

  const persistWeek = async (p, per, adj) => {
    try { await store.set("fishtank-lastweek-v4", JSON.stringify({ players: p, period: per, weekAdj: adj })); } catch (e) {}
  };
  const setAdj = (adj) => { setWeekAdj(adj); persistWeek(players, period, adj); };

  const onFile = async (file) => {
    setErr("");
    try {
      const buf = await file.arrayBuffer();
      const { players: p, period: per } = parseWorkbook(buf);
      if (players && period && period !== per && model) {
        try { await archiveWeek("fishtank", "Fish Tank", period, async () => { await downloadWorkbook(model, period, cfg); }); } catch (e) {}
      }
      const names = { ...cfg.names };
      p.forEach((x) => { names[x.memberId] = x.name; if (x.saId !== "-") names[x.saId] = x.saName; });
      await persistCfg({ ...cfg, names });
      setPlayers(p); setPeriod(per); setWeekAdj({}); setTab("settle");
      persistWeek(p, per, {});
      try { await idbPut(rawKey("fishtank", per), { site: "fishtank", siteName: "Fish Tank", period: per, name: file.name, buf }); } catch (e) {}
    } catch (e) { setErr(e.message || String(e)); }
  };
  const archiveNow = async () => {
    if (!model || !period) return;
    await archiveWeek("fishtank", "Fish Tank", period, async () => { await downloadWorkbook(model, period, cfg); });
    setSaveNote(`Archived ${period} — see the Archive tab.`);
  };

  const model = useMemo(() => (players ? buildModel(players, cfg, weekAdj, period) : null), [players, cfg, weekAdj, period]);

  const needsSetup = useMemo(() => {
    if (!model) return { deals: [], assigns: [] };
    const deals = [
      ...model.looseSAs.filter((e) => !cfg.confirmedSAs[e.id]),
      ...model.umbEntities.flatMap((u) => u.subgroups.filter((s) => !cfg.confirmedSAs[s.id])),
      ...model.indEntities.filter((e) => !cfg.confirmedPlayers[e.id]),
    ];
    return { deals, assigns: model.unassigned };
  }, [model, cfg]);

  const finalized = cfg.finalizedPeriods?.[period];
  // "Lock" = the old one-way Finalize, but reversible: locking snapshots each
  // backed makeup player's entering balance and rolls it forward (unless it's
  // tracked on a unified staking deal, which stays frozen either way);
  // unlocking puts that entering balance back so edits recompute cleanly,
  // and a re-lock re-snapshots off whatever the (possibly edited) numbers now are.
  const lockWeek = () => {
    const backed = { ...cfg.backed };
    const snapshot = {};
    model.backedEntities.forEach((e) => {
      if (e.dealType !== "makeup") return;
      const k = e.key.slice(2);
      if (backed[k]?.unifiedDealId) return; // makeup tracked on the shared staking deal instead — leave local balance frozen
      snapshot[k] = Math.round(e.makeupBefore * 100) / 100;
      if (backed[k]) backed[k] = { ...backed[k], makeup: Math.round(e.makeupAfter * 100) / 100 };
    });
    up({ backed, finalizedPeriods: { ...cfg.finalizedPeriods, [period]: { snapshot } } });
  };
  const unlockWeek = () => {
    const snap = cfg.finalizedPeriods?.[period];
    if (!snap) return;
    const backed = { ...cfg.backed };
    Object.entries(snap.snapshot || {}).forEach(([k, entering]) => { if (backed[k]) backed[k] = { ...backed[k], makeup: entering }; });
    const fp = { ...cfg.finalizedPeriods };
    delete fp[period];
    up({ backed, finalizedPeriods: fp });
  };

  if (!loaded) return <div style={{ fontFamily: "Georgia, serif", padding: 40, color: "#8A7E6C" }}>Loading saved setup…</div>;

  const theme = cfg.theme === "dark" ? "dark" : "light";
  const rawMode = cfg.mode || "fishtank";
  const mode = rawMode === "allamerican" ? "oc:allamerican" : rawMode;
  const activeClub = mode.startsWith("oc:") ? clubs.find((c) => c.id === mode.slice(3)) : null;

  return (
    <div style={{ ...PALETTES[theme], minHeight: "100vh", background: C.paper, color: C.ink, fontFamily: "'Avenir Next', 'Segoe UI', system-ui, sans-serif", colorScheme: theme }}>
      <div style={{ background: C.bar, padding: "18px 26px", display: "flex", alignItems: "center", gap: 18, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 24, color: "var(--barText)" }}>
          AK's Book <span style={{ color: "var(--barGold)", fontSize: 15 }}>• weekly accounting</span>
        </div>
        <div style={{ display: "flex", gap: 2, background: "rgba(255,255,255,0.08)", borderRadius: 7, padding: 3 }}>
          {[["fishtank", "Fish Tank"], ...clubs.map((c) => ["oc:" + c.id, c.name]), ["agent", "My Clubs"], ["book", "Book"], ["tabs", "Tabs"], ["archive", "Archive"]].map(([k, label]) => (
            <button key={k} onClick={() => up({ mode: k })} style={{
              border: "none", cursor: "pointer", borderRadius: 5, padding: "5px 14px", fontSize: 12.5, fontWeight: 700,
              background: mode === k ? "var(--gold)" : "transparent",
              color: mode === k ? "var(--onGold)" : "var(--barMute)" }}>
              {label}
            </button>
          ))}
          <button onClick={addClub} title="Add an owner club (same weekly export format as All American / Bazaar)" style={{ border: "none", cursor: "pointer", borderRadius: 5, padding: "5px 9px", fontSize: 13, fontWeight: 700, background: "transparent", color: "var(--barMute)" }}>+</button>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 10, alignItems: "center" }}>
          {mode === "fishtank" && period && <span style={{ color: "var(--barMute)", fontSize: 12.5 }}>{period}{finalized ? " · finalized" : ""}</span>}
          {mode === "fishtank" && period && model && <button onClick={archiveNow} title="Store this week's raw upload + generated workbook in the Archive now"
            style={{ background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 12 }}>Archive this week</button>}
          <input ref={seedRef} type="file" accept=".json,application/json" style={{ display: "none" }}
            onChange={(e) => { const f = e.target.files?.[0]; if (f) onSeedFile(f); e.target.value = ""; }} />
          <button onClick={exportSnapshot} title="Download everything (deals, names, clubs, tabs) as a seed file — share it, back it up, or bake it into the app"
            style={{ background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 12 }}>
            Export data
          </button>
          <button onClick={() => seedRef.current?.click()} title="Load a seed file — fills in deals, names, clubs, and tabs"
            style={{ background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 12 }}>
            Import data
          </button>
          <button onClick={() => up({ theme: theme === "dark" ? "light" : "dark" })} title="Toggle dark mode"
            style={{ background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 14 }}>
            {theme === "dark" ? "☀" : "☾"}
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />
          {mode === "fishtank" && <Btn tone="gold" small onClick={() => fileRef.current?.click()}>Upload weekly export</Btn>}
        </div>
      </div>

      {mode === "agent" ? (
        <AgentClubs theme={theme} />
      ) : mode === "book" ? (
        <BookSection />
      ) : activeClub ? (
        <AllAmerican key={activeClub.id} club={activeClub} clubs={clubs} saveClubs={saveClubs} onDeleteClub={deleteClub} />
      ) : mode.startsWith("oc:") ? (
        <div style={{ padding: 40, color: C.mute }}>That club no longer exists. <button onClick={() => up({ mode: "tabs" })} style={{ color: C.goldDark, background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Go to Tabs</button></div>
      ) : mode === "tabs" ? (
        <TabsLedger clubs={clubs} />
      ) : mode === "archive" ? (
        <ArchiveView clubs={clubs} />
      ) : (
      <>
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap" }}>
        {[["settle", "Settlements"], ["agents", "Agent & umbrella reports"], ["backed", "House-backed"], ["recon", "Ak / Jon"], ["deals", "Deals & setup"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: tab === k ? C.card : "transparent", color: tab === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2,
            boxShadow: tab === k ? "0 -1px 4px rgba(0,0,0,0.1)" : "none" }}>
            {label}
            {k === "deals" && (needsSetup.deals.length + needsSetup.assigns.length > 0) && (
              <span style={{ marginLeft: 6, background: C.red, color: "#fff", borderRadius: 9, padding: "1px 7px", fontSize: 10.5 }}>
                {needsSetup.deals.length + needsSetup.assigns.length}
              </span>
            )}
          </button>
        ))}
      </div>

      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {err && <div style={{ background: "var(--errBg)", color: C.red, padding: "10px 14px", borderRadius: 6, marginBottom: 14, fontSize: 13.5 }}>{err}</div>}
        {saveNote && <div style={{ background: C.banner, color: C.goldDark, padding: "8px 14px", borderRadius: 6, marginBottom: 14, fontSize: 12.5 }}>{saveNote}</div>}

        {!players && (
          <div style={{ background: C.card, border: `1px dashed ${C.gold}`, borderRadius: 10, padding: "50px 30px", textAlign: "center" }}>
            <div style={{ fontFamily: "Georgia, serif", fontSize: 20, marginBottom: 8 }}>Start the week</div>
            <div style={{ color: C.mute, fontSize: 14, marginBottom: 18 }}>Upload the club's weekly .xlsx export. Deals, umbrellas, backed-player ledgers, and fees are saved and apply automatically.</div>
            <Btn onClick={() => fileRef.current?.click()}>Choose file</Btn>
          </div>
        )}

        {players && model && tab === "settle" && <SettleTab model={model} cfg={cfg} needsSetup={needsSetup} goDeals={() => setTab("deals")} period={period} />}
        {players && model && tab === "agents" && <AgentsTab model={model} expanded={expanded} setExpanded={setExpanded} period={period} />}
        {players && model && tab === "backed" && <BackedTab model={model} cfg={cfg} up={up} finalized={finalized} lockWeek={lockWeek} unlockWeek={unlockWeek} />}
        {players && model && tab === "recon" && <ReconTab model={model} cfg={cfg} up={up} period={period} />}
        {players && model && tab === "deals" && <DealsTab model={model} cfg={cfg} up={up} needsSetup={needsSetup} weekAdj={weekAdj} setAdj={setAdj} />}
      </div>
      </>
      )}
    </div>
  );
}

// ———————————————— Settlements ————————————————
function SettleTab({ model, cfg, needsSetup, goDeals, period }) {
  const t = model.totals;
  const pending = needsSetup.deals.length + needsSetup.assigns.length;
  const [exportData, setExportData] = useState(null);
  const exportCsv = () =>
    setExportData({
      title: `Settlements · ${period || "this week"}`,
      text: toTSV(["Deal", "Type", "Hands", "Winnings", "Tips", "Avg TB %", "Tipback", "Settlement"],
        model.entities.map((e) => [e.name, e.type, e.hands, e.pnl.toFixed(2), e.fee.toFixed(2), e.fee ? ((e.tipback / e.fee) * 100).toFixed(1) : "", e.tipback.toFixed(2), e.settlement.toFixed(2)])),
    });

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      {pending > 0 && (
        <div style={{ background: C.banner, border: `1px solid ${C.gold}`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ fontSize: 13.5 }}>
            <b>{pending} item{pending > 1 ? "s" : ""} need review this week</b> — {needsSetup.deals.length > 0 && `${needsSetup.deals.length} deal${needsSetup.deals.length > 1 ? "s" : ""} to confirm`}{needsSetup.deals.length > 0 && needsSetup.assigns.length > 0 && ", "}{needsSetup.assigns.length > 0 && `${needsSetup.assigns.length} unassigned to Ak/Jon`}. Unconfirmed deals use the default {cfg.defaultTB}%.
          </div>
          <div style={{ marginLeft: "auto" }}><Btn tone="gold" small onClick={goDeals}>Review now</Btn></div>
        </div>
      )}

      <div style={{ display: "flex", alignItems: "baseline", marginBottom: 10 }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>What every deal owes</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <Btn tone="gold" small onClick={() => {
            if (model.unassigned.length > 0) {
              window.alert(`Assign ${model.unassigned.length} remaining deal${model.unassigned.length > 1 ? "s" : ""} to Ak or Jon first (Ak / Jon tab). The workbook includes each owner's collection list and the final Ak↔Jon transfer, so it needs every deal assigned.`);
              return;
            }
            downloadWorkbook(model, period, cfg);
          }}>Download Excel workbook</Btn>
          <Btn tone="ghost" small onClick={exportCsv}>Copy table</Btn>
        </div>
      </div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>
        <span style={{ color: C.green, fontWeight: 700 }}>Green</span> = you pay them · <span style={{ color: C.red, fontWeight: 700 }}>red</span> = they pay you.
        Umbrellas and super agents settle as one line; no-SA players settle individually. Makeup players' settlement is their share of profit above makeup only (RB is a credit inside their net); action buys settle the player's share of P&L + RB. Owner accounts are excluded (see Ak / Jon).
      </div>

      <div style={{ background: C.card, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 6px rgba(0,0,0,0.15)" }}>
        <div style={{ background: C.bar, color: "var(--barText)", display: "grid", gridTemplateColumns: "minmax(200px,1.5fr) repeat(6, 1fr)", padding: "13px 10px", alignItems: "center" }}>
          <div style={{ paddingLeft: 10, fontFamily: "Georgia, serif", fontSize: 16 }}>Grand Total</div>
          {[fmtI(t.hands), fmt(t.pnl), fmt(t.fee), t.fee ? ((t.tipback / t.fee) * 100).toFixed(0) + "%" : "—", fmt(t.tipback)].map((v, i) => (
            <div key={i} style={{ textAlign: "right", paddingRight: 10, fontVariantNumeric: "tabular-nums", fontSize: 14.5 }}>{v}</div>
          ))}
          <div style={{ textAlign: "right", paddingRight: 10, fontVariantNumeric: "tabular-nums", fontSize: 14.5, color: t.settlement >= 0 ? "var(--barGreen)" : "var(--barRed)", fontWeight: 700 }}>{fmt(t.settlement)}</div>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Deal name</th>
            <th style={th}>Hands</th><th style={th}>Winnings</th><th style={th}>Tips</th><th style={th}>Avg TB %</th><th style={th}>Tipback</th><th style={th}>Settlement</th>
          </tr></thead>
          <tbody>
            {model.entities.map((e, i) => (
              <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                <td style={tdL}>
                  <span style={{ fontWeight: 600 }}>{e.name}</span> {typePill(e)}
                  {e.adjusted && <span style={{ marginLeft: 6 }}><Pill tone="blue">mid-week deal</Pill></span>}
                  {e.type === "backed" && e.dealType === "makeup" && e.inMakeup && <span style={{ marginLeft: 6 }}><Pill tone="red">in makeup</Pill></span>}
                  {e.type === "player" && e.members[0].actionTaxPct ? <span style={{ marginLeft: 6 }}><Pill tone="blue">action {e.members[0].actionTaxPct}%</Pill></span> : null}
                </td>
                <td style={td}>{fmtI(e.hands)}</td>
                <td style={td}>{fmt(e.pnl)}</td>
                <td style={td}>{fmt(e.fee)}</td>
                <td style={td}>{e.fee ? ((e.tipback / e.fee) * 100).toFixed(0) + "%" : "—"}</td>
                <td style={td}>{fmt(e.tipback)}</td>
                <td style={td}>{money(e.settlement)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ———————————————— Agent & umbrella reports ————————————————
function AgentsTab({ model, expanded, setExpanded, period }) {
  const groups = [...model.umbEntities, ...model.looseSAs];
  const [exportData, setExportData] = useState(null);
  const exportOne = (e, members) => {
    const rows = members.map((m) => [m.name, m.memberId, m.saName, m.agentName, m.hands, m.pnl.toFixed(2), m.fee.toFixed(2), m.tbPct, m.tipback.toFixed(2), m.settlement.toFixed(2)]);
    const total = ["TOTAL", "", "", "", e.hands, e.pnl.toFixed(2), e.fee.toFixed(2), e.fee ? ((e.tipback / e.fee) * 100).toFixed(1) : "", e.tipback.toFixed(2), e.settlement.toFixed(2)];
    setExportData({
      title: `${e.name} · ${period || "this week"}`,
      text: toTSV(["Player", "Device ID", "Super Agent", "Agent", "Hands", "Winnings", "Tips", "TB %", "Tipback", "Settlement"], [...rows, total]),
    });
  };

  const memberTable = (members, total) => (
    <table style={{ width: "100%", borderCollapse: "collapse" }}>
      <thead><tr>
        <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Device ID</th><th style={{ ...th, textAlign: "left" }}>Agent</th>
        <th style={th}>Hands</th><th style={th}>Winnings</th><th style={th}>Tips</th><th style={th}>TB %</th><th style={th}>Tipback</th><th style={th}>Settlement</th>
      </tr></thead>
      <tbody>
        {[...members].sort((a, b) => b.fee - a.fee).map((m, i) => (
          <tr key={m.memberId + i} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
            <td style={{ ...tdL, fontWeight: 600 }}>{m.name}{m.actionTaxPct ? <span style={{ marginLeft: 6 }}><Pill tone="blue">action {m.actionTaxPct}%</Pill></span> : null}</td>
            <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{m.memberId}</td>
            <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{m.agentName}</td>
            <td style={td}>{fmtI(m.hands)}</td><td style={td}>{fmt(m.pnl)}</td><td style={td}>{fmt(m.fee)}</td>
            <td style={td}>{m.tbPct}%</td><td style={td}>{fmt(m.tipback)}</td><td style={td}>{money(m.settlement)}</td>
          </tr>
        ))}
        <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
          <td style={{ ...tdL, fontWeight: 700 }} colSpan={3}>Total</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmtI(total.hands)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(total.pnl)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(total.fee)}</td>
          <td style={td}>{total.fee ? ((total.tipback / total.fee) * 100).toFixed(0) + "%" : "—"}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(total.tipback)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{money(total.settlement)}</td>
        </tr>
      </tbody>
    </table>
  );

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ fontFamily: "Georgia, serif", fontSize: 19, marginBottom: 4 }}>Reports for each deal</div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 14 }}>Click a row to expand; export a CSV to send them. Umbrella reports break out each super agent inside.</div>
      {groups.map((e) => {
        const open = expanded[e.key];
        return (
          <div key={e.key} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
            <div onClick={() => setExpanded({ ...expanded, [e.key]: !open })} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 16px", cursor: "pointer", background: C.cream }}>
              <span style={{ color: C.goldDark, fontSize: 12, width: 12 }}>{open ? "▼" : "►"}</span>
              <span style={{ fontWeight: 700, fontSize: 14.5 }}>{e.name}</span>
              {typePill(e)}
              <span style={{ color: C.mute, fontSize: 12 }}>{e.members.length} player{e.members.length !== 1 ? "s" : ""} · {fmtI(e.hands)} hands</span>
              <span style={{ marginLeft: "auto", fontSize: 13.5 }}>settlement {money(e.settlement)}</span>
              <Btn tone="gold" small onClick={(ev) => { ev.stopPropagation(); downloadDealExcel(e, period); }}>Excel</Btn>
              <Btn tone="ghost" small onClick={(ev) => { ev.stopPropagation(); exportOne(e, e.members); }}>Copy</Btn>
            </div>
            {open && (e.type === "umbrella"
              ? e.subgroups.map((s) => (
                  <div key={s.key} style={{ borderTop: `1px solid ${C.line}` }}>
                    <div style={{ padding: "8px 16px", fontSize: 13, fontWeight: 700, color: C.goldDark, background: C.rowAlt }}>
                      {s.name} — settlement {money(s.settlement)}
                    </div>
                    {memberTable(s.members, s)}
                  </div>
                ))
              : memberTable(e.members, e))}
          </div>
        );
      })}
    </div>
  );
}

// ———————————————— House-backed tab ————————————————
function BackedTab({ model, cfg, up, finalized, lockWeek, unlockWeek }) {
  const [newName, setNewName] = useState("");
  const [newDeal, setNewDeal] = useState("makeup");
  const addBacked = () => {
    const k = newName.trim().toLowerCase();
    if (!k) return;
    const base = newDeal === "action"
      ? { name: newName.trim(), deal: "action", actionPct: 50, rbPct: 100, backer: "jon" }
      : { name: newName.trim(), deal: "makeup", rbNormal: cfg.defaultTB, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: "split" };
    up({ backed: { ...cfg.backed, [k]: base } });
    setNewName("");
  };
  const setB = (k, patch) => up({ backed: { ...cfg.backed, [k]: { ...cfg.backed[k], ...patch } } });
  const removeB = (k) => { const b = { ...cfg.backed }; delete b[k]; up({ backed: b }); };
  const findE = (k) => model.backedEntities.find((x) => x.key === `b:${k}`);
  const backerSel = (k, b, withSplit) => (
    <select value={b.backer} onChange={(e) => setB(k, { backer: e.target.value })} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
      {withSplit && <option value="split">Ak & Jon 50/50</option>}
      <option value="ak">Ak</option><option value="jon">Jon</option>
    </select>
  );

  const makeupPlayers = Object.entries(cfg.backed).filter(([, b]) => b.deal !== "action");
  const actionPlayers = Object.entries(cfg.backed).filter(([, b]) => b.deal === "action");

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 4, gap: 10, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>House-backed players</div>
        <div style={{ marginLeft: "auto" }}>
          {finalized
            ? <Btn tone="ghost" small onClick={unlockWeek}>🔒 Week locked · unlock to edit</Btn>
            : <Btn tone="gold" small onClick={lockWeek}>Lock week — roll makeup forward</Btn>}
        </div>
      </div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 16 }}>
        <b>Makeup deals</b>: the week's net = P&L + RB credit. Above makeup, the player is paid their % of the excess and the backer books the rest; below, no cash moves and the net accrues to makeup on the backer's book. The margin on their fees stays in split club profit. RB rate follows makeup <b>entering</b> the week; Lock once to roll it forward. If you need to fix something after locking, unlock — the roll reverts so you can edit — then lock again to re-snapshot off the corrected numbers. If this week was already accepted into Tabs, re-sync it from Tabs → Bookkeeping afterward so the ledger picks up the correction cleanly (no double-posting).
        {" "}<b>Action buys</b>: the backer owns their % of (P&L + rakeback); the player settles the remainder.
        {" "}Paste a Tabs → Staking makeup deal's ID into <b>Unified deal</b> on a row to fold that player's weekly net into one shared, ongoing pool across sites (and manual/external games) instead of tracking makeup locally here — the local balance freezes and the deal's own % staked / chop handle the tab. Use a "settle per session" deal (the default) so this week's report and any mid-week manual results each update the running makeup as they're logged, in the order they happened, rather than getting batched into one lump sum.
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Makeup deals</div>
      <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Makeup entering</th><th style={{ ...th, textAlign: "center" }}>Status</th>
            <th style={th}>RB % normal</th><th style={th}>RB % makeup</th><th style={th}>Player profit %</th>
            <th style={th}>Tips</th><th style={th}>P&L</th><th style={th}>RB credit</th><th style={th}>Net</th><th style={th}>Settlement</th><th style={th}>Makeup after</th><th style={th}>Unified deal</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {makeupPlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>
                    {b.name}
                    {!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}
                  </td>
                  <td style={tdL}>{backerSel(k, b, true)}</td>
                  <td style={td}><NumInput width={85} value={b.makeup} onChange={(v) => setB(k, { makeup: v })} /></td>
                  <td style={{ ...td, textAlign: "center" }}>{(b.makeup || 0) > 0.005 ? <Pill tone="red">in makeup</Pill> : <Pill tone="green">clear</Pill>}</td>
                  <td style={td}>
                    <span style={{ opacity: (b.makeup || 0) > 0.005 ? 0.35 : 1 }}>
                      <PctInput width={46} max={999} value={b.rbNormal} onChange={(v) => v != null && setB(k, { rbNormal: v })} />
                    </span>
                    {(b.makeup || 0) <= 0.005 && <div style={{ fontSize: 9.5, color: C.green, fontWeight: 700, marginTop: 2 }}>ACTIVE</div>}
                  </td>
                  <td style={td}>
                    <span style={{ opacity: (b.makeup || 0) > 0.005 ? 1 : 0.35 }}>
                      <PctInput width={46} max={999} value={b.rbMakeup} onChange={(v) => v != null && setB(k, { rbMakeup: v })} />
                    </span>
                    {(b.makeup || 0) > 0.005 && <div style={{ fontSize: 9.5, color: C.green, fontWeight: 700, marginTop: 2 }}>ACTIVE</div>}
                  </td>
                  <td style={td}><PctInput width={46} value={b.playerProfitPct ?? 50} onChange={(v) => v != null && setB(k, { playerProfitPct: v })} /></td>
                  <td style={td}>{e ? fmt(e.fee) : "—"}</td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td}>{e ? <>{fmt(e.tipback)} <span style={{ color: C.mute, fontSize: 11 }}>@{e.rb}%</span></> : "—"}</td>
                  <td style={td}>{e ? money(e.net) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? fmt(e.makeupAfter) : fmt(b.makeup || 0)}</td>
                  <td style={td}>
                    <input placeholder="deal id" value={b.unifiedDealId || ""} onChange={(ev) => setB(k, { unifiedDealId: ev.target.value.trim() })} style={{ ...inputS, width: 88, fontSize: 11 }} title="Paste a Tabs → Staking makeup deal's ID to fold this player's weekly net into that shared, ongoing pool instead of tracking makeup locally here. Use a 'settle per session' deal so it updates alongside any mid-week manual results in the order they happened." />
                    {b.unifiedDealId && <div style={{ fontSize: 9.5, color: C.goldDark, fontWeight: 700, marginTop: 2 }}>UNIFIED</div>}
                  </td>
                  <td style={td}><button onClick={() => removeB(k)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Action buys</div>
      <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Backer action %</th><th style={th}>RB %</th>
            <th style={th}>Tips</th><th style={th}>Week P&L</th><th style={th}>Player settlement</th><th style={th}>Backer book</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {actionPlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>
                    {b.name}
                    {!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}
                  </td>
                  <td style={tdL}>{backerSel(k, b, false)}</td>
                  <td style={td}><PctInput width={46} value={b.actionPct} onChange={(v) => v != null && setB(k, { actionPct: v })} /></td>
                  <td style={td}><PctInput width={46} max={999} value={b.rbPct ?? 100} onChange={(v) => v != null && setB(k, { rbPct: v })} /></td>
                  <td style={td}>{e ? fmt(e.fee) : "—"}</td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? money(e.backerBook) : "—"}</td>
                  <td style={td}><button onClick={() => removeB(k)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <input placeholder="Add backed player by exact nickname…" value={newName} onChange={(e) => setNewName(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && addBacked()} style={{ ...inputS, width: 260 }} />
        <select value={newDeal} onChange={(e) => setNewDeal(e.target.value)} style={{ ...inputS, fontSize: 12 }}>
          <option value="makeup">Makeup deal</option><option value="action">Action buy</option>
        </select>
        <Btn tone="ghost" small onClick={addBacked}>+ Add</Btn>
        <div style={{ marginLeft: "auto", fontSize: 13 }}>
          Backed books this week — Ak: <b>{money(model.backedBook.ak)}</b> · Jon: <b>{money(model.backedBook.jon)}</b>
        </div>
      </div>
    </div>
  );
}

// ———————————————— Ak / Jon reconciliation ————————————————
function ReconTab({ model, cfg, up, period }) {
  const m = model;
  const assign = (key, who) => up({ assignments: { ...cfg.assignments, [key]: who } });
  const assignAll = (who) => { const a = { ...cfg.assignments }; m.unassigned.forEach((e) => (a[e.key] = who)); up({ assignments: a }); };
  const owePos = m.akOwesJon > 0.005, oweNeg = m.akOwesJon < -0.005;

  const row = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "6px 0", borderBottom: opts.rule ? `1px solid ${C.line}` : "none", fontSize: 13.5 }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{typeof val === "number" ? money(val) : val}</span>
    </div>
  );

  return (
    <div>
      <div style={{ background: C.bar, borderRadius: 12, padding: "26px 30px", marginBottom: 18, textAlign: "center", color: "var(--barText)" }}>
        <div style={{ fontSize: 12, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--barGold)", marginBottom: 6 }}>Owner balance · {period || "this week"}</div>
        {m.unassigned.length > 0 ? (
          <div style={{ fontSize: 17 }}>Assign the {m.unassigned.length} remaining deal{m.unassigned.length > 1 ? "s" : ""} below to get the final number.</div>
        ) : (
          <div style={{ fontFamily: "Georgia, serif", fontSize: 30 }}>
            {owePos && <>Ak pays Jon <span style={{ color: "var(--barGold)" }}>{fmt(m.akOwesJon)}</span></>}
            {oweNeg && <>Jon pays Ak <span style={{ color: "var(--barGold)" }}>{fmt(-m.akOwesJon)}</span></>}
            {!owePos && !oweNeg && <>Perfectly even — no transfer needed</>}
          </div>
        )}
        <div style={{ fontSize: 12, color: "var(--barSubtle)", marginTop: 8 }}>
          After this transfer each of you nets exactly: half of net club profit + your own accounts' P&L with 100% feeback + your backed books{m.feeRows.some((f) => f.recipient !== "external") ? " + any fee routed to you" : ""}.
          {m.balanceOk ? " Balance check: ✓ books tie out." : m.unassigned.length === 0 ? " ⚠ Balance check failed — review setup." : ""}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(310px, 1fr))", gap: 14, marginBottom: 18 }}>
        <Card title="Club economics">
          {row("Total tips collected (all accounts)", m.clubRevenue)}
          {row("Tipbacks to agents & players", -m.extTipbacks)}
          {row("Backed players' RB (cash + credits)", -m.backedRB)}
          {row("Owner accounts' 100% feeback", -m.ownFeeback, { rule: true })}
          {row("Club profit", m.clubProfit, { bold: true })}
          {m.feeRows.map((f) => row(`${f.label} · ${f.kind === "fixed" ? "fixed" : `${f.pct}% of ${cfg.feeBase === "gross" ? "tips" : "profit"}`}${f.recipient !== "external" ? ` → ${f.recipient === "ak" ? "Ak" : "Jon"}` : ""}`, -f.amount))}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("Net profit to split", m.netProfit, { bold: true })}
          {row("Each owner's half", m.netProfit / 2)}
        </Card>
        {["ak", "jon"].map((w) => (
          <Card key={w} title={`${w === "ak" ? "Ak" : "Jon"}'s position`}>
            {m.own.filter((p) => p.owner === w).map((p) => row(`${p.name} · P&L ${fmt(p.pnl)} + feeback ${fmt(p.feeback)}`, p.position))}
            {m.own.filter((p) => p.owner === w).length === 0 && <div style={{ color: C.mute, fontSize: 12.5 }}>No activity from these accounts this week.</div>}
            <div style={{ borderTop: `1px solid ${C.line}` }} />
            {row("Own accounts (P&L + 100% feeback)", m.ownPosition[w], { bold: true })}
            {row("Backed books", m.backedBook[w])}
            {row("Action-buy tax book", m.taxBook[w])}
            {row("Half of net club profit", m.netProfit / 2)}
            {row("Entitlement (all-in)", m.entitle[w], { bold: true })}
            {row("Actual cash from assigned settlements", m.actual[w])}
          </Card>
        ))}
      </div>

      <Card title="Who settles with whom" right={m.unassigned.length > 0 && (
        <span style={{ display: "flex", gap: 8 }}>
          <Btn tone="ghost" small onClick={() => assignAll("ak")}>Rest → Ak</Btn>
          <Btn tone="ghost" small onClick={() => assignAll("jon")}>Rest → Jon</Btn>
        </span>
      )}>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>Mark who physically settles each deal. Remembered for future weeks.</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Deal</th><th style={th}>Settlement</th><th style={{ ...th, textAlign: "center" }}>Settled by</th>
          </tr></thead>
          <tbody>
            {m.entities.map((e, i) => {
              const who = cfg.assignments[e.key];
              return (
                <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={tdL}><b>{e.name}</b> {typePill(e)}</td>
                  <td style={td}>{money(e.settlement)}</td>
                  <td style={{ ...td, textAlign: "center" }}>
                    {["ak", "jon"].map((w) => (
                      <button key={w} onClick={() => assign(e.key, w)} style={{
                        margin: "0 3px", padding: "4px 14px", borderRadius: 5, cursor: "pointer", fontSize: 12, fontWeight: 700,
                        border: `1px solid ${who === w ? C.goldDark : C.line}`,
                        background: who === w ? C.gold : C.surface, color: who === w ? "var(--onGold)" : C.mute }}>
                        {w === "ak" ? "Ak" : "Jon"}
                      </button>
                    ))}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ———————————————— Deals & setup ————————————————
function DealsTab({ model, cfg, up, needsSetup, weekAdj, setAdj }) {
  const [showAllSA, setShowAllSA] = useState({});
  const [umbName, setUmbName] = useState("");
  const [adjTarget, setAdjTarget] = useState("");
  const [taxTarget, setTaxTarget] = useState("");
  const taxTargets = [...new Map([...model.saEntities, ...model.indEntities].flatMap((e) => e.members).map((m) => [m.memberId, m])).values()].sort((a, b) => b.fee - a.fee);
  const addTax = () => {
    if (!taxTarget) return;
    up({ actionTax: { ...(cfg.actionTax || {}), [taxTarget]: { pct: 20, backer: "split" } } });
    setTaxTarget("");
  };

  const allSAs = [...model.looseSAs, ...model.umbEntities.flatMap((u) => u.subgroups)].sort((a, b) => b.fee - a.fee);
  const inds = model.indEntities;
  const omit = (o, k) => { const x = { ...o }; delete x[k]; return x; };

  const setSA = (id, pct) => up({ saDeals: pct == null ? omit(cfg.saDeals, id) : { ...cfg.saDeals, [id]: pct }, confirmedSAs: { ...cfg.confirmedSAs, [id]: true } });
  const setPlayer = (id, pct, isInd) => {
    const patch = { playerDeals: pct == null ? omit(cfg.playerDeals, id) : { ...cfg.playerDeals, [id]: pct } };
    if (isInd) patch.confirmedPlayers = { ...cfg.confirmedPlayers, [id]: true };
    up(patch);
  };
  const confirmAll = () => {
    const cs = { ...cfg.confirmedSAs }, cp = { ...cfg.confirmedPlayers };
    allSAs.forEach((e) => (cs[e.id] = true)); inds.forEach((e) => (cp[e.id] = true));
    up({ confirmedSAs: cs, confirmedPlayers: cp });
  };

  const addUmbrella = () => { if (!umbName.trim()) return; up({ umbrellas: [...cfg.umbrellas, { id: "u" + Date.now(), name: umbName.trim(), saIds: [] }] }); setUmbName(""); };
  const toggleSAinUmb = (uid, saId) => {
    up({ umbrellas: cfg.umbrellas.map((u) => {
      if (u.id === uid) return { ...u, saIds: u.saIds.includes(saId) ? u.saIds.filter((x) => x !== saId) : [...u.saIds, saId] };
      return { ...u, saIds: u.saIds.filter((x) => x !== saId) };
    }) });
  };
  const saOptions = allSAs.map((e) => ({ id: e.id, name: e.name }));
  Object.entries(cfg.names).forEach(([id, name]) => {
    if (cfg.saDeals[id] !== undefined && !saOptions.find((o) => o.id === id)) saOptions.push({ id, name });
  });

  const adjTargets = [
    ...allSAs.map((e) => ({ key: e.key, label: `${e.name} (SA)`, fee: e.fee })),
    ...inds.map((e) => ({ key: e.key, label: `${e.name} (player)`, fee: e.fee })),
    ...model.backedEntities.map((e) => ({ key: e.key, label: `${e.name} (backed)`, fee: e.fee })),
  ];
  const addAdj = () => {
    if (!adjTarget) return;
    setAdj({ ...weekAdj, [adjTarget]: { amtA: 0, rateA: cfg.defaultTB, rateB: cfg.defaultTB } });
    setAdjTarget("");
  };
  const setAdjField = (key, patch) => setAdj({ ...weekAdj, [key]: { ...weekAdj[key], ...patch } });
  const removeAdj = (key) => { const a = { ...weekAdj }; delete a[key]; setAdj(a); };

  const newBadge = (isNew) => isNew && <Pill tone="red">confirm</Pill>;

  return (
    <div>
      {needsSetup.deals.length > 0 && (
        <div style={{ background: C.banner, border: `1px solid ${C.gold}`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center" }}>
          <div style={{ fontSize: 13.5 }}><b>New this week:</b> {needsSetup.deals.map((e) => e.name).join(", ")} — set or confirm their deals below.</div>
          <div style={{ marginLeft: "auto" }}><Btn tone="gold" small onClick={confirmAll}>Everything's right — confirm all</Btn></div>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(330px, 1fr))", gap: 14, marginBottom: 16 }}>
        <Card title="Defaults & fees">
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, fontSize: 13.5 }}>
            Default tipback for anyone without a deal
            <span style={{ marginLeft: "auto" }}><PctInput max={999} value={cfg.defaultTB} onChange={(v) => v != null && up({ defaultTB: v })} /></span>
          </div>
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.08em" }}>Cut fees</div>
          {cfg.fees.map((f, i) => (
            <div key={f.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, fontSize: 13 }}>
              <input value={f.label} onChange={(e) => { const fees = [...cfg.fees]; fees[i] = { ...f, label: e.target.value }; up({ fees }); }} style={{ ...inputS, flex: 1 }} />
              <select value={f.kind === "fixed" ? "fixed" : "pct"} onChange={(e) => { const fees = [...cfg.fees]; fees[i] = { ...f, kind: e.target.value === "fixed" ? "fixed" : "pct" }; up({ fees }); }} style={{ ...inputS, padding: "5px 6px", fontSize: 12 }}>
                <option value="pct">%</option><option value="fixed">$</option>
              </select>
              {f.kind === "fixed"
                ? <NumInput width={72} value={f.amount ?? 0} onChange={(v) => { const fees = [...cfg.fees]; fees[i] = { ...f, amount: v }; up({ fees }); }} />
                : <PctInput width={52} value={f.pct} onChange={(v) => { if (v == null) return; const fees = [...cfg.fees]; fees[i] = { ...f, pct: v }; up({ fees }); }} />}
              <select value={f.recipient} onChange={(e) => { const fees = [...cfg.fees]; fees[i] = { ...f, recipient: e.target.value }; up({ fees }); }} style={{ ...inputS, padding: "5px 6px", fontSize: 12 }}>
                <option value="external">→ outside</option><option value="ak">→ Ak</option><option value="jon">→ Jon</option>
              </select>
              <button onClick={() => up({ fees: cfg.fees.filter((_, j) => j !== i) })} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
            </div>
          ))}
          <div style={{ display: "flex", gap: 10, alignItems: "center", marginTop: 8 }}>
            <Btn tone="ghost" small onClick={() => up({ fees: [...cfg.fees, { id: "f" + Date.now(), label: "New fee", pct: 1, recipient: "external", paidBy: "split" }] })}>+ Add fee</Btn>
            <label style={{ marginLeft: "auto", fontSize: 12, color: C.mute }}>
              Fees are % of{" "}
              <select value={cfg.feeBase} onChange={(e) => up({ feeBase: e.target.value })} style={{ ...inputS, padding: "3px 6px", fontSize: 12 }}>
                <option value="net">club profit (after tipbacks)</option>
                <option value="gross">gross tips collected</option>
              </select>
            </label>
          </div>
        </Card>

        <Card title="Owner accounts">
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 10 }}>House play: 100% feeback, excluded from settlements, P&L + feeback credited to the owner. One nickname per line. Backed players are managed on the House-backed tab.</div>
          {["ak", "jon"].map((w) => (
            <div key={w} style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{w === "ak" ? "Ak" : "Jon"}</div>
              <textarea defaultValue={cfg.ownAccounts[w].join("\n")}
                onBlur={(e) => up({ ownAccounts: { ...cfg.ownAccounts, [w]: e.target.value.split("\n").map((s) => s.trim()).filter(Boolean) } })}
                rows={2}
                style={{ ...inputS, width: "100%", boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
            </div>
          ))}
        </Card>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Card title="Umbrella groups" right={
          <span style={{ display: "flex", gap: 8 }}>
            <input placeholder="New umbrella name…" value={umbName} onChange={(e) => setUmbName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUmbrella()} style={{ ...inputS, width: 180 }} />
            <Btn tone="ghost" small onClick={addUmbrella}>+ Create</Btn>
          </span>
        }>
          <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>Group super agents under one name — they settle as a single line and get one combined report. Each SA keeps its own tipback rate.</div>
          {cfg.umbrellas.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No umbrellas yet.</div>}
          {cfg.umbrellas.map((u) => (
            <div key={u.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <input value={u.name} onChange={(e) => up({ umbrellas: cfg.umbrellas.map((x) => x.id === u.id ? { ...x, name: e.target.value } : x) })} style={{ ...inputS, fontWeight: 700, width: 200 }} />
                <Pill tone="blue">{u.saIds.length} SAs</Pill>
                <button onClick={() => up({ umbrellas: cfg.umbrellas.filter((x) => x.id !== u.id) })} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>× delete</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {saOptions.map((o) => {
                  const inThis = u.saIds.includes(o.id);
                  const inOther = !inThis && cfg.umbrellas.some((x) => x.id !== u.id && x.saIds.includes(o.id));
                  return (
                    <button key={o.id} onClick={() => !inOther && toggleSAinUmb(u.id, o.id)} style={{
                      padding: "4px 12px", borderRadius: 14, fontSize: 12, fontWeight: 600, cursor: inOther ? "default" : "pointer",
                      border: `1px solid ${inThis ? C.goldDark : C.line}`,
                      background: inThis ? C.gold : C.surface, color: inThis ? "var(--onGold)" : inOther ? "var(--chipOff)" : C.mute, opacity: inOther ? 0.6 : 1 }}>
                      {o.name}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </Card>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Card title="Mid-week deal changes (this week only)" right={
          <span style={{ display: "flex", gap: 8 }}>
            <select value={adjTarget} onChange={(e) => setAdjTarget(e.target.value)} style={{ ...inputS, fontSize: 12, maxWidth: 220 }}>
              <option value="">Pick a deal…</option>
              {adjTargets.filter((t) => !weekAdj[t.key]).map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}
            </select>
            <Btn tone="ghost" small onClick={addAdj}>+ Add change</Btn>
          </span>
        }>
          <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>
            For the rare case a deal changes partway through the week: the first $X of their tips settles at the old rate, the rest at the new rate. These clear automatically when you upload the next week's file.
          </div>
          {Object.keys(weekAdj).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None this week.</div>}
          {Object.entries(weekAdj).map(([key, a]) => {
            const t = adjTargets.find((x) => x.key === key);
            const tipback = Math.min(a.amtA, t?.fee || 0) * a.rateA / 100 + Math.max(0, (t?.fee || 0) - a.amtA) * a.rateB / 100;
            return (
              <div key={key} style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", borderTop: `1px solid ${C.line}`, padding: "10px 0", fontSize: 13 }}>
                <b style={{ minWidth: 140 }}>{t?.label || key}</b>
                <span style={{ color: C.mute }}>first</span>
                <NumInput width={90} value={a.amtA} onChange={(v) => setAdjField(key, { amtA: v })} />
                <span style={{ color: C.mute }}>of tips @</span>
                <PctInput width={50} value={a.rateA} onChange={(v) => v != null && setAdjField(key, { rateA: v })} />
                <span style={{ color: C.mute }}>· remaining {t ? fmt(Math.max(0, t.fee - a.amtA)) : "—"} @</span>
                <PctInput width={50} value={a.rateB} onChange={(v) => v != null && setAdjField(key, { rateB: v })} />
                <span style={{ marginLeft: "auto", fontWeight: 700 }}>tipback → {fmt(tipback)}</span>
                <button onClick={() => removeAdj(key)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
              </div>
            );
          })}
        </Card>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Card title="Action buys on players (tax wins / rebate losses)" right={
          <span style={{ display: "flex", gap: 8 }}>
            <select value={taxTarget} onChange={(e) => setTaxTarget(e.target.value)} style={{ ...inputS, fontSize: 12, maxWidth: 220 }}>
              <option value="">Pick a player…</option>
              {taxTargets.filter((t) => !(cfg.actionTax || {})[t.memberId]).map((t) => <option key={t.memberId} value={t.memberId}>{t.name}{t.saName !== "-" ? ` (${t.saName})` : ""}</option>)}
            </select>
            <Btn tone="ghost" small onClick={addTax}>+ Add</Btn>
          </span>
        }>
          <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>
            The set % applies to the player's net after rakeback (P&L + tipback): the house takes that % when the net is positive and gives back the same % when it's negative. The house's cut lands on the chosen book in Ak / Jon.
          </div>
          {Object.keys(cfg.actionTax || {}).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None yet.</div>}
          {Object.entries(cfg.actionTax || {}).map(([mid, a]) => {
            const t = taxTargets.find((x) => x.memberId === mid);
            const nm = t?.name || cfg.names[mid] || mid;
            const cut = t ? ((t.pnl + t.tipback) * a.pct) / 100 : null;
            return (
              <div key={mid} style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", borderTop: `1px solid ${C.line}`, padding: "9px 0", fontSize: 13 }}>
                <b style={{ minWidth: 130 }}>{nm}</b>
                {!t && <Pill tone="gold">no play this week</Pill>}
                <span style={{ color: C.mute }}>house %</span>
                <PctInput width={50} value={a.pct} onChange={(v) => v != null && up({ actionTax: { ...cfg.actionTax, [mid]: { ...a, pct: v } } })} />
                <span style={{ color: C.mute }}>book</span>
                <select value={a.backer || "split"} onChange={(e) => up({ actionTax: { ...cfg.actionTax, [mid]: { ...a, backer: e.target.value } } })} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
                  <option value="split">Ak & Jon 50/50</option><option value="ak">Ak</option><option value="jon">Jon</option>
                </select>
                {t && <span style={{ marginLeft: "auto" }}>net {money(t.pnl + t.tipback)} → house cut {money(cut)}</span>}
                <button onClick={() => { const at = { ...cfg.actionTax }; delete at[mid]; up({ actionTax: at }); }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
              </div>
            );
          })}
        </Card>
      </div>

      <div style={{ background: C.card, borderRadius: 10, padding: "16px 20px", boxShadow: "0 1px 5px rgba(0,0,0,0.15)", marginBottom: 16 }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 16, marginBottom: 4 }}>Super agent deals</div>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>The rate applies to every player under the super agent. Expand to override a specific player.</div>
        {allSAs.map((e) => {
          const isNew = !cfg.confirmedSAs[e.id];
          const open = showAllSA[e.id];
          const umb = cfg.umbrellas.find((u) => u.saIds.includes(e.id));
          return (
            <div key={e.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{e.name}</span>
                {umb && <Pill tone="blue">{umb.name}</Pill>}
                <span style={{ color: C.mute, fontSize: 12 }}>{e.members.length} player{e.members.length !== 1 ? "s" : ""} · tips {fmt(e.fee)}</span>
                {newBadge(isNew)}
                <span style={{ marginLeft: "auto" }}>
                  <PctInput max={999} value={cfg.saDeals[e.id] ?? cfg.defaultTB} onChange={(v) => setSA(e.id, v ?? cfg.defaultTB)} />
                </span>
                <button onClick={() => setShowAllSA({ ...showAllSA, [e.id]: !open })} style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontSize: 12.5, fontWeight: 700 }}>
                  {open ? "hide players" : "player overrides"}
                </button>
              </div>
              {open && (
                <div style={{ marginTop: 8, marginLeft: 10, display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 6 }}>
                  {e.members.map((mm) => (
                    <div key={mm.memberId} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, background: C.rowAlt, borderRadius: 6, padding: "5px 10px" }}>
                      <span>{mm.name}</span>
                      <span style={{ marginLeft: "auto" }}>
                        <PctInput width={50} max={999} value={cfg.playerDeals[mm.memberId] ?? ""} onChange={(v) => setPlayer(mm.memberId, v, false)} />
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div style={{ background: C.card, borderRadius: 10, padding: "16px 20px", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 16, marginBottom: 4 }}>Individual players (no super agent)</div>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>Settled one by one; each has their own tipback deal. House-backed players are managed on their own tab.</div>
        {inds.map((e) => (
          <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 10, borderTop: `1px solid ${C.line}`, padding: "9px 0" }}>
            <span style={{ fontWeight: 700, fontSize: 14 }}>{e.name}</span>
            <span style={{ color: C.mute, fontSize: 12 }}>{e.id} · tips {fmt(e.fee)}</span>
            {newBadge(!cfg.confirmedPlayers[e.id])}
            <span style={{ marginLeft: "auto" }}>
              <PctInput max={999} value={cfg.playerDeals[e.id] ?? cfg.defaultTB} onChange={(v) => setPlayer(e.id, v ?? cfg.defaultTB, true)} />
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}


// ———— Cross-module helpers ————
async function loadFishTankModel() {
  try {
    const [c, d] = await Promise.all([store.get("fishtank-config-v4"), store.get("fishtank-lastweek-v4")]);
    if (!d?.value) return null;
    const wkData = JSON.parse(d.value);
    if (!wkData.players?.length) return null;
    const saved = c?.value ? JSON.parse(c.value) : {};
    const ftCfg = { ...DEFAULT_CONFIG, ...saved,
      ownAccounts: saved.ownAccounts || DEFAULT_CONFIG.ownAccounts,
      fees: saved.fees || DEFAULT_CONFIG.fees,
      backed: saved.backed || DEFAULT_CONFIG.backed,
      umbrellas: saved.umbrellas || DEFAULT_CONFIG.umbrellas };
    return { model: buildModel(wkData.players, ftCfg, wkData.weekAdj || {}, wkData.period || ""), period: wkData.period || "", cfg: ftCfg };
  } catch (e) { return null; }
}
// rows for a set of usernames out of the fish tank week (ext + backed players)
function ftRowsForNames(ft, nameSet) {
  if (!ft) return [];
  const rows = [];
  const label = `Fish Tank${ft.period ? ` (${ft.period})` : ""}`;
  ft.model.entities.forEach((e) => {
    if (e.type === "backed") {
      if (nameSet.has(e.name.trim().toLowerCase())) rows.push({ id: "ft-" + e.key, clubName: label, name: e.name, customDeal: e.dealType === "action" ? `action buy ${e.actionPct}%` : `makeup · RB ${e.rb}%`, pnl: e.pnl, tips: e.fee, tipback: e.tipback, settlement: e.settlement, margin: 0, played: true });
    } else {
      e.members.forEach((mm) => {
        if (nameSet.has(mm.name.trim().toLowerCase())) rows.push({ id: "ft-" + mm.memberId, clubName: label, name: mm.name, customDeal: `TB ${mm.tbPct}%${mm.actionTaxPct ? ` · action ${mm.actionTaxPct}%` : ""}`, pnl: mm.pnl, tips: mm.fee, tipback: mm.tipback, settlement: mm.settlement, margin: 0, played: true });
      });
    }
  });
  return rows;
}
// Persons (Tabs → Player data): one person, many usernames across sites.
// An alias with a site only matches that site; a blank site matches anywhere.
const normPersons = (persons) => (persons || []).map((p) => {
  const aliases = Array.isArray(p.aliases) ? p.aliases : (p.usernames || []).map((u) => ({ name: u, site: "" }));
  return { ...p, aliases, usernames: aliases.map((a) => a.name), notes: p.notes || "", kind: p.kind || "player" };
});
const makeNameMapper = (persons) => {
  const exact = {}, any = {};
  normPersons(persons).forEach((per) => per.aliases.forEach((a) => {
    const k = String(a.name).trim().toLowerCase();
    if (!k) return;
    if (a.site) exact[k + "|" + a.site.toLowerCase()] = per.name; else any[k] = per.name;
  }));
  return (name, site) => {
    const k = String(name).trim().toLowerCase();
    return (site && exact[k + "|" + String(site).toLowerCase()]) || any[k] || name;
  };
};
async function loadPersons() {
  try { const c = await store.get("tabs-v1"); if (c?.value) { const d = JSON.parse(c.value); if (Array.isArray(d.persons)) return normPersons(d.persons); } } catch (e) {}
  // Not migrated yet — fall back to My Clubs bundles.
  try { const c = await store.get("agentclubs-v3"); if (c?.value) { const a = JSON.parse(c.value); return normPersons(a.personAliases || []); } } catch (e) {}
  return [];
}

// ════════════════════════════════════════════════════════════════
// MY CLUBS — personal agent downlines (manual weekly entry)
// ════════════════════════════════════════════════════════════════
// Per player:  tipback = tips × TB%   (if TR on)
//              net     = P&L + tipback
//              TR      = TR% × base   (if TR on) — taxes wins, rebates losses
//                        base: net (P&L+tipback) · P&L only · P&L+tips (gross)
//              settle  = (net − TR) × (1 − rebate%) × conversion
// Per club (what YOU get from the club, the revenue side): same shape with the
// club's TB% and TR% — your margin is the difference.

const uid = () => Math.random().toString(36).slice(2, 9);
const SEED_CLUB_NAMES = ["Pumpkin", "Blackwater", "Betflix", "Socal", "Vans", "Don't Tilt", "Rafolini", "Bouncy Castle", "Aces Fortune", "Straddle Up", "Straddle Up (Xander)", "Ace Chasers", "Pineapple PC", "Trap City", "TPA Tiny", "OneTime", "45th Street"];
const FORMULA_VARS = ["pnl", "tips", "tb", "tr", "rebate", "tipback", "net", "gross"];
const FORMULA_HELP = "pnl · tips · tb · tr · tipback (tips×tb%) · net (pnl+tipback) · gross (pnl+tips) — plus min, max, abs, round";
const DEFAULT_FORMULA = "net - tr/100*net";
const NEW_CLUB = (name) => ({
  id: uid(), name, conv: 100, actionBase: "net", useFormula: false, formula: DEFAULT_FORMULA,
  clubTB: 80, clubAction: 0,
  players: [],
});
// Older saves used toggles + a separate rebate; fold everything into tb/tr.
function normalizeAgent(a) {
  const clubs = (a.clubs || []).map((c) => {
    const clubTB = c.useClubTB === false ? 0 : c.clubTB || 0;
    const clubAction = c.useClubAction === undefined ? c.clubAction || 0 : (c.useClubAction ? c.clubAction || 0 : 0);
    const players = (c.players || []).map((p) => {
      const tb = p.useTB === false ? 0 : p.tb || 0;
      let tr = p.useAction === undefined ? p.tr ?? p.action ?? 0 : (p.useAction ? p.action || 0 : 0);
      if (!tr && p.rebate) tr = p.rebate; // same knob under the old split naming
      // id/name/tb/tr are rebuilt above from legacy/alt field names, but any
      // other current field (unifiedDealId being the one that matters) has to
      // be carried through as-is — this function reruns on every load, so
      // anything not explicitly kept here gets silently wiped on the next
      // remount/reload instead of just once during a real migration.
      return { ...p, id: p.id, name: p.name, tb, tr };
    });
    const { def, useClubTB, useClubAction, ...rest } = c;
    return { ...rest, clubTB, clubAction, players };
  });
  return { ...a, clubs };
}
const AGENT_DEFAULT = { clubs: SEED_CLUB_NAMES.map((n) => ({ ...NEW_CLUB(n), id: n.toLowerCase().replace(/[^a-z0-9]+/g, "-") })), umbrellas: [], weeks: {}, currentWeek: "", personAliases: [], myAccounts: [] };

const dealLabel = (p) =>
  [(p.tb || 0) > 0 ? `TB ${p.tb}%` : "no TB", (p.tr || 0) > 0 ? `TR ${p.tr}%` : null].filter(Boolean).join(" · ");

const TR_BASES = [["net", "of net (P&L + tip back)"], ["pnl", "of P&L only"], ["gross", "of P&L + tips"]];
const trBaseLabel = (b) => (b === "pnl" ? "P&L" : b === "gross" ? "P&L+tips" : "net");

// Custom formula support. Variables are plain numbers; percentages come in as
// percent values (tb 75 means 75%). Returns the settlement before conversion.
function compileFormula(expr) {
  if (!expr || !expr.trim()) return null;
  if (!/^[0-9a-zA-Z_+\-*/().,%\s<>=?:&|!]+$/.test(expr)) throw new Error("Invalid character in formula");
  const allowed = new Set([...FORMULA_VARS, "min", "max", "abs", "round"]);
  for (const id of expr.match(/[a-zA-Z_][a-zA-Z0-9_]*/g) || []) {
    if (!allowed.has(id)) throw new Error(`Unknown name: ${id}`);
  }
  const fn = new Function(...FORMULA_VARS, "min", "max", "abs", "round", `"use strict"; return (${expr});`);
  const probe = fn(100, 100, 75, 10, 0, 75, 175, 200, Math.min, Math.max, Math.abs, Math.round);
  if (typeof probe !== "number" || !isFinite(probe)) throw new Error("Formula must produce a number");
  return (v) => fn(v.pnl, v.tips, v.tb, v.tr, v.rebate, v.tipback, v.net, v.gross, Math.min, Math.max, Math.abs, Math.round);
}

function settleLine(pnl, tips, d, club, conv) {
  const tipback = (tips * (d.tb || 0)) / 100;
  const net = pnl + tipback;
  const gross = pnl + tips;
  const tr = d.tr || 0;
  if (club._fn) {
    let s;
    try { s = club._fn({ pnl, tips, tb: d.tb || 0, tr, rebate: 0, tipback, net, gross }); }
    catch (e) { s = net; }
    if (!isFinite(s)) s = net;
    return { tipback, net, actionCut: net - s, settlement: s * conv };
  }
  const trBase = club.actionBase === "pnl" ? pnl : club.actionBase === "gross" ? gross : net;
  const actionCut = (trBase * tr) / 100;
  return { tipback, net, actionCut, settlement: (net - actionCut) * conv };
}

function computeAgent(acfg, week) {
  const entries = week?.entries || {};
  const adjustments = week?.adjustments || [];
  const myAccSet = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
  const clubs = acfg.clubs.map((club) => {
    const conv = (club.conv ?? 100) / 100;
    const base = club.actionBase || "net";
    let _fn = null, formulaError = null;
    if (club.useFormula) {
      try { _fn = compileFormula(club.formula); } catch (e) { formulaError = e.message; }
    }
    const ctx = { ...club, actionBase: base, _fn };
    const players = club.players.map((p) => {
      const e = entries[p.id] || {};
      const pnl = +e.pnl || 0, tips = +e.tips || 0;
      const played = pnl !== 0 || tips !== 0;
      const isMine = myAccSet.has(p.name.trim().toLowerCase());
      const clubDeal = { tb: club.clubTB || 0, tr: club.clubAction || 0 };
      const eff = isMine ? clubDeal : p; // your own accounts ride the club's deal automatically
      const mine = settleLine(pnl, tips, eff, ctx, conv);
      // what the club pays you for this player's action
      const clubSide = settleLine(pnl, tips, clubDeal, ctx, conv);
      return { ...p, tb: eff.tb, tr: eff.tr, isMine, clubId: club.id, clubName: club.name, pnl, tips, played,
        tipback: mine.tipback, net: mine.net, actionCut: mine.actionCut * conv,
        settlement: mine.settlement, clubValue: clubSide.settlement,
        margin: clubSide.settlement - mine.settlement,
        rakeMargin: ((tips * ((club.clubTB || 0) - (eff.tb || 0))) / 100) * conv,
        // margin contribution: what you charge the player minus what the club charges you
        actionMargin: (mine.actionCut - clubSide.actionCut) * conv };
    });
    const sum = (f) => players.reduce((a, p) => a + p[f], 0);
    const clubAdj = adjustments.filter((a) => a.clubId === club.id).reduce((a, x) => a + (+x.amount || 0), 0);
    const clubSettlement = sum("clubValue") + clubAdj;
    return { ...club, conv, base, formulaError, playersC: players, pnl: sum("pnl"), tips: sum("tips"),
      settlements: sum("settlement"), clubValue: sum("clubValue"), rakeMargin: sum("rakeMargin"),
      actionMargin: sum("actionMargin"), margin: sum("margin"),
      clubAdj, clubSettlement, active: players.some((p) => p.played) || clubAdj !== 0 };
  });
  const allPlayers = clubs.flatMap((c) => c.playersC);
  const umbrellas = (acfg.umbrellas || []).map((u) => {
    const members = allPlayers.filter((p) => u.playerIds.includes(p.id));
    const played = members.filter((m) => m.played);
    return { ...u, members, played, settlement: played.reduce((a, m) => a + m.settlement, 0) };
  });
  const inUmbrella = new Set((acfg.umbrellas || []).flatMap((u) => u.playerIds));
  const globalAdj = adjustments.filter((a) => !a.clubId);
  const globalAdjTotal = globalAdj.reduce((a, x) => a + (+x.amount || 0), 0);
  const T = (f) => clubs.reduce((a, c) => a + c[f], 0);
  const totals = { pnl: T("pnl"), tips: T("tips"), settlements: T("settlements"), clubValue: T("clubValue"),
    rakeMargin: T("rakeMargin"), actionMargin: T("actionMargin"), margin: T("margin"),
    clubSettlements: T("clubSettlement"), globalAdjTotal };
  totals.cashNet = totals.clubSettlements - totals.settlements + totals.globalAdjTotal;
  return { clubs, allPlayers, umbrellas, inUmbrella, totals, globalAdj };
}

// v1 (flat fields) or v2 (deal models) → v3 inline fields
function migrateAgent(old) {
  const models = old.models || null;
  const clubs = (old.clubs || []).map((c) => {
    const players = (c.players || []).map((p) => {
      let tb = p.tb, action = p.action ?? 0, rebate = p.rebate ?? 0;
      if (models) {
        const m = models.find((x) => x.id === p.modelId) || models.find((x) => x.id === c.defaultModelId) || {};
        tb = p.tbOverride ?? m.tb ?? 75; action = m.action ?? 0; rebate = m.rebate ?? 0;
      }
      tb = tb ?? 75;
      return { id: p.id || uid(), name: p.name, tb, action, rebate, useTB: (tb || 0) > 0, useAction: (action || 0) > 0 };
    });
    return { ...NEW_CLUB(c.name), id: c.id, name: c.name, conv: c.conv ?? 100,
      clubTB: c.clubTB ?? 80, clubAction: 0, players };
  });
  return { clubs, umbrellas: old.umbrellas || [], weeks: old.weeks || {}, currentWeek: old.currentWeek || "" };
}

function AgentClubs({ theme }) {
  const [acfg, setAcfg] = useState(AGENT_DEFAULT);
  const [tab, setTab] = useState("entry");
  const [loaded, setLoaded] = useState(false);
  const [exportData, setExportData] = useState(null);
  const [ft, setFt] = useState(null);
  const [persons, setPersons] = useState([]);
  useEffect(() => { (async () => { setFt(await loadFishTankModel()); setPersons(await loadPersons()); })(); }, []);

  useEffect(() => {
    (async () => {
      for (const [key, needsMig] of [["agentclubs-v3", false], ["agentclubs-v2", true], ["agentclubs-v1", true]]) {
        try {
          const c = await store.get(key);
          if (c?.value) {
            const parsed = JSON.parse(c.value);
            const next = normalizeAgent(needsMig ? migrateAgent(parsed) : { ...AGENT_DEFAULT, ...parsed });
            setAcfg(next);
            if (needsMig) { try { await store.set("agentclubs-v3", JSON.stringify(next)); } catch (e) {} }
            setLoaded(true); return;
          }
        } catch (e) {}
      }
      setLoaded(true);
    })();
  }, []);
  const save = async (next) => { setAcfg(next); try { await store.set("agentclubs-v3", JSON.stringify(next)); } catch (e) {} };
  const up = (patch) => save({ ...acfg, ...patch });

  const weekKeys = Object.keys(acfg.weeks);
  const wk = acfg.currentWeek && acfg.weeks[acfg.currentWeek] ? acfg.currentWeek : weekKeys[weekKeys.length - 1] || "";
  const week = acfg.weeks[wk];
  const model = useMemo(() => computeAgent(acfg, week), [acfg, week]);

  const newWeek = () => {
    const label = window.prompt("Week label (e.g. 07/20 - 07/26):");
    if (!label || acfg.weeks[label]) return;
    up({ weeks: { ...acfg.weeks, [label]: { entries: {}, adjustments: [] } }, currentWeek: label });
  };
  const setEntry = (pid, field, v) => {
    const w = acfg.weeks[wk]; if (!w) return;
    up({ weeks: { ...acfg.weeks, [wk]: { ...w, entries: { ...w.entries, [pid]: { ...(w.entries[pid] || {}), [field]: v } } } } });
  };
  const setAdjs = (adjustments) => {
    const w = acfg.weeks[wk]; if (!w) return;
    up({ weeks: { ...acfg.weeks, [wk]: { ...w, adjustments } } });
  };
  // Lock is just a soft freeze on this week's entry/adjustment inputs (My
  // Clubs has no local makeup balance to snapshot/roll like Fish Tank/AA) —
  // toggle it off to fix a number, then back on when you're done.
  const setLocked = (locked) => { const w = acfg.weeks[wk]; if (!w) return; up({ weeks: { ...acfg.weeks, [wk]: { ...w, locked } } }); };

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap", alignItems: "center" }}>
        {[["entry", "Weekly entry"], ["summary", "Summary & reports"], ["clubs", "Clubs & deals"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: tab === k ? C.card : "transparent", color: tab === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2 }}>{label}</button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", paddingBottom: 6 }}>
          <select value={wk} onChange={(e) => up({ currentWeek: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
            {weekKeys.length === 0 && <option value="">No weeks yet</option>}
            {weekKeys.map((k) => <option key={k} value={k}>{k}</option>)}
          </select>
          {wk && <button title="Delete this week" onClick={() => {
            if (!window.confirm(`Delete week "${wk}" and all its entered data? This can't be undone.`)) return;
            const weeks = { ...acfg.weeks }; delete weeks[wk];
            const rest = Object.keys(weeks);
            up({ weeks, currentWeek: rest[rest.length - 1] || "" });
          }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>}
          {week && (week.locked
            ? <Btn tone="ghost" small onClick={() => setLocked(false)}>🔒 Week locked · unlock to edit</Btn>
            : <Btn tone="ghost" small onClick={() => setLocked(true)}>Lock week</Btn>)}
          <Btn tone="gold" small onClick={newWeek}>+ New week</Btn>
        </div>
      </div>
      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {!week && tab !== "clubs" && (
          <div style={{ background: C.card, border: `1px dashed ${C.gold}`, borderRadius: 10, padding: "44px 30px", textAlign: "center" }}>
            <div style={{ fontFamily: "Georgia, serif", fontSize: 20, marginBottom: 8 }}>Start a week</div>
            <div style={{ color: C.mute, fontSize: 14, marginBottom: 16 }}>Create a week, then type each player's P&L and tips.</div>
            <Btn onClick={newWeek}>+ New week</Btn>
          </div>
        )}
        {week && tab === "entry" && <AgentEntry model={model} setEntry={setEntry} setAdjs={setAdjs} week={week} acfg={acfg} />}
        {week && tab === "summary" && <AgentSummary model={model} wk={wk} setExportData={setExportData} acfg={acfg} up={up} ft={ft} persons={persons} />}
        {tab === "clubs" && <AgentClubsSetup acfg={acfg} up={up} />}
      </div>
    </div>
  );
}

function AgentEntry({ model, setEntry, setAdjs, week, acfg }) {
  const [adjClub, setAdjClub] = useState("");
  const adjustments = week.adjustments || [];
  const locked = !!week.locked;
  return (
    <div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>
        Enter each player's <b>P&L</b> and <b>Tips</b> in club currency; blank = no play. <span style={{ color: C.green, fontWeight: 700 }}>Green</span> = you pay them · <span style={{ color: C.red, fontWeight: 700 }}>red</span> = they pay you. <b>Margin</b> is what you keep after the club pays you for that player.
        {locked && <span style={{ color: C.goldDark, fontWeight: 700 }}> 🔒 This week is locked — unlock it above to edit.</span>}
      </div>
      {model.clubs.filter((c) => c.players.length > 0).map((c) => (
        <div key={c.id} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", background: C.cream, flexWrap: "wrap" }}>
            <span style={{ fontWeight: 700, fontSize: 14.5 }}>{c.name}</span>
            {c.conv !== 1 && <Pill tone="blue">conv {n2(c.conv * 100)}%</Pill>}
            <span style={{ color: C.mute, fontSize: 12 }}>
              from club: TB {c.clubTB || 0}%{(c.clubAction || 0) > 0 ? ` · TR ${c.clubAction}%` : ""}
            </span>
            {c.useFormula && <Pill tone={c.formulaError ? "red" : "blue"}>{c.formulaError ? "formula error" : "custom formula"}</Pill>}
            <span style={{ marginLeft: "auto", fontSize: 12.5 }}>players {money(c.settlements)} · club {money(c.clubSettlement)} · margin {money(c.margin)}</span>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={{ ...th, textAlign: "left" }}>Player</th>
              <th style={{ ...th, textAlign: "left" }}>Deal</th>
              <th style={th}>P&L</th><th style={th}>Tips</th>
              <th style={th}>Tipback</th><th style={th}>Settlement</th><th style={th}>Your margin</th>
            </tr></thead>
            <tbody>
              {c.playersC.map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}{model.inUmbrella.has(p.id) && <span style={{ marginLeft: 6 }}><Pill tone="blue">{(acfg.umbrellas.find((u) => u.playerIds.includes(p.id)) || {}).name}</Pill></span>}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 11.5 }}>{dealLabel(p)}</td>
                  <td style={td}><NumInput width={92} value={(week.entries[p.id] || {}).pnl ?? ""} onChange={(v) => setEntry(p.id, "pnl", v)} disabled={locked} /></td>
                  <td style={td}><NumInput width={82} value={(week.entries[p.id] || {}).tips ?? ""} onChange={(v) => setEntry(p.id, "tips", v)} disabled={locked} /></td>
                  <td style={td}>{p.played ? fmt(p.tipback) : "—"}</td>
                  <td style={td}>{p.played ? money(p.settlement) : "—"}</td>
                  <td style={td}>{p.played ? money(p.margin) : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      <Card title="Adjustments (stakes, transfers, one-offs)" right={
        <span style={{ display: "flex", gap: 8 }}>
          <select value={adjClub} onChange={(e) => setAdjClub(e.target.value)} disabled={locked} style={{ ...inputS, fontSize: 12 }}>
            <option value="">Not club-specific</option>
            {acfg.clubs.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <Btn tone="ghost" small disabled={locked} onClick={() => setAdjs([...adjustments, { id: uid(), clubId: adjClub || null, label: "New item", amount: 0 }])}>+ Add</Btn>
        </span>
      }>
        {adjustments.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None this week. Positive = money to you, negative = money out.</div>}
        {adjustments.map((a) => (
          <div key={a.id} style={{ display: "flex", gap: 10, alignItems: "center", borderTop: `1px solid ${C.line}`, padding: "8px 0" }}>
            <input value={a.label} disabled={locked} onChange={(e) => setAdjs(adjustments.map((x) => x.id === a.id ? { ...x, label: e.target.value } : x))} style={{ ...inputS, flex: 1, opacity: locked ? 0.45 : 1 }} />
            <span style={{ color: C.mute, fontSize: 12 }}>{a.clubId ? (acfg.clubs.find((c) => c.id === a.clubId)?.name || "?") : "general"}</span>
            <NumInput width={100} value={a.amount} onChange={(v) => setAdjs(adjustments.map((x) => x.id === a.id ? { ...x, amount: v } : x))} disabled={locked} />
            {!locked && <button onClick={() => setAdjs(adjustments.filter((x) => x.id !== a.id))} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>}
          </div>
        ))}
      </Card>
    </div>
  );
}

function AgentSummary({ model, wk, setExportData, acfg, up, ft, persons }) {
  const t = model.totals;
  const active = model.clubs.filter((c) => c.active);
  const [reportSel, setReportSel] = useState("");
  const bundled = new Set(persons.flatMap((p) => p.usernames.map((u) => u.trim().toLowerCase())));
  const ftNames = ft ? [...new Set(ft.model.entities.flatMap((e) => e.type === "backed" ? [e.name] : e.members.map((m) => m.name)))] : [];
  const rawNames = [...new Set([...model.allPlayers.map((p) => p.name.trim()), ...ftNames.map((n) => n.trim())])]
    .filter((n) => !bundled.has(n.toLowerCase())).sort((a, b) => a.localeCompare(b));

  // resolve selection → username set + display name
  let reportLabel = "", nameSet = new Set();
  if (reportSel.startsWith("person:")) {
    const per = persons.find((p) => p.id === reportSel.slice(7));
    if (per) { reportLabel = per.name; nameSet = new Set(per.usernames.map((u) => u.trim().toLowerCase())); }
  } else if (reportSel.startsWith("u:")) {
    reportLabel = reportSel.slice(2); nameSet = new Set([reportLabel.toLowerCase()]);
  }
  const mcRows = model.allPlayers.filter((p) => nameSet.has(p.name.trim().toLowerCase()) && p.played);
  const ftRows = ftRowsForNames(ft, nameSet);
  // clubs owned by this person also fold into their report (sign flipped: positive = you pay them)
  const clubRows = model.clubs.filter((c) => c.active && (c.owner || "").trim() && (nameSet.has(c.owner.trim().toLowerCase()) || c.owner.trim().toLowerCase() === reportLabel.toLowerCase()))
    .map((c) => ({ id: "club-" + c.id, clubName: c.name, name: c.owner, customDeal: "club settlement", pnl: c.pnl, tips: c.tips, tipback: 0, settlement: -c.clubSettlement, margin: 0, played: true }));
  const reportRows = [...ftRows, ...mcRows, ...clubRows];
  const reportTotal = reportRows.reduce((a, p) => a + p.settlement, 0);
  // my play + fish tank side for the total card
  const myAcc = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
  const myPlay = model.allPlayers.filter((p) => p.played && myAcc.has(p.name.trim().toLowerCase()));
  const myPlayTotal = myPlay.reduce((a, p) => a + p.settlement, 0);
  const ftEntitle = ft ? ft.model.entitle.ak : null;
  const weekTotal = (ftEntitle || 0) + t.margin + myPlayTotal + t.globalAdjTotal;
  const row = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "6px 0", fontSize: 13.5 }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{money(val)}</span>
    </div>
  );
  const copyRows = (title, rows) => setExportData({ title, text: toTSV(["Club", "Player", "Deal", "P&L", "Tips", "Tipback", "Settlement", "Your margin"], rows.map((p) => [p.clubName, p.name, p.customDeal || dealLabel(p), p.pnl.toFixed(2), p.tips.toFixed(2), p.tipback.toFixed(2), p.settlement.toFixed(2), (p.margin || 0).toFixed(2)])) });

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>Week summary · {wk}</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <Btn tone="gold" small onClick={() => downloadAgentWorkbook(model, wk)}>Download Excel</Btn>
          <Btn tone="ghost" small onClick={() => copyRows(`My clubs · ${wk}`, model.allPlayers.filter((p) => p.played))}>Copy</Btn>
        </div>
      </div>
      <div style={{ marginBottom: 14 }}>
        <Card title={`Total week P&L${ft?.period ? ` · Fish Tank ${ft.period}` : ""}`}>
          {ftEntitle == null
            ? <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 4 }}>No Fish Tank week loaded — upload one in the Fish Tank mode and it appears here automatically.</div>
            : row("Fish Tank — your entitlement (own play + backed books + ½ profit)", ftEntitle)}
          {row("My Clubs — margin on players", t.margin)}
          {row(`My play on other clubs${myPlay.length ? ` (${myPlay.map((p) => p.name).join(", ")})` : ""}`, myPlayTotal)}
          {myAcc.size === 0 && <div style={{ color: C.mute, fontSize: 11.5, margin: "2px 0 4px" }}>List your own usernames under "My accounts" in Clubs & deals to count your personal play here.</div>}
          {row("General adjustments", t.globalAdjTotal)}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("TOTAL", weekTotal, { bold: true })}
        </Card>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14, marginBottom: 16 }}>
        <Card title="Your week">
          {row("Club side (clubs pay you)", t.clubSettlements)}
          {row("Player settlements", t.settlements)}
          {row("General adjustments", t.globalAdjTotal)}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("Cash net", t.cashNet, { bold: true })}
        </Card>
        <Card title="Where the margin came from">
          {row("TB margin (club TB − player TB)", t.rakeMargin)}
          {row("TR margin (player TR − club TR)", t.actionMargin)}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("Total margin on players", t.margin, { bold: true })}
        </Card>
      </div>
      <div style={{ marginBottom: 16 }}>
        <Card title="Player report" right={
          <span style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <select value={reportSel} onChange={(e) => setReportSel(e.target.value)} style={{ ...inputS, fontSize: 12.5, minWidth: 180 }}>
              <option value="">Pick a player…</option>
              {persons.length > 0 && <optgroup label="People (bundles)">
                {persons.map((p) => <option key={p.id} value={"person:" + p.id}>{p.name} ({p.usernames.length})</option>)}
              </optgroup>}
              <optgroup label="Usernames">
                {rawNames.map((n) => <option key={n} value={"u:" + n}>{n}</option>)}
              </optgroup>
            </select>
            {reportRows.length > 0 && <Btn tone="gold" small onClick={() => downloadPlayerExcel(reportLabel, reportRows, wk)}>Excel</Btn>}
            {reportRows.length > 0 && <Btn tone="ghost" small onClick={() => copyRows(`${reportLabel} · ${wk}`, reportRows)}>Copy</Btn>}
          </span>
        }>
          {!reportSel && <div style={{ color: C.mute, fontSize: 13 }}>One report across Fish Tank and every club — pick a username or a bundled person (bundles are managed in Tabs → Player data). Clubs whose owner matches also fold in.</div>}
          {reportSel && reportRows.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No activity recorded for {reportLabel} this week.</div>}
          {reportRows.length > 0 && (
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <tbody>
                {reportRows.map((p, i) => (
                  <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                    <td style={{ ...tdL, fontWeight: 600 }}>{p.clubName}{p.name && p.name.toLowerCase() !== reportLabel.toLowerCase() ? <span style={{ color: C.mute, fontWeight: 400, fontSize: 11 }}> · {p.name}</span> : null}</td>
                    <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{p.customDeal || dealLabel(p)}</td>
                    <td style={td}>P&L {fmt(p.pnl)}</td>
                    <td style={td}>tips {fmt(p.tips)}</td>
                    <td style={td}>{money(p.settlement)}</td>
                  </tr>
                ))}
                <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
                  <td style={{ ...tdL, fontWeight: 700 }} colSpan={4}>TOTAL</td>
                  <td style={{ ...td, fontWeight: 700 }}>{money(reportTotal)}</td>
                </tr>
              </tbody>
            </table>
          )}
        </Card>
      </div>
      {model.umbrellas.filter((u) => u.played.length > 0).map((u) => (
        <div key={u.id} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", background: C.cream }}>
            <span style={{ fontWeight: 700 }}>{u.name}</span>
            <Pill tone="blue">umbrella · {u.played.length} lines</Pill>
            <span style={{ marginLeft: "auto", fontSize: 12.5 }}>settles as one {money(u.settlement)}</span>
            <Btn tone="gold" small onClick={() => downloadPlayerExcel(u.name, u.played, wk)}>Excel</Btn>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              {u.played.map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{p.clubName}</td>
                  <td style={td}>P&L {fmt(p.pnl)}</td>
                  <td style={td}>tips {fmt(p.tips)}</td>
                  <td style={td}>{money(p.settlement)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      {active.map((c) => (
        <div key={c.id} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", background: C.cream, flexWrap: "wrap" }}>
            <span style={{ fontWeight: 700 }}>{c.name}</span>
            <span style={{ color: C.mute, fontSize: 12 }}>tips {fmt(c.tips)} · P&L {fmt(c.pnl)}</span>
            <span style={{ marginLeft: "auto", fontSize: 12.5 }}>club {money(c.clubSettlement)} · margin {money(c.margin)}</span>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              {c.playersC.filter((p) => p.played).map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600, width: "24%" }}>{p.name}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{dealLabel(p)}</td>
                  <td style={td}>P&L {fmt(p.pnl)}</td>
                  <td style={td}>tips {fmt(p.tips)}</td>
                  <td style={td}>{money(p.settlement)}</td>
                  <td style={td}>margin {money(p.margin)}</td>
                </tr>
              ))}
              {c.clubAdj !== 0 && (
                <tr style={{ borderTop: `1px solid ${C.line}` }}><td style={{ ...tdL, color: C.mute }} colSpan={5}>Club adjustments</td><td style={td}>{money(c.clubAdj)}</td></tr>
              )}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}

// ———— Clubs & deals: inline fields, toggles, club-side revenue ————
const Toggle = ({ on, onClick, label }) => (
  <button onClick={onClick} style={{
    border: `1px solid ${on ? C.goldDark : C.line}`, background: on ? C.gold : C.surface,
    color: on ? "var(--onGold)" : C.mute, borderRadius: 12, padding: "2px 10px",
    fontSize: 11, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>{label}</button>
);

function FormulaEditor({ club, setClub }) {
  const [txt, setTxt] = useState(club.formula || DEFAULT_FORMULA);
  useEffect(() => setTxt(club.formula || DEFAULT_FORMULA), [club.formula]);
  let err = null, preview = null;
  try {
    const fn = compileFormula(txt);
    if (fn) {
      const pnl = -1000, tips = 500, tb = club.clubTB || 0, tr = club.clubAction || 0;
      const tipback = (tips * tb) / 100;
      preview = fn({ pnl, tips, tb, tr, rebate: 0, tipback, net: pnl + tipback, gross: pnl + tips });
    }
  } catch (e) { err = e.message; }
  return (
    <div style={{ padding: "10px 16px", borderBottom: `1px solid ${C.line}`, background: C.rowAlt }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.06em", minWidth: 120 }}>Formula</span>
        <input value={txt} onChange={(e) => setTxt(e.target.value)}
          onBlur={() => setClub(club.id, { formula: txt })}
          onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
          spellCheck={false}
          style={{ ...inputS, flex: 1, minWidth: 260, fontFamily: "ui-monospace, Menlo, monospace", fontSize: 12.5, borderColor: err ? C.red : C.gold }} />
        {err
          ? <span style={{ color: C.red, fontSize: 11.5 }}>{err}</span>
          : <span style={{ color: C.mute, fontSize: 11.5 }}>P&L −1,000 / tips 500 → <b style={{ color: C.ink }}>{preview == null ? "—" : fmt(preview)}</b></span>}
      </div>
      <div style={{ color: C.mute, fontSize: 11, marginTop: 5 }}>
        Variables: {FORMULA_HELP}. This IS the club's TR formula — players and the club row run through it with their own tb / tr. Percentages are numbers, so use <code>tr/100</code>.
      </div>
      <div style={{ color: C.mute, fontSize: 11, marginTop: 3 }}>
        Examples: <code>net - tr/100*net</code> · <code>net - tr/100*gross</code> · <code>net - tr/100*max(0, net)</code>
      </div>
    </div>
  );
}

function AgentClubsSetup({ acfg, up }) {
  const [newClub, setNewClub] = useState("");
  const [umbName, setUmbName] = useState("");
  const setClub = (id, patch) => up({ clubs: acfg.clubs.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const addClub = () => { if (!newClub.trim()) return; up({ clubs: [...acfg.clubs, NEW_CLUB(newClub.trim())] }); setNewClub(""); };
  const delClub = (cid) => { if (window.confirm("Delete this club and its roster?")) up({ clubs: acfg.clubs.filter((c) => c.id !== cid) }); };
  const addPlayer = (c) => setClub(c.id, { players: [...c.players, { id: uid(), name: "New player", tb: c.clubTB || 0, tr: c.clubAction || 0 }] });
  const setPlayer = (cid, pid, patch) => setClub(cid, { players: acfg.clubs.find((c) => c.id === cid).players.map((p) => (p.id === pid ? { ...p, ...patch } : p)) });
  const delPlayer = (cid, pid) => setClub(cid, { players: acfg.clubs.find((c) => c.id === cid).players.filter((p) => p.id !== pid) });
  const myAccSet = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
  const addUmb = () => { if (!umbName.trim()) return; up({ umbrellas: [...(acfg.umbrellas || []), { id: uid(), name: umbName.trim(), playerIds: [] }] }); setUmbName(""); };
  const togglePlayerInUmb = (uid_, pid) => up({
    umbrellas: acfg.umbrellas.map((u) => u.id === uid_
      ? { ...u, playerIds: u.playerIds.includes(pid) ? u.playerIds.filter((x) => x !== pid) : [...u.playerIds, pid] }
      : { ...u, playerIds: u.playerIds.filter((x) => x !== pid) }) });
  const allP = acfg.clubs.flatMap((c) => c.players.map((p) => ({ ...p, clubName: c.name })));

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 14, marginBottom: 16, alignItems: "start" }}>
        <Card title="Umbrellas (bunch players into one settlement)" right={
          <span style={{ display: "flex", gap: 8 }}>
            <input placeholder="New umbrella name…" value={umbName} onChange={(e) => setUmbName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUmb()} style={{ ...inputS, width: 180 }} />
            <Btn tone="ghost" small onClick={addUmb}>+ Create</Btn>
          </span>
        }>
          {(acfg.umbrellas || []).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None yet. Groups players across any clubs into one settlement and one report.</div>}
          {(acfg.umbrellas || []).map((u) => (
            <div key={u.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <input value={u.name} onChange={(e) => up({ umbrellas: acfg.umbrellas.map((x) => x.id === u.id ? { ...x, name: e.target.value } : x) })} style={{ ...inputS, fontWeight: 700, width: 200 }} />
                <Pill tone="blue">{u.playerIds.length} players</Pill>
                <button onClick={() => up({ umbrellas: acfg.umbrellas.filter((x) => x.id !== u.id) })} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>× delete</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {allP.map((p) => {
                  const inThis = u.playerIds.includes(p.id);
                  const inOther = !inThis && acfg.umbrellas.some((x) => x.id !== u.id && x.playerIds.includes(p.id));
                  return (
                    <button key={p.id} onClick={() => !inOther && togglePlayerInUmb(u.id, p.id)} style={{
                      padding: "3px 10px", borderRadius: 12, fontSize: 11.5, fontWeight: 600, cursor: inOther ? "default" : "pointer",
                      border: `1px solid ${inThis ? C.goldDark : C.line}`,
                      background: inThis ? C.gold : C.surface, color: inThis ? "var(--onGold)" : inOther ? "var(--chipOff)" : C.mute, opacity: inOther ? 0.55 : 1 }}>
                      {p.name} · {p.clubName}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </Card>
        <Card title="My accounts">
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 8 }}>Your own usernames across these clubs — their play counts as "my play" in the week summary. One per line.</div>
          <textarea defaultValue={(acfg.myAccounts || []).join("\n")}
            onBlur={(e) => up({ myAccounts: e.target.value.split("\n").map((s) => s.trim()).filter(Boolean) })}
            rows={5} style={{ ...inputS, width: "100%", boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
        </Card>
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 12, alignItems: "center", flexWrap: "wrap" }}>
        <input placeholder="New club name…" value={newClub} onChange={(e) => setNewClub(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addClub()} style={{ ...inputS, width: 240 }} />
        <Btn tone="gold" small onClick={addClub}>+ Add club</Btn>
        <div style={{ marginLeft: "auto", color: C.mute, fontSize: 12 }}><b>TB</b> = tip back · <b>TR</b> = tax rebate (taxes wins, rebates losses). Toggle either off to drop it from the deal. Conversion: value of 1 club unit as % (100 = 1:1). Paste a Tabs → Staking makeup deal's ID into a player's <b>Unified deal</b> to fold their weekly net (P&L + rakeback) here into a shared pool with their Fish Tank/AA stake instead of settling it on its own (not supported for players inside an umbrella).</div>
      </div>

      {acfg.clubs.map((c) => {
        return (
          <div key={c.id} style={{ background: C.card, borderRadius: 10, marginBottom: 12, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
            <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", padding: "10px 16px", background: C.cream }}>
              <input value={c.name} onChange={(e) => setClub(c.id, { name: e.target.value })} style={{ ...inputS, fontWeight: 700, width: 165 }} />
              <span style={{ fontSize: 12, color: C.mute }}>owner</span>
              <input value={c.owner || ""} placeholder="who's behind it" onChange={(e) => setClub(c.id, { owner: e.target.value })} style={{ ...inputS, width: 130, fontSize: 12 }} />
              <span style={{ fontSize: 12, color: C.mute }} title="1 club unit in USD, as a percent — 100 = 1:1">FX → USD</span>
              <NumInput width={64} value={c.conv} onChange={(v) => setClub(c.id, { conv: v || 100 })} />
              <span style={{ fontSize: 11.5, color: C.mute }}>%</span>
              <span style={{ fontSize: 12, color: C.mute }}>TR base</span>
              <select value={c.actionBase || "net"} disabled={!!c.useFormula} onChange={(e) => setClub(c.id, { actionBase: e.target.value })} style={{ ...inputS, padding: "4px 6px", fontSize: 12, opacity: c.useFormula ? 0.4 : 1 }}>
                {TR_BASES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
              <Toggle on={!!c.useFormula} onClick={() => setClub(c.id, { useFormula: !c.useFormula, formula: c.formula || DEFAULT_FORMULA })} label="custom formula" />
              <button onClick={() => delClub(c.id)} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 13 }}>× delete club</button>
            </div>

            {c.useFormula && <FormulaEditor club={c} setClub={setClub} />}

            {/* what the club gives YOU — type a number to turn it on, 0 = off */}
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", padding: "9px 16px", borderBottom: `1px solid ${C.line}`, background: C.rowAlt }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.06em", minWidth: 120 }}>From the club</span>
              <span style={{ fontSize: 11.5, color: C.mute }}>TB</span>
              <PctInput width={52} max={999} value={c.clubTB ?? 0} onChange={(v) => v != null && setClub(c.id, { clubTB: v })} />
              <span style={{ fontSize: 11.5, color: C.mute }}>TR</span>
              <PctInput width={52} value={c.clubAction ?? 0} onChange={(v) => v != null && setClub(c.id, { clubAction: v })} />
              <span style={{ marginLeft: "auto", color: C.mute, fontSize: 11.5 }}>
                your revenue side · new players start on this deal · your accounts always ride it{(c.clubAction || 0) > 0 ? ` · TR ${trBaseLabel(c.actionBase || "net")}` : ""}
              </span>
            </div>

            {c.players.length > 0 && (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>
                  <th style={{ ...th, textAlign: "left" }}>Player</th>
                  <th style={th}>TB %</th>
                  <th style={th}>TR %</th>
                  <th style={{ ...th, textAlign: "left" }}>Formula</th><th style={th}>Unified deal</th><th style={th}></th>
                </tr></thead>
                <tbody>
                  {c.players.map((p, i) => {
                    const isMine = myAccSet.has((p.name || "").trim().toLowerCase());
                    const inUmb = (acfg.umbrellas || []).some((u) => u.playerIds.includes(p.id));
                    return (
                    <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                      <td style={tdL}>
                        <input value={p.name} onChange={(e) => setPlayer(c.id, p.id, { name: e.target.value })} style={{ ...inputS, width: 150, fontSize: 12.5 }} />
                        {isMine && <span style={{ marginLeft: 6 }}><Pill tone="gold">you · club deal</Pill></span>}
                      </td>
                      <td style={td}>{isMine
                        ? <span style={{ color: C.mute, fontSize: 12 }}>{c.clubTB || 0}%</span>
                        : <PctInput width={52} max={999} value={p.tb ?? 0} onChange={(v) => v != null && setPlayer(c.id, p.id, { tb: v })} />}</td>
                      <td style={td}>{isMine
                        ? <span style={{ color: C.mute, fontSize: 12 }}>{c.clubAction || 0}%</span>
                        : <PctInput width={52} value={p.tr ?? 0} onChange={(v) => v != null && setPlayer(c.id, p.id, { tr: v })} />}</td>
                      <td style={{ ...tdL, color: C.goldDark, fontSize: 11, fontFamily: "ui-monospace, Menlo, monospace" }}>
                        {c.useFormula
                          ? c.formula
                          : `(P&L${(isMine ? c.clubTB : p.tb) ? ` + tips×${isMine ? c.clubTB : p.tb}%` : ""}${(isMine ? c.clubAction : p.tr) ? ` − ${isMine ? c.clubAction : p.tr}%×${trBaseLabel(c.actionBase || "net")}` : ""})`}
                      </td>
                      <td style={td}>
                        {inUmb ? <span style={{ color: C.mute, fontSize: 11 }} title="Not supported for players inside an umbrella — umbrella members settle as one group line.">n/a · in umbrella</span> : (
                          <>
                            <input placeholder="deal id" value={p.unifiedDealId || ""} onChange={(e) => setPlayer(c.id, p.id, { unifiedDealId: e.target.value.trim() })} style={{ ...inputS, width: 88, fontSize: 11 }} title="Paste a weekly makeup deal's ID from Tabs → Staking to fold this player's My Clubs weekly net (P&L + rakeback) into that shared pool instead of settling it here on its own." />
                            {p.unifiedDealId && <div style={{ fontSize: 9.5, color: C.goldDark, fontWeight: 700, marginTop: 2 }}>UNIFIED</div>}
                          </>
                        )}
                      </td>
                      <td style={{ ...td, width: 36 }}><button onClick={() => delPlayer(c.id, p.id)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                    </tr>
                  );})}
                </tbody>
              </table>
            )}
            <div style={{ padding: "8px 16px" }}>
              <button onClick={() => addPlayer(c)} style={{ border: `1px dashed ${C.gold}`, background: "transparent", color: C.goldDark, borderRadius: 6, padding: "6px 12px", cursor: "pointer", fontSize: 12.5, fontWeight: 700 }}>+ Add player</button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function playerSheetRows(ws, rows, withClub) {
  rows.forEach((p, i) => {
    const r = ws.addRow([]);
    xText(r.getCell(1), withClub ? p.clubName : withId(p.name, p.id), { bold: true });
    xText(r.getCell(2), withClub ? withId(p.name, p.id) : (p.customDeal || dealLabel(p)), { mute: true });
    xText(r.getCell(3), withClub ? (p.customDeal || dealLabel(p)) : "", { mute: true });
    xMoney(r.getCell(4), p.pnl, { colorSign: false });
    xMoney(r.getCell(5), p.tips, { colorSign: false });
    xMoney(r.getCell(6), p.tipback, { colorSign: false });
    xMoney(r.getCell(7), p.settlement);
    xMoney(r.getCell(8), p.margin || 0);
    if (i % 2 === 1) for (let j = 1; j <= 8; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const tr = ws.addRow([]);
  xText(tr.getCell(1), "TOTAL", { bold: true });
  xMoney(tr.getCell(7), rows.reduce((a, p) => a + p.settlement, 0), { bold: true });
  xMoney(tr.getCell(8), rows.reduce((a, p) => a + (p.margin || 0), 0), { bold: true });
  for (let j = 1; j <= 8; j++) tr.getCell(j).fill = fillOf(XLC.cream);
}
const AG_HEAD = ["Club", "Player", "Deal", "P&L", "Tips", "Tipback", "Settlement", "Your margin"];
async function downloadPlayerExcel(name, rows, wk) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(safeSheetName(name, wb));
  [18, 18, 22, 12, 12, 12, 13, 13].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `${name} — ${wk}`);
  xHeader(ws, AG_HEAD, 3);
  playerSheetRows(ws, rows, true);
  await saveWb(wb, `${name.replace(/[^\w]+/g, "_")}_${wk.replace(/[^\d]/g, "_").replace(/^_+|_+$/g, "") || "week"}.xlsx`);
}
async function downloadAgentWorkbook(model, wk) {
  const wb = new ExcelJS.Workbook();
  const active = model.clubs.filter((c) => c.active);
  const t = model.totals;
  const ws = wb.addWorksheet("Summary");
  [22, 13, 13, 15, 15, 13].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `My clubs — ${wk}`);
  const hr = xHeader(ws, ["Club", "Players P&L", "Tips", "Player settlements", "Club settlement", "Your margin"], 1);
  ws.views = [{ state: "frozen", ySplit: hr }];
  active.forEach((c, i) => {
    const r = ws.addRow([]);
    xText(r.getCell(1), c.name, { bold: true });
    xMoney(r.getCell(2), c.pnl, { colorSign: false });
    xMoney(r.getCell(3), c.tips, { colorSign: false });
    xMoney(r.getCell(4), c.settlements);
    xMoney(r.getCell(5), c.clubSettlement);
    xMoney(r.getCell(6), c.margin);
    if (i % 2 === 1) for (let j = 1; j <= 6; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const gr = ws.addRow([]);
  xText(gr.getCell(1), "TOTAL", { bold: true, white: true });
  [t.pnl, t.tips, t.settlements, t.clubSettlements, t.margin].forEach((v, i) => xMoney(gr.getCell(i + 2), v, { bold: true, white: true, colorSign: false }));
  for (let j = 1; j <= 6; j++) gr.getCell(j).fill = fillOf(XLC.bar);
  ws.addRow([]);
  [["Cash net (club side − players + adjustments)", t.cashNet], ["TB margin", t.rakeMargin], ["TR margin", t.actionMargin]].forEach(([lbl, v]) => {
    const r = ws.addRow([]); xText(r.getCell(1), lbl, { bold: true }); xMoney(r.getCell(2), v, { bold: true });
  });
  model.umbrellas.filter((u) => u.played.length > 0).forEach((u) => {
    const w2 = wb.addWorksheet(safeSheetName(u.name, wb));
    [18, 18, 22, 12, 12, 12, 13, 13].forEach((w, i) => (w2.getColumn(i + 1).width = w));
    xTitle(w2, `${u.name} (umbrella) — ${wk}`);
    xHeader(w2, AG_HEAD, 3);
    playerSheetRows(w2, u.played, true);
  });
  active.forEach((c) => {
    const w2 = wb.addWorksheet(safeSheetName(c.name, wb));
    [18, 22, 12, 12, 12, 12, 13, 13].forEach((w, i) => (w2.getColumn(i + 1).width = w));
    xTitle(w2, `${c.name} — ${wk}`);
    xHeader(w2, ["Player", "Deal", "", "P&L", "Tips", "Tipback", "Settlement", "Your margin"], 3);
    playerSheetRows(w2, c.playersC.filter((p) => p.played), false);
    if (c.clubAdj !== 0) { const ar = w2.addRow([]); xText(ar.getCell(1), "Club adjustments", { mute: true }); xMoney(ar.getCell(7), c.clubAdj); }
    const cr = w2.addRow([]);
    xText(cr.getCell(1), "CLUB SETTLEMENT (you ↔ club)", { bold: true });
    xMoney(cr.getCell(7), c.clubSettlement, { bold: true });
  });
  await saveWb(wb, `MyClubs_${wk.replace(/[^\d]/g, "_").replace(/^_+|_+$/g, "") || "week"}.xlsx`);
}

// ════════════════════════════════════════════════════════════════
// TABS — running ledger of who owes whom (players, clubs, other)
// ════════════════════════════════════════════════════════════════
// Convention everywhere: POSITIVE = they owe you · NEGATIVE = you owe them.
// Weekly settlements push in with one click; everything else is manual
// (crypto vig, sales, outside staking...). Entries flagged as P&L feed
// the side-income tracker.

// ════════════════════════════════════════════════════════════════
// ALL AMERICAN UNION — 3-owner club (AkFish · LeaderZay · Laye)
// ════════════════════════════════════════════════════════════════
// Economics:
//   · Every line's rakeback is paid OUT (to the agent or the player). The
//     MARGIN (tips − rakeback) is the profit on that line.
//   · Agent lines: margin goes into the union pool, split 50/50 AkFish↔Laye.
//   · Personal lines: margin goes 100% to that owner — his player, his spread.
//   · Own accounts: each owner lists his own usernames; their play is logged
//     as that owner's P&L (with full feeback — the union doesn't profit off
//     an owner's own play). No cash to collect from yourself.
//   · Bad beat jackpot: read from the export's "bad beat" line item (editable).
//     Splits 1/3 · 1/3 · 1/3 as PROFIT for every owner, but the cash sits in
//     the jackpot pool — held by no one — so it's excluded from the transfer.
//   · Stake (makeup) / action-buy deals ride outside union profit and the
//     settle-up; only their rake margin stays in the pool.
//   · Settle-up: each owner's collectable profit vs what he actually
//     collected gives exact A-pays-B transfers.
const AA_CFG_KEY = "allamerican-v1";
const AA_WEEK_KEY = "allamerican-lastweek-v1";
// Owner clubs — All American (Bazaar) is the first one; more can be added from
// the header. Each club has its own owner list, deal, storage keys, and week.
const OC_LIST_KEY = "ownerclubs-v1";
const OC_LEGACY_CLUB = { id: "allamerican", name: "All American", cfgKey: AA_CFG_KEY, weekKey: AA_WEEK_KEY, meId: "akfish",
  owners: [{ id: "akfish", label: "AkFish", poolPct: 50, jpPct: 33.34 }, { id: "laye", label: "Laye", poolPct: 50, jpPct: 33.33 }, { id: "leaderzay", label: "LeaderZay", poolPct: 0, jpPct: 33.33 }] };
async function loadOwnerClubs() {
  try {
    const c = await store.get(OC_LIST_KEY);
    if (c?.value) { const v = JSON.parse(c.value); if (Array.isArray(v.clubs)) return v.clubs; }
  } catch (e) {}
  // First run: seed from the legacy All American keys (present or not — it's the default club).
  const clubs = [OC_LEGACY_CLUB];
  try { await store.set(OC_LIST_KEY, JSON.stringify({ clubs })); } catch (e) {}
  return clubs;
}
async function saveOwnerClubs(clubs) { try { await store.set(OC_LIST_KEY, JSON.stringify({ clubs })); } catch (e) {} }
const newOwnerClub = (name) => { const id = "oc-" + uid(); return { id, name, cfgKey: `oc-cfg-${id}`, weekKey: `oc-week-${id}`, meId: "me", owners: [{ id: "me", label: "Ak", poolPct: 100, jpPct: 100 }] }; };
// Per-club helpers: owner ids, labels, and how the pool / jackpot / split-backer books divide.
function ocHelpers(club) {
  const owners = (club && club.owners) || [];
  const ownerIds = owners.map((o) => o.id);
  const find = (o) => owners.find((x) => x.id === o);
  const poolOwners = owners.filter((o) => (+o.poolPct || 0) > 0);
  const poolTot = poolOwners.reduce((a, o) => a + (+o.poolPct || 0), 0) || 1;
  const poolShare = (o) => { const x = find(o); return x && (+x.poolPct || 0) > 0 ? (+x.poolPct) / poolTot : 0; };
  const jpTot = owners.reduce((a, o) => a + (+o.jpPct || 0), 0) || 1;
  const jpShareOf = (o) => { const x = find(o); return x ? (+x.jpPct || 0) / jpTot : 0; };
  const shareOf = (backer, o) => (backer === "split" ? poolShare(o) : backer === o ? 1 : 0);
  const zero = () => Object.fromEntries(ownerIds.map((o) => [o, 0]));
  const pctS = (f) => { const v = Math.round(f * 1000) / 10; return (v % 1 === 0 ? v.toFixed(0) : String(v)) + "%"; };
  const splitLabel = poolOwners.length > 1 ? poolOwners.map((o) => o.label).join(" & ") + " " + poolOwners.map((o) => pctS(poolShare(o.id))).join(" / ") : poolOwners.length === 1 ? `${poolOwners[0].label} 100%` : "Split";
  const lbl = (o) => (o === "split" ? splitLabel : (find(o) || {}).label || o || "?");
  const poolLabel = poolOwners.length > 1 ? `pool ${poolOwners.map((o) => pctS(poolShare(o.id))).join("/")}` : poolOwners.length === 1 ? `pool → ${poolOwners[0].label}` : "pool (unassigned)";
  const tags = [["agent", `Agent · ${poolLabel}`], ...owners.map((o) => [o.id, `Personal · ${o.label}`])];
  return { owners, ownerIds, lbl, poolShare, jpShareOf, shareOf, zero, splitLabel, poolLabel, tags, pctS, hasPool: poolOwners.length > 0, poolOwners, meId: club?.meId };
}
const aaPctN = (v) => Math.round(v * 100) / 100; // keep decimals like 72.5
const aaPctS = (v) => { const r = aaPctN(v); return (r % 1 === 0 ? r.toFixed(0) : String(r)) + "%"; };
// TR (take rate) is a separate cut off a line's own net (P&L + rakeback),
// on top of whatever margin the rakeback spread already generates — e.g. a
// "70/10" deal: 70% TB (rakeback) plus a 10% TR cut of (P&L+RB). Defaults to
// 0 so existing clubs (whose numbers were already correct without it) are
// completely unaffected; a club that runs every player on a TR deal (e.g.
// one with a single owner and no separate staking setup per player) sets
// defaultTR once and every line picks it up automatically.
const AA_DEFAULT_CFG = { defaultTB: 80, defaultTR: 0, deals: {}, ownDeals: {}, tr: {}, playerTr: {}, agentTr: {}, owners: {}, jackpots: {}, jpModes: {}, collectors: {}, backed: {}, ownAccounts: {}, finalizedPeriods: {}, names: {}, dlUmbrellas: [], jpHolds: [] };

// Minimal-transfer settle: positive delta = collected more than their share → pays out.
function aaTransfers(delta, ownerIds) {
  const payers = [], receivers = [];
  ownerIds.forEach((o) => { const v = delta[o] || 0; if (v > 0.005) payers.push([o, v]); else if (v < -0.005) receivers.push([o, -v]); });
  const out = []; let i = 0, j = 0;
  while (i < payers.length && j < receivers.length) {
    const amt = Math.min(payers[i][1], receivers[j][1]);
    out.push({ from: payers[i][0], to: receivers[j][0], amount: amt });
    payers[i][1] -= amt; receivers[j][1] -= amt;
    if (payers[i][1] < 0.005) i++;
    if (receivers[j][1] < 0.005) j++;
  }
  return out;
}

function buildAAModel(players, cfg, period, club) {
  const H = ocHelpers(club || OC_LEGACY_CLUB);
  const { ownerIds, shareOf, zero } = H;
  const deals = cfg.deals || {};
  const ownDeals = cfg.ownDeals || {};
  const playerTr = cfg.playerTr || {};
  const agentTr = cfg.agentTr || {};
  const backedMap = cfg.backed || {};
  const finEntry = period ? (cfg.finalizedPeriods || {})[period] : null;
  const snapMakeup = finEntry && typeof finEntry === "object" ? finEntry.snapshot || null : null;

  // Own accounts first, then backed, then everything else — an owner's own
  // username never lands in a tree or on a deal.
  const ownMap = {};
  ownerIds.forEach((o) => ((cfg.ownAccounts || {})[o] || []).forEach((n) => { const k = String(n).trim().toLowerCase(); if (k) ownMap[k] = o; }));
  const ownRows = [], backedRows = [], normalRows = [];
  players.forEach((p) => {
    const lo = p.name.trim().toLowerCase();
    if (ownMap[lo]) ownRows.push({ ...p, owner: ownMap[lo] });
    else if (backedMap[lo]) backedRows.push({ ...p, backedKey: lo });
    else normalRows.push(p);
  });

  // Deal cascade: an explicit player-level override always wins (this is the
  // VIP case — a player under an agent whose rate shouldn't follow the
  // agent's rate). Failing that, fall back to the player's agent's rate,
  // then the super agent's rate, then the union default.
  //
  // ownDeals is a separate, higher-priority override reserved for a super
  // agent or agent's OWN personal-play row (their memberId is the same ID as
  // their sa/agent line). Without it, that one row would have no way to run
  // a VIP rate of its own without also moving deals[saId]/deals[agentId] —
  // the exact same key every other player under them falls back to — which
  // would silently change everyone else's rate too. ownDeals never matches
  // an ordinary player's memberId (only ever set for a sa/agent's own row),
  // so it's safe to check first for every row.
  const rows = normalRows.map((p) => {
    const tbPct = ownDeals[p.memberId] ?? deals[p.memberId] ?? (p.agentId !== "-" ? deals[p.agentId] : undefined) ?? (p.saId !== "-" ? deals[p.saId] : undefined) ?? cfg.defaultTB;
    return { ...p, tbPct, tipback: (p.fee * tbPct) / 100 };
  });

  // Hierarchy: Super Agent → Agent → Player, same tree shape as the union
  // export itself. Most rows carry both a Super Agent ID and an Agent ID;
  // some sit directly under a Super Agent with no agent between; some sit
  // under an Agent that has no Super Agent above it at all — those used to
  // fall through to "individual player" because only saId was checked. Now
  // they roll up onto their agent's line instead. Managers and Masters are
  // outside this hierarchy entirely (the export gives them "-"/"-" for both
  // IDs) — they get their own categories rather than being swept in with
  // ordinary individual players.
  const saMap = new Map(); const agentMap = new Map();
  const managerMap = new Map(); const masterMap = new Map(); const individuals = [];
  for (const p of rows) {
    const role = (p.role || "").trim().toLowerCase();
    if (role === "manager") {
      if (!managerMap.has(p.memberId)) managerMap.set(p.memberId, { key: `mgr:${p.memberId}`, type: "manager", id: p.memberId, name: p.name, members: [] });
      managerMap.get(p.memberId).members.push(p);
    } else if (role === "master") {
      if (!masterMap.has(p.memberId)) masterMap.set(p.memberId, { key: `mstr:${p.memberId}`, type: "master", id: p.memberId, name: p.name, members: [] });
      masterMap.get(p.memberId).members.push(p);
    } else if (p.saId !== "-") {
      if (!saMap.has(p.saId)) saMap.set(p.saId, { key: `sa:${p.saId}`, type: "sa", id: p.saId, name: p.saName, members: [] });
      saMap.get(p.saId).members.push(p);
    } else if (p.agentId !== "-") {
      if (!agentMap.has(p.agentId)) agentMap.set(p.agentId, { key: `ag:${p.agentId}`, type: "agent", id: p.agentId, name: p.agentName, members: [] });
      agentMap.get(p.agentId).members.push(p);
    } else individuals.push({ key: `p:${p.memberId}`, type: "player", id: p.memberId, name: p.name, members: [p] });
  }
  const agg = (e) => { const s = (f) => e.members.reduce((a, m) => a + m[f], 0);
    return { ...e, hands: s("hands"), pnl: s("pnl"), fee: s("fee"), tipback: s("tipback") }; };

  const baseEntities = [
    ...[...saMap.values()].map(agg),
    ...[...agentMap.values()].map(agg),
    ...[...managerMap.values()].map(agg),
    ...[...masterMap.values()].map(agg),
    ...individuals.map(agg),
  ];
  // Every super agent / agent / unlinked player / manager / master line,
  // before any DL umbrella merge — this is the pick-list the DL Umbrellas
  // tab assigns chips from. Managers and Masters are included too: a club
  // owner's own line (recorded under one of those roles) can sit in an
  // umbrella right alongside ordinary agents and players.
  const dlAssignable = baseEntities.filter((e) => e.type === "sa" || e.type === "agent" || e.type === "player" || e.type === "manager" || e.type === "master").sort((a, b) => Math.abs(b.fee) - Math.abs(a.fee));

  // DL umbrellas: some super agents / agents / unlinked players are, in
  // reality, all settled through one person — fold those lines into a single
  // combined line (one class, one collector) same as Fish Tank's umbrella
  // groups, while each member keeps its own rate and player overrides intact
  // underneath (visible via "player overrides" on the merged line).
  const dlUmbrellas = cfg.dlUmbrellas || [];
  const dlOf = {};
  dlUmbrellas.forEach((u) => (u.memberKeys || []).forEach((k) => (dlOf[k] = u)));
  const dlMap = new Map(); const nonDl = [];
  for (const e of baseEntities) {
    const u = dlOf[e.key];
    if (u) {
      if (!dlMap.has(u.id)) dlMap.set(u.id, { key: `dl:${u.id}`, type: "dlUmbrella", id: u.id, name: u.name, subgroups: [], members: [] });
      const m = dlMap.get(u.id); m.subgroups.push(e); m.members.push(...e.members);
    } else nonDl.push(e);
  }
  const dlEntities = [...dlMap.values()].map((e) => {
    const s = (f) => e.subgroups.reduce((a, g) => a + g[f], 0);
    return { ...e, hands: s("hands"), pnl: s("pnl"), fee: s("fee"), tipback: s("tipback") };
  });

  const entities = [...dlEntities, ...nonDl]
    .map((e) => {
      // Super agent AND agent lines both default into the owner pool, same
      // as before (an agent with no SA above still runs like an agent line).
      // A DL umbrella defaults the same way — it's still fundamentally
      // agent-type lines, just billed through one person. Managers/Masters/
      // standalone players still need a manual class, same as individuals
      // always have — they're too varied to default — UNLESS the club only
      // has one owner, in which case there's no actual ambiguity: every line
      // is either the pool (which is 100% that owner anyway) or personal to
      // that owner, so default straight to them instead of making Ak click
      // through every single line by hand.
      const soleOwner = (club?.owners || []).length === 1 ? club.owners[0].id : null;
      const tag = (cfg.owners || {})[e.key] || (e.type === "sa" || e.type === "agent" || e.type === "dlUmbrella" ? "agent" : soleOwner);
      // TR: a straight cut of this line's net (P&L + rakeback), on top of the
      // ordinary rake margin below — e.g. a "70/10" deal (70% TB, 10% TR).
      // Defaults to 0 (a no-op, same as before TR existed) unless the club
      // sets its own defaultTR or this specific line is overridden. One dial
      // per line (even a DL umbrella — e.id is the umbrella's own id here,
      // not any one subgroup's), same as before.
      const tr = (cfg.tr || {})[e.id] ?? cfg.defaultTR ?? 0;
      // TR cascades the same way TB already does: a nested agent (one that
      // sits inside a super agent's tree, e.g. Jump4Joy under Chapo) can run
      // its own TR default via agentTr, keyed by that agent's own id — every
      // one of their players (present or future) inherits it automatically,
      // same as agentTr's TB sibling (`deals[agentId]`) already does. A
      // standalone top-level agent never hits this: their own id already IS
      // e.id, so m.agentId === e.id there and the nested-agent check below
      // is skipped — they just use the line's own dial, unchanged.
      // playerTr lets one specific player (including an agent's own play —
      // their "VIP deal") run a different TR than the rest of their line,
      // exactly like the VIP TB override — same idea, separate map, keyed
      // the same way, checked first. Cuts are rounded per player (real
      // settlement amounts) then summed, so the line's total always
      // reconciles exactly with what each player's own report row shows.
      const members = e.members.map((m) => {
        const nestedAgentId = m.agentId && m.agentId !== "-" && m.agentId !== e.id ? m.agentId : null;
        const agentDefault = nestedAgentId ? agentTr[nestedAgentId] : undefined;
        const mtr = playerTr[m.memberId] ?? agentDefault ?? tr;
        return { ...m, tr: mtr, trCut: r2(((m.pnl + m.tipback) * mtr) / 100) };
      });
      const trCut = members.reduce((a, m) => a + m.trCut, 0);
      const net = e.pnl + e.tipback;
      const margin = e.fee - e.tipback + trCut; // profit on this line: tips − rakeback paid out, plus any TR cut
      const settlement = net - trCut; // cash the counterparty is owed by the club
      const collector = tag && tag !== "agent" ? tag : (cfg.collectors || {})[e.key] || null;
      return { ...e, members, tag, tr, trCut, margin, settlement, unionCash: -settlement, collector };
    })
    .sort((a, b) => Math.abs(b.fee) - Math.abs(a.fee));

  // Own accounts: position = P&L with 100% feeback; no rakeback leaves, no cash to collect.
  const own = ownRows.map((p) => ({ ...p, position: p.pnl + p.fee }));
  const ownPosition = zero();
  own.forEach((p) => (ownPosition[p.owner] = (ownPosition[p.owner] || 0) + p.position));

  // Stake / action deals — same math as Fish Tank's house-backed engine.
  const backedEntities = backedRows.map((p) => {
    const b = backedMap[p.backedKey];
    const key = `b:${p.backedKey}`;
    if (b.deal === "action") {
      const rbCredit = (p.fee * (b.rbPct ?? 100)) / 100;
      const net = p.pnl + rbCredit;
      const backerBook = (net * b.actionPct) / 100;
      return { key, type: "backed", dealType: "action", id: p.memberId, name: p.name, members: [p],
        hands: p.hands, pnl: p.pnl, fee: p.fee, rbCredit, net,
        margin: p.fee - rbCredit, settlement: net - backerBook, backerBook,
        unionCash: -(p.pnl + rbCredit), backer: b.backer, actionPct: b.actionPct };
    }
    const entering = snapMakeup && snapMakeup[p.backedKey] != null ? snapMakeup[p.backedKey] : (b.makeup || 0);
    const inMakeup = entering > 0.005;
    const rb = inMakeup ? b.rbMakeup : b.rbNormal;
    const rbCredit = (p.fee * (rb ?? 100)) / 100;
    const net = p.pnl + rbCredit;
    const excess = Math.max(0, net - entering);
    const playerCash = (excess * (b.playerProfitPct ?? 50)) / 100;
    const makeupAfter = Math.max(0, entering - net);
    return { key, type: "backed", dealType: "makeup", id: p.memberId, name: p.name, members: [p],
      hands: p.hands, pnl: p.pnl, fee: p.fee, rbCredit, net, rb, inMakeup,
      margin: p.fee - rbCredit, settlement: playerCash, backerBook: net - playerCash,
      unionCash: -(p.pnl + rbCredit), backer: b.backer,
      makeupBefore: entering, makeupAfter };
  });

  const untagged = entities.filter((e) => !e.tag);
  const uncollected = entities.filter((e) => e.tag === "agent" && !e.collector);

  // ——— Profit (per owner) ———
  // The pool is ONLY the agent-line margins, split by each owner's pool %.
  // Stake/action players belong to their backer: the rake margin on those
  // lines goes to that owner personally (split backer → by pool %), never into the pool.
  const pool = entities.filter((e) => e.tag === "agent").reduce((a, e) => a + e.margin, 0);
  const personalMargin = zero();
  entities.forEach((e) => { if (e.tag && e.tag !== "agent" && personalMargin[e.tag] != null) personalMargin[e.tag] += e.margin; });
  const dealMargin = zero();
  backedEntities.forEach((e) => ownerIds.forEach((o) => (dealMargin[o] += e.margin * shareOf(e.backer, o))));
  const jackpot = +((cfg.jackpots || {})[period] ?? 0) || 0;
  const jpSharesRaw = Object.fromEntries(ownerIds.map((o) => [o, jackpot * H.jpShareOf(o)]));
  // BBJ holds: an owner who isn't ready to collect their jackpot share yet
  // can have another owner hold it for them — that owner's share is fully
  // redirected to the holder for every downstream calc (profit, cash target,
  // the transfer) starting the period the hold began. Everything else about
  // the redirected-away owner (personal lines, own accounts, stake books)
  // is untouched — they're only pulled out of the jackpot line. The running
  // total of what's been redirected over time lives in the BBJ holds card
  // on Owners & setup (computed from cfg.jackpots history, not here).
  const jpHolds = cfg.jpHolds || [];
  const jpHoldMoves = jpHolds
    .filter((h) => ownerIds.includes(h.holderId) && ownerIds.includes(h.forOwnerId) && (!h.sincePeriod || period >= h.sincePeriod))
    .map((h) => ({ id: h.id, holderId: h.holderId, forOwnerId: h.forOwnerId, amount: jpSharesRaw[h.forOwnerId] || 0 }));
  const jpShares = { ...jpSharesRaw };
  jpHoldMoves.forEach((m) => { jpShares[m.forOwnerId] -= m.amount; jpShares[m.holderId] += m.amount; });
  const poolShares = Object.fromEntries(ownerIds.map((o) => [o, pool * H.poolShare(o)]));
  // Where's the jackpot cash? "pool" = external jackpot escrow, held by no one,
  // excluded from the transfer. "collections" = the contribution was deducted
  // from players' P&L, so the line collectors are holding it — the shares
  // belong in the settle-up. Tell-tale: books off by exactly the jackpot.
  const jpInCollections = ((cfg.jpModes || {})[period] || "pool") === "collections";
  const profit = Object.fromEntries(ownerIds.map((o) => [o, poolShares[o] + personalMargin[o] + dealMargin[o] + ownPosition[o] + jpShares[o]]));

  // ——— Collections (actuals) ———
  // The jackpot cash sits in the jackpot pool (held by no one), so the
  // settle-up compares collections against profit MINUS the jackpot shares.
  // Own-account P&L stays IN the transfer: an owner's winnings are funded by
  // the week's collections, same as the Fish Tank Ak/Jon recon.
  const actual = zero();
  entities.forEach((e) => { if (e.collector && actual[e.collector] != null) actual[e.collector] += e.unionCash; });
  backedEntities.forEach((e) => ownerIds.forEach((o) => (actual[o] += e.unionCash * shareOf(e.backer, o))));

  const cashProfit = Object.fromEntries(ownerIds.map((o) => [o, profit[o] - (jpInCollections ? 0 : jpShares[o])]));
  const delta = Object.fromEntries(ownerIds.map((o) => [o, actual[o] - cashProfit[o]]));
  const transfers = aaTransfers(delta, ownerIds);
  const ready = untagged.length === 0 && uncollected.length === 0;
  const imbalance = ownerIds.reduce((a, o) => a + delta[o], 0);
  const balanceOk = ready && Math.abs(imbalance) < 0.02;

  // ——— Stake / action books, per backer (settled separately) ———
  const backedBook = zero();
  backedEntities.forEach((e) => ownerIds.forEach((o) => (backedBook[o] += e.backerBook * shareOf(e.backer, o))));

  const clubRevenue = entities.reduce((a, e) => a + e.fee, 0) + backedEntities.reduce((a, e) => a + e.fee, 0) + own.reduce((a, p) => a + p.fee, 0);
  const agentRB = entities.filter((e) => e.tag === "agent").reduce((a, e) => a + e.tipback, 0);
  const personalRB = entities.filter((e) => e.tag && e.tag !== "agent").reduce((a, e) => a + e.tipback, 0);
  const backedRB = backedEntities.reduce((a, e) => a + e.rbCredit, 0);
  const ownFeeback = own.reduce((a, p) => a + p.fee, 0);
  const totalPersonalMargin = ownerIds.reduce((a, o) => a + personalMargin[o], 0);
  const totalDealMargin = ownerIds.reduce((a, o) => a + dealMargin[o], 0);
  // Backwards-compat: single jpShare when everyone gets the same slice.
  const jpShare = ownerIds.length ? jackpot / ownerIds.length : 0;

  return { entities, backedEntities, own, ownPosition, untagged, uncollected, ready, H, ownerIds, dlAssignable,
    pool, poolShares, personalMargin, dealMargin, totalPersonalMargin, totalDealMargin, jackpot, jpShare, jpShares, jpSharesRaw, jpHoldMoves, jpInCollections, profit, cashProfit, actual, delta, transfers, imbalance, balanceOk,
    backedBook, clubRevenue, agentRB, personalRB, backedRB, ownFeeback };
}

async function loadAAModel(club) {
  try {
    const [c, d] = await Promise.all([store.get(club.cfgKey), store.get(club.weekKey)]);
    if (!d?.value) return null;
    const wk = JSON.parse(d.value);
    if (!wk.players?.length) return null;
    const cfg = { ...AA_DEFAULT_CFG, ...(c?.value ? JSON.parse(c.value) : {}) };
    return { model: buildAAModel(wk.players, cfg, wk.period || "", club), period: wk.period || "", cfg, club };
  } catch (e) { return null; }
}
async function loadAllOwnerClubModels() {
  const clubs = await loadOwnerClubs();
  const out = [];
  for (const club of clubs) { const m = await loadAAModel(club); if (m) out.push(m); }
  return out;
}
// My Clubs (agent) model, standalone — same pattern as loadFishTankModel/loadAAModel.
async function loadAgentModel() {
  try {
    const c = await store.get("agentclubs-v3");
    if (!c?.value) return null;
    const acfg = normalizeAgent({ ...AGENT_DEFAULT, ...JSON.parse(c.value) });
    const weekKeys = Object.keys(acfg.weeks || {}).sort();
    const wk = acfg.currentWeek && acfg.weeks[acfg.currentWeek] ? acfg.currentWeek : weekKeys[weekKeys.length - 1];
    if (!wk) return null;
    const model = computeAgent(acfg, acfg.weeks[wk]);
    return { model, acfg, wk };
  } catch (e) { return null; }
}

// ——— Styled Excel — matches the app's dark/gold theme ———
// AA's exports now share Fish Tank's exact visual language — same XLC palette,
// same Arial font, same plain bold title line, same gold header / striped rows /
// cream TOTAL row, same visible gridlines. The old dark-banner/Calibri look is
// gone; only the AA-specific sheet shapes (owner splits, KV summaries, TR %)
// remain, since Fish Tank has no equivalent for those.
const AA_PCT_FMT = '0.##"%"';
const aaFile = (base, period) => `AA-${String(base).replace(/[^\w-]+/g, "_")}-${(period || "week").replace(/[^\w-]+/g, "_")}.xlsx`;

function aaWs(wb, name, widths) {
  const ws = wb.addWorksheet(name);
  ws.columns = widths.map((w) => ({ width: w }));
  ws._span = widths.length;
  return ws;
}
function aaBand(ws, title, sub) {
  xTitle(ws, sub ? `${title} — ${sub}` : title);
}
function aaSection(ws, text) {
  ws.addRow([]);
  const r = ws.addRow([text.toUpperCase()]);
  r.getCell(1).font = { ...FBASE, size: 9, bold: true, color: { argb: XLC.mute } };
  for (let i = 1; i <= ws._span; i++) r.getCell(i).border = { bottom: { style: "thin", color: { argb: XLC.gold } } };
}
// spec per column: "tl" text-left · "tm" text-left-mute · "i" integer · "m" money · "p" percent · "" blank
function aaCellStyle(c, kind, { alt = false, total = false } = {}) {
  c.font = { ...FBASE, bold: total, color: { argb: kind === "tm" ? XLC.mute : XLC.ink } };
  if (total) c.fill = fillOf(XLC.cream);
  else if (alt) c.fill = fillOf(XLC.rowAlt);
  c.alignment = { horizontal: kind === "tl" || kind === "tm" ? "left" : "right" };
  if (kind === "m") {
    c.numFmt = MONEY_FMT;
    const v = typeof c.value === "number" ? c.value : 0;
    c.font = { ...FBASE, bold: total, color: { argb: v > 0.005 ? XLC.green : v < -0.005 ? XLC.red : XLC.ink } };
  }
  if (kind === "p") c.numFmt = AA_PCT_FMT;
  if (kind === "i") c.numFmt = "#,##0";
}
// `spec` (same array passed to aaData/aaTotal below it) drives left/right
// alignment per column, same idea as Fish Tank's xHeader leftCols — just
// derived from the real column kind instead of a fixed count, since AA's
// tables don't all have the same number of leading text columns.
function aaHeader(ws, labels, spec) {
  const r = ws.addRow(labels);
  r.eachCell({ includeEmpty: true }, (c, col) => {
    if (col > ws._span) return;
    const kind = spec ? spec[col - 1] : col === 1 ? "tl" : "";
    c.font = { name: "Arial", size: 9, bold: true, color: { argb: XLC.white } };
    c.fill = fillOf(XLC.gold);
    c.alignment = { horizontal: kind === "tl" || kind === "tm" ? "left" : "right" };
  });
  for (let i = 1; i <= ws._span; i++) { const c = r.getCell(i); if (!c.fill) c.fill = fillOf(XLC.gold); }
  return r.number;
}
function aaData(ws, vals, spec, alt) {
  const r = ws.addRow(vals);
  spec.forEach((k, idx) => aaCellStyle(r.getCell(idx + 1), k, { alt }));
  for (let i = spec.length + 1; i <= ws._span; i++) aaCellStyle(r.getCell(i), "", { alt });
  return r;
}
function aaTotal(ws, vals, spec) {
  const r = ws.addRow(vals);
  spec.forEach((k, idx) => aaCellStyle(r.getCell(idx + 1), k, { total: true }));
  for (let i = spec.length + 1; i <= ws._span; i++) aaCellStyle(r.getCell(i), "", { total: true });
  return r;
}
function aaKV(ws, label, v, { bold = false, gold = false } = {}) {
  const r = ws.addRow([label, typeof v === "number" ? n2(v) : v]);
  const cl = r.getCell(1), cv = r.getCell(2);
  cl.font = { ...FBASE, bold, color: { argb: bold ? XLC.ink : XLC.mute } };
  cl.alignment = { horizontal: "left" };
  if (typeof v === "number") {
    cv.numFmt = MONEY_FMT;
    cv.font = { ...FBASE, bold: true, color: { argb: gold ? XLC.gold : v > 0.005 ? XLC.green : v < -0.005 ? XLC.red : XLC.ink } };
  } else cv.font = { ...FBASE, bold: true, color: { argb: XLC.ink } };
  cv.alignment = { horizontal: "right" };
  if (bold) for (let i = 1; i <= ws._span; i++) r.getCell(i).fill = fillOf(XLC.cream);
  return r;
}
function aaBarKV(ws, label, v) {
  const r = ws.addRow([label, n2(v)]);
  for (let i = 1; i <= ws._span; i++) r.getCell(i).fill = fillOf(XLC.cream);
  r.getCell(1).font = { ...FBASE, bold: true };
  r.getCell(1).alignment = { horizontal: "left" };
  const cv = r.getCell(2);
  cv.numFmt = MONEY_FMT;
  cv.font = { ...FBASE, bold: true, color: { argb: v > 0.005 ? XLC.green : v < -0.005 ? XLC.red : XLC.ink } };
  cv.alignment = { horizontal: "right" };
  return r;
}

async function downloadAAWorkbook(model, period, club) {
  const wb = new ExcelJS.Workbook();
  const per = period || "this week";
  const H = model.H, ids = model.ownerIds, lbl = H.lbl;
  const clubName = club?.name || "All American";

  const ws = aaWs(wb, "Owner split", [40, ...ids.map(() => 15)]);
  aaBand(ws, clubName, `Weekly settlement · ${per}`);
  aaSection(ws, "Profit per owner");
  const P4 = ["tl", ...ids.map(() => "m")];
  aaHeader(ws, ["", ...ids.map((o) => lbl(o))], P4);
  const rowOf = (m) => ids.map((o) => n2(m[o] || 0));
  aaData(ws, [`Pool share (${H.poolLabel})`, ...rowOf(model.poolShares)], P4, false);
  aaData(ws, ["Personal-line margins (100%)", ...rowOf(model.personalMargin)], P4, true);
  aaData(ws, ["Stake/action rake margins (backer's)", ...rowOf(model.dealMargin)], P4, false);
  aaData(ws, ["Own accounts P&L (full feeback)", ...rowOf(model.ownPosition)], P4, true);
  aaData(ws, ["Bad beat contribution (by JP %)", ...rowOf(model.jpShares)], P4, true);
  aaTotal(ws, [`PROFIT — ${clubName.toUpperCase()}`, ...rowOf(model.profit)], P4);
  aaSection(ws, "Settle-up");
  aaData(ws, ["Collectable share" + (model.jpInCollections ? " (incl. jackpot shares)" : " (jackpot pool excluded)"), ...rowOf(model.cashProfit)], P4, false);
  aaData(ws, ["Actually collected", ...rowOf(model.actual)], P4, true);
  aaTotal(ws, ["OVER / (UNDER) COLLECTED", ...rowOf(model.delta)], P4);
  ws.addRow([]);
  if (model.transfers.length === 0) aaKV(ws, "Transfers", "perfectly even — none needed");
  model.transfers.forEach((t) => aaBarKV(ws, `${lbl(t.from)} pays ${lbl(t.to)}`, t.amount));
  aaSection(ws, "Club economics");
  aaKV(ws, "Rake collected", model.clubRevenue);
  aaKV(ws, "Rakeback → agents (pool lines)", -model.agentRB);
  aaKV(ws, "Rakeback → players on personal lines", -model.personalRB);
  aaKV(ws, "RB credits → stake/action deals", -model.backedRB);
  aaKV(ws, "Feeback → owners' own accounts", -model.ownFeeback);
  aaKV(ws, `Pool (agent lines only, ${H.poolLabel})`, model.pool, { bold: true });
  aaKV(ws, "Personal margins (owner-routed)", model.totalPersonalMargin, { bold: true });
  aaKV(ws, "Stake/action rake margins → backer", model.totalDealMargin, { bold: true });
  aaKV(ws, model.jpInCollections ? "Bad beat contribution (in the collections)" : "Bad beat contribution (sits in JP pool)", model.jackpot, { gold: true });

  const L2 = aaWs(wb, "Lines", [26, 20, 13, 9, 13, 13, 8, 13, 13, 15, 13]);
  aaBand(L2, "Every line", `Class, margin routing, and collections · ${per}`);
  const S2 = ["tl", "tm", "tm", "i", "m", "m", "p", "m", "m", "tm", "m"];
  aaHeader(L2, ["Line", "Class", "Collected by", "Hands", "Winnings", "Tips", "TB %", "Rakeback", "Margin", "Margin goes to", "Club cash"], S2);
  model.entities.forEach((e, i) => {
    aaData(L2, [e.type === "dlUmbrella" ? e.name : withId(e.name, e.id), e.tag ? (e.tag === "agent" ? "Agent · pool" : `Personal · ${lbl(e.tag)}`) : "UNTAGGED",
      e.collector ? lbl(e.collector) : "", e.hands, n2(e.pnl), n2(e.fee),
      e.fee ? aaPctN((e.tipback / e.fee) * 100) : "", n2(e.tipback), n2(e.margin),
      e.tag === "agent" ? H.poolLabel : e.tag ? `${lbl(e.tag)} 100%` : "", n2(e.unionCash)], S2, i % 2 === 1);
  });
  model.own.forEach((p, i) => {
    aaData(L2, [withId(p.name, p.memberId), `Own account · ${lbl(p.owner)}`, "—", p.hands, n2(p.pnl), n2(p.fee), 100, n2(p.fee), n2(p.position), `${lbl(p.owner)} P&L`, 0], S2, (model.entities.length + i) % 2 === 1);
  });

  if (model.backedEntities.length) {
    const L3 = aaWs(wb, "Stake & action", [24, 15, 16, 13, 13, 13, 13, 15, 15, 14]);
    aaBand(L3, "Stake & action deals", `Settled between backer and player · ${per}`);
    const S3 = ["tl", "tm", "tm", "m", "m", "m", "m", "m", "m", "m"];
    aaHeader(L3, ["Player", "Deal", "Backer", "Winnings", "Tips", "RB credit", "Net", "Player settlement", "Backer book", "Makeup after"], S3);
    model.backedEntities.forEach((e, i) => {
      aaData(L3, [withId(e.name, e.id), e.dealType === "action" ? `action ${e.actionPct}%` : "stake · makeup",
        lbl(e.backer),
        n2(e.pnl), n2(e.fee), n2(e.rbCredit), n2(e.net), n2(e.settlement), n2(e.backerBook),
        e.dealType === "makeup" ? n2(e.makeupAfter) : ""], S3, i % 2 === 1);
    });
  }
  await saveWb(wb, aaFile(clubName, period));
}

// Report for one line — an agent tree, a player, or a staked/action player.
async function downloadAALineExcel(e, period, model) {
  const wb = new ExcelJS.Workbook();
  const per = period || "this week";
  const H = model.H, lbl = H.lbl;
  const cls = e.type === "backed"
    ? (e.dealType === "action" ? `Action buy · ${e.actionPct}% to ${lbl(e.backer)}` : `Stake deal · backer ${lbl(e.backer)}`)
    : e.tag === "agent" ? `Agent · ${H.poolLabel}` : e.tag ? `Personal · ${lbl(e.tag)}` : "Untagged";

  if (e.type === "backed") {
    const ws = aaWs(wb, "Report", [40, 16]);
    aaBand(ws, withId(e.name, e.id), `${cls} · ${per}`);
    aaSection(ws, "Week result");
    aaKV(ws, "Winnings (P&L)", e.pnl);
    aaKV(ws, "Tips", e.fee);
    aaKV(ws, `RB credit${e.rb != null ? ` @ ${aaPctN(e.rb)}%` : ""}`, e.rbCredit);
    aaKV(ws, "Week net (P&L + RB)", e.net, { bold: true });
    aaSection(ws, "Deal settlement");
    if (e.dealType === "makeup") {
      aaKV(ws, "Makeup entering", e.makeupBefore);
      aaKV(ws, "Makeup after", e.makeupAfter);
    }
    aaBarKV(ws, "Player settlement", e.settlement);
    aaBarKV(ws, "Backer book (settles separately)", e.backerBook);
    aaKV(ws, `Rake margin → ${lbl(e.backer)} personally`, e.margin);
  } else {
    const ws = aaWs(wb, "Report", [24, 15, 9, 13, 13, 8, 13, 8, 14]);
    aaBand(ws, e.type === "dlUmbrella" ? e.name : withId(e.name, e.id), `${cls} · ${per}`);
    const S = ["tl", "tm", "i", "m", "m", "p", "m", "p", "m"];
    aaHeader(ws, ["Player", "Agent", "Hands", "Winnings", "Tips", "TB %", "Rakeback", "TR %", "Settlement"], S);
    [...e.members].sort((a, b) => b.fee - a.fee).forEach((m, i) => {
      aaData(ws, [withId(m.name, m.memberId), m.agentName !== "-" ? withId(m.agentName, m.agentId) : "—", m.hands, n2(m.pnl), n2(m.fee), aaPctN(m.tbPct), n2(m.tipback), aaPctN(m.tr || 0), n2(m.pnl + m.tipback - (m.trCut || 0))], S, i % 2 === 1);
    });
    const eNet = e.pnl + e.tipback;
    aaTotal(ws, ["TOTAL", "", e.hands, n2(e.pnl), n2(e.fee), e.fee ? aaPctN((e.tipback / e.fee) * 100) : "", n2(e.tipback), eNet ? aaPctN((e.trCut / eNet) * 100) : "", n2(e.settlement)], ["tl", "", "i", "m", "m", "p", "m", "p", "m"]);
    aaSection(ws, "Notes");
    aaKV(ws, "Settlement — owed to the line by the club (negative: they pay in)", e.settlement, { bold: true });
    aaKV(ws, `Margin (tips − rakeback) → ${e.tag === "agent" ? `club pool, ${H.poolLabel}` : e.tag ? `${lbl(e.tag)} 100%` : "unassigned"}`, e.margin);
    if (e.collector) aaKV(ws, "Collected by", lbl(e.collector));
  }
  await saveWb(wb, aaFile(e.name, period));
}

// Full weekly report for one owner — his own play, his personal lines, and the shared agents. Never another owner's players.
async function downloadAAOwnerExcel(model, o, period, club) {
  const H = model.H, lbl = H.lbl;
  const L = lbl(o);
  const per = period || "this week";
  const clubName = club?.name || "All American";
  const inPool = H.poolShare(o) > 0;
  const wb = new ExcelJS.Workbook();

  const ws = aaWs(wb, "Summary", [44, 16]);
  aaBand(ws, L, `${clubName} · ${per}`);
  aaSection(ws, "Profit build-up");
  if (inPool) aaKV(ws, `Pool share (${H.pctS(H.poolShare(o))})`, model.poolShares[o]);
  aaKV(ws, "Personal-line margins (100%)", model.personalMargin[o]);
  aaKV(ws, "Stake/action rake margins (his deals)", model.dealMargin[o]);
  aaKV(ws, "Own accounts P&L (full feeback)", model.ownPosition[o]);
  aaKV(ws, "Bad beat jackpot share", model.jpShares[o]);
  aaBarKV(ws, `PROFIT — ${clubName.toUpperCase()}`, model.profit[o]);
  aaSection(ws, "Settle-up");
  aaKV(ws, "Collectable share", model.cashProfit[o]);
  aaKV(ws, "Actually collected", model.actual[o]);
  aaKV(ws, "Over / (under) collected", model.delta[o], { bold: true });
  const mine = model.transfers.filter((t) => t.from === o || t.to === o);
  if (mine.length === 0) aaKV(ws, "Transfers", "none — even");
  mine.forEach((t) => aaBarKV(ws, `${lbl(t.from)} pays ${lbl(t.to)}`, t.amount));
  if (Math.abs(model.backedBook[o]) > 0.005) {
    aaSection(ws, "Stake & action");
    aaKV(ws, "Deal book — settles separately, not in the transfer", model.backedBook[o], { bold: true });
  }

  const shared = inPool ? model.entities.filter((e) => e.tag === "agent") : [];
  if (shared.length) {
    const wsS = aaWs(wb, "Shared agents", [26, 9, 13, 13, 8, 13, 13, 14]);
    aaBand(wsS, "Shared agents", `${H.poolLabel} · ${per}`);
    const S = ["tl", "i", "m", "m", "p", "m", "m", "m"];
    aaHeader(wsS, ["Agent line", "Hands", "Winnings", "Tips", "TB %", "Rakeback", "Margin → pool", `${L}'s share`], S);
    shared.forEach((e, i) => aaData(wsS, [e.type === "dlUmbrella" ? e.name : withId(e.name, e.id), e.hands, n2(e.pnl), n2(e.fee), e.fee ? aaPctN((e.tipback / e.fee) * 100) : "", n2(e.tipback), n2(e.margin), n2(e.margin * H.poolShare(o))], S, i % 2 === 1));
    aaTotal(wsS, ["POOL", "", "", "", "", "", n2(model.pool), n2(model.poolShares[o])], ["tl", "", "", "", "", "", "m", "m"]);
  }

  const personal = model.entities.filter((e) => e.tag === o);
  if (personal.length) {
    const ws2 = aaWs(wb, "Personal lines", [26, 9, 13, 13, 8, 13, 14, 14]);
    aaBand(ws2, "Personal players & agents", `Margin 100% to ${L} · ${per}`);
    const S = ["tl", "i", "m", "m", "p", "m", "m", "m"];
    aaHeader(ws2, ["Line", "Hands", "Winnings", "Tips", "TB %", "Rakeback paid", "Margin → " + L, "Club cash"], S);
    personal.forEach((e, i) => aaData(ws2, [e.type === "dlUmbrella" ? e.name : withId(e.name, e.id), e.hands, n2(e.pnl), n2(e.fee), e.fee ? aaPctN((e.tipback / e.fee) * 100) : "", n2(e.tipback), n2(e.margin), n2(e.unionCash)], S, i % 2 === 1));
    aaTotal(ws2, ["TOTAL PERSONAL MARGIN", "", "", "", "", "", n2(model.personalMargin[o]), ""], ["tl", "", "", "", "", "", "m", ""]);
  }

  const collected = model.entities.filter((e) => e.collector === o);
  const backedMine = model.backedEntities.filter((e) => H.shareOf(e.backer, o) > 0);
  if (collected.length || backedMine.length) {
    const ws3 = aaWs(wb, "Collections", [26, 26, 16]);
    aaBand(ws3, "Collections", `Lines ${L} collects this week · ${per}`);
    const S = ["tl", "tm", "m"];
    aaHeader(ws3, ["Line", "Class", "Club cash collected"], S);
    let i = 0;
    collected.forEach((e) => aaData(ws3, [e.type === "dlUmbrella" ? e.name : withId(e.name, e.id), e.tag === "agent" ? "Agent" : `Personal · ${L}`, n2(e.unionCash)], S, i++ % 2 === 1));
    backedMine.forEach((e) => aaData(ws3, [withId(e.name, e.id), e.backer === "split" ? `Stake/action (${H.pctS(H.shareOf(e.backer, o))} — split backer)` : "Stake/action (backer collects)", n2(e.unionCash * H.shareOf(e.backer, o))], S, i++ % 2 === 1));
    aaTotal(ws3, ["TOTAL COLLECTED" + (model.jpInCollections ? " (jackpot inside line cash)" : ""), "", n2(model.actual[o])], ["tl", "", "m"]);
  }

  const ownMine = model.own.filter((p) => p.owner === o);
  if (ownMine.length) {
    const ws4 = aaWs(wb, "Own accounts", [26, 9, 13, 13, 13, 15]);
    aaBand(ws4, "Personal play", `${L}'s own accounts · ${per}`);
    const S = ["tl", "i", "m", "m", "m", "m"];
    aaHeader(ws4, ["Account", "Hands", "Winnings", "Tips", "Feeback", "Position → P&L"], S);
    ownMine.forEach((p, i) => aaData(ws4, [withId(p.name, p.memberId), p.hands, n2(p.pnl), n2(p.fee), n2(p.fee), n2(p.position)], S, i % 2 === 1));
    aaTotal(ws4, ["TOTAL", "", "", "", "", n2(model.ownPosition[o])], ["tl", "", "", "", "", "m"]);
  }

  if (backedMine.length) {
    const ws5 = aaWs(wb, "Stake & action", [24, 18, 13, 13, 13, 15, 16, 15, 14]);
    aaBand(ws5, "Stake & action book", `Settles separately · ${per}`);
    const S = ["tl", "tm", "m", "m", "m", "m", "m", "m", "m"];
    aaHeader(ws5, ["Player", "Deal", "Winnings", "RB credit", "Net", "Player settlement", `${L}'s book share`, `Rake margin → ${L}`, "Makeup after"], S);
    backedMine.forEach((e, i) => {
      const sh = H.shareOf(e.backer, o);
      aaData(ws5, [withId(e.name, e.id), (e.dealType === "action" ? `action ${e.actionPct}%` : "stake · makeup") + (e.backer === "split" ? " · split" : ""), n2(e.pnl), n2(e.rbCredit), n2(e.net), n2(e.settlement), n2(e.backerBook * sh), n2(e.margin * sh), e.dealType === "makeup" ? n2(e.makeupAfter) : ""], S, i % 2 === 1);
    });
    aaTotal(ws5, ["BOOK TOTAL — SETTLES SEPARATELY", "", "", "", "", "", n2(model.backedBook[o]), n2(model.dealMargin[o]), ""], ["tl", "", "", "", "", "", "m", "m", ""]);
  }

  await saveWb(wb, aaFile(L, period));
}

function AAOwnAccounts({ model, cfg, up }) {
  const ownerIds = model.ownerIds, lbl = model.H.lbl;
  const [drafts, setDrafts] = useState(() => Object.fromEntries(ownerIds.map((o) => [o, ((cfg.ownAccounts || {})[o] || []).join(", ")])));
  useEffect(() => { setDrafts(Object.fromEntries(ownerIds.map((o) => [o, ((cfg.ownAccounts || {})[o] || []).join(", ")]))); }, [cfg.ownAccounts, ownerIds.join("|")]);
  const commit = (o) => {
    const list = drafts[o].split(",").map((s) => s.trim()).filter(Boolean);
    up({ ownAccounts: { ...(cfg.ownAccounts || {}), [o]: list } });
  };
  return (
    <div style={{ marginTop: 18 }}>
      <Card title="Own accounts — the owners' personal play" right={<Pill tone="gold">P&L logged to owner · full feeback</Pill>}>
        <div style={{ fontSize: 12.5, color: C.mute, marginBottom: 10 }}>
          List each owner's own usernames (comma-separated, exact nicknames). Their play is pulled off the lines above and logged straight into that owner's profit as P&L + 100% feeback — the union doesn't profit off an owner's own play, and the position settles through the weekly transfer like everything else.
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 12 }}>
          {ownerIds.map((o) => (
            <div key={o}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, marginBottom: 4 }}>{lbl(o)}</div>
              <input value={drafts[o] || ""} onChange={(e) => setDrafts({ ...drafts, [o]: e.target.value })}
                onBlur={() => commit(o)} onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
                placeholder="nickname1, nickname2…" style={{ ...inputS, width: "100%", boxSizing: "border-box" }} />
            </div>
          ))}
        </div>
        {model.own.length > 0 && (
          <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 12 }}>
            <thead><tr style={{ background: C.cream }}>
              <th style={{ ...th, textAlign: "left" }}>Account</th><th style={{ ...th, textAlign: "left" }}>Owner</th>
              <th style={th}>Winnings</th><th style={th}>Tips</th><th style={th}>Feeback</th><th style={th}>Position → owner P&L</th>
            </tr></thead>
            <tbody>
              {model.own.map((p, i) => (
                <tr key={p.memberId} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}</td>
                  <td style={tdL}><Pill tone="green">{lbl(p.owner)}</Pill></td>
                  <td style={td}>{money(p.pnl)}</td>
                  <td style={td}>{fmt(p.fee)}</td>
                  <td style={td}>{fmt(p.fee)}</td>
                  <td style={td}><b>{money(p.position)}</b></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

function AABackedTab({ model, cfg, up, period }) {
  const [newName, setNewName] = useState("");
  const [newDeal, setNewDeal] = useState("stake");
  const H = model.H, ownerIds = model.ownerIds, lbl = H.lbl;
  const me = H.meId && ownerIds.includes(H.meId) ? H.meId : ownerIds[0] || "split";
  const addBacked = () => {
    const k = newName.trim().toLowerCase();
    if (!k) return;
    const base = newDeal === "action"
      ? { name: newName.trim(), deal: "action", actionPct: 50, rbPct: 100, backer: me }
      : { name: newName.trim(), deal: "makeup", rbNormal: cfg.defaultTB, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: me };
    up({ backed: { ...cfg.backed, [k]: base } });
    setNewName("");
  };
  const setB = (k, patch) => up({ backed: { ...cfg.backed, [k]: { ...cfg.backed[k], ...patch } } });
  const removeB = (k) => { const b = { ...cfg.backed }; delete b[k]; up({ backed: b }); };
  const findE = (k) => model.backedEntities.find((x) => x.key === `b:${k}`);
  const backerSel = (k, b) => (
    <select value={b.backer} onChange={(e) => setB(k, { backer: e.target.value })} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
      {ownerIds.map((o) => <option key={o} value={o}>{lbl(o)}</option>)}
      {H.hasPool && <option value="split">{H.splitLabel}</option>}
    </select>
  );
  const finalized = !!(cfg.finalizedPeriods || {})[period];
  // "Lock" = the old one-way Finalize, but reversible — see Fish Tank's lockWeek/unlockWeek for the same pattern.
  const lockWeek = () => {
    const backed = { ...cfg.backed };
    const snapshot = {};
    model.backedEntities.forEach((e) => {
      if (e.dealType !== "makeup") return;
      const k = e.key.slice(2);
      if (backed[k]?.unifiedDealId) return; // makeup tracked on the shared staking deal instead — leave local balance frozen
      snapshot[k] = Math.round(e.makeupBefore * 100) / 100;
      if (backed[k]) backed[k] = { ...backed[k], makeup: Math.round(e.makeupAfter * 100) / 100 };
    });
    up({ backed, finalizedPeriods: { ...cfg.finalizedPeriods, [period]: { snapshot } } });
  };
  const unlockWeek = () => {
    const snap = (cfg.finalizedPeriods || {})[period];
    if (!snap) return;
    const backed = { ...cfg.backed };
    Object.entries(snap.snapshot || {}).forEach(([k, entering]) => { if (backed[k]) backed[k] = { ...backed[k], makeup: entering }; });
    const fp = { ...cfg.finalizedPeriods };
    delete fp[period];
    up({ backed, finalizedPeriods: fp });
  };

  const stakePlayers = Object.entries(cfg.backed || {}).filter(([, b]) => b.deal !== "action");
  const actionPlayers = Object.entries(cfg.backed || {}).filter(([, b]) => b.deal === "action");

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 4, gap: 10, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>Stake & action-buy deals</div>
        <div style={{ marginLeft: "auto" }}>
          {finalized
            ? <Btn tone="ghost" small onClick={unlockWeek}>🔒 Week locked · unlock to edit</Btn>
            : <Btn tone="gold" small disabled={!period} onClick={lockWeek}>Lock week — roll makeup forward</Btn>}
        </div>
      </div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 16 }}>
        Matched by exact username — a staked player inside an agent tree is pulled out of the tree onto their deal automatically.
        {" "}<b>Stake (makeup)</b>: week net = P&L + RB credit; above makeup the player is paid their % of the excess, below it the net accrues to makeup. RB rate follows makeup <b>entering</b> the week; Lock once to roll it forward. Unlock to fix something, then lock again to re-snapshot — if this week's already in Tabs, re-sync it from Tabs → Bookkeeping afterward.
        {" "}<b>Action buy</b>: the backer owns their % of (P&L + rakeback); the player settles the remainder.
        {" "}The rake margin on these lines goes to the <b>backer personally</b> (never the pool), and the deal P&L below is <b>between backer and player only</b> — it never enters the owner settle-up.
        {" "}Paste a Tabs → Staking makeup deal's ID into <b>Unified deal</b> on a row to fold that player's weekly net into one shared, ongoing pool across sites (and manual/external games) instead of tracking makeup locally here — the local balance freezes and the deal's own % staked / chop handle the tab. Use a "settle per session" deal (the default) so this week's report and any mid-week manual results each update the running makeup as they're logged, in the order they happened, rather than getting batched into one lump sum.
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Stake deals (makeup)</div>
      <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Makeup entering</th><th style={{ ...th, textAlign: "center" }}>Status</th>
            <th style={th}>RB % normal</th><th style={th}>RB % makeup</th><th style={th}>Player profit %</th>
            <th style={th}>Tips</th><th style={th}>P&L</th><th style={th}>RB credit</th><th style={th}>Net</th><th style={th}>Player gets</th><th style={th}>Backer book</th><th style={th}>Makeup after</th><th style={th}>Unified deal</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {stakePlayers.length === 0 && <tr><td colSpan={16} style={{ ...tdL, color: C.mute, padding: 14 }}>No stake deals yet — add one below.</td></tr>}
            {stakePlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{b.name}{!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}</td>
                  <td style={tdL}>{backerSel(k, b)}</td>
                  <td style={td}><NumInput width={85} value={b.makeup} onChange={(v) => setB(k, { makeup: v })} /></td>
                  <td style={{ ...td, textAlign: "center" }}>{(b.makeup || 0) > 0.005 ? <Pill tone="red">in makeup</Pill> : <Pill tone="green">clear</Pill>}</td>
                  <td style={td}><PctInput width={46} max={999} value={b.rbNormal} onChange={(v) => v != null && setB(k, { rbNormal: v })} /></td>
                  <td style={td}><PctInput width={46} max={999} value={b.rbMakeup} onChange={(v) => v != null && setB(k, { rbMakeup: v })} /></td>
                  <td style={td}><PctInput width={46} value={b.playerProfitPct ?? 50} onChange={(v) => v != null && setB(k, { playerProfitPct: v })} /></td>
                  <td style={td}>{e ? fmt(e.fee) : "—"}</td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td}>{e ? <>{fmt(e.rbCredit)} <span style={{ color: C.mute, fontSize: 11 }}>@{e.rb}%</span></> : "—"}</td>
                  <td style={td}>{e ? money(e.net) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? money(e.backerBook) : "—"}</td>
                  <td style={td}>{e ? fmt(e.makeupAfter) : fmt(b.makeup || 0)}</td>
                  <td style={td}>
                    <input placeholder="deal id" value={b.unifiedDealId || ""} onChange={(ev) => setB(k, { unifiedDealId: ev.target.value.trim() })} style={{ ...inputS, width: 88, fontSize: 11 }} title="Paste a Tabs → Staking makeup deal's ID to fold this player's weekly net into that shared, ongoing pool instead of tracking makeup locally here. Use a 'settle per session' deal so it updates alongside any mid-week manual results in the order they happened." />
                    {b.unifiedDealId && <div style={{ fontSize: 9.5, color: C.goldDark, fontWeight: 700, marginTop: 2 }}>UNIFIED</div>}
                  </td>
                  <td style={td}><button onClick={() => removeB(k)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Action buys</div>
      <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Backer action %</th><th style={th}>RB %</th>
            <th style={th}>Tips</th><th style={th}>Week P&L</th><th style={th}>Player settlement</th><th style={th}>Backer book</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {actionPlayers.length === 0 && <tr><td colSpan={9} style={{ ...tdL, color: C.mute, padding: 14 }}>No action buys yet — add one below.</td></tr>}
            {actionPlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{b.name}{!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}</td>
                  <td style={tdL}>{backerSel(k, b)}</td>
                  <td style={td}><PctInput width={46} value={b.actionPct} onChange={(v) => v != null && setB(k, { actionPct: v })} /></td>
                  <td style={td}><PctInput width={46} max={999} value={b.rbPct ?? 100} onChange={(v) => v != null && setB(k, { rbPct: v })} /></td>
                  <td style={td}>{e ? fmt(e.fee) : "—"}</td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? money(e.backerBook) : "—"}</td>
                  <td style={td}><button onClick={() => removeB(k)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <input placeholder="Add player by exact nickname…" value={newName} onChange={(e) => setNewName(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && addBacked()} style={{ ...inputS, width: 260 }} />
        <select value={newDeal} onChange={(e) => setNewDeal(e.target.value)} style={{ ...inputS, fontSize: 12 }}>
          <option value="stake">Stake deal (makeup)</option><option value="action">Action buy</option>
        </select>
        <Btn tone="ghost" small onClick={addBacked}>+ Add</Btn>
        <div style={{ marginLeft: "auto", fontSize: 13, textAlign: "right" }}>
          <div>Deal books this week — settle these separately: {ownerIds.map((o, i) => <span key={o}>{i ? " · " : ""}{lbl(o)} <b>{money(model.backedBook[o])}</b></span>)}</div>
          <div style={{ color: C.mute, fontSize: 12, marginTop: 2 }}>Rake margins → backer (in profit, not the books above): {ownerIds.map((o, i) => <span key={o}>{i ? " · " : ""}{lbl(o)} <b>{money(model.dealMargin[o])}</b></span>)}</div>
        </div>
      </div>
    </div>
  );
}

function AAReportsTab({ model, cfg, period, expanded, setExpanded, club }) {
  const [exportData, setExportData] = useState(null);
  const toggle = (k) => setExpanded({ ...expanded, [k]: !expanded[k] });

  const miniHead = (labels) => (
    <thead><tr>{labels.map((l, i) => <th key={i} style={{ ...th, textAlign: i === 0 ? "left" : "right" }}>{l}</th>)}</tr></thead>
  );
  const sectionTitle = (t) => (
    <div style={{ fontSize: 11.5, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", margin: "14px 0 6px" }}>{t}</div>
  );

  // Each member carries its own tr/trCut (buildAAModel applies playerTr
  // overrides per player, falling back to the line's own TR dial) — so a
  // player with a different TR deal than the rest of their line shows the
  // correct cut right on their own row, not just at the line total.
  const memberTable = (e) => {
    const eNet = e.pnl + e.tipback;
    return (
    <table style={{ width: "100%", borderCollapse: "collapse" }}>
      {miniHead(["Player", "Agent", "Hands", "Winnings", "Tips", "TB %", "Rakeback", "TR %", "Settlement"])}
      <tbody>
        {[...e.members].sort((a, b) => b.fee - a.fee).map((m, i) => (
          <tr key={m.memberId + i} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
            <td style={{ ...tdL, fontWeight: 600 }}>{m.name}</td>
            <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{m.agentName !== "-" ? m.agentName : "—"}</td>
            <td style={td}>{fmtI(m.hands)}</td><td style={td}>{fmt(m.pnl)}</td><td style={td}>{fmt(m.fee)}</td>
            <td style={td}>{m.tbPct}%</td><td style={td}>{fmt(m.tipback)}</td><td style={td}>{aaPctS(m.tr || 0)}</td><td style={td}>{money(m.pnl + m.tipback - (m.trCut || 0))}</td>
          </tr>
        ))}
        <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
          <td style={{ ...tdL, fontWeight: 700 }} colSpan={2}>Total</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmtI(e.hands)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(e.pnl)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(e.fee)}</td>
          <td style={td}>{e.fee ? aaPctS((e.tipback / e.fee) * 100) : "—"}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(e.tipback)}</td>
          <td style={td}>{eNet ? aaPctS((e.trCut / eNet) * 100) : "—"}</td>
          <td style={{ ...td, fontWeight: 700 }}>{money(e.settlement)}</td>
        </tr>
      </tbody>
    </table>
  );};

  const copyLine = (e) => { const eNet = e.pnl + e.tipback; setExportData({
    title: `${e.name} · ${period || "this week"}`,
    text: toTSV(["Player", "Agent", "Hands", "Winnings", "Tips", "TB %", "Rakeback", "TR %", "Settlement"],
      [...e.members.map((m) => [m.name, m.agentName, m.hands, m.pnl.toFixed(2), m.fee.toFixed(2), m.tbPct, m.tipback.toFixed(2), aaPctN(m.tr || 0), (m.pnl + m.tipback - (m.trCut || 0)).toFixed(2)]),
       ["TOTAL", "", e.hands, e.pnl.toFixed(2), e.fee.toFixed(2), e.fee ? aaPctN((e.tipback / e.fee) * 100) : "", e.tipback.toFixed(2), eNet ? aaPctN((e.trCut / eNet) * 100) : "", e.settlement.toFixed(2)]]),
  }); };

  const rowShell = (key, header, body) => (
    <div key={key} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
      {header}
      {expanded[key] && <div style={{ padding: "4px 16px 16px" }}>{body}</div>}
    </div>
  );

  // ——— Owner report rows ———
  const H = model.H, lbl = H.lbl;
  const ownerRow = (o) => {
    const L = lbl(o);
    const key = `rep-owner-${o}`;
    const inPool = H.poolShare(o) > 0;
    const ownAcc = model.own.filter((p) => p.owner === o);
    const personal = model.entities.filter((e) => e.tag === o);
    const shared = inPool ? model.entities.filter((e) => e.tag === "agent") : [];
    const deals = model.backedEntities.filter((e) => H.shareOf(e.backer, o) > 0);
    const header = (
      <div onClick={() => toggle(key)} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 16px", cursor: "pointer", background: C.cream }}>
        <span style={{ color: C.goldDark, fontSize: 12, width: 12 }}>{expanded[key] ? "▼" : "►"}</span>
        <span style={{ fontWeight: 700, fontSize: 14.5 }}>{L}</span>
        <Pill tone="gold">owner</Pill>
        <span style={{ color: C.mute, fontSize: 12 }}>
          {personal.length} personal line{personal.length !== 1 ? "s" : ""}{ownAcc.length ? ` · ${ownAcc.length} own account${ownAcc.length !== 1 ? "s" : ""}` : ""}{shared.length ? ` · ${shared.length} shared agent${shared.length !== 1 ? "s" : ""}` : ""}
        </span>
        <span style={{ marginLeft: "auto", fontSize: 13.5 }}>profit {money(model.profit[o])}</span>
        <Btn tone="gold" small onClick={(ev) => { ev.stopPropagation(); downloadAAOwnerExcel(model, o, period, club); }}>Excel</Btn>
      </div>
    );
    const body = (
      <>
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap", background: C.banner, borderRadius: 8, padding: "10px 14px", marginTop: 10, fontSize: 13 }}>
          {inPool && <span>Pool share <b>{fmt(model.poolShares[o])}</b></span>}
          <span>Personal margins <b>{fmt(model.personalMargin[o])}</b></span>
          <span>Deal rake margins <b>{fmt(model.dealMargin[o])}</b></span>
          <span>Own play P&L <b>{fmt(model.ownPosition[o])}</b></span>
          <span>Jackpot share <b>{fmt(model.jpShares[o])}</b></span>
          <span style={{ marginLeft: "auto" }}>Profit <b style={{ color: model.profit[o] >= 0 ? C.green : C.red }}>{fmt(model.profit[o])}</b></span>
        </div>

        {ownAcc.length > 0 && <>
          {sectionTitle("Personal play — own accounts")}
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            {miniHead(["Account", "Hands", "Winnings", "Tips", "Feeback", "Position → P&L"])}
            <tbody>
              {ownAcc.map((p, i) => (
                <tr key={p.memberId} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}</td>
                  <td style={td}>{fmtI(p.hands)}</td><td style={td}>{fmt(p.pnl)}</td><td style={td}>{fmt(p.fee)}</td>
                  <td style={td}>{fmt(p.fee)}</td><td style={td}><b>{money(p.position)}</b></td>
                </tr>
              ))}
            </tbody>
          </table>
        </>}

        {personal.length > 0 && <>
          {sectionTitle(`Personal players & agents — margin 100% to ${L}`)}
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            {miniHead(["Line", "Hands", "Winnings", "Tips", "TB %", "Rakeback paid", "Margin → " + L])}
            <tbody>
              {personal.map((e, i) => (
                <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{e.name} {typePill(e)}</td>
                  <td style={td}>{fmtI(e.hands)}</td><td style={td}>{fmt(e.pnl)}</td><td style={td}>{fmt(e.fee)}</td>
                  <td style={td}>{e.fee ? aaPctS((e.tipback / e.fee) * 100) : "—"}</td>
                  <td style={td}>{fmt(e.tipback)}</td><td style={td}><b>{money(e.margin)}</b></td>
                </tr>
              ))}
              <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
                <td style={{ ...tdL, fontWeight: 700 }} colSpan={6}>Total personal margin</td>
                <td style={{ ...td, fontWeight: 700 }}>{money(model.personalMargin[o])}</td>
              </tr>
            </tbody>
          </table>
        </>}

        {shared.length > 0 && <>
          {sectionTitle(`Shared agents — ${H.poolLabel}`)}
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            {miniHead(["Agent line", "Hands", "Winnings", "Tips", "Rakeback", "Margin → pool", `${L}'s share`])}
            <tbody>
              {shared.map((e, i) => (
                <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{e.name} {typePill(e)}</td>
                  <td style={td}>{fmtI(e.hands)}</td><td style={td}>{fmt(e.pnl)}</td><td style={td}>{fmt(e.fee)}</td>
                  <td style={td}>{fmt(e.tipback)}</td><td style={td}>{fmt(e.margin)}</td><td style={td}><b>{money(e.margin * H.poolShare(o))}</b></td>
                </tr>
              ))}
              <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
                <td style={{ ...tdL, fontWeight: 700 }} colSpan={5}>Club pool</td>
                <td style={{ ...td, fontWeight: 700 }}>{fmt(model.pool)}</td>
                <td style={{ ...td, fontWeight: 700 }}>{money(model.poolShares[o])}</td>
              </tr>
            </tbody>
          </table>
        </>}

        {deals.length > 0 && <>
          {sectionTitle("Stake & action book — settles separately")}
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            {miniHead(["Player", "Deal", "Winnings", "RB credit", "Net", `${L}'s book share`, `Rake margin → ${L}`])}
            <tbody>
              {deals.map((e, i) => (
                <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{e.name}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{e.dealType === "action" ? `action ${e.actionPct}%` : "stake"}{e.backer === "split" ? " · split" : ""}</td>
                  <td style={td}>{fmt(e.pnl)}</td><td style={td}>{fmt(e.rbCredit)}</td><td style={td}>{fmt(e.net)}</td>
                  <td style={td}><b>{money(e.backerBook * H.shareOf(e.backer, o))}</b></td>
                  <td style={td}>{money(e.margin * H.shareOf(e.backer, o))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </>}
      </>
    );
    return rowShell(key, header, body);
  };

  // ——— Agent / player report rows (only union lines the owners share or run; each shows just its own players) ———
  const lineRow = (e) => {
    const key = `rep-${e.key}`;
    const header = (
      <div onClick={() => toggle(key)} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 16px", cursor: "pointer", background: C.cream }}>
        <span style={{ color: C.goldDark, fontSize: 12, width: 12 }}>{expanded[key] ? "▼" : "►"}</span>
        <span style={{ fontWeight: 700, fontSize: 14.5 }}>{e.name}</span>
        {typePill(e)}
        {e.tag && e.tag !== "agent" && <Pill tone="green">personal · {lbl(e.tag)}</Pill>}
        {e.tag === "agent" && <Pill>pool 50/50</Pill>}
        <span style={{ color: C.mute, fontSize: 12 }}>{e.members.length} player{e.members.length !== 1 ? "s" : ""} · {fmtI(e.hands)} hands</span>
        <span style={{ marginLeft: "auto", fontSize: 13.5 }}>settlement {money(e.settlement)}</span>
        <Btn tone="gold" small onClick={(ev) => { ev.stopPropagation(); downloadAALineExcel(e, period, model); }}>Excel</Btn>
        <Btn tone="ghost" small onClick={(ev) => { ev.stopPropagation(); copyLine(e); }}>Copy</Btn>
      </div>
    );
    return rowShell(key, header, <div style={{ marginTop: 8 }}>{memberTable(e)}</div>);
  };

  const dealRow = (e) => {
    const key = `rep-${e.key}`;
    const header = (
      <div onClick={() => toggle(key)} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 16px", cursor: "pointer", background: C.cream }}>
        <span style={{ color: C.goldDark, fontSize: 12, width: 12 }}>{expanded[key] ? "▼" : "►"}</span>
        <span style={{ fontWeight: 700, fontSize: 14.5 }}>{e.name}</span>
        <Pill tone="blue">{e.dealType === "action" ? `action ${e.actionPct}%` : "stake"}</Pill>
        {e.dealType === "makeup" && e.inMakeup && <Pill tone="red">in makeup</Pill>}
        <span style={{ color: C.mute, fontSize: 12 }}>backer {lbl(e.backer)}</span>
        <span style={{ marginLeft: "auto", fontSize: 13.5 }}>player gets {money(e.settlement)}</span>
        <Btn tone="gold" small onClick={(ev) => { ev.stopPropagation(); downloadAALineExcel(e, period, model); }}>Excel</Btn>
      </div>
    );
    const body = (
      <div style={{ display: "flex", gap: 22, flexWrap: "wrap", background: C.banner, borderRadius: 8, padding: "10px 14px", marginTop: 10, fontSize: 13 }}>
        <span>P&L <b>{fmt(e.pnl)}</b></span>
        <span>Tips <b>{fmt(e.fee)}</b></span>
        <span>RB credit <b>{fmt(e.rbCredit)}</b>{e.rb != null ? <span style={{ color: C.mute }}> @{e.rb}%</span> : null}</span>
        <span>Net <b>{fmt(e.net)}</b></span>
        {e.dealType === "makeup" && <span>Makeup <b>{fmt(e.makeupBefore)}</b> → <b>{fmt(e.makeupAfter)}</b></span>}
        <span>Player <b>{fmt(e.settlement)}</b></span>
        <span>Backer book <b>{fmt(e.backerBook)}</b></span>
      </div>
    );
    return rowShell(key, header, body);
  };

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ fontFamily: "Georgia, serif", fontSize: 19, marginBottom: 4 }}>Reports</div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 14 }}>
        Click a row to expand; Excel downloads a styled workbook to send. Owner reports show only that owner's world — his own play, his personal players, and the shared agents — never another owner's personal players.
      </div>

      <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 8 }}>Owners</div>
      {model.ownerIds.map(ownerRow)}

      <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", margin: "20px 0 8px" }}>Agents & players</div>
      {model.entities.map(lineRow)}
      {model.backedEntities.length > 0 && <>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", margin: "20px 0 8px" }}>Stake & action players</div>
        {model.backedEntities.map(dealRow)}
      </>}
    </div>
  );
}


// ——— Owner club setup: name, owners (label · pool % · JP %), which owner is you, delete club ———
function OwnerClubSetup({ club, clubs, saveClubs, onDeleteClub }) {
  const [name, setName] = useState(club.name);
  useEffect(() => setName(club.name), [club.id, club.name]);
  const patch = (p) => saveClubs(clubs.map((c) => (c.id === club.id ? { ...c, ...p } : c)));
  const setOwner = (id, p) => patch({ owners: club.owners.map((o) => (o.id === id ? { ...o, ...p } : o)) });
  const addOwner = () => { const id = "o-" + uid(); patch({ owners: [...club.owners, { id, label: "New owner", poolPct: 0, jpPct: 0 }] }); };
  const delOwner = (id) => { if (club.owners.length <= 1) return; if (!window.confirm("Remove this owner? Lines tagged to them will need re-tagging.")) return; patch({ owners: club.owners.filter((o) => o.id !== id), meId: club.meId === id ? club.owners.find((o) => o.id !== id)?.id : club.meId }); };
  const H = ocHelpers(club);
  return (
    <div style={{ marginTop: 6 }}>
      <Card title="Club" right={<Pill tone="gold">owner club</Pill>}>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <span style={{ fontSize: 13, color: C.mute }}>Name</span>
          <input value={name} onChange={(e) => setName(e.target.value)} onBlur={() => name.trim() && name.trim() !== club.name && patch({ name: name.trim() })}
            onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()} style={{ ...inputS, width: 220, fontWeight: 700 }} />
          <span style={{ fontSize: 12, color: C.mute }}>Same weekly export format as All American. Deals, tags, and makeup live with this club only.</span>
          <span style={{ marginLeft: "auto" }}>
            <Btn tone="ghost" small onClick={() => { if (window.confirm(`Delete club "${club.name}" and all its saved weeks, deals, and tags? Archived files stay in Archive.`)) onDeleteClub(club.id); }}>Delete club</Btn>
          </span>
        </div>
      </Card>
      <div style={{ height: 14 }} />
      <Card title="Owners" right={<Btn tone="ghost" small onClick={addOwner}>+ Owner</Btn>}>
        <div style={{ fontSize: 12.5, color: C.mute, marginBottom: 10 }}>
          <b>Pool %</b> — how the agent-line margin pool splits (0 = not in the pool; shares are normalized, so 50/50/0 and 1/1/0 mean the same). <b>JP %</b> — how the bad beat contribution splits. <b>Me</b> marks your own seat: your share of stake/action books feeds Tabs → Staking. Owners can be just you.
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Owner</th><th style={th}>Pool %</th><th style={{ ...th }}>Pool share</th><th style={th}>JP %</th><th style={th}>JP share</th><th style={{ ...th, textAlign: "center" }}>Me</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {club.owners.map((o, i) => (
              <tr key={o.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                <td style={tdL}><input value={o.label} onChange={(e) => setOwner(o.id, { label: e.target.value })} style={{ ...inputS, width: 160, fontWeight: 600 }} /></td>
                <td style={td}><NumInput width={70} value={o.poolPct ?? 0} onChange={(v) => setOwner(o.id, { poolPct: v ?? 0 })} /></td>
                <td style={{ ...td, color: C.mute }}>{H.pctS(H.poolShare(o.id))}</td>
                <td style={td}><NumInput width={70} value={o.jpPct ?? 0} onChange={(v) => setOwner(o.id, { jpPct: v ?? 0 })} /></td>
                <td style={{ ...td, color: C.mute }}>{H.pctS(H.jpShareOf(o.id))}</td>
                <td style={{ ...td, textAlign: "center" }}><input type="radio" name={"me-" + club.id} checked={club.meId === o.id} onChange={() => patch({ meId: o.id })} /></td>
                <td style={{ ...td, width: 30 }}><button onClick={() => delOwner(o.id)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

const aaMemberOverrideRow = (m, ctx) => (
  <div key={m.memberId} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, background: C.rowAlt, borderRadius: 6, padding: "5px 10px" }}>
    <span>{m.name}</span>
    <span style={{ marginLeft: "auto", color: C.mute, fontSize: 10.5 }}>{ctx.tbFallback != null ? `TB ${aaPctS(ctx.tbFallback)}` : ""}</span>
    <PctInput width={44} max={999} value={ctx.deals[m.memberId] ?? ""} onChange={(v) => ctx.setDeal(m.memberId, v)} />
    <span style={{ color: C.mute, fontSize: 10.5 }}>{ctx.trFallback != null ? `TR ${aaPctS(ctx.trFallback)}` : ""}</span>
    <PctInput width={40} value={ctx.playerTr[m.memberId] ?? ""} onChange={(v) => ctx.setPlayerTr(m.memberId, v)} />
  </div>
);

// A super agent or agent's own personal-play row, pinned to a VIP TB rate of
// its own via ownDeals — a separate map from `deals` on purpose. That row's
// memberId is the same ID as the sa/agent line itself, and `deals[that id]`
// is also what every other player under them falls back to when they have
// no closer override; if this VIP dial wrote into `deals` too, pinning the
// top person's own rate would silently move the whole group's default along
// with it. Blank here just means "follow that same default", same as today.
// The TR dial next to it writes into the shared `playerTr` map — TR never
// had that collision problem (the line's own TR dial lives in a completely
// separate `cfg.tr`, keyed by the line, not by any one member), so any
// player — including this one — can just pin their own TR straight in it.
const aaOwnOverrideRow = (m, ctx) => (
  <div key={"own:" + m.memberId} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, background: C.banner, borderRadius: 6, padding: "5px 10px", border: `1px dashed ${C.gold}`, flexWrap: "wrap" }}>
    <span style={{ fontWeight: 700 }}>{m.name}</span>
    <Pill tone="gold">own play · VIP</Pill>
    <span style={{ marginLeft: "auto", color: C.mute, fontSize: 10.5 }}>{ctx.tbFallback != null ? `TB ${aaPctS(ctx.tbFallback)}` : ""}</span>
    <PctInput width={44} max={999} value={ctx.ownDeals[m.memberId] ?? ""} onChange={(v) => ctx.setOwnDeal(m.memberId, v)} />
    <span style={{ color: C.mute, fontSize: 10.5 }}>{ctx.trFallback != null ? `TR ${aaPctS(ctx.trFallback)}` : ""}</span>
    <PctInput width={40} value={ctx.playerTr[m.memberId] ?? ""} onChange={(v) => ctx.setPlayerTr(m.memberId, v)} />
  </div>
);

// Body of the override panel for one sa/agent entity — no outer padding, so
// it can be reused both as a top-level line's panel and nested inside a DL
// umbrella's per-subgroup breakdown. `trFallback` is the LINE's own TR dial,
// passed in from the top (AALineOverrides) rather than derived from `e.id`
// here, because for a DL umbrella that's the umbrella's own id, not any one
// subgroup's, so it has to be resolved by the caller who knows which one
// it's building the panel for. A nested agent (inside an SA's tree) can then
// override that default for just their own branch via agentTr — see below.
function AAEntityOverrideBody({ e, deals, ownDeals, playerTr, agentTr, trFallback, cfg, setDeal, setOwnDeal, setPlayerTr, setAgentTr }) {
  if (e.type === "agent") {
    // Top-level agent (no super agent above it) — its players, plus (if the
    // agent also has their own personal-play row) a VIP dial for just that
    // row. The agent's default TB and TR rates for everyone else are still
    // the top-level dials (the row's own TB% input, and trFallback) — there's
    // no separate agentTr level to set here since this agent already IS the
    // top of the line.
    const self = e.members.find((m) => m.memberId === e.id);
    const members = e.members.filter((m) => m.memberId !== e.id);
    if (!members.length && !self) return null;
    const fallback = deals[e.id] ?? cfg.defaultTB;
    const ctx = { tbFallback: fallback, deals, setDeal, ownDeals, setOwnDeal, trFallback, playerTr, setPlayerTr };
    return (
      <div>
        {self && <div style={{ marginBottom: members.length ? 8 : 0 }}>{aaOwnOverrideRow(self, ctx)}</div>}
        {members.length > 0 && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 6 }}>
            {members.map((m) => aaMemberOverrideRow(m, ctx))}
          </div>
        )}
      </div>
    );
  }

  if (e.type !== "sa") return null;

  // Super agent tree: split into agent sub-groups (each gets its own TB and
  // TR rate, falling back to the SA's rate) and players sitting directly
  // under the SA. The SA's own play (if any) gets its own VIP dial, shown
  // separately up top. A nested agent's TR dial (agentTr) is a real cascade
  // level, same as their TB dial already is — set it once and every one of
  // that agent's players (present or future) inherits it automatically,
  // unless a specific player is pinned with their own playerTr override.
  const own = e.members.find((m) => m.memberId === e.id);
  const byAgent = new Map(); const direct = [];
  e.members.forEach((m) => {
    if (m.memberId === e.id) return; // the SA's own play — shown separately above
    if (m.agentId && m.agentId !== "-") {
      if (!byAgent.has(m.agentId)) byAgent.set(m.agentId, { agentId: m.agentId, agentName: m.agentName, members: [] });
      byAgent.get(m.agentId).members.push(m);
    } else direct.push(m);
  });
  const saFallback = deals[e.id] ?? cfg.defaultTB;
  const saCtx = { tbFallback: saFallback, deals, setDeal, ownDeals, setOwnDeal, trFallback, playerTr, setPlayerTr };

  return (
    <div>
      {own && <div style={{ marginBottom: 8 }}>{aaOwnOverrideRow(own, saCtx)}</div>}
      {[...byAgent.values()].map((g) => {
        const agentFallback = deals[g.agentId] ?? saFallback;
        const agentTrFallback = agentTr[g.agentId] ?? trFallback;
        const agentCtx = { ...saCtx, tbFallback: agentFallback, trFallback: agentTrFallback };
        const agentSelf = g.members.find((m) => m.memberId === g.agentId);
        const players = g.members.filter((m) => m.memberId !== g.agentId);
        return (
          <div key={g.agentId} style={{ marginBottom: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, fontWeight: 700, marginBottom: 4, flexWrap: "wrap" }}>
              <span>{g.agentName}</span>
              <Pill tone="gold">agent</Pill>
              <span style={{ color: C.mute, fontWeight: 400, fontSize: 11 }}>TB default {aaPctS(saFallback)}</span>
              <PctInput width={44} max={999} value={deals[g.agentId] ?? ""} onChange={(v) => setDeal(g.agentId, v)} />
              <span style={{ color: C.mute, fontWeight: 400, fontSize: 11 }}>TR default {aaPctS(trFallback)}</span>
              <PctInput width={40} value={agentTr[g.agentId] ?? ""} onChange={(v) => setAgentTr(g.agentId, v)} />
              <span style={{ marginLeft: "auto", color: C.mute, fontSize: 10.5 }}>→ every {g.agentName} player inherits this, unless overridden below</span>
            </div>
            <div style={{ marginLeft: 14 }}>
              {agentSelf && <div style={{ marginBottom: players.length ? 6 : 0 }}>{aaOwnOverrideRow(agentSelf, agentCtx)}</div>}
              {players.length > 0 && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 6 }}>
                  {players.map((m) => aaMemberOverrideRow(m, agentCtx))}
                </div>
              )}
            </div>
          </div>
        );
      })}
      {direct.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 6 }}>
          {direct.map((m) => aaMemberOverrideRow(m, saCtx))}
        </div>
      )}
    </div>
  );
}

// Player-override panel for a Super Agent, Agent, or DL umbrella line on the
// Lines & ownership table — same idea as Fish Tank's SA deal / player-
// override list, extended with the extra Agent level a union tree can have,
// and with DL umbrellas (several sa/agent/player lines billed through one
// person) showing each of their subgroups broken out underneath. A blank
// input means "no override, inherit the rate shown as the default"; typing
// a number pins that player (or that whole agent, or a whole subgroup) to a
// rate of its own — this is the VIP-deal override. TR's top dial lives here
// too — one per line (per umbrella, not per subgroup inside it) — resolved
// once and handed down as the fallback every nested override row starts
// from; a nested agent can then pin their own TR default (agentTr, set
// inside AAEntityOverrideBody) that cascades to just their own players.
function AALineOverrides({ e, cfg, setDeal, setOwnDeal, setPlayerTr, setAgentTr }) {
  const deals = cfg.deals || {};
  const ownDeals = cfg.ownDeals || {};
  const playerTr = cfg.playerTr || {};
  const agentTr = cfg.agentTr || {};

  if (e.type === "dlUmbrella") {
    const trFallback = (cfg.tr || {})[e.id] ?? cfg.defaultTR ?? 0;
    return (
      <div style={{ padding: "10px 16px 14px 40px" }}>
        {e.subgroups.map((sub) => (
          <div key={sub.key} style={{ marginBottom: 10, borderBottom: `1px solid ${C.line}`, paddingBottom: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, marginBottom: 6 }}>
              <span>{sub.name}</span>
              {typePill(sub)}
              <span style={{ color: C.mute, fontWeight: 400, fontSize: 11.5 }}>tips {fmt(sub.fee)}</span>
              <span style={{ marginLeft: "auto" }} />
              <PctInput width={50} max={999} value={deals[sub.id] ?? cfg.defaultTB} onChange={(v) => setDeal(sub.id, v)} />
            </div>
            <div style={{ marginLeft: 14 }}>
              <AAEntityOverrideBody e={sub} deals={deals} ownDeals={ownDeals} playerTr={playerTr} agentTr={agentTr} trFallback={trFallback} cfg={cfg} setDeal={setDeal} setOwnDeal={setOwnDeal} setPlayerTr={setPlayerTr} setAgentTr={setAgentTr} />
            </div>
          </div>
        ))}
      </div>
    );
  }

  const trFallback = (cfg.tr || {})[e.id] ?? cfg.defaultTR ?? 0;
  const body = <AAEntityOverrideBody e={e} deals={deals} ownDeals={ownDeals} playerTr={playerTr} agentTr={agentTr} trFallback={trFallback} cfg={cfg} setDeal={setDeal} setOwnDeal={setOwnDeal} setPlayerTr={setPlayerTr} setAgentTr={setAgentTr} />;
  if (!body) return null;
  return <div style={{ padding: "10px 16px 14px 40px" }}>{body}</div>;
}

// DL umbrellas — group super agents, top-level agents, and unlinked players
// that are all actually settled through one person into a single combined
// line (one class, one collector), exactly like Fish Tank's umbrella groups.
// Membership is tracked by each entity's full key (sa:/ag:/p:) rather than a
// raw ID, since those three id-spaces can otherwise collide.
function AADLUmbrellasTab({ model, cfg, up }) {
  const [name, setName] = useState("");
  const umbrellas = cfg.dlUmbrellas || [];
  const assignable = model.dlAssignable || [];

  const addUmbrella = () => {
    if (!name.trim()) return;
    up({ dlUmbrellas: [...umbrellas, { id: "dl" + Date.now(), name: name.trim(), memberKeys: [] }] });
    setName("");
  };
  const renameUmbrella = (id, v) => up({ dlUmbrellas: umbrellas.map((u) => (u.id === id ? { ...u, name: v } : u)) });
  const deleteUmbrella = (id) => up({ dlUmbrellas: umbrellas.filter((u) => u.id !== id) });
  const toggleMember = (uid, key) => {
    up({ dlUmbrellas: umbrellas.map((u) => {
      if (u.id === uid) return { ...u, memberKeys: u.memberKeys.includes(key) ? u.memberKeys.filter((k) => k !== key) : [...u.memberKeys, key] };
      return { ...u, memberKeys: u.memberKeys.filter((k) => k !== key) };
    }) });
  };

  const groups = [
    ["Super agents", assignable.filter((e) => e.type === "sa")],
    ["Agents", assignable.filter((e) => e.type === "agent")],
    ["Unlinked players", assignable.filter((e) => e.type === "player")],
    ["Managers", assignable.filter((e) => e.type === "manager")],
    ["Masters", assignable.filter((e) => e.type === "master")],
  ].filter(([, list]) => list.length > 0);

  return (
    <div>
      <div style={{ fontFamily: "Georgia, serif", fontSize: 19, marginBottom: 4 }}>DL umbrellas</div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 16 }}>
        Some super agents, agents, unlinked players, managers, and masters — including a club owner's own Manager/Master line — are all actually settled through the same person. Group those lines into one DL umbrella and they combine into a single line on Lines & ownership — one class, one collector — while each member keeps its own rate; expand "player overrides" on the merged line to override any subgroup or player underneath.
      </div>
      <Card title="Umbrella groups" right={
        <span style={{ display: "flex", gap: 8 }}>
          <input placeholder="New umbrella name…" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUmbrella()} style={{ ...inputS, width: 200 }} />
          <Btn tone="ghost" small onClick={addUmbrella}>+ Create</Btn>
        </span>
      }>
        {umbrellas.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No DL umbrellas yet.</div>}
        {umbrellas.map((u) => (
          <div key={u.id} style={{ borderTop: `1px solid ${C.line}`, padding: "12px 0" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
              <input value={u.name} onChange={(e) => renameUmbrella(u.id, e.target.value)} style={{ ...inputS, fontWeight: 700, width: 200 }} />
              <Pill tone="blue">{u.memberKeys.length} member{u.memberKeys.length !== 1 ? "s" : ""}</Pill>
              <button onClick={() => deleteUmbrella(u.id)} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>× delete</button>
            </div>
            {groups.map(([label, list]) => (
              <div key={label} style={{ marginBottom: 8 }}>
                <div style={{ fontSize: 11, color: C.mute, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{label}</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {list.map((o) => {
                    const inThis = u.memberKeys.includes(o.key);
                    const inOther = !inThis && umbrellas.some((x) => x.id !== u.id && x.memberKeys.includes(o.key));
                    return (
                      <button key={o.key} onClick={() => !inOther && toggleMember(u.id, o.key)} style={{
                        padding: "4px 12px", borderRadius: 14, fontSize: 12, fontWeight: 600, cursor: inOther ? "default" : "pointer",
                        border: `1px solid ${inThis ? C.goldDark : C.line}`,
                        background: inThis ? C.gold : C.surface, color: inThis ? "var(--onGold)" : inOther ? "var(--chipOff)" : C.mute, opacity: inOther ? 0.6 : 1 }}>
                        {o.name}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
            {groups.length === 0 && <div style={{ color: C.mute, fontSize: 12.5 }}>No super agents, agents, or unlinked players on this week's export yet.</div>}
          </div>
        ))}
      </Card>
    </div>
  );
}

// BBJ holds — when an owner isn't ready to collect their bad beat jackpot
// share, another owner holds it for them instead. The weekly redirect
// itself happens in buildAAModel (see jpHoldMoves); this card is the
// running ledger of what's been held, rebuilt from the club's own jackpot
// history (cfg.jackpots — already persisted per period regardless of this
// feature) times the held owner's JP %, plus any payouts logged once some
// or all of it actually changes hands.
function aaJpHoldLedger(hold, cfg, H) {
  const jackpots = cfg.jackpots || {};
  const rows = Object.entries(jackpots)
    .filter(([period, amt]) => (!hold.sincePeriod || period >= hold.sincePeriod) && Math.abs(+amt || 0) > 0.005)
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([period, amt]) => ({ period, jackpot: +amt || 0, share: (+amt || 0) * H.jpShareOf(hold.forOwnerId) }));
  const held = rows.reduce((a, r) => a + r.share, 0);
  const paidOut = (hold.payouts || []).reduce((a, p) => a + (+p.amount || 0), 0);
  return { rows, held, paidOut, running: held - paidOut };
}

function AABBJHoldsCard({ model, cfg, up, period }) {
  const ownerIds = model?.ownerIds || [];
  const H = model?.H;
  const lbl = H ? H.lbl : (o) => o;
  const holds = cfg.jpHolds || [];
  const [holderId, setHolderId] = useState(ownerIds[0] || "");
  const [forOwnerId, setForOwnerId] = useState(ownerIds[1] || ownerIds[0] || "");
  const [open, setOpen] = useState({});
  const [payoutDrafts, setPayoutDrafts] = useState({});

  useEffect(() => {
    if (!ownerIds.includes(holderId)) setHolderId(ownerIds[0] || "");
    if (!ownerIds.includes(forOwnerId) || forOwnerId === holderId) setForOwnerId(ownerIds.find((o) => o !== holderId) || ownerIds[0] || "");
  }, [ownerIds.join("|")]);

  if (!model || ownerIds.length < 2) return null;

  const addHold = () => {
    if (!holderId || !forOwnerId || holderId === forOwnerId) return;
    up({ jpHolds: [...holds, { id: "jph" + Date.now(), holderId, forOwnerId, sincePeriod: period || "", payouts: [] }] });
  };
  const deleteHold = (id) => { if (!window.confirm("Delete this BBJ hold? Its running total goes with it — log any final payout first if you need the record.")) return; up({ jpHolds: holds.filter((h) => h.id !== id) }); };
  const setSince = (id, v) => up({ jpHolds: holds.map((h) => (h.id === id ? { ...h, sincePeriod: v } : h)) });
  const addPayout = (id) => {
    const d = payoutDrafts[id] || {};
    const amt = parseFloat(d.amount);
    if (!isFinite(amt) || amt <= 0) return;
    up({ jpHolds: holds.map((h) => (h.id === id ? { ...h, payouts: [...(h.payouts || []), { id: "po" + Date.now(), date: d.date || today(), amount: amt, note: d.note || "" }] } : h)) });
    setPayoutDrafts({ ...payoutDrafts, [id]: { date: "", amount: "", note: "" } });
  };
  const removePayout = (holdId, payoutId) => up({ jpHolds: holds.map((h) => (h.id === holdId ? { ...h, payouts: (h.payouts || []).filter((p) => p.id !== payoutId) } : h)) });

  return (
    <Card title="BBJ holds — holding a jackpot share for another owner" right={
      <span style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12.5 }}>
        <select value={holderId} onChange={(e) => setHolderId(e.target.value)} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
          {ownerIds.map((o) => <option key={o} value={o}>{lbl(o)}</option>)}
        </select>
        <span style={{ color: C.mute }}>holds for</span>
        <select value={forOwnerId} onChange={(e) => setForOwnerId(e.target.value)} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
          {ownerIds.filter((o) => o !== holderId).map((o) => <option key={o} value={o}>{lbl(o)}</option>)}
        </select>
        <Btn tone="ghost" small onClick={addHold}>+ Add</Btn>
      </span>
    }>
      <div style={{ fontSize: 12.5, color: C.mute, marginBottom: 12 }}>
        When an owner isn't ready to collect their bad beat jackpot share, another owner can hold it for them each week — on Settlements, that week's share is fully redirected to the holder, while the held owner keeps settling everything else (personal lines, own accounts) normally. The running total below is what the holder currently owes back, built from the jackpot amount recorded each week since the hold started; log a payout whenever some or all of it actually changes hands.
      </div>
      {holds.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No BBJ holds set up yet.</div>}
      {holds.map((h) => {
        const { rows, held, paidOut, running } = aaJpHoldLedger(h, cfg, H);
        const isOpen = !!open[h.id];
        const draft = payoutDrafts[h.id] || {};
        return (
          <div key={h.id} style={{ borderTop: `1px solid ${C.line}`, padding: "12px 0" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 6 }}>
              <b style={{ fontSize: 14 }}>{lbl(h.holderId)}</b>
              <span style={{ color: C.mute, fontSize: 13 }}>holds for</span>
              <b style={{ fontSize: 14 }}>{lbl(h.forOwnerId)}</b>
              <span style={{ color: C.mute, fontSize: 12 }}>since</span>
              <input value={h.sincePeriod || ""} onChange={(e) => setSince(h.id, e.target.value)} placeholder="the beginning" style={{ ...inputS, width: 190, fontSize: 12 }} />
              <span style={{ marginLeft: "auto", fontWeight: 700 }}>Running total held: {money(running)}</span>
              <button onClick={() => setOpen({ ...open, [h.id]: !isOpen })} style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontSize: 12, fontWeight: 700 }}>
                {isOpen ? "hide detail" : "show detail"}
              </button>
              <button onClick={() => deleteHold(h.id)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
            </div>
            {isOpen && (
              <div style={{ marginLeft: 4 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 10 }}>
                  <thead><tr style={{ background: C.cream }}>
                    <th style={{ ...th, textAlign: "left" }}>Period</th><th style={th}>Jackpot</th><th style={th}>{lbl(h.forOwnerId)}'s share</th>
                  </tr></thead>
                  <tbody>
                    {rows.length === 0 && <tr><td colSpan={3} style={{ ...tdL, color: C.mute, padding: 10 }}>No recorded jackpot weeks yet since this hold started.</td></tr>}
                    {rows.map((r, i) => (
                      <tr key={r.period} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                        <td style={tdL}>{r.period}</td><td style={td}>{fmt(r.jackpot)}</td><td style={td}>{fmt(r.share)}</td>
                      </tr>
                    ))}
                    <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
                      <td style={{ ...tdL, fontWeight: 700 }} colSpan={2}>Total held</td>
                      <td style={{ ...td, fontWeight: 700 }}>{fmt(held)}</td>
                    </tr>
                  </tbody>
                </table>
                <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Payouts</div>
                {(h.payouts || []).length === 0 && <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 8 }}>None logged yet.</div>}
                {(h.payouts || []).map((p) => (
                  <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12.5, padding: "4px 0" }}>
                    <span style={{ color: C.mute }}>{p.date}</span>
                    <span style={{ fontWeight: 700 }}>{fmt(p.amount)}</span>
                    {p.note && <span style={{ color: C.mute }}>{p.note}</span>}
                    <button onClick={() => removePayout(h.id, p.id)} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 14 }}>×</button>
                  </div>
                ))}
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                  <input type="date" value={draft.date || ""} onChange={(e) => setPayoutDrafts({ ...payoutDrafts, [h.id]: { ...draft, date: e.target.value } })} style={{ ...inputS, fontSize: 12 }} />
                  <NumInput width={90} value={draft.amount ?? ""} onChange={(v) => setPayoutDrafts({ ...payoutDrafts, [h.id]: { ...draft, amount: v } })} />
                  <input placeholder="note (optional)" value={draft.note || ""} onChange={(e) => setPayoutDrafts({ ...payoutDrafts, [h.id]: { ...draft, note: e.target.value } })} style={{ ...inputS, flex: 1, minWidth: 140, fontSize: 12 }} />
                  <Btn tone="ghost" small onClick={() => addPayout(h.id)}>+ Log payout</Btn>
                </div>
                {paidOut > 0.005 && <div style={{ color: C.mute, fontSize: 11.5, marginTop: 6 }}>Paid out so far: {fmt(paidOut)}</div>}
              </div>
            )}
          </div>
        );
      })}
    </Card>
  );
}

function AllAmerican({ club, clubs, saveClubs, onDeleteClub }) {
  const H = ocHelpers(club);
  const { ownerIds, lbl } = H;
  const [cfg, setCfg] = useState(AA_DEFAULT_CFG);
  const [players, setPlayers] = useState(null);
  const [period, setPeriod] = useState("");
  const [jpFromExport, setJpFromExport] = useState(null);
  const [tab, setTab] = useState("settle");
  const [err, setErr] = useState("");
  const [saveNote, setSaveNote] = useState("");
  const [loaded, setLoaded] = useState(false);
  const [exportData, setExportData] = useState(null);
  const [repExpanded, setRepExpanded] = useState({});
  const [showMembers, setShowMembers] = useState({});
  const fileRef = useRef(null);

  useEffect(() => { (async () => {
    try { const c = await store.get(club.cfgKey); if (c?.value) setCfg({ ...AA_DEFAULT_CFG, ...JSON.parse(c.value) }); } catch (e) {}
    try { const d = await store.get(club.weekKey); if (d?.value) { const s = JSON.parse(d.value); setPlayers(s.players || null); setPeriod(s.period || ""); setJpFromExport(s.jackpotFromExport ?? null); } } catch (e) {}
    setLoaded(true);
  })(); }, [club.id]);

  const up = async (patch) => {
    const next = { ...cfg, ...patch };
    setCfg(next);
    try { await store.set(club.cfgKey, JSON.stringify(next)); setSaveNote(""); }
    catch (e) { setSaveNote("Change couldn't be saved — it still applies this session."); }
  };

  const onFile = async (file) => {
    setErr("");
    try {
      const buf = await file.arrayBuffer();
      const { players: p, period: per, jackpot: jp, jackpotFound } = parseWorkbook(buf);
      // Archive the week being replaced (raw upload + generated workbook) before it's gone.
      if (players && period && period !== per) {
        try { await archiveWeek(club.id, club.name, period, async () => { await downloadAAWorkbook(model, period, club); for (const o of ownerIds) await downloadAAOwnerExcel(model, o, period, club); }); } catch (e) {}
      }
      const names = { ...cfg.names };
      p.forEach((x) => { names[x.memberId] = x.name; if (x.saId !== "-") names[x.saId] = x.saName; });
      const patch = { names };
      if (jackpotFound) patch.jackpots = { ...cfg.jackpots, [per]: Math.round(jp * 100) / 100 };
      await up(patch);
      setPlayers(p); setPeriod(per); setJpFromExport(jackpotFound ? Math.round(jp * 100) / 100 : null); setTab("settle");
      try { await store.set(club.weekKey, JSON.stringify({ players: p, period: per, jackpotFromExport: jackpotFound ? jp : null })); } catch (e) {}
      try { await idbPut(rawKey(club.id, per), { site: club.id, siteName: club.name, period: per, name: file.name, buf }); } catch (e) {}
    } catch (e) { setErr(e.message || String(e)); }
  };

  const model = useMemo(() => (players ? buildAAModel(players, cfg, period, club) : null), [players, cfg, period, club]);
  const setTag = (key, tag) => up({ owners: { ...cfg.owners, [key]: tag } });
  const setDeal = (id, v) => up({ deals: { ...cfg.deals, [id]: v } });
  const setOwnDeal = (id, v) => up({ ownDeals: { ...(cfg.ownDeals || {}), [id]: v } });
  const setTr = (id, v) => up({ tr: { ...cfg.tr, [id]: v } });
  const setPlayerTr = (id, v) => up({ playerTr: { ...(cfg.playerTr || {}), [id]: v } });
  const setAgentTr = (id, v) => up({ agentTr: { ...(cfg.agentTr || {}), [id]: v } });
  const setCollector = (key, who) => up({ collectors: { ...cfg.collectors, [key]: who } });
  const setJackpot = (v) => up({ jackpots: { ...cfg.jackpots, [period]: v } });
  const setJpMode = (mode) => up({ jpModes: { ...(cfg.jpModes || {}), [period]: mode } });

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading saved setup…</div>;

  const jackpot = model ? model.jackpot : 0;
  const pendingCount = model ? model.untagged.length + model.uncollected.length : 0;

  const exportCsv = () => model && setExportData({
    title: `All American · ${period || "this week"}`,
    text: toTSV(["Line", "Class", "Collected by", "Winnings", "Tips", "TB %", "Rakeback paid", "Margin", "Margin goes to", "Union cash"],
      [...model.entities.map((e) => [e.name, e.tag ? (e.tag === "agent" ? "agent" : `personal ${e.tag}`) : "untagged", e.collector || "", e.pnl.toFixed(2), e.fee.toFixed(2), e.fee ? aaPctN((e.tipback / e.fee) * 100) : "", e.tipback.toFixed(2), e.margin.toFixed(2), e.tag === "agent" ? "pool" : e.tag || "", e.unionCash.toFixed(2)]),
       ...model.own.map((p) => [p.name, `own · ${p.owner}`, "", p.pnl.toFixed(2), p.fee.toFixed(2), "100", p.fee.toFixed(2), p.position.toFixed(2), `${p.owner} pnl`, "0"])]),
  });

  const ownerRow = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "6px 0", borderBottom: opts.rule ? `1px solid ${C.line}` : "none", fontSize: 13.5 }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{typeof val === "number" ? money(val) : val}</span>
    </div>
  );

  const collectorSel = (e) => (
    <select value={e.collector || ""} onChange={(ev) => setCollector(e.key, ev.target.value)} style={{ ...inputS, padding: "4px 6px", fontSize: 12, borderColor: e.collector ? C.line : C.red }}>
      {!e.collector && <option value="">— who collects? —</option>}
      {ownerIds.map((o) => <option key={o} value={o}>{lbl(o)}</option>)}
    </select>
  );

  return (
    <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />

      {err && <div style={{ background: "var(--errBg)", color: C.red, padding: "10px 14px", borderRadius: 6, marginBottom: 14, fontSize: 13.5 }}>{err}</div>}
      {saveNote && <div style={{ background: C.banner, color: C.goldDark, padding: "8px 14px", borderRadius: 6, marginBottom: 14, fontSize: 12.5 }}>{saveNote}</div>}

      {!players && (
        <div style={{ background: C.card, border: `1px dashed ${C.gold}`, borderRadius: 10, padding: "50px 30px", textAlign: "center" }}>
          <div style={{ fontFamily: "Georgia, serif", fontSize: 20, marginBottom: 8 }}>{club.name}</div>
          <div style={{ color: C.mute, fontSize: 14, marginBottom: 18 }}>
            Upload the club's weekly .xlsx export. The margin on every line (tips − rakeback) is the profit: agent lines feed the owner pool ({H.poolLabel}), personal lines route 100% to their owner, own accounts log straight into each owner's P&L, and the bad beat contribution is read from the export and split by each owner's JP %. Tags, deals, and makeup balances persist week to week.
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
            <Btn onClick={() => fileRef.current?.click()}>Choose file</Btn>
            <Btn tone="ghost" onClick={() => setTab("setup")}>Owners & setup</Btn>
          </div>
          {tab === "setup" && <div style={{ textAlign: "left", marginTop: 20 }}><OwnerClubSetup club={club} clubs={clubs} saveClubs={saveClubs} onDeleteClub={onDeleteClub} /></div>}
        </div>
      )}

      {players && model && (
        <>
          <div style={{ display: "flex", gap: 4, marginBottom: 16, borderBottom: `2px solid ${C.line}`, flexWrap: "wrap" }}>
            {[["settle", "Lines & ownership"], ["dlumbrellas", "DL Umbrellas"], ["backed", "Stake & action"], ["reports", "Reports"], ["owners", "Settlements"], ["setup", "Owners & setup"]].map(([k, label]) => (
              <button key={k} onClick={() => setTab(k)} style={{
                border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
                background: tab === k ? C.card : "transparent", color: tab === k ? C.ink : C.mute,
                borderRadius: "8px 8px 0 0", marginBottom: -2, boxShadow: tab === k ? "0 -1px 4px rgba(0,0,0,0.1)" : "none" }}>
                {label}
                {k === "settle" && pendingCount > 0 && (
                  <span style={{ marginLeft: 6, background: C.red, color: "#fff", borderRadius: 9, padding: "1px 7px", fontSize: 10.5 }}>{pendingCount}</span>
                )}
              </button>
            ))}
            <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", paddingBottom: 6 }}>
              <span style={{ color: C.mute, fontSize: 12.5 }}>{period || "this week"}</span>
              <Btn tone="ghost" small onClick={async () => { await archiveWeek(club.id, club.name, period, async () => { await downloadAAWorkbook(model, period, club); for (const o of ownerIds) await downloadAAOwnerExcel(model, o, period, club); }); setSaveNote(`Archived ${period} — see the Archive tab.`); }}>Archive this week</Btn>
              <Btn tone="gold" small onClick={() => fileRef.current?.click()}>Upload weekly export</Btn>
            </div>
          </div>

          {pendingCount > 0 && (
            <div style={{ background: C.banner, border: `1px solid ${C.gold}`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, fontSize: 13.5 }}>
              <b>{pendingCount} item{pendingCount > 1 ? "s" : ""} need review</b> — {model.untagged.length > 0 && `${model.untagged.length} line${model.untagged.length > 1 ? "s" : ""} untagged (nothing lands on anyone's book until each is classed)`}{model.untagged.length > 0 && model.uncollected.length > 0 && "; "}{model.uncollected.length > 0 && `${model.uncollected.length} agent line${model.uncollected.length > 1 ? "s" : ""} missing a collector (personal and staked lines auto-assign)`}. The settle-up stays provisional until both are done.
            </div>
          )}

          {tab === "settle" && (
            <>
              <div style={{ display: "flex", alignItems: "baseline", marginBottom: 10 }}>
                <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>Every line: class, margin routing, and who collects it</div>
                <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
                  <span style={{ fontSize: 12, color: C.mute }}>Default TR %</span>
                  <PctInput value={cfg.defaultTR ?? 0} onChange={(v) => v != null && up({ defaultTR: v })} width={50} />
                  <Btn tone="gold" small onClick={() => {
                    if (!model.ready) { window.alert(`Finish review first — ${model.untagged.length} untagged line(s), ${model.uncollected.length} without a collector. The workbook includes the final settle-up, so everything needs classing.`); return; }
                    downloadAAWorkbook(model, period, club);
                  }}>Download Excel workbook</Btn>
                  <Btn tone="ghost" small onClick={exportCsv}>Copy table</Btn>
                </div>
              </div>
              <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>
                <b>Margin</b> = tips − rakeback = the profit on the line, plus any TR cut. The rakeback itself is always paid out to the agent or player at their TB%. <b>TR</b> is an extra cut of that line's net (P&L + rakeback) on top of the rake margin — e.g. a "70/10" deal (70% TB, 10% TR) — off by default; set "Default TR %" above for a club where every line runs on one (skips needing a separate stake per player), or override it per line. <b>Agent</b> margins feed the owner pool ({H.poolLabel}); <b>personal</b> margins go 100% to that owner (his player, his spread), and he collects the line himself. Owners' own accounts are handled below; staked/action players sit on the Stake & action tab. Every player rolls up onto their super agent or, lacking one, their agent — "player overrides" on a line lets you pin one player (or a whole agent) to their own rate, e.g. a VIP deal that shouldn't follow their agent's rate. <b>Managers</b> and <b>Masters</b> sit outside that hierarchy and show up as their own lines.
              </div>

              <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: C.cream }}>
                    <th style={{ ...th, textAlign: "left" }}>Line</th>
                    <th style={{ ...th, textAlign: "left" }}>Class</th>
                    <th style={{ ...th, textAlign: "left" }}>Collected by</th>
                    <th style={th}>TB %</th>
                    <th style={th}>TR %</th>
                    <th style={th}>Winnings</th><th style={th}>Tips</th>
                    <th style={th}>Rakeback paid</th>
                    <th style={{ ...th }}>Margin → to</th><th style={th}>Union cash</th>
                  </tr></thead>
                  <tbody>
                    {model.entities.map((e, i) => {
                      const treeable = e.type === "dlUmbrella" || ((e.type === "sa" || e.type === "agent") && e.members.length > 1);
                      const open = !!showMembers[e.key];
                      return (
                      <React.Fragment key={e.key}>
                      <tr style={{ background: !e.tag ? "var(--errBg)" : i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                        <td style={tdL}>
                          <span style={{ fontWeight: 600 }}>{e.name}</span> {typePill(e)}
                          {treeable && (
                            <button onClick={() => setShowMembers({ ...showMembers, [e.key]: !open })}
                              style={{ marginLeft: 8, border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>
                              {open ? "hide players" : "player overrides"}
                            </button>
                          )}
                        </td>
                        <td style={{ ...tdL, whiteSpace: "nowrap" }}>
                          <select value={e.tag || ""} onChange={(ev) => setTag(e.key, ev.target.value)} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
                            {!e.tag && <option value="">— choose class —</option>}
                            {H.tags.map(([v, label]) => <option key={v} value={v}>{label}</option>)}
                          </select>
                        </td>
                        <td style={{ ...tdL, whiteSpace: "nowrap" }}>
                          {e.tag === "agent" ? collectorSel(e) : e.tag ? <Pill tone="green">{lbl(e.tag)} · auto</Pill> : <span style={{ color: C.mute }}>—</span>}
                        </td>
                        <td style={td}><PctInput value={cfg.deals[e.id] ?? cfg.defaultTB} onChange={(v) => setDeal(e.id, v)} width={54} max={999} /></td>
                        <td style={td}><PctInput value={cfg.tr?.[e.id] ?? cfg.defaultTR ?? 0} onChange={(v) => v != null && setTr(e.id, v)} width={54} /></td>
                        <td style={td}>{fmt(e.pnl)}</td>
                        <td style={td}>{fmt(e.fee)}</td>
                        <td style={td}>{fmt(e.tipback)}</td>
                        <td style={{ ...td, whiteSpace: "nowrap" }}>
                          {money(e.margin)}{" "}
                          <span style={{ fontSize: 11, color: e.tag && e.tag !== "agent" ? C.goldDark : C.mute, fontWeight: e.tag && e.tag !== "agent" ? 700 : 400 }}>
                            {e.tag === "agent" ? "→ pool" : e.tag ? `→ ${lbl(e.tag)} 100%` : ""}
                          </span>
                        </td>
                        <td style={td}>{e.tag ? money(e.unionCash) : <span style={{ color: C.mute }}>—</span>}</td>
                      </tr>
                      {treeable && open && (
                        <tr>
                          <td colSpan={10} style={{ padding: 0, background: "var(--paper)" }}>
                            <AALineOverrides e={e} cfg={cfg} setDeal={setDeal} setOwnDeal={setOwnDeal} setPlayerTr={setPlayerTr} setAgentTr={setAgentTr} />
                          </td>
                        </tr>
                      )}
                      </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
                <div style={{ background: C.bar, color: "var(--barText)", display: "flex", padding: "12px 16px", fontSize: 13.5, gap: 24, flexWrap: "wrap" }}>
                  <span>Rake <b style={{ color: "var(--barGold)" }}>{fmt(model.clubRevenue)}</b></span>
                  <span>RB paid <b style={{ color: "var(--barGold)" }}>{fmt(model.agentRB + model.personalRB + model.backedRB)}</b></span>
                  <span>Pool <b style={{ color: model.pool >= 0 ? "var(--barGreen)" : "var(--barRed)" }}>{fmt(model.pool)}</b> <span style={{ color: "var(--barSubtle)" }}>→ {H.poolOwners.map((o) => `${o.label} ${fmt(model.poolShares[o.id])}`).join(" · ")}</span></span>
                  <span style={{ marginLeft: "auto" }}>Personal margins — {ownerIds.map((o, i) => <span key={o}>{i ? " · " : ""}{lbl(o)} <b style={{ color: "var(--barGold)" }}>{fmt(model.personalMargin[o])}</b></span>)}</span>
                </div>
              </div>
              <div style={{ color: C.mute, fontSize: 12, marginTop: 8 }}>
                A line's TB% applies to the whole tree. Assign-all collectors: {" "}
                {ownerIds.map((o) => (
                  <button key={o} onClick={() => { const c = { ...cfg.collectors }; model.uncollected.forEach((e) => (c[e.key] = o)); up({ collectors: c }); }}
                    style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontWeight: 700, fontSize: 12, textDecoration: "underline", padding: "0 4px" }}>
                    rest → {lbl(o)}
                  </button>
                ))}
              </div>

              <AAOwnAccounts model={model} cfg={cfg} up={up} />
            </>
          )}

          {tab === "dlumbrellas" && <AADLUmbrellasTab model={model} cfg={cfg} up={up} />}

          {tab === "backed" && <AABackedTab model={model} cfg={cfg} up={up} period={period} />}

          {tab === "reports" && <AAReportsTab model={model} cfg={cfg} period={period} expanded={repExpanded} setExpanded={setRepExpanded} club={club} />}
          {tab === "setup" && (
            <>
              <OwnerClubSetup club={club} clubs={clubs} saveClubs={saveClubs} onDeleteClub={onDeleteClub} />
              <div style={{ height: 14 }} />
              <AABBJHoldsCard model={model} cfg={cfg} up={up} period={period} />
            </>
          )}

          {tab === "owners" && (
            <>
              <div style={{ background: C.bar, borderRadius: 12, padding: "26px 30px", marginBottom: 18, textAlign: "center", color: "var(--barText)" }}>
                <div style={{ fontSize: 12, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--barGold)", marginBottom: 6 }}>Profit per owner — {club.name} · {period || "this week"}</div>
                <div style={{ fontFamily: "Georgia, serif", fontSize: 26, display: "flex", justifyContent: "center", gap: 40, flexWrap: "wrap" }}>
                  {ownerIds.map((o) => (
                    <span key={o}>{lbl(o)} <span style={{ color: model.profit[o] >= 0 ? "var(--barGreen)" : "var(--barRed)" }}>{fmt(model.profit[o])}</span></span>
                  ))}
                </div>
                <div style={{ borderTop: "1px solid rgba(255,255,255,0.15)", margin: "16px auto 0", paddingTop: 14, maxWidth: 680 }}>
                  {!model.ready ? (
                    <div style={{ fontSize: 15 }}>Settle-up pending — finish tagging and collector assignment on the Lines tab.</div>
                  ) : model.transfers.length === 0 ? (
                    <div style={{ fontFamily: "Georgia, serif", fontSize: 20 }}>Perfectly even — no transfer needed</div>
                  ) : (
                    model.transfers.map((t, i) => (
                      <div key={i} style={{ fontFamily: "Georgia, serif", fontSize: 20, marginTop: i ? 4 : 0 }}>
                        {lbl(t.from)} pays {lbl(t.to)} <span style={{ color: "var(--barGold)" }}>{fmt(t.amount)}</span>
                      </div>
                    ))
                  )}
                  <div style={{ fontSize: 12, color: "var(--barSubtle)", marginTop: 8 }}>
                    Transfers compare what each owner collected against his collectable share. {model.jpInCollections ? "The jackpot shares settle through the transfer (contribution came out of players' P&L, so the collectors hold the cash). " : "Excluded on purpose: the jackpot shares (that cash sits in the jackpot pool, held by no one). "}Stake/action books are always excluded (settled separately). Own-account P&L IS in the transfer — an owner's winnings get paid out of the week's collections.
                    {model.ready && (model.balanceOk ? " Balance check: ✓ books tie out." : ` ⚠ Books off by ${fmt(model.imbalance)} — the export's P&L doesn't net to its rake (promos, uncollected pots, or jackpot drop inside P&L).`)}
                  </div>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(310px, 1fr))", gap: 14, marginBottom: 18 }}>
                <Card title="Bad beat jackpot" right={<Pill tone="gold">by JP % · profit, no holder</Pill>}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8, flexWrap: "wrap" }}>
                    <span style={{ fontSize: 13.5, color: C.mute }}>Contribution this week</span>
                    <NumInput value={jackpot} onChange={setJackpot} width={100} />
                    {jpFromExport != null && Math.abs(jpFromExport - jackpot) < 0.01 && <Pill tone="green">read from export</Pill>}
                    {jpFromExport != null && Math.abs(jpFromExport - jackpot) >= 0.01 && <Pill tone="blue">export said {fmt(jpFromExport)}</Pill>}
                    {jpFromExport == null && <Pill>not found in export — enter manually</Pill>}
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 10, fontSize: 12.5 }}>
                    {[["pool", "Cash sits in the jackpot pool (held by no one) — thirds excluded from the transfer"],
                      ["collections", "Contribution was deducted from players' P&L — the collectors are holding it, so the shares settle through the transfer"]].map(([v, label]) => (
                      <label key={v} style={{ display: "flex", gap: 8, alignItems: "flex-start", cursor: "pointer", color: model.jpInCollections === (v === "collections") ? C.ink : C.mute }}>
                        <input type="radio" name="aa-jp-mode" checked={model.jpInCollections === (v === "collections")} onChange={() => setJpMode(v)} style={{ marginTop: 2 }} />
                        <span>{label}</span>
                      </label>
                    ))}
                  </div>
                  {!model.jpInCollections && Math.abs(model.imbalance - model.jackpot) < 0.02 && Math.abs(model.jackpot) > 0.005 && model.ready && (
                    <div style={{ background: C.banner, color: C.goldDark, borderRadius: 6, padding: "8px 10px", fontSize: 12, marginBottom: 8 }}>
                      Books are off by exactly the jackpot ({fmt(model.jackpot)}) — that's the signature of the contribution being deducted from players' P&L. Switching to the second option will tie the books and route each owner his share through the transfer.
                    </div>
                  )}
                  {ownerIds.map((o, i) => {
                    const heldAway = model.jpHoldMoves.find((m) => m.forOwnerId === o);
                    const heldFor = model.jpHoldMoves.filter((m) => m.holderId === o);
                    const isLastRow = i === ownerIds.length - 1 && heldFor.length === 0;
                    return (
                      <React.Fragment key={o}>
                        {ownerRow(`${lbl(o)} share (${H.pctS(H.jpShareOf(o))})${heldAway ? ` — held by ${lbl(heldAway.holderId)}` : ""}`, model.jpSharesRaw[o], { rule: isLastRow })}
                        {heldFor.map((m) => ownerRow(`+ holding ${lbl(m.forOwnerId)}'s BBJ this week`, m.amount, { rule: i === ownerIds.length - 1 && m === heldFor[heldFor.length - 1] }))}
                      </React.Fragment>
                    );
                  })}
                  {ownerRow("Total", model.jackpot, { bold: true })}
                  {model.jpHoldMoves.length > 0 && (
                    <div style={{ fontSize: 12, color: C.mute, marginTop: 8 }}>
                      BBJ holds in effect this week — set up or edit these on Owners & setup. The held owner's other settlements (personal lines, own accounts) are unaffected; only their jackpot line moves.
                    </div>
                  )}
                </Card>

                <Card title="Union economics">
                  {ownerRow("Rake collected", model.clubRevenue)}
                  {ownerRow("Rakeback → agents (pool lines)", -model.agentRB)}
                  {ownerRow("Rakeback → players on personal lines", -model.personalRB)}
                  {ownerRow("RB credits → stake/action deals", -model.backedRB)}
                  {ownerRow("Feeback → owners' own accounts", -model.ownFeeback, { rule: true })}
                  {ownerRow(`Pool (agent lines only, ${H.poolLabel})`, model.pool, { bold: true })}
                  {ownerRow("Personal margins (owner-routed)", model.totalPersonalMargin, { bold: true })}
                  {ownerRow("Stake/action rake margins → backer", model.totalDealMargin, { bold: true })}
                  <div style={{ fontSize: 12, color: C.mute, marginTop: 8 }}>
                    The pool is agent-line margins only. Personal-line margins and the rake margin on staked/action players skip the pool — they land straight on their owner's / backer's card.
                  </div>
                </Card>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(310px, 1fr))", gap: 14 }}>
                {ownerIds.map((o) => (
                  <Card key={o} title={`${lbl(o)} — profit`} right={H.poolShare(o) <= 0 ? <Pill>margin + own + JP</Pill> : <Pill>pool + margin + own + JP</Pill>}>
                    {H.poolShare(o) > 0 && ownerRow(`Pool share (${H.pctS(H.poolShare(o))})`, model.poolShares[o])}
                    {ownerRow("Personal-line margin · 100%", model.personalMargin[o])}
                    {ownerRow("Stake/action rake margin · his deals", model.dealMargin[o])}
                    {ownerRow("Own accounts P&L", model.ownPosition[o])}
                    {(() => {
                      const heldAway = model.jpHoldMoves.find((m) => m.forOwnerId === o);
                      const heldFor = model.jpHoldMoves.filter((m) => m.holderId === o);
                      const label = heldAway ? `Jackpot share — held by ${lbl(heldAway.holderId)}` : heldFor.length ? `Jackpot share (${H.pctS(H.jpShareOf(o))}) + holding ${heldFor.map((m) => lbl(m.forOwnerId)).join(", ")}'s BBJ` : `Jackpot share (${H.pctS(H.jpShareOf(o))})`;
                      return ownerRow(label, model.jpShares[o], { rule: true });
                    })()}
                    {ownerRow(`Profit — ${club.name}`, model.profit[o], { bold: true })}
                    <div style={{ marginTop: 10 }}>
                      {ownerRow("Collectable share", model.cashProfit[o])}
                      {ownerRow("Collected", model.actual[o])}
                      {ownerRow(model.delta[o] > 0.005 ? "Over-collected → pays out" : model.delta[o] < -0.005 ? "Under-collected → receives" : "Even", Math.abs(model.delta[o]) < 0.005 ? "—" : Math.abs(model.delta[o]), { bold: true, rule: true })}
                    </div>
                    {Math.abs(model.backedBook[o]) > 0.005 && (
                      <div style={{ fontSize: 12, color: C.mute, marginTop: 8 }}>
                        + {money(model.backedBook[o])} on stake/action books — settled separately, not in the transfer above.
                      </div>
                    )}
                    <div style={{ marginTop: 12 }}>
                      <Btn tone="ghost" small onClick={() => downloadAAOwnerExcel(model, o, period, club)}>Download {lbl(o)} report</Btn>
                    </div>
                  </Card>
                ))}
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}

const TABS_KEY = "tabs-v1";
const TABS_EMPTY = { counterparties: [], entries: [], pushed: {}, stakingPushed: {}, unifiedPushed: {}, persons: [], vigRate: 5, staking: { deals: [], results: [], imported: [] }, misc: [], settleChecklist: [] };
const today = () => new Date().toISOString().slice(0, 10);
const METHOD_SUGGESTIONS = ["zelle", "venmo", "cash app", "paypal", "crypto", "cash", "wire", "apple pay"];
const isCrypto = (m) => String(m || "").trim().toLowerCase() === "crypto";
const r2 = (x) => Math.round((+x + Number.EPSILON) * 100) / 100;

// Crypto vig: vig amount = |transaction| × rate, always. That's logged in the
// Vig section with a sign based on direction — positive when Ak is sending,
// negative when he's receiving. The post-vig transaction amount is always
// |transaction| + vig amount (vig always adds to the magnitude, whichever
// direction), then the original sign of the entered amount is reapplied.
function cryptoEffect(base, rate) {
  const b = +base || 0, r = (+rate || 0) / 100;
  const vigAmount = Math.abs(b) * r;
  const vig = b >= 0 ? vigAmount : -vigAmount;
  const postVig = Math.abs(b) + vigAmount;
  const amount = b >= 0 ? postVig : -postVig;
  return { amount: r2(amount), vig: r2(vig) };
}

async function loadTabs() {
  let raw = null;
  try { const c = await store.get(TABS_KEY); if (c?.value) raw = JSON.parse(c.value); } catch (e) {}
  const d = { ...TABS_EMPTY, ...(raw || {}) };
  d.staking = { ...TABS_EMPTY.staking, ...(d.staking || {}) };
  d.stakingPushed = d.stakingPushed || {};
  d.unifiedPushed = d.unifiedPushed || {};
  d.misc = d.misc || [];
  d.settleChecklist = d.settleChecklist || [];
  let changed = false;
  // First run of the new Player data tab: pull bundles over from My Clubs.
  if (!Array.isArray(raw?.persons)) {
    try { const c = await store.get("agentclubs-v3"); if (c?.value) { const a = JSON.parse(c.value); d.persons = normPersons(a.personAliases || []); changed = true; } } catch (e) {}
  }
  d.persons = normPersons(d.persons);
  // Old "count as P&L" entries → Misc P&L (the tab entry itself stays put).
  if (!d.miscMigrated) {
    d.entries.filter((e) => e.pnl).forEach((e) => d.misc.push({ id: "mig-" + e.id, date: e.date, amount: e.amount, note: e.note || "", category: e.category || "uncategorized" }));
    d.entries = d.entries.map((e) => (e.pnl ? { ...e, pnl: undefined, category: undefined } : e));
    d.miscMigrated = true; changed = true;
  }
  if (changed) { try { await store.set(TABS_KEY, JSON.stringify(d)); } catch (e) {} }
  return d;
}
async function saveTabs(data) { try { await store.set(TABS_KEY, JSON.stringify(data)); } catch (e) {} }

function applyWeekToTabsData(data, sourceKey, weekLabel, items) {
  if (data.pushed[sourceKey]) return data;
  const byKey = {}; data.counterparties.forEach((cp) => (byKey[cp.kind + "|" + cp.name.toLowerCase()] = cp));
  const d = today();
  const checklist = data.settleChecklist || (data.settleChecklist = []);
  const onList = new Set(checklist.map((x) => x.cpId));
  items.filter((it) => Math.abs(it.amount) > 0.005).forEach((it) => {
    const key = it.kind + "|" + it.name.toLowerCase();
    let cp = byKey[key];
    if (!cp) { cp = { id: uid(), name: it.name, kind: it.kind }; byKey[key] = cp; data.counterparties.push(cp); }
    data.entries.push({ id: uid(), date: d, cpId: cp.id, amount: r2(it.amount), note: `${it.note} · ${weekLabel}`, source: "week", week: weekLabel, sourceKey });
    // Club-settlement-driven change (weekly fold-in, not an ordinary mid-week
    // transaction) — add to the settle checklist so it doesn't get missed.
    // Only lines Ak himself is on the hook to settle: Fish Tank umbrellas
    // assigned to Ak, AA SA/Agent/individual lines that are Ak's personally
    // or a split line collected by Ak, and every My Clubs umbrella/player.
    if (it.kind === "player" && it.checklist && !onList.has(cp.id)) { checklist.push({ id: uid(), cpId: cp.id, name: cp.name, addedAt: d, week: weekLabel, done: false, sourceKey }); onList.add(cp.id); }
  });
  data.pushed[sourceKey] = true;
  return data;
}

// Fold a source's house-backed results (your share only) into the Staking log — once per source.
function applyStakingImport(data, sourceKey, items) {
  if (data.stakingPushed[sourceKey]) return data;
  const d = today();
  (items || []).forEach((it) => data.staking.imported.push({ id: uid(), sourceKey, date: d, ...it }));
  data.stakingPushed[sourceKey] = true;
  return data;
}

// Auto-feed a house-backed makeup player's weekly net (P&L + RB credit) into a
// shared staking deal — used when a Fish Tank / owner-club backed player is
// linked (via "Unified deal" on their backed row) to a manual weekly makeup
// deal, so results from multiple sites (and manual/external games) all land
// on ONE combined makeup pool with weekly cadence grouping the sum, instead
// of each site tracking its own separate makeup balance. Dedup is per-deal
// AND per-game (by scanning already-posted results for this sourceKey +
// dealId + game), not just per-deal — the same shared deal can be fed by
// MULTIPLE feeds under one sourceKey at once (e.g. one player linked to it
// from two different My Clubs sub-clubs in the same week); keying dedup on
// dealId alone would treat that second feed as "already posted" the moment
// the first one lands, silently dropping it. Not a single once-per-source
// flag either — a source can be accepted before any player is linked, then
// have a player linked to a unified deal afterward, and the backfill below
// needs to still be able to post THAT deal's feed later without re-posting
// ones already applied. A since-deleted deal id is skipped quietly rather
// than erroring.
function applyUnifiedStakingResults(data, sourceKey, feeds) {
  const feedKey = (dealId, game) => dealId + "|" + game;
  const already = new Set(data.staking.results.filter((r) => r.sourceKey === sourceKey).map((r) => feedKey(r.dealId, r.game)));
  const byKey = {}; data.counterparties.forEach((cp) => (byKey[cp.kind + "|" + cp.name.toLowerCase()] = cp));
  (feeds || []).forEach((f) => {
    if (!data.staking.deals.some((dl) => dl.id === f.dealId)) return;
    if (!already.has(feedKey(f.dealId, f.game))) data.staking.results.push({ id: uid(), dealId: f.dealId, date: f.date, game: f.game, pnl: f.pnl, holder: f.holder, note: f.note, sourceKey });
    // This player's weekly net now lives on the shared staking deal instead of
    // this site's own local settlement — if this source was already accepted
    // into Tabs before the link existed (or before it reliably persisted —
    // see the normalizeAgent fix above), a normal ledger line for this exact
    // player/role is likely already sitting there double-counting the same
    // result. Retire it now that the unified feed is taking over. Runs even
    // when the feed itself was already posted in an earlier pass (before this
    // cleanup existed), so it still catches a stale line left behind by that.
    // Matched on the note's un-dated prefix, not just cp+sourceKey, so a
    // person's OTHER non-staked line for the same week/source (e.g. AA's
    // separate "(personal · ...)" row) is left alone.
    if (f.cp && f.noteMatch) {
      const cp = byKey[f.cp.kind + "|" + f.cp.name.toLowerCase()];
      if (cp) data.entries = data.entries.filter((e) => !(e.sourceKey === sourceKey && e.cpId === cp.id && (e.note || "").startsWith(f.noteMatch)));
    }
  });
  data.unifiedPushed[sourceKey] = true;
  return data;
}

// Undo everything a source previously posted — its ledger entries, staking
// imports/results, and any not-yet-done checklist line it added (a checked-
// off line is left alone; someone already acted on it) — and clear its
// pushed flags so it can be re-applied cleanly. Used when a Fish Tank/AA/My
// Clubs week gets unlocked, corrected, and relocked after it was already
// accepted into Tabs: re-syncing removes the stale amount first, so the
// corrected one replaces it instead of stacking on top or netting oddly.
function clearSource(data, sourceKey) {
  return {
    ...data,
    entries: data.entries.filter((e) => e.sourceKey !== sourceKey),
    staking: { ...data.staking,
      imported: data.staking.imported.filter((it) => it.sourceKey !== sourceKey),
      results: data.staking.results.filter((r) => r.sourceKey !== sourceKey) },
    settleChecklist: (data.settleChecklist || []).filter((x) => x.sourceKey !== sourceKey || x.done),
    pushed: { ...data.pushed, [sourceKey]: false },
    stakingPushed: { ...data.stakingPushed, [sourceKey]: false },
    unifiedPushed: { ...data.unifiedPushed, [sourceKey]: false },
  };
}

// Your share of a house-backed line's book this week: action → net; makeup → what was chopped to you.
function backedShareRows(entities, shareFn, site, mapName) {
  const out = [];
  entities.forEach((e) => {
    const sh = shareFn(e.backer);
    if (!(sh > 0)) return;
    if (e.dealType === "action") {
      out.push({ name: mapName(e.name, site), rawName: e.name, site, kind: "action", net: r2(e.backerBook * sh), chopped: 0, share: sh, dealKey: e.key });
    } else {
      const recovered = e.makeupBefore - e.makeupAfter; // >0 = makeup shrank, <0 = grew
      const chopped = e.backerBook - recovered;           // the excess that was chopped to the backer
      out.push({ name: mapName(e.name, site), rawName: e.name, site, kind: "makeup", net: r2(e.backerBook * sh), chopped: r2(chopped * sh), makeupBefore: r2(e.makeupBefore * sh), makeupAfter: r2(e.makeupAfter * sh), share: sh, dealKey: e.key });
    }
  });
  return out;
}

// Every username seen on the last imported week for Fish Tank + each owner
// club, plus every player across every My Clubs sub-club — the full roster
// for the Player Data chip picker, not just names that already have tab entries.
async function loadAllSiteUsernames(clubs) {
  const out = [];
  // `hint` is a cosmetic label only (e.g. "club owner") — the alias's actual
  // `site` must match the literal site string resolve()/mapName() use at
  // fold-in time, or the alias silently stops matching future weeks.
  const push = (name, site, hint) => { if (name && String(name).trim()) out.push({ name: String(name).trim(), site, hint }); };
  // Fish Tank — top-of-upline names only: SA/umbrella groups fold their downline into
  // one settled entity already (buildModel), same as what Settlements shows.
  try { const ft = await loadFishTankModel(); if (ft) { ft.model.entities.forEach((e) => push(e.name, "Fish Tank")); ft.model.backedEntities.forEach((e) => push(e.name, "Fish Tank")); } } catch (e) {}
  // All American / other owner clubs — same idea: players under an agent or super
  // agent (or folded into a DL umbrella) are represented by that upline's entity only.
  for (const club of clubs || []) {
    try { const aa = await loadAAModel(club); if (aa) { aa.model.entities.forEach((e) => push(e.name, club.name)); aa.model.backedEntities.forEach((e) => push(e.name, club.name)); } } catch (e) {}
  }
  // My Clubs — players bundled into a My Clubs umbrella are represented by the
  // umbrella's name only; everyone else (not in any umbrella) shows individually.
  try {
    const c = await store.get("agentclubs-v3");
    if (c?.value) {
      const acfg = normalizeAgent({ ...AGENT_DEFAULT, ...JSON.parse(c.value) });
      const weekKeys = Object.keys(acfg.weeks || {});
      const wk = acfg.currentWeek && acfg.weeks[acfg.currentWeek] ? acfg.currentWeek : weekKeys[weekKeys.length - 1];
      if (wk) {
        const m = computeAgent(acfg, acfg.weeks[wk]);
        m.umbrellas.forEach((u) => push(u.name, "My Clubs"));
        m.allPlayers.filter((p) => !m.inUmbrella.has(p.id)).forEach((p) => push(p.name, "My Clubs"));
      } else {
        // No week entered yet for this site — fall back to the raw roster so a
        // brand-new club still shows up in the picker.
        (acfg.clubs || []).forEach((cl) => (cl.players || []).forEach((p) => push(p.name, "My Clubs")));
      }
      // Club owners — whoever's behind a My Clubs sub-club settles as their own
      // "username" too, which is what lets an owner who also plays get bundled
      // with their player alias(es) into one tab (Player data → Tab shows as: Club).
      // Site stays "My Clubs" (same as players) so the alias still matches at
      // fold-in time; "club owner" is just a cosmetic hint for the picker.
      (acfg.clubs || []).forEach((cl) => push(cl.owner, "My Clubs", "club owner"));
    }
  } catch (e) {}
  const seen = new Set();
  return out.filter((x) => { const k = x.site + "|" + x.name.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; }).sort((a, b) => a.name.localeCompare(b.name));
}

// Build the week's tab items from Fish Tank + owner clubs + My Clubs, with
// usernames folded into persons and clubs folded into their owners.
async function collectPendingWeeks(persons, clubs) {
  const pending = [];
  let acfg = null;
  try { const c = await store.get("agentclubs-v3"); if (c?.value) acfg = normalizeAgent({ ...AGENT_DEFAULT, ...JSON.parse(c.value) }); } catch (e) {}
  const mapName = makeNameMapper(persons);
  // Someone bundled in Player Data can have a fixed "tab shows as" kind (e.g. a
  // My Clubs club owner who also plays somewhere) — once their alias maps a raw
  // name to that person, route the item to their chosen kind instead of whatever
  // kind the source naturally produced, so player-side and club-side settlements
  // land on the same one counterparty every week, not two permanently-separate ones.
  const personKindOf = {};
  normPersons(persons).forEach((p) => { personKindOf[p.name.trim().toLowerCase()] = p.kind || "player"; });
  const resolve = (raw, site, naturalKind) => {
    const name = mapName(raw, site);
    return { name, kind: personKindOf[name.trim().toLowerCase()] || naturalKind };
  };
  // A backed makeup player can be linked (per-site, via "Unified deal" on
  // their backed row) to one shared manual staking deal — when linked, their
  // weekly net (P&L + RB credit) is auto-fed as a synthetic result on that
  // deal instead of being tracked on this site's own local makeup balance,
  // so Fish Tank + owner-club (+ manual/external) results all combine into
  // one ongoing pool. Dated to today (when the report is accepted/finalized
  // and this net actually becomes known) so it sorts correctly alongside any
  // mid-week manual/external results on a "settle per session" deal — each
  // event updates the running makeup as it comes in, rather than everything
  // in a calendar week being batched into one lump settle.
  const unifiedIdOf = (backedCfg, key) => (backedCfg?.[key.slice(2)]?.unifiedDealId || "").trim();
  const live = []; // current makeup balances on imported makeup deals (your share)
  const ft = await loadFishTankModel();
  if (ft && ft.period) {
    const ftAssignments = ft.cfg?.assignments || {};
    // Fish Tank's own `entities` list (unlike AA's) already merges backed
    // players in — exclude backed makeup players linked to a unified deal so
    // their local (frozen) settlement doesn't ALSO post here on top of the
    // unifiedFeeds entry below.
    const items = ft.model.entities.filter((e) => !(e.type === "backed" && e.dealType === "makeup" && unifiedIdOf(ft.cfg?.backed, e.key))).map((e) => ({ ...resolve(e.name, "Fish Tank", "player"), amount: -e.settlement, note: `Fish Tank · ${e.name}`, checklist: ftAssignments[e.key] === "ak" }));
    const ftShare = (backer) => (backer === "ak" ? 1 : backer === "split" ? 0.5 : 0);
    const stakingItems = backedShareRows(ft.model.backedEntities, ftShare, "Fish Tank", mapName);
    const unifiedFeeds = [];
    ft.model.backedEntities.forEach((e) => {
      if (e.dealType !== "makeup") return;
      const dealId = unifiedIdOf(ft.cfg?.backed, e.key);
      if (!dealId) return;
      unifiedFeeds.push({ dealId, date: today(), game: "Fish Tank", pnl: r2(e.net), holder: "ak", note: `auto · Fish Tank · ${ft.period}`, cp: resolve(e.name, "Fish Tank", "player"), noteMatch: `Fish Tank · ${e.name}` });
    });
    pending.push({ sourceKey: `ft:${ft.period}`, label: `Fish Tank · ${ft.period}`, items, stakingItems, unifiedFeeds });
    Object.entries(ft.cfg?.backed || {}).forEach(([k, b]) => { const sh = ftShare(b.backer); if (b.deal !== "action" && sh > 0) live.push({ name: mapName(b.name, "Fish Tank"), rawName: b.name, site: "Fish Tank", makeup: r2((b.makeup || 0) * sh), share: sh, rb: `RB ${b.rbNormal}% / ${b.rbMakeup}% · player ${b.playerProfitPct ?? 50}%` }); });
  }
  for (const club of clubs || []) {
    const aa = await loadAAModel(club);
    if (!aa || !aa.period) continue;
    const H = ocHelpers(club);
    const site = club.name;
    // Same convention as Fish Tank: amount is what the counterparty owes you.
    // "Settled by Ak" — SA/Agent/individual/Manager/Master/DL-umbrella lines
    // tagged personally to Ak or split-and-collected-by-Ak, and house-backed/
    // staked accounts where Ak holds any share of the backing (own or split).
    const aaChecklistEligible = (e) => e.tag === H.meId || (e.tag === "agent" && e.collector === H.meId);
    // Backed makeup players linked to a unified deal are excluded here — their
    // weekly net posts through that deal instead (see unifiedFeeds below), so
    // it isn't double-counted against this site's own (now-vestigial) local makeup.
    const items = [
      ...aa.model.entities.map((e) => ({ ...resolve(e.name, site, "player"), amount: -e.settlement, note: `${site} · ${e.name}${e.tag && e.tag !== "agent" ? ` (personal · ${H.lbl(e.tag)})` : ""}`, checklist: aaChecklistEligible(e) })),
      ...aa.model.backedEntities.filter((e) => !unifiedIdOf(aa.cfg?.backed, e.key)).map((e) => ({ ...resolve(e.name, site, "player"), amount: -e.settlement, note: `${site} · ${e.name} (${e.dealType === "action" ? "action buy" : "stake"})`, checklist: H.meId ? H.shareOf(e.backer, H.meId) > 0 : false })),
    ];
    const shareFn = (backer) => (H.meId ? H.shareOf(backer, H.meId) : 0);
    const stakingItems = backedShareRows(aa.model.backedEntities, shareFn, site, mapName);
    const unifiedFeeds = [];
    aa.model.backedEntities.forEach((e) => {
      if (e.dealType !== "makeup") return;
      const dealId = unifiedIdOf(aa.cfg?.backed, e.key);
      if (!dealId) return;
      unifiedFeeds.push({ dealId, date: today(), game: site, pnl: r2(e.net), holder: "ak", note: `auto · ${site} · ${aa.period}`, cp: resolve(e.name, site, "player"), noteMatch: `${site} · ${e.name} (stake)` });
    });
    pending.push({ sourceKey: `aa:${club.id === "allamerican" ? "" : club.id + ":"}${aa.period}`, label: `${site} · ${aa.period}`, items, stakingItems, unifiedFeeds });
    Object.entries(aa.cfg?.backed || {}).forEach(([k, b]) => { const sh = shareFn(b.backer); if (b.deal !== "action" && sh > 0) live.push({ name: mapName(b.name, site), rawName: b.name, site, makeup: r2((b.makeup || 0) * sh), share: sh, rb: `RB ${b.rbNormal}% / ${b.rbMakeup}% · player ${b.playerProfitPct ?? 50}%` }); });
  }
  if (acfg) {
    const weekKeys = Object.keys(acfg.weeks || {});
    const wk = acfg.currentWeek && acfg.weeks[acfg.currentWeek] ? acfg.currentWeek : weekKeys[weekKeys.length - 1];
    if (wk) {
      const m = computeAgent(acfg, acfg.weeks[wk]);
      const items = [];
      m.umbrellas.filter((u) => u.played.length > 0).forEach((u) => items.push({ ...resolve(u.name, "My Clubs", "player"), amount: -u.settlement, note: `My Clubs umbrella`, checklist: true }));
      // A player also staked elsewhere (linked via "Unified deal") skips settling
      // here on their own — their weekly net (P&L + rakeback) feeds that shared
      // deal instead (see unifiedFeeds below), same as a Fish Tank/AA backed row.
      m.allPlayers.filter((p) => p.played && !m.inUmbrella.has(p.id) && !(p.unifiedDealId || "").trim()).forEach((p) => items.push({ ...resolve(p.name, "My Clubs", "player"), amount: -p.settlement, note: `My Clubs · ${p.name} (${p.clubName})`, checklist: true }));
      m.clubs.filter((c) => c.active).forEach((c) => items.push({ ...resolve((c.owner || "").trim() || c.name, "My Clubs", "club"), amount: c.clubSettlement, note: `Club ${c.name}` }));
      const unifiedFeeds = [];
      m.allPlayers.forEach((p) => {
        if (!p.played || m.inUmbrella.has(p.id)) return;
        const dealId = (p.unifiedDealId || "").trim();
        if (!dealId) return;
        unifiedFeeds.push({ dealId, date: today(), game: `My Clubs · ${p.clubName}`, pnl: r2(p.net), holder: "ak", note: `auto · My Clubs · ${p.clubName} · ${wk}`, cp: resolve(p.name, "My Clubs", "player"), noteMatch: `My Clubs · ${p.name} (${p.clubName})` });
      });
      pending.push({ sourceKey: `mc:${wk}`, label: `My Clubs · ${wk}`, items, stakingItems: [], unifiedFeeds });
    }
  }
  return { pending, live };
}

// ——— Staking engine ———
// Deal: { id, cpId, type: 'action'|'makeup', pct, akChopPct, cadence: 'session'|'weekly', makeupStart, note }
// Result: { id, dealId, date, game, pnl, pct, holder: 'ak'|'player', note }
// net (your share) = pnl × pct%. The player's entitlement E = his unstaked share
// (pnl − net) + his chop; what he actually holds H = pnl on entries where he held
// the money. His tab moves by H − E — the same formula for wins, losses, either holder.
const weekOf = (date) => { // Monday of that ISO week, as YYYY-MM-DD
  const d = new Date(date + "T00:00:00"); if (isNaN(d)) return date;
  const day = (d.getDay() + 6) % 7; d.setDate(d.getDate() - day);
  return d.toISOString().slice(0, 10);
};
function computeStaking(data, cps, live, persons) {
  const st = data.staking || TABS_EMPTY.staking;
  // Imported rows are stored once (no double-counting) and never revisited, but
  // their raw name/site were kept alongside — re-map through CURRENT aliases
  // every time so a bundle added after the fact still folds historical rows in,
  // instead of leaving them stuck under whatever name was live when imported.
  const mapImportedName = makeNameMapper(persons || []);
  const cpName = (id) => cps.find((c) => c.id === id)?.name || "?";
  const byDeal = {};
  const derived = []; // tab entries produced by staking
  const dealOut = [];
  st.deals.forEach((deal) => {
    const rs = st.results.filter((r) => r.dealId === deal.id).sort((a, b) => a.date.localeCompare(b.date) || st.results.indexOf(a) - st.results.indexOf(b));
    const rows = [];
    let netAB = 0, chopped = 0, makeup = r2(deal.makeupStart || 0);
    const akPct = deal.akChopPct ?? 50;
    const settle = (net, pnl, held, unstaked, when, refId, label) => {
      let recovered = 0, excess = 0, akChop = 0, playerChop = 0;
      if (deal.type === "makeup") {
        if (net < 0) makeup = r2(makeup - net);
        else { recovered = Math.min(net, makeup); makeup = r2(makeup - recovered); excess = net - recovered; akChop = r2(excess * akPct / 100); playerChop = r2(excess - akChop); }
        chopped = r2(chopped + akChop);
      } else netAB = r2(netAB + net);
      const playerEnt = unstaked + playerChop;
      const tab = r2(held - playerEnt);
      if (Math.abs(tab) > 0.005) derived.push({ id: "stk-" + refId, date: when, cpId: deal.cpId, amount: tab, note: label, source: "staking", dealId: deal.id, refId });
      return { recovered, excess, akChop, playerChop, tab, makeupAfter: makeup };
    };
    if (deal.type === "action" || (deal.cadence || "session") === "session") {
      rs.forEach((r) => {
        const pct = r.pct ?? deal.pct ?? 100;
        const net = r2((+r.pnl || 0) * pct / 100);
        const held = r.holder === "player" ? +r.pnl || 0 : 0;
        const label = `${deal.type === "action" ? "Action buy" : "Makeup"} · ${r.game || "game"} · P&L ${fmt(+r.pnl || 0)} @ ${pct}%${r.holder === "player" ? " (player held)" : " (Ak held)"}`;
        const x = settle(net, +r.pnl || 0, held, (+r.pnl || 0) - net, r.date, r.id, label);
        rows.push({ ...r, pct, net, ...x, weekKey: null });
      });
    } else {
      // weekly cadence: makeup / chop computed on the week's net. A result's
      // own date normally decides which week it lands in, but that's overridable
      // per-result (r.groupWeek) — needed because a re-synced/backfilled auto
      // result is dated the day it was posted, not the day the game happened, so
      // going back to correct an old week's numbers would otherwise bump that
      // result into whatever week "today" falls in instead of staying put.
      const weeks = {};
      rs.forEach((r) => { const w = weekOf(r.groupWeek || r.date); (weeks[w] = weeks[w] || []).push(r); });
      Object.keys(weeks).sort().forEach((w) => {
        const list = weeks[w];
        let net = 0, held = 0, unstaked = 0;
        const partial = list.map((r) => { const pct = r.pct ?? deal.pct ?? 100; const n = r2((+r.pnl || 0) * pct / 100); net += n; if (r.holder === "player") held += +r.pnl || 0; unstaked += (+r.pnl || 0) - n; return { ...r, pct, net: n }; });
        const last = list[list.length - 1];
        const x = settle(r2(net), 0, r2(held), r2(unstaked), last.date, deal.id + "-" + w, `Makeup · week of ${w} · net ${fmt(net)}`);
        partial.forEach((p, i) => rows.push({ ...p, weekKey: w, weekNet: r2(net), ...(i === partial.length - 1 ? x : { recovered: null, excess: null, akChop: null, playerChop: null, tab: null, makeupAfter: null }) }));
      });
    }
    const out = { deal, name: cpName(deal.cpId), rows, netActionBuy: netAB, chopped, makeup, inMakeup: makeup > 0.005 };
    byDeal[deal.id] = out; dealOut.push(out);
  });
  // imported (Fish Tank / owner clubs) — your share
  const imp = st.imported || [];
  const impByName = {};
  imp.forEach((it) => { const nm = it.rawName != null ? mapImportedName(it.rawName, it.site) : it.name; const k = nm.toLowerCase(); impByName[k] = impByName[k] || { name: nm, netActionBuy: 0, chopped: 0, rows: [] }; impByName[k].rows.push(it); if (it.kind === "action") impByName[k].netActionBuy = r2(impByName[k].netActionBuy + it.net); else impByName[k].chopped = r2(impByName[k].chopped + (it.chopped || 0)); });
  (live || []).forEach((l) => { const k = l.name.toLowerCase(); impByName[k] = impByName[k] || { name: l.name, netActionBuy: 0, chopped: 0, rows: [] }; (impByName[k].live = impByName[k].live || []).push(l); });
  const importedList = Object.values(impByName).map((x) => ({ ...x, makeup: r2((x.live || []).reduce((a, l) => a + l.makeup, 0)) }));
  // Ended stakes settle out of the ongoing totals below and land in their own
  // running tally instead — see endedTotals.
  const activeDealOut = dealOut.filter((d) => !d.deal.ended);
  const totals = {
    netActionBuy: r2(activeDealOut.reduce((a, d) => a + d.netActionBuy, 0) + importedList.reduce((a, x) => a + x.netActionBuy, 0)),
    chopped: r2(activeDealOut.reduce((a, d) => a + d.chopped, 0) + importedList.reduce((a, x) => a + x.chopped, 0)),
    makeup: r2(activeDealOut.reduce((a, d) => a + d.makeup, 0) + importedList.reduce((a, x) => a + x.makeup, 0)),
  };
  // per-counterparty makeup (manual deals + imported by name) — ended deals no
  // longer count as owed once they're closed out.
  const makeupByCp = {};
  activeDealOut.forEach((d) => { if (d.inMakeup) makeupByCp[d.deal.cpId] = r2((makeupByCp[d.deal.cpId] || 0) + d.makeup); });
  importedList.forEach((x) => { if (x.makeup > 0.005) { const cp = cps.find((c) => c.name.toLowerCase() === x.name.toLowerCase()); if (cp) makeupByCp[cp.id] = r2((makeupByCp[cp.id] || 0) + x.makeup); } });
  // Ended stakes — separate running log + lifetime total, kept out of the
  // ongoing figures above. A makeup deal's ending result is (total chopped −
  // outstanding makeup): positive nets a win, negative nets a loss (the write-
  // off — the difference between what was still owed and what was already
  // recouped via chops). "signed" is the same amount but +/- for summing.
  const endedDeals = dealOut.filter((d) => d.deal.ended).map((d) => ({ ...d, signed: d.deal.endOutcome === "loss" ? -d.deal.endAmount : d.deal.endAmount }));
  const endedMakeupOnly = r2(endedDeals.filter((d) => d.deal.type === "makeup").reduce((a, d) => a + d.signed, 0));
  // The "include action buys" toggle isn't about ended action-buy deals —
  // those barely exist, action buys just run on indefinitely rather than
  // getting "ended". It's the running net action buy off every currently-open
  // manual action-buy deal in the deals area above (not gated on that same
  // player also having an ended makeup stake — most won't), and never
  // staking.imported (house-backed) rows, which are still bugged and excluded
  // for now. A deal that HAS been individually ended is skipped here — its
  // result already lands in endedDeals above, so counting it again here
  // would double it up.
  const liveActionBuyRows = dealOut.filter((d) => d.deal.type === "action" && !d.deal.ended);
  const liveActionBuyTotal = r2(liveActionBuyRows.reduce((a, d) => a + d.netActionBuy, 0));
  const endedTotals = {
    makeupOnly: endedMakeupOnly,
    all: r2(endedMakeupOnly + liveActionBuyTotal),
  };
  return { deals: dealOut, byDeal, derived, imported: importedList, totals, makeupByCp, endedDeals, endedTotals, liveActionBuyRows };
}

const TABS_VIEWS = [["balances", "Balances"], ["bookkeeping", "Bookkeeping"], ["ledger", "Ledger"], ["vig", "Vig"], ["staking", "Staking"], ["misc", "Misc. P&L"], ["players", "Player data"]];
const sortDateDesc = (list, indexOf) => [...list].sort((a, b) => (b.date || "").localeCompare(a.date || "") || indexOf(b) - indexOf(a));
const dateInput = (v, onChange, w = 118) => <input type="date" value={v || ""} onChange={(e) => onChange(e.target.value)} style={{ ...inputS, width: w, fontSize: 12.5 }} />;
const iconBtn = (label, onClick, color, title) => <button onClick={onClick} title={title} style={{ border: "none", background: "none", color, cursor: "pointer", fontSize: 13, padding: "0 4px" }}>{label}</button>;

// ———————————————— Book: club weekly P&L summary + reminders ————————————————
// A generic Fish-Tank-style weekly report for the ledger, separate from Tabs
// (counterparty balances/staking). Loads its own snapshot of Fish Tank, every
// owner club, and My Clubs from storage — the same pattern loadFishTankModel /
// loadAAModel / loadAgentModel already use elsewhere in the app.
const BOOK_CHECKLIST_KEY = "book-checklist-v1";
async function loadBookChecklist() {
  try { const c = await store.get(BOOK_CHECKLIST_KEY); if (c?.value) { const v = JSON.parse(c.value); if (Array.isArray(v.items)) return v.items; } } catch (e) {}
  return [];
}
async function saveBookChecklist(items) { try { await store.set(BOOK_CHECKLIST_KEY, JSON.stringify({ items })); } catch (e) {} }

// Saved weekly snapshots of the Book summary (label → totals), same key/shape as the aks-book site.
const BOOK_WEEKS_KEY = "book-weeks-v1";
async function loadBookWeeks() {
  try { const c = await store.get(BOOK_WEEKS_KEY); if (c?.value) { const v = JSON.parse(c.value); if (v && typeof v.weeks === "object") return v.weeks; } } catch (e) {}
  return {};
}
async function saveBookWeeks(weeks) { try { await store.set(BOOK_WEEKS_KEY, JSON.stringify({ weeks })); } catch (e) {} }

function BookSection() {
  const [subtab, setSubtab] = useState("summary");
  const [loaded, setLoaded] = useState(false);
  const [ft, setFt] = useState(null);
  const [ownerClubs, setOwnerClubs] = useState([]);
  const [agent, setAgent] = useState(null);
  const [persons, setPersons] = useState([]);
  const [checklist, setChecklist] = useState([]);
  const [weeks, setWeeks] = useState({});
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => { (async () => {
    setLoaded(false);
    const [ftM, ocM, agM, ppl, cl, wk] = await Promise.all([loadFishTankModel(), loadAllOwnerClubModels(), loadAgentModel(), loadPersons(), loadBookChecklist(), loadBookWeeks()]);
    setFt(ftM); setOwnerClubs(ocM); setAgent(agM); setPersons(ppl); setChecklist(cl); setWeeks(wk);
    setLoaded(true);
  })(); }, [refreshKey]);

  const saveChecklist = async (items) => { setChecklist(items); await saveBookChecklist(items); };
  const saveWeeks = async (w) => { setWeeks(w); await saveBookWeeks(w); };
  // Snapshot the live summary under a label (e.g. "09/01 - 09/07") so it shows in Saved weeks.
  const saveCurrentWeek = async () => {
    const t = bookTotals({ ft, ownerClubs, agent });
    const label = window.prompt("Save this week's totals under what label? (e.g. 09/01 - 09/07)", t.period || "");
    if (!label) return;
    if (weeks[label] && !window.confirm(`"${label}" is already saved — overwrite it with today's numbers?`)) return;
    await saveWeeks({ ...weeks, [label]: { ...t, savedAt: new Date().toISOString() } });
    setSubtab("weeks");
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap", alignItems: "center" }}>
        {[["summary", "Summary"], ["weeks", `Saved weeks${Object.keys(weeks).length ? ` (${Object.keys(weeks).length})` : ""}`], ["checklist", "Checklist"]].map(([k, label]) => (
          <button key={k} onClick={() => setSubtab(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: subtab === k ? C.card : "transparent", color: subtab === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2,
            boxShadow: subtab === k ? "0 -1px 4px rgba(0,0,0,0.1)" : "none" }}>
            {label}
          </button>
        ))}
        {subtab === "summary" && loaded && <button onClick={saveCurrentWeek} title="Save these totals as a week in Saved weeks" style={{ marginLeft: "auto", marginBottom: 6, border: `1px solid ${C.line}`, background: "none", color: C.ink, cursor: "pointer", borderRadius: 6, padding: "4px 10px", fontSize: 12, fontWeight: 600 }}>Save week</button>}
        <button onClick={() => setRefreshKey((k) => k + 1)} title="Reload Fish Tank / owner club / My Clubs data" style={{ marginLeft: subtab === "summary" && loaded ? 6 : "auto", marginBottom: 6, border: `1px solid ${C.line}`, background: "none", color: C.mute, cursor: "pointer", borderRadius: 6, padding: "4px 10px", fontSize: 12 }}>↻ Refresh</button>
      </div>
      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {!loaded ? <div style={{ color: C.mute, padding: 20 }}>Loading…</div> : (
          <>
            {subtab === "summary" && <BookSummary ft={ft} ownerClubs={ownerClubs} agent={agent} />}
            {subtab === "weeks" && <BookSavedWeeks weeks={weeks} save={saveWeeks} />}
            {subtab === "checklist" && <BookChecklist items={checklist} save={saveChecklist} persons={persons} />}
          </>
        )}
      </div>
    </div>
  );
}

// A backed player's book, split into what's actually been realized vs what's
// still just an open marker. "Chopped profit" = action-buy net (realized
// immediately, every session) + a makeup deal's cut of any excess win once
// the player's makeup is cleared (also realized — the backer actually keeps
// it). "Makeup" = the CURRENT outstanding balance on makeup deals still in
// the red — not a loss yet, just an open marker until the stake ends (see
// Tabs → Staking → End stake), so it's reported as its own figure rather
// than netted in as a negative.
function backedChopMakeup(entities, shareFn) {
  let chop = 0, makeup = 0;
  (entities || []).forEach((e) => {
    const share = shareFn(e.backer);
    if (!share) return;
    if (e.dealType === "action") { chop += e.backerBook * share; return; }
    const recovered = e.makeupBefore - e.makeupAfter;
    chop += (e.backerBook - recovered) * share;
    if (e.inMakeup) makeup += e.makeupAfter * share;
  });
  return { chop: r2(chop), makeup: r2(makeup) };
}

// Same math as BookSummary, flattened into the book-weeks-v1 snapshot shape.
function bookTotals({ ft, ownerClubs, agent }) {
  const ftPersonal = ft ? ft.model.ownPosition.ak : 0;
  const ftFee = ft ? r2(ft.model.entitle.ak - ftPersonal - ft.model.backedBook.ak) : 0;
  const ocRows = ownerClubs.map((oc) => {
    const meId = oc.club.meId;
    const personal = oc.model.ownPosition[meId] || 0;
    return { name: oc.club.name, period: oc.period, personal, fee: r2((oc.model.profit[meId] || 0) - personal) };
  });
  const ocPersonalTotal = r2(ocRows.reduce((a, o) => a + o.personal, 0));
  const ocFeeTotal = r2(ocRows.reduce((a, o) => a + o.fee, 0));
  const ocLabel = ocRows.length === 1 ? ocRows[0].name : ocRows.length > 1 ? "Owner clubs" : "Owner club";
  const ocFeeLabel = ocRows.length === 1 ? `${ocRows[0].name} ownership share (pool + BBJ share + personal margin + stake margin)` : `${ocLabel} ownership share`;
  const myAcc = agent ? new Set((agent.acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean)) : new Set();
  const myPlayRows = agent ? agent.model.allPlayers.filter((p) => p.played && myAcc.has(p.name.trim().toLowerCase())) : [];
  const myPlayTotal = r2(myPlayRows.reduce((a, p) => a + p.settlement, 0));
  const mcMargin = agent ? agent.model.totals.margin : 0;
  const mcAdj = agent ? agent.model.totals.globalAdjTotal : 0;
  const personalTotal = r2(ftPersonal + ocPersonalTotal + myPlayTotal);
  const rakeProfitTotal = r2(ftFee + ocFeeTotal + mcMargin);
  const clubTotal = r2(personalTotal + rakeProfitTotal);
  return { period: ft?.period || "", ftPersonal, ftFee, ocRows, ocPersonalTotal, ocFeeTotal, ocLabel, ocFeeLabel, myPlayTotal, myPlayNames: [...new Set(myPlayRows.map((p) => p.name))], mcMargin, mcAdj, personalTotal, rakeProfitTotal, clubTotal, grandTotal: clubTotal, hasFt: !!ft, hasAgent: !!agent };
}

// Read-only view of saved weeks; renders whichever fields a snapshot has (older site snapshots carry staking/vig/misc too).
function BookSavedWeeks({ weeks, save }) {
  const labels = Object.keys(weeks).sort((a, b) => (weeks[b]?.savedAt || "").localeCompare(weeks[a]?.savedAt || ""));
  const [sel, setSel] = useState(labels[0] || "");
  const w = weeks[sel];
  const row = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "5px 0", fontSize: 13.5, paddingLeft: opts.indent ? 16 : 0, borderTop: opts.line ? `1px solid ${C.line}` : "none" }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{money(val || 0)}</span>
    </div>
  );
  const del = async () => {
    if (!sel || !window.confirm(`Delete the saved week "${sel}"? This can't be undone.`)) return;
    const next = { ...weeks }; delete next[sel];
    await save(next); setSel(Object.keys(next)[0] || "");
  };
  if (!labels.length) return <Card title="No saved weeks yet"><div style={{ color: C.mute, fontSize: 13 }}>Use <b>Save week</b> on the Summary tab to snapshot a week's totals.</div></Card>;
  const peopleRows = (title, rows, key) => rows?.length > 0 && (
    <Card title={title}>
      {rows.map((r) => (
        <div key={r.name} style={{ display: "flex", gap: 10, padding: "5px 0", fontSize: 12.5, borderTop: `1px solid ${C.line}`, flexWrap: "wrap" }}>
          <b>{r.name}</b><span style={{ color: C.mute }}>{(r.sites || []).join(", ")}</span>
          <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{key === "makeup" ? <>chop {money(r.chop)} · accrued makeup <b style={{ color: C.goldDark }}>{fmt(r.accrued)}</b></> : <>net {money(r.net)}</>}</span>
        </div>
      ))}
    </Card>
  );
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <select value={sel} onChange={(e) => setSel(e.target.value)} style={{ padding: "6px 8px", borderRadius: 6, border: `1px solid ${C.line}`, background: C.card, color: C.ink, fontSize: 13 }}>
          {labels.map((l) => <option key={l} value={l}>{l}</option>)}
        </select>
        {w?.savedAt && <span style={{ color: C.mute, fontSize: 12 }}>saved {new Date(w.savedAt).toLocaleString()}{w.reconstructed ? " · reconstructed" : ""}</span>}
        <button onClick={del} style={{ marginLeft: "auto", border: `1px solid ${C.line}`, background: "none", color: C.red, cursor: "pointer", borderRadius: 6, padding: "4px 10px", fontSize: 12 }}>Delete week</button>
      </div>
      {w && (<>
        <Card title={`Week ${sel}`}>
          {w.period && <div style={{ color: C.mute, fontSize: 12, marginBottom: 6 }}>{w.period}</div>}
          {row("GRAND TOTAL", w.grandTotal ?? w.clubTotal, { bold: true })}
          {row("Club weekly P&L", w.clubTotal, { line: true, bold: true })}
          {row("Personal play total", w.personalTotal, { indent: true })}
          {row("All in Fish Tank", w.ftPersonal, { indent: true })}
          {row(w.ocLabel || "Owner clubs", w.ocPersonalTotal, { indent: true })}
          {row(`Remaining clubs${w.myPlayNames?.length ? ` (${w.myPlayNames.join(", ")})` : ""}`, w.myPlayTotal, { indent: true })}
          {row("Fee margin total", w.rakeProfitTotal, { indent: true })}
          {row("All in Fish Tank ownership share", w.ftFee, { indent: true })}
          {row(w.ocFeeLabel || "Owner clubs ownership share", w.ocFeeTotal, { indent: true })}
          {row("Personal DL margin", w.mcMargin, { indent: true })}
          {w.stakingVigMiscTotal != null && <>
            {row("Staking + vig + misc", w.stakingVigMiscTotal, { line: true, bold: true })}
            {row("Makeup chopped profit", w.makeupChopTotal, { indent: true })}
            {row("Action buy net", w.actionNetTotal, { indent: true })}
            {row(`Vig${w.priorWeekStart ? ` (${w.priorWeekStart} – ${w.priorWeekEnd})` : ""}`, w.vigWeekTotal, { indent: true })}
            {row("Misc P&L", w.miscWeekTotal, { indent: true })}
          </>}
          {w.makeupAccruedTotal != null && <div style={{ color: C.mute, fontSize: 12, marginTop: 6 }}>Accrued makeup (open marker, not in totals): <b style={{ color: C.goldDark }}>{fmt(w.makeupAccruedTotal)}</b></div>}
        </Card>
        {w.ocRows?.length > 1 && (
          <Card title="Owner clubs — by club">
            {w.ocRows.map((o) => <div key={o.name} style={{ display: "flex", gap: 10, padding: "5px 0", fontSize: 12.5, borderTop: `1px solid ${C.line}`, flexWrap: "wrap" }}><b>{o.name}</b><span style={{ color: C.mute }}>{o.period}</span><span style={{ marginLeft: "auto" }}>personal play {money(o.personal)} · fee margin {money(o.fee)}</span></div>)}
          </Card>
        )}
        {peopleRows("Makeup stakes", w.makeupRows, "makeup")}
        {peopleRows("Action buys", w.actionRows, "action")}
        {(w.vigRows?.length > 0 || w.miscRows?.length > 0) && (
          <Card title="Vig & misc">
            {(w.vigRows || []).map((r) => row(`Vig · ${r.name}`, r.vig))}
            {(w.miscRows || []).map((r, i) => <div key={i}>{row(`Misc · ${r.category || "uncategorized"}`, r.amount)}</div>)}
          </Card>
        )}
      </>)}
    </div>
  );
}

function BookSummary({ ft, ownerClubs, agent }) {
  const row = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "6px 0", fontSize: 13.5, paddingLeft: opts.indent ? 16 : 0 }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{money(val)}</span>
    </div>
  );
  const rowCM = (label, chop, makeup) => (
    <div style={{ display: "flex", padding: "6px 0", fontSize: 13.5, alignItems: "baseline", flexWrap: "wrap", rowGap: 2 }}>
      <span style={{ color: C.mute }}>{label}</span>
      <span style={{ marginLeft: "auto", display: "flex", gap: 18, fontVariantNumeric: "tabular-nums" }}>
        <span>chopped profit {money(chop)}</span>
        <span>makeup <b style={{ color: C.goldDark }}>{fmt(makeup)}</b></span>
      </span>
    </div>
  );

  // Fish Tank: entitle.ak bundles personal play + ½ profit + staking together —
  // split it apart and drop the staking piece (backed books).
  const ftPersonal = ft ? ft.model.ownPosition.ak : 0;
  const ftStaking = ft ? ft.model.backedBook.ak : 0;
  const ftFee = ft ? r2(ft.model.entitle.ak - ftPersonal - ftStaking) : 0;
  const ftCM = ft ? backedChopMakeup(ft.model.backedEntities, (backer) => (backer === "split" ? 0.5 : backer === "ak" ? 1 : 0)) : { chop: 0, makeup: 0 };

  // Owner clubs (Midnight Bazaar + any others added) — Ak's share on each.
  // Unlike Fish Tank, "profit" here already excludes the deal/backed book, so
  // the fee margin is just profit minus own-account play.
  const ocRows = ownerClubs.map((oc) => {
    const meId = oc.club.meId;
    const personal = oc.model.ownPosition[meId] || 0;
    const staking = oc.model.backedBook[meId] || 0;
    const fee = r2((oc.model.profit[meId] || 0) - personal);
    const cm = backedChopMakeup(oc.model.backedEntities, (backer) => oc.model.H.shareOf(backer, meId));
    return { name: oc.club.name, period: oc.period, personal, staking, fee, cm };
  });
  const ocPersonalTotal = r2(ocRows.reduce((a, o) => a + o.personal, 0));
  const ocFeeTotal = r2(ocRows.reduce((a, o) => a + o.fee, 0));
  const ocChopTotal = r2(ocRows.reduce((a, o) => a + o.cm.chop, 0));
  const ocMakeupTotal = r2(ocRows.reduce((a, o) => a + o.cm.makeup, 0));
  // Only Midnight Bazaar existed as an owner club when these labels were first
  // written, so its name got hardcoded — now that more can be added (e.g.
  // Honey Pot), name it when there's exactly one, and fall back to a generic
  // label (with the per-club breakdown card below) once there's more than one.
  const ocLabel = ocRows.length === 1 ? ocRows[0].name : ocRows.length > 1 ? "Owner clubs" : "Owner club";
  const ocFeeLabel = ocRows.length === 1 ? `${ocRows[0].name} ownership share (pool + BBJ share + personal margin + stake margin)` : `${ocLabel} ownership share`;

  // My Clubs (the "agent" section) — margin on players is fee income; personal
  // play is whatever's tagged as "My accounts" there, settling at other clubs.
  const myAcc = agent ? new Set((agent.acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean)) : new Set();
  const myPlayRows = agent ? agent.model.allPlayers.filter((p) => p.played && myAcc.has(p.name.trim().toLowerCase())) : [];
  const myPlayTotal = r2(myPlayRows.reduce((a, p) => a + p.settlement, 0));
  const mcMargin = agent ? agent.model.totals.margin : 0;
  const mcAdj = agent ? agent.model.totals.globalAdjTotal : 0;

  const personalTotal = r2(ftPersonal + ocPersonalTotal + myPlayTotal);
  const feeTotal = r2(ftFee + ocFeeTotal + mcMargin);
  const total = r2(personalTotal + feeTotal);

  return (
    <div>
      <div style={{ marginBottom: 14 }}>
        <Card title="Total club weekly P&L">
          <div style={{ color: C.mute, fontSize: 12, marginBottom: 8 }}>Personal play + fee margins only — staking (backed books / deal books, chop, makeup) is tracked separately in Tabs → Staking and excluded here.</div>
          {row("TOTAL", total, { bold: true })}
          <div style={{ borderTop: `1px solid ${C.line}`, margin: "10px 0 4px" }} />
          <div style={{ fontWeight: 700, fontSize: 13, marginTop: 6, marginBottom: 2 }}>Personal play</div>
          {row(`All in Fish Tank${ft?.period ? ` · ${ft.period}` : ""}`, ftPersonal, { indent: true })}
          {!ft && <div style={{ color: C.mute, fontSize: 11.5, paddingLeft: 16 }}>No Fish Tank week loaded.</div>}
          {row(ocLabel, ocPersonalTotal, { indent: true })}
          {ocRows.length === 0 && <div style={{ color: C.mute, fontSize: 11.5, paddingLeft: 16 }}>No owner club week loaded.</div>}
          {row(`Remaining clubs${myPlayRows.length ? ` (${[...new Set(myPlayRows.map((p) => p.name))].join(", ")})` : ""}`, myPlayTotal, { indent: true })}
          {!agent && <div style={{ color: C.mute, fontSize: 11.5, paddingLeft: 16 }}>No My Clubs week loaded.</div>}
          {row("Personal play total", personalTotal, { bold: true })}
          <div style={{ borderTop: `1px solid ${C.line}`, margin: "10px 0 4px" }} />
          <div style={{ fontWeight: 700, fontSize: 13, marginTop: 6, marginBottom: 2 }}>Fee margin profits</div>
          {row(`All in Fish Tank ownership share${ft?.period ? ` · ${ft.period}` : ""}`, ftFee, { indent: true })}
          {row(ocFeeLabel, ocFeeTotal, { indent: true })}
          {row("Personal DL margin", mcMargin, { indent: true })}
          {row("Fee margin total", feeTotal, { bold: true })}
        </Card>
      </div>

      {ocRows.length > 1 && (
        <div style={{ marginBottom: 14 }}>
          <Card title="Owner clubs — by club">
            {ocRows.map((o) => (
              <div key={o.name} style={{ borderTop: `1px solid ${C.line}`, padding: "6px 0", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", fontSize: 12.5 }}>
                <b>{o.name}</b><span style={{ color: C.mute }}>{o.period}</span>
                <span style={{ marginLeft: "auto" }}>personal play {money(o.personal)} · fee margin {money(o.fee)}</span>
              </div>
            ))}
          </Card>
        </div>
      )}

      <Card title="Not included above — tracked separately">
        <div style={{ color: C.mute, fontSize: 12, marginBottom: 8 }}>Staking and My Clubs' general adjustments aren't part of the club P&L above — shown here for reference only. Makeup is the current outstanding balance, not a loss — it only becomes one if a stake ends still in the red (Tabs → Staking → End stake).</div>
        {rowCM("Fish Tank — backed books (staking)", ftCM.chop, ftCM.makeup)}
        {rowCM(`${ocLabel} — deal books (staking)`, ocChopTotal, ocMakeupTotal)}
        {row("My Clubs — general adjustments", mcAdj)}
      </Card>
    </div>
  );
}

function BookChecklist({ items, save, persons }) {
  const [draft, setDraft] = useState({ player: "", note: "", due: today() });
  const addItem = () => {
    if (!draft.note.trim()) return;
    save([...items, { id: uid(), player: draft.player.trim(), note: draft.note.trim(), due: draft.due || today(), done: false }]);
    setDraft({ player: "", note: "", due: today() });
  };
  const toggleDone = (id) => save(items.map((x) => (x.id === id ? { ...x, done: !x.done } : x)));
  const removeItem = (id) => save(items.filter((x) => x.id !== id));
  const clearDone = () => save(items.filter((x) => !x.done));
  const open = [...items.filter((x) => !x.done)].sort((a, b) => (a.due || "").localeCompare(b.due || ""));
  const done = items.filter((x) => x.done);
  const allNames = [...new Set(persons.map((p) => p.name))].sort((a, b) => a.localeCompare(b));
  const isOverdue = (x) => !x.done && x.due && x.due < today();

  return (
    <div>
      <div style={{ marginBottom: 14 }}>
        <Card title="New reminder">
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <input list="book-player-names" placeholder="Player (optional)" value={draft.player} onChange={(e) => setDraft({ ...draft, player: e.target.value })} style={{ ...inputS, width: 160 }} />
            <datalist id="book-player-names">{allNames.map((n) => <option key={n} value={n} />)}</datalist>
            <input placeholder="Reminder…" value={draft.note} onChange={(e) => setDraft({ ...draft, note: e.target.value })} style={{ ...inputS, flex: 1, minWidth: 160 }} onKeyDown={(e) => { if (e.key === "Enter") addItem(); }} />
            {dateInput(draft.due, (v) => setDraft({ ...draft, due: v }))}
            <Btn tone="gold" small onClick={addItem} disabled={!draft.note.trim()}>+ Add</Btn>
          </div>
        </Card>
      </div>
      <Card title={`Reminders${open.length ? ` · ${open.length} open` : ""}`} right={done.length > 0 ? <Btn tone="ghost" small onClick={clearDone}>Clear checked ({done.length})</Btn> : null}>
        {items.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No reminders yet — add one above.</div>}
        {[...open, ...done].map((x) => (
          <div key={x.id} style={{ display: "flex", alignItems: "center", gap: 8, borderTop: `1px solid ${C.line}`, padding: "7px 0", fontSize: 13, opacity: x.done ? 0.5 : 1 }}>
            <input type="checkbox" checked={!!x.done} onChange={() => toggleDone(x.id)} />
            {x.player && <span style={{ fontWeight: 600 }}>{x.player}</span>}
            <span style={{ textDecoration: x.done ? "line-through" : "none" }}>{x.note}</span>
            <span style={{ marginLeft: "auto", fontSize: 11, color: isOverdue(x) ? C.red : C.mute, fontWeight: isOverdue(x) ? 700 : 400, whiteSpace: "nowrap" }}>{x.due}{isOverdue(x) ? " · overdue" : ""}</span>
            {iconBtn("×", () => removeItem(x.id), C.red, "Remove")}
          </div>
        ))}
      </Card>
    </div>
  );
}

function TabsLedger({ clubs }) {
  const [data, setData] = useState(TABS_EMPTY);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("balances");
  const [exportData, setExportData] = useState(null);
  const [filterCp, setFilterCp] = useState("");
  const [showZero, setShowZero] = useState(false);
  const [pending, setPending] = useState([]);
  const [live, setLive] = useState([]);
  const [siteUsernames, setSiteUsernames] = useState([]);
  const [tweaks, setTweaks] = useState({});
  const [editId, setEditId] = useState(null);
  const [editDraft, setEditDraft] = useState({});
  const [note, setNote] = useState("");
  const [expandedCp, setExpandedCp] = useState(null);
  const [confirmDeleteCp, setConfirmDeleteCp] = useState(null);
  const [showPosted, setShowPosted] = useState(false);

  // Pulls the current computed numbers off Fish Tank/AA/My Clubs (via
  // collectPendingWeeks). Only runs on mount by default — call it again with
  // the "↻ Refresh" button in Bookkeeping after unlocking/editing/relocking a
  // week elsewhere, so a re-sync there picks up the corrected numbers instead
  // of whatever was cached when this page first loaded.
  const refreshPending = async () => {
    let d = await loadTabs();
    const { pending: p, live: l } = await collectPendingWeeks(d.persons, clubs || []);
    // Backfill: weeks already folded into tabs before the Staking section existed.
    let changed = false;
    p.forEach((s) => { if (d.pushed[s.sourceKey] && !d.stakingPushed[s.sourceKey]) { d = applyStakingImport({ ...d, staking: { ...d.staking, imported: [...d.staking.imported] }, stakingPushed: { ...d.stakingPushed } }, s.sourceKey, s.stakingItems); changed = true; } });
    // Backfill: a backed player's row got linked to a unified deal after their
    // current week was already accepted — pick that week's feed up retroactively.
    // (applyUnifiedStakingResults dedupes per-deal against existing results, so
    // this safely re-checks every load instead of only firing once per source.)
    p.forEach((s) => { if (d.pushed[s.sourceKey] && (s.unifiedFeeds || []).length) { const beforeR = d.staking.results.length, beforeE = d.entries.length; d = applyUnifiedStakingResults({ ...d, staking: { ...d.staking, results: [...d.staking.results] }, unifiedPushed: { ...d.unifiedPushed } }, s.sourceKey, s.unifiedFeeds); if (d.staking.results.length !== beforeR || d.entries.length !== beforeE) changed = true; } });
    if (changed) await saveTabs(d);
    const su = await loadAllSiteUsernames(clubs || []);
    setData(d); setPending(p); setLive(l); setSiteUsernames(su); setLoaded(true);
  };
  useEffect(() => { refreshPending(); }, []);
  const save = async (next) => { setData(next); await saveTabs(next); };
  const flash = (m) => { setNote(m); setTimeout(() => setNote(""), 3500); };

  const cps = useMemo(() => [...data.counterparties].sort((a, b) => a.name.localeCompare(b.name)), [data]);
  const cpById = (id) => cps.find((c) => c.id === id);
  const staking = useMemo(() => computeStaking(data, cps, live, data.persons), [data, cps, live]);
  const allEntries = useMemo(() => [...data.entries, ...staking.derived], [data, staking]);
  const balances = useMemo(() => { const m = {}; allEntries.forEach((e) => (m[e.cpId] = (m[e.cpId] || 0) + e.amount)); return m; }, [allEntries]);
  const totalAll = Object.values(balances).reduce((a, v) => a + v, 0);
  const vigEntries = data.entries.filter((e) => isCrypto(e.method));
  const vigTotal = r2(vigEntries.reduce((a, e) => a + (+e.vig || 0), 0));
  const miscByCat = useMemo(() => { const m = {}; data.misc.forEach((e) => { const k = e.category || "uncategorized"; m[k] = r2((m[k] || 0) + e.amount); }); return m; }, [data]);
  const miscTotal = r2(data.misc.reduce((a, e) => a + e.amount, 0));
  const otherTotals = [["Crypto vig (total)", vigTotal], ["Net action buy", staking.totals.netActionBuy], ["Total chopped on stake", staking.totals.chopped], ["Total makeup (owed to you in play)", staking.totals.makeup], ...Object.entries(miscByCat).map(([k, v]) => [`Misc · ${k}`, v])];

  // ——— counterparties ———
  const findOrCreateCp = (list, name, kind) => {
    const key = kind + "|" + name.trim().toLowerCase();
    let cp = list.find((c) => c.kind + "|" + c.name.toLowerCase() === key);
    if (cp) return [list, cp];
    cp = { id: uid(), name: name.trim(), kind };
    return [[...list, cp], cp];
  };
  const mergeCps = (next, fromIds, intoId) => {
    const set = new Set(fromIds.filter((id) => id !== intoId));
    next.entries = next.entries.map((e) => (set.has(e.cpId) ? { ...e, cpId: intoId } : e));
    next.staking = { ...next.staking, deals: next.staking.deals.map((d) => (set.has(d.cpId) ? { ...d, cpId: intoId } : d)) };
    next.counterparties = next.counterparties.filter((c) => !set.has(c.id));
    if (next.settleChecklist) next.settleChecklist = next.settleChecklist.map((x) => (set.has(x.cpId) ? { ...x, cpId: intoId } : x));
    return next;
  };
  const renameCp = async (cp) => {
    const name = window.prompt(`Rename "${cp.name}" to:`, cp.name);
    if (!name || !name.trim() || name.trim() === cp.name) return;
    const dup = cps.find((c) => c.id !== cp.id && c.kind === cp.kind && c.name.toLowerCase() === name.trim().toLowerCase());
    let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries] };
    if (dup) {
      if (!window.confirm(`"${dup.name}" already exists. Merge ${cp.name}'s entries into ${dup.name}?`)) return;
      next = mergeCps(next, [cp.id], dup.id);
      if (filterCp === cp.id) setFilterCp(dup.id);
    } else next.counterparties = next.counterparties.map((c) => (c.id === cp.id ? { ...c, name: name.trim() } : c));
    await save(next);
  };
  const settleUp = async (cp) => {
    const bal = balances[cp.id] || 0;
    if (Math.abs(bal) < 0.005) return;
    if (!window.confirm(`Log a settling entry of ${fmt(-bal)} for ${cp.name}?`)) return;
    await save({ ...data, entries: [...data.entries, { id: uid(), date: today(), cpId: cp.id, amount: r2(-bal), note: "Settled up", source: "manual" }] });
  };

  // ——— settle checklist (Bookkeeping) ———
  const settleChecklist = data.settleChecklist || [];
  const toggleChecklistDone = (id) => save({ ...data, settleChecklist: settleChecklist.map((x) => (x.id === id ? { ...x, done: !x.done } : x)) });
  const removeChecklistItem = (id) => save({ ...data, settleChecklist: settleChecklist.filter((x) => x.id !== id) });
  const clearDoneChecklist = () => save({ ...data, settleChecklist: settleChecklist.filter((x) => !x.done) });

  // ——— pending weeks review ———
  const setTweak = (srcKey, rowKey, patch) => setTweaks((t) => ({ ...t, [srcKey]: { ...(t[srcKey] || {}), [rowKey]: { ...(t[srcKey]?.[rowKey] || {}), ...patch } } }));
  const aggRows = (s) => {
    const agg = {};
    s.items.filter((it) => Math.abs(it.amount) > 0.005).forEach((it) => {
      const k = it.kind + "|" + it.name.toLowerCase();
      if (!agg[k]) agg[k] = { key: k, name: it.name, kind: it.kind, amount: 0, n: 0, notes: [], checklist: false };
      agg[k].amount += it.amount; agg[k].n += 1; if (it.note) agg[k].notes.push(it.note); if (it.checklist) agg[k].checklist = true;
    });
    return Object.values(agg).sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
  };
  const rowState = (s, r) => {
    const t = tweaks[s.sourceKey]?.[r.key] || {};
    const amount = t.amount !== undefined ? t.amount : r2(r.amount);
    const name = (t.name !== undefined && t.name.trim()) ? t.name.trim() : r.name;
    const edited = Math.abs(amount - r.amount) > 0.005 || name !== r.name;
    let nt = r.notes.length <= 1 ? (r.notes[0] || "") : r.notes.join(" + ");
    if (nt.length > 140) nt = nt.slice(0, 137) + "…";
    if (edited) nt = (nt ? nt + " · " : "") + "edited on review";
    return { key: r.key, kind: r.kind, skip: !!t.skip, amount, name, note: nt, edited, checklist: r.checklist, orig: r };
  };
  const acceptSources = async (sources, resync) => {
    let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries], pushed: { ...data.pushed }, stakingPushed: { ...data.stakingPushed }, unifiedPushed: { ...data.unifiedPushed }, staking: { ...data.staking, imported: [...data.staking.imported], results: [...data.staking.results] }, settleChecklist: [...(data.settleChecklist || [])] };
    sources.forEach((s) => {
      // Re-sync (a week that was already accepted, then unlocked/edited/relocked):
      // clear its old posted entries/staking rows first, so the corrected numbers
      // replace them cleanly instead of stacking on top or netting oddly. Naively
      // that means EVERY line for this source — not just the one you actually
      // went back to fix — comes back as a brand-new entry dated today, even
      // players/clubs whose number didn't change at all. So snapshot the old
      // entries/checklist rows (by counterparty) and the old unified-feed
      // results (by deal) before clearing; after the corrected batch is
      // reapplied, any line whose amount+note came back identical gets its
      // original id/date restored instead of looking like a fresh transaction
      // that happened today — only genuinely changed lines (or new/removed
      // ones) end up looking new.
      // Both preservedDates and preservedEntries key on more than just
      // dealId/cpId — a single deal or counterparty can have MORE THAN ONE
      // line under the same sourceKey at once (e.g. one player linked to a
      // shared deal from two different My Clubs sub-clubs in one week, each
      // its own result/entry). Keying on dealId/cpId alone would collapse
      // those into one snapshot slot and only ever restore one of them,
      // leaving the other looking like a fresh line every resync.
      let preservedDates = null, preservedEntries = null, preservedChecklist = null;
      if (resync) {
        preservedDates = {};
        next.staking.results.filter((r) => r.sourceKey === s.sourceKey).forEach((r) => { preservedDates[r.dealId + "|" + r.game] = { date: r.date, groupWeek: r.groupWeek || null }; });
        preservedEntries = {};
        next.entries.filter((e) => e.sourceKey === s.sourceKey).forEach((e) => { preservedEntries[e.cpId + "|" + e.note] = { id: e.id, date: e.date, amount: e.amount }; });
        preservedChecklist = {};
        (next.settleChecklist || []).filter((x) => x.sourceKey === s.sourceKey && !x.done).forEach((x) => { preservedChecklist[x.cpId] = { id: x.id, addedAt: x.addedAt }; });
        next = clearSource(next, s.sourceKey);
      }
      const items = aggRows(s).map((r) => rowState(s, r)).filter((r) => !r.skip && Math.abs(r.amount) > 0.005).map((r) => ({ name: r.name, kind: r.kind, amount: r.amount, note: r.note, checklist: r.checklist }));
      next = applyWeekToTabsData(next, s.sourceKey, s.label, items);
      next = applyStakingImport(next, s.sourceKey, s.stakingItems);
      next = applyUnifiedStakingResults(next, s.sourceKey, s.unifiedFeeds);
      if (preservedDates) {
        next = { ...next, staking: { ...next.staking, results: next.staking.results.map((r) => {
          const p = r.sourceKey === s.sourceKey ? preservedDates[r.dealId + "|" + r.game] : null;
          return p ? { ...r, date: p.date, groupWeek: p.groupWeek } : r;
        }) } };
      }
      if (preservedEntries) {
        next = { ...next, entries: next.entries.map((e) => {
          if (e.sourceKey !== s.sourceKey) return e;
          const p = preservedEntries[e.cpId + "|" + e.note];
          return (p && Math.abs(p.amount - e.amount) < 0.005) ? { ...e, id: p.id, date: p.date } : e;
        }) };
      }
      if (preservedChecklist) {
        next = { ...next, settleChecklist: (next.settleChecklist || []).map((x) => {
          if (x.sourceKey !== s.sourceKey || x.done) return x;
          const p = preservedChecklist[x.cpId];
          return p ? { ...x, id: p.id, addedAt: p.addedAt } : x;
        }) };
      }
    });
    await save(next);
  };

  // ——— ledger entries ———
  const [f, setF] = useState({ kind: "player", name: "", amount: "", note: "", date: today(), method: "", rate: null });
  const fRate = f.rate ?? data.vigRate ?? 5;
  const fPreview = isCrypto(f.method) && f.amount !== "" ? cryptoEffect(parseFloat(String(f.amount).replace(/,/g, "")), fRate) : null;
  const addEntry = async () => {
    const amt = parseFloat(String(f.amount).replace(/,/g, ""));
    if (!f.name.trim() || isNaN(amt) || amt === 0) return;
    const [counterparties, cp] = findOrCreateCp(data.counterparties, f.name, f.kind);
    const entry = { id: uid(), date: f.date || today(), cpId: cp.id, amount: r2(amt), note: f.note.trim(), source: "manual", method: f.method.trim() || undefined };
    if (isCrypto(f.method)) { const x = cryptoEffect(amt, fRate); entry.baseAmount = r2(amt); entry.vigRate = +fRate; entry.vig = x.vig; entry.amount = x.amount; }
    await save({ ...data, counterparties, entries: [...data.entries, entry] });
    setF({ ...f, name: "", amount: "", note: "", rate: null });
  };

  // ——— swaps: settle one player's tab against another with no cash moving ———
  const [addMode, setAddMode] = useState("entry");
  const [sw, setSw] = useState({ from: "", fromKind: "player", to: "", toKind: "player", amount: "", note: "", date: today(), method: "" });
  const addSwap = async () => {
    const amt = parseFloat(String(sw.amount).replace(/,/g, ""));
    const from = sw.from.trim(), to = sw.to.trim();
    if (!from || !to || isNaN(amt) || amt <= 0 || (from.toLowerCase() === to.toLowerCase() && sw.fromKind === sw.toKind)) return;
    let counterparties = data.counterparties;
    const [cps1, fromCp] = findOrCreateCp(counterparties, from, sw.fromKind);
    counterparties = cps1;
    const [cps2, toCp] = findOrCreateCp(counterparties, to, sw.toKind);
    counterparties = cps2;
    const swapId = uid();
    const d = sw.date || today();
    const note = sw.note.trim();
    const method = sw.method.trim() || undefined;
    const entries = [
      { id: uid(), date: d, cpId: fromCp.id, amount: r2(-amt), note: `Swap → ${toCp.name}${note ? " · " + note : ""}`, source: "swap", swapId, swapWith: toCp.id, method },
      { id: uid(), date: d, cpId: toCp.id, amount: r2(amt), note: `Swap ← ${fromCp.name}${note ? " · " + note : ""}`, source: "swap", swapId, swapWith: fromCp.id, method },
    ];
    await save({ ...data, counterparties, entries: [...data.entries, ...entries] });
    setSw({ ...sw, from: "", to: "", amount: "", note: "" });
  };
  // ——— undo: a short-lived stack of deleted ledger entries (this session only) ———
  const [undoStack, setUndoStack] = useState([]);
  const pushUndo = (label, removed) => setUndoStack((st) => [...st.slice(-19), { id: uid(), label, entries: removed }]);
  const undoLastDelete = async () => {
    const last = undoStack[undoStack.length - 1];
    if (!last) return;
    setUndoStack((st) => st.slice(0, -1));
    await save({ ...data, entries: [...data.entries, ...last.entries] });
    flash(`Restored: ${last.label}`);
  };
  const delEntry = async (id) => {
    const e = data.entries.find((x) => x.id === id);
    if (!e) return;
    if (e.source === "swap" && e.swapId) {
      if (!window.confirm("Delete this swap? Both sides — the sender's and receiver's tabs — will be removed together.")) return;
      const removed = data.entries.filter((x) => x.swapId === e.swapId);
      await save({ ...data, entries: data.entries.filter((x) => x.swapId !== e.swapId) });
      const [a, b] = removed;
      pushUndo(`Swap · ${cpById(a?.cpId)?.name || "?"} ↔ ${cpById(b?.cpId)?.name || "?"}`, removed);
      return;
    }
    if (!window.confirm("Delete this entry?")) return;
    await save({ ...data, entries: data.entries.filter((x) => x.id !== id) });
    pushUndo(`${cpById(e.cpId)?.name || "?"} · ${fmt(e.amount)}`, [e]);
  };
  const startEdit = (e) => { setEditId(e.id); setEditDraft({ date: e.date, amount: e.baseAmount ?? e.amount, note: e.note || "", method: e.method || "", rate: e.vigRate ?? data.vigRate ?? 5, cpId: e.cpId }); };
  const commitEdit = async () => {
    const e = data.entries.find((x) => x.id === editId); if (!e) return;
    const amt = parseFloat(String(editDraft.amount).replace(/,/g, ""));
    if (isNaN(amt)) return;
    let upd = { ...e, date: editDraft.date || e.date, note: editDraft.note, method: editDraft.method.trim() || undefined, cpId: editDraft.cpId || e.cpId };
    if (isCrypto(editDraft.method)) { const x = cryptoEffect(amt, editDraft.rate); upd = { ...upd, baseAmount: r2(amt), vigRate: +editDraft.rate, vig: x.vig, amount: x.amount }; }
    else { upd = { ...upd, amount: r2(amt) }; delete upd.baseAmount; delete upd.vigRate; delete upd.vig; }
    await save({ ...data, entries: data.entries.map((x) => (x.id === editId ? upd : x)) });
    setEditId(null);
  };
  const setEntryRate = async (id, rate) => {
    const e = data.entries.find((x) => x.id === id); if (!e) return;
    const x = cryptoEffect(e.baseAmount ?? e.amount, rate);
    await save({ ...data, entries: data.entries.map((y) => (y.id === id ? { ...y, baseAmount: y.baseAmount ?? y.amount, vigRate: +rate, vig: x.vig, amount: x.amount } : y)) });
  };

  // ——— staking ———
  const [sd, setSd] = useState({ name: "", type: "action", pct: 50, akChopPct: 50, cadence: "session", makeupStart: 0, note: "" });
  const [sr, setSr] = useState({ dealId: "", date: today(), game: "", pnl: "", pct: "", holder: "ak", note: "" });
  const [openDeal, setOpenDeal] = useState({});
  const [editRes, setEditRes] = useState(null);
  const [resDraft, setResDraft] = useState({});
  const [endedIncludeAction, setEndedIncludeAction] = useState(false);
  const addDeal = async () => {
    if (!sd.name.trim()) return;
    const [counterparties, cp] = findOrCreateCp(data.counterparties, sd.name, "player");
    const deal = { id: uid(), cpId: cp.id, type: sd.type, pct: +sd.pct || 0, akChopPct: +sd.akChopPct || 0, cadence: sd.cadence, makeupStart: r2(sd.makeupStart || 0), note: sd.note.trim() };
    await save({ ...data, counterparties, staking: { ...data.staking, deals: [...data.staking.deals, deal] } });
    setSd({ ...sd, name: "", note: "" });
  };
  const setDeal = async (id, patch) => save({ ...data, staking: { ...data.staking, deals: data.staking.deals.map((d) => (d.id === id ? { ...d, ...patch } : d)) } });
  const delDeal = async (id) => { if (!window.confirm("Delete this deal and all its logged results? Tab entries it produced disappear too.")) return; await save({ ...data, staking: { ...data.staking, deals: data.staking.deals.filter((d) => d.id !== id), results: data.staking.results.filter((r) => r.dealId !== id) } }); };
  // Ending a stake settles its lifetime P&L into a separate ended-stakes log
  // (see computeStaking) and locks the deal — no more results can be logged
  // against it. Makeup deals: result = total chopped − outstanding makeup
  // (positive = win, negative = the write-off/loss — makeup was never posted
  // to the tab during play, so ending it just stops it counting as owed).
  // Action-buy deals: result = lifetime net action buy.
  const endStake = async (id) => {
    const out = staking.byDeal[id]; if (!out || out.deal.ended) return;
    const isMakeup = out.deal.type === "makeup";
    const net = r2(isMakeup ? out.chopped - out.makeup : out.netActionBuy);
    const outcome = net >= 0 ? "win" : "loss";
    const amount = Math.abs(net);
    const msg = isMakeup
      ? `End ${out.name}'s stake?\n\nChopped ${money(out.chopped)} − outstanding makeup ${money(out.makeup)} = ${outcome} of ${money(amount)}.\n\nThis logs that result to Ended Stakes and locks the deal — no more results can be added.`
      : `End ${out.name}'s stake?\n\nLifetime net action buy ${money(out.netActionBuy)} logs as a ${outcome} of ${money(amount)}.\n\nThis logs that result to Ended Stakes and locks the deal — no more results can be added.`;
    if (!window.confirm(msg)) return;
    await save({ ...data, staking: { ...data.staking, deals: data.staking.deals.map((d) => (d.id === id ? { ...d, ended: true, endedAt: today(), endOutcome: outcome, endAmount: amount, endChopped: isMakeup ? out.chopped : null, endMakeup: isMakeup ? out.makeup : null, endNetActionBuy: isMakeup ? null : out.netActionBuy } : d)) } });
  };
  const reopenStake = async (id) => {
    if (!window.confirm("Reopen this stake? It moves back to ongoing deals and the ended-stakes entry is removed (its logged results stay intact).")) return;
    await save({ ...data, staking: { ...data.staking, deals: data.staking.deals.map((d) => (d.id === id ? { ...d, ended: false, endedAt: null, endOutcome: null, endAmount: null, endChopped: null, endMakeup: null, endNetActionBuy: null } : d)) } });
  };
  const addResult = async () => {
    const deal = data.staking.deals.find((d) => d.id === sr.dealId); if (!deal || deal.ended) return;
    const pnl = parseFloat(String(sr.pnl).replace(/,/g, "")); if (isNaN(pnl)) return;
    const pct = sr.pct === "" ? deal.pct : +sr.pct;
    const res = { id: uid(), dealId: deal.id, date: sr.date || today(), game: sr.game.trim(), pnl: r2(pnl), pct, holder: sr.holder, note: sr.note.trim() };
    await save({ ...data, staking: { ...data.staking, results: [...data.staking.results, res] } });
    setSr({ ...sr, game: "", pnl: "", note: "" });
    setOpenDeal({ ...openDeal, [deal.id]: true });
  };
  const delResult = async (id) => { if (!window.confirm("Delete this result?")) return; await save({ ...data, staking: { ...data.staking, results: data.staking.results.filter((r) => r.id !== id) } }); };
  const commitRes = async () => {
    const pnl = parseFloat(String(resDraft.pnl).replace(/,/g, "")); if (isNaN(pnl)) return;
    await save({ ...data, staking: { ...data.staking, results: data.staking.results.map((r) => (r.id === editRes ? { ...r, date: resDraft.date, groupWeek: resDraft.groupWeek || null, game: resDraft.game, pnl: r2(pnl), pct: +resDraft.pct, holder: resDraft.holder, note: resDraft.note } : r)) } });
    setEditRes(null);
  };
  const delImported = async (id) => { if (!window.confirm("Remove this imported line from staking totals?")) return; await save({ ...data, staking: { ...data.staking, imported: data.staking.imported.filter((r) => r.id !== id) } }); };

  // ——— misc ———
  const [mf, setMf] = useState({ date: today(), amount: "", note: "", category: "" });
  const [miscSort, setMiscSort] = useState("date");
  const [miscCatFilter, setMiscCatFilter] = useState("");
  const [editMisc, setEditMisc] = useState(null);
  const [miscDraft, setMiscDraft] = useState({});
  const addMisc = async () => {
    const amt = parseFloat(String(mf.amount).replace(/,/g, "")); if (isNaN(amt) || amt === 0) return;
    await save({ ...data, misc: [...data.misc, { id: uid(), date: mf.date || today(), amount: r2(amt), note: mf.note.trim(), category: mf.category.trim() || "uncategorized" }] });
    setMf({ ...mf, amount: "", note: "" });
  };
  const delMisc = async (id) => { if (!window.confirm("Delete this entry?")) return; await save({ ...data, misc: data.misc.filter((e) => e.id !== id) }); };
  const commitMisc = async () => {
    const amt = parseFloat(String(miscDraft.amount).replace(/,/g, "")); if (isNaN(amt)) return;
    await save({ ...data, misc: data.misc.map((e) => (e.id === editMisc ? { ...e, date: miscDraft.date, amount: r2(amt), note: miscDraft.note, category: miscDraft.category.trim() || "uncategorized" } : e)) });
    setEditMisc(null);
  };

  // ——— persons ———
  const SITES = ["", "Fish Tank", ...(clubs || []).map((c) => c.name), "My Clubs"];
  const [newPerson, setNewPerson] = useState("");
  const [addAlias, setAddAlias] = useState({});
  const [pickerOpen, setPickerOpen] = useState({});
  const persons = data.persons || [];
  const setPersons = async (persons) => save({ ...data, persons: normPersons(persons) });
  const addPerson = async () => { if (!newPerson.trim()) return; await setPersons([...persons, { id: uid(), name: newPerson.trim(), kind: "player", aliases: [], notes: "" }]); setNewPerson(""); };
  const patchPerson = (id, p) => setPersons(persons.map((x) => (x.id === id ? { ...x, ...p } : x)));
  // Merges across EVERY kind (player/club/other) — this is what lets someone who
  // plays under a username AND owns/settles a club get folded onto one tab: their
  // "player" tab and their "club" tab both merge into a single cp of per.kind.
  const mergePerson = async (per) => {
    const kind = per.kind || "player";
    const names = new Set([per.name.toLowerCase(), ...per.aliases.map((a) => a.name.trim().toLowerCase())]);
    const matches = cps.filter((c) => names.has(c.name.toLowerCase()) && !(c.kind === kind && c.name.toLowerCase() === per.name.toLowerCase()));
    if (!matches.length) { flash(`No separate tab entries found under ${per.name}'s aliases.`); return; }
    if (!window.confirm(`Merge ${matches.map((m) => `${m.name} (${m.kind})`).join(", ")} into "${per.name}" (${kind})? Their ledger entries and staking deals move under the one name.`)) return;
    let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries] };
    const [cpsNext, target] = findOrCreateCp(next.counterparties, per.name, kind);
    next.counterparties = cpsNext;
    next = mergeCps(next, matches.map((m) => m.id), target.id);
    await save(next);
    flash(`Merged into ${per.name}.`);
  };
  const allSeenNames = useMemo(() => [...new Set([...cps.map((c) => c.name), ...pending.flatMap((s) => s.items.map((it) => it.name))])].sort((a, b) => a.localeCompare(b)), [cps, pending]);

  // ——— username picker (Player data) — every username seen across every site, DL-umbrella style ———
  const usernameGroups = useMemo(() => {
    const sites = [...new Set(siteUsernames.map((u) => u.site))];
    return sites.map((site) => [site, siteUsernames.filter((u) => u.site === site).sort((a, b) => a.name.localeCompare(b.name))]).filter(([, list]) => list.length > 0);
  }, [siteUsernames]);
  const personHasUsername = (per, name, site) => per.name.toLowerCase() === name.toLowerCase() || per.aliases.some((a) => a.name.toLowerCase() === name.toLowerCase() && (!a.site || a.site === site));
  const toggleUsername = (per, name, site) => {
    if (per.name.toLowerCase() === name.toLowerCase()) return;
    const has = per.aliases.some((a) => a.name.toLowerCase() === name.toLowerCase() && (!a.site || a.site === site));
    patchPerson(per.id, { aliases: has ? per.aliases.filter((a) => !(a.name.toLowerCase() === name.toLowerCase() && (!a.site || a.site === site))) : [...per.aliases, { name, site }] });
  };

  // ——— counterparty notes / delete (Balances dropdown) ———
  const findPersonForCp = (cp) => persons.find((per) => (per.kind || "player") === cp.kind && (per.name.toLowerCase() === cp.name.toLowerCase() || per.aliases.some((a) => a.name.toLowerCase() === cp.name.toLowerCase())));
  const setCpNotes = (id, notes) => save({ ...data, counterparties: data.counterparties.map((c) => (c.id === id ? { ...c, notes } : c)) });
  const bundleIntoPlayerData = async (cp) => {
    if (persons.some((per) => (per.kind || "player") === cp.kind && per.name.toLowerCase() === cp.name.toLowerCase())) { flash(`${cp.name} is already in Player Data.`); return; }
    await setPersons([...persons, { id: uid(), name: cp.name, kind: cp.kind, aliases: [], notes: cp.notes || "" }]);
    flash(`Added ${cp.name} to Player Data.`);
  };
  const deleteCpImpact = (cp) => {
    const bal = balances[cp.id] || 0;
    const dealIds = new Set(data.staking.deals.filter((d) => d.cpId === cp.id).map((d) => d.id));
    const entryCount = data.entries.filter((e) => e.cpId === cp.id).length;
    const dealCount = dealIds.size;
    const resultCount = data.staking.results.filter((r) => dealIds.has(r.dealId)).length;
    const impCount = data.staking.imported.filter((it) => it.name.toLowerCase() === cp.name.toLowerCase()).length;
    const parts = [];
    if (entryCount) parts.push(`${entryCount} ledger entr${entryCount === 1 ? "y" : "ies"}`);
    if (dealCount) parts.push(`${dealCount} staking deal${dealCount === 1 ? "" : "s"} (${resultCount} logged result${resultCount === 1 ? "" : "s"})`);
    if (impCount) parts.push(`${impCount} imported staking line${impCount === 1 ? "" : "s"}`);
    return { bal, dealIds, parts };
  };
  const deleteCp = async (cp) => {
    const { dealIds } = deleteCpImpact(cp);
    await save({
      ...data,
      counterparties: data.counterparties.filter((c) => c.id !== cp.id),
      entries: data.entries.filter((e) => e.cpId !== cp.id),
      staking: {
        ...data.staking,
        deals: data.staking.deals.filter((d) => d.cpId !== cp.id),
        results: data.staking.results.filter((r) => !dealIds.has(r.dealId)),
        imported: data.staking.imported.filter((it) => it.name.toLowerCase() !== cp.name.toLowerCase()),
      },
      settleChecklist: (data.settleChecklist || []).filter((x) => x.cpId !== cp.id),
    });
    if (filterCp === cp.id) setFilterCp("");
    setExpandedCp((x) => (x === cp.id ? null : x));
    setConfirmDeleteCp(null);
  };

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;

  const exportAll = () => setExportData({ title: "Tabs ledger", text: toTSV(["Date", "Counterparty", "Kind", "Amount", "Method", "Vig", "Note", "Source"],
    sortDateDesc(allEntries, (e) => allEntries.indexOf(e)).map((e) => [e.date, cpById(e.cpId)?.name || "?", cpById(e.cpId)?.kind || "", e.amount.toFixed(2), e.method || "", e.vig != null ? e.vig.toFixed(2) : "", e.note || "", e.source === "week" ? e.week : e.source])) });

  const kindSection = (kind, title, extraRows) => {
    const list = cps.filter((c) => c.kind === kind).filter((c) => showZero || Math.abs(balances[c.id] || 0) > 0.005 || staking.makeupByCp[c.id]);
    const subtotal = list.reduce((a, c) => a + (balances[c.id] || 0), 0);
    return (
      <Card title={`${title} · ${fmt(subtotal)}`}>
        {list.length === 0 && !extraRows && <div style={{ color: C.mute, fontSize: 13 }}>Nothing open.</div>}
        {list.map((c) => {
          const open = expandedCp === c.id;
          const per = findPersonForCp(c);
          return (
            <div key={c.id} style={{ borderTop: `1px solid ${C.line}` }}>
              <div onClick={() => { setExpandedCp(open ? null : c.id); if (open) setConfirmDeleteCp((x) => (x === c.id ? null : x)); }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 0", fontSize: 13.5, cursor: "pointer" }}>
                <span style={{ color: C.mute, fontSize: 10, width: 10, display: "inline-block", flexShrink: 0 }}>{open ? "▾" : "▸"}</span>
                <span style={{ color: C.ink, fontWeight: 600, fontSize: 13.5 }}>{c.name}</span>
                {iconBtn("✎", (e) => { e.stopPropagation(); renameCp(c); }, C.mute, "Rename (renaming onto an existing name merges)")}
                {staking.makeupByCp[c.id] > 0.005 && <Pill tone="red">in makeup {fmt(staking.makeupByCp[c.id])}</Pill>}
                <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(balances[c.id] || 0)}</span>
                <button onClick={(e) => { e.stopPropagation(); settleUp(c); }} title="Log settling entry" style={{ border: `1px solid ${C.line}`, background: C.surface, color: C.mute, borderRadius: 5, cursor: "pointer", fontSize: 10.5, padding: "2px 8px" }}>settle</button>
              </div>
              {open && (
                <div onClick={(e) => e.stopPropagation()} style={{ padding: "6px 8px 12px 20px", background: C.surface, borderRadius: 6, marginBottom: 8, display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <Btn tone="ghost" small onClick={() => { setFilterCp(c.id); setView("ledger"); }}>View transactions</Btn>
                    {confirmDeleteCp !== c.id && (
                      <button onClick={() => setConfirmDeleteCp(c.id)} style={{ border: `1px solid ${C.red}`, background: "transparent", color: C.red, borderRadius: 6, padding: "5px 12px", fontSize: 12, fontWeight: 700, letterSpacing: "0.03em", cursor: "pointer" }}>Delete this name…</button>
                    )}
                  </div>
                  {confirmDeleteCp === c.id && (() => {
                    const { bal, parts } = deleteCpImpact(c);
                    return (
                      <div style={{ border: `1px solid ${C.red}`, borderRadius: 6, padding: "9px 11px", background: "rgba(200,50,50,0.07)" }}>
                        <div style={{ fontSize: 12.5, color: C.red, fontWeight: 700, marginBottom: 4 }}>Delete "{c.name}" permanently?</div>
                        <div style={{ fontSize: 11.5, color: C.mute, marginBottom: 8 }}>
                          Current balance {fmt(bal)}.{parts.length > 0 ? ` This also removes ${parts.join(", ")}.` : ""} This can't be undone.
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                          <button onClick={() => deleteCp(c)} style={{ border: "none", background: C.red, color: "#fff", borderRadius: 6, padding: "6px 14px", fontSize: 12, fontWeight: 700, letterSpacing: "0.03em", cursor: "pointer" }}>Yes, delete permanently</button>
                          <Btn tone="ghost" small onClick={() => setConfirmDeleteCp(null)}>Cancel</Btn>
                        </div>
                      </div>
                    );
                  })()}
                  {per ? (
                    <div>
                      <div style={{ fontSize: 11, color: C.goldDark, marginBottom: 4 }}>Linked to Player Data · {per.name}{c.name.toLowerCase() !== per.name.toLowerCase() ? ` (alias)` : ""}</div>
                      <textarea value={per.notes || ""} onChange={(e) => patchPerson(per.id, { notes: e.target.value })} placeholder="Notes — payment methods, Telegram / Discord, deal reminders…" rows={2} style={{ ...inputS, width: "100%", boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
                    </div>
                  ) : (
                    <div>
                      <textarea value={c.notes || ""} onChange={(e) => setCpNotes(c.id, e.target.value)} placeholder="Notes — payment methods, communication, etc." rows={2} style={{ ...inputS, width: "100%", boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
                      <div style={{ marginTop: 6 }}><Btn tone="ghost" small onClick={() => bundleIntoPlayerData(c)}>Bundle into Player Data</Btn></div>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
        {extraRows && extraRows.length > 0 && (
          <div style={{ marginTop: list.length ? 10 : 0 }}>
            <div style={{ fontSize: 10.5, letterSpacing: "0.08em", textTransform: "uppercase", color: C.goldDark, fontWeight: 700, marginBottom: 2 }}>Running totals · read-only · not in net position</div>
            {extraRows.map(([label, v]) => (
              <div key={label} style={{ display: "flex", padding: "5px 0", borderTop: `1px solid ${C.line}`, fontSize: 13 }}>
                <span style={{ color: C.mute }}>{label}</span>
                <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(v)}</span>
              </div>
            ))}
          </div>
        )}
      </Card>
    );
  };

  const methodInput = (value, onChange, w = 110) => (
    <>
      <input list="tabs-methods" placeholder="Method" value={value} onChange={(e) => onChange(e.target.value)} style={{ ...inputS, width: w }} />
      <datalist id="tabs-methods">{[...new Set([...METHOD_SUGGESTIONS, ...data.entries.map((e) => e.method).filter(Boolean)])].map((m) => <option key={m} value={m} />)}</datalist>
    </>
  );

  const ledgerRows = sortDateDesc(allEntries.filter((e) => !filterCp || e.cpId === filterCp), (e) => allEntries.indexOf(e));
  const stakingDealsAll = staking.deals.filter((d) => !d.deal.ended);
  const dealsByPlayer = {};
  stakingDealsAll.forEach((d) => { (dealsByPlayer[d.deal.cpId] = dealsByPlayer[d.deal.cpId] || []).push(d); });
  const endedDealsSorted = [...staking.endedDeals].sort((a, b) => (b.deal.endedAt || "").localeCompare(a.deal.endedAt || ""));
  const holderName = (deal) => (cpById(deal.cpId)?.name || "player");
  const miscCats = [...new Set(data.misc.map((e) => e.category || "uncategorized"))].sort();
  const miscRows = (() => {
    let list = data.misc.filter((e) => !miscCatFilter || (e.category || "uncategorized") === miscCatFilter);
    if (miscSort === "category") return [...list].sort((a, b) => (a.category || "").localeCompare(b.category || "") || (b.date || "").localeCompare(a.date || ""));
    return sortDateDesc(list, (e) => data.misc.indexOf(e));
  })();

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap", alignItems: "center" }}>
        {TABS_VIEWS.map(([k, label]) => (
          <button key={k} onClick={() => setView(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: view === k ? C.card : "transparent", color: view === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2 }}>
            {label}
          </button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", paddingBottom: 6 }}>
          <span style={{ fontSize: 13, color: C.mute }}>Net position: <b style={{ color: totalAll >= 0 ? C.green : C.red }}>{fmt(totalAll)}</b></span>
          <Btn tone="gold" small onClick={() => downloadTabsExcel(data, cps, balances, allEntries, staking, vigEntries, vigTotal, miscByCat, miscTotal, otherTotals)}>Download Excel</Btn>
          <Btn tone="ghost" small onClick={exportAll}>Copy</Btn>
        </div>
      </div>
      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {note && <div style={{ background: C.banner, color: C.goldDark, padding: "8px 14px", borderRadius: 6, marginBottom: 14, fontSize: 12.5 }}>{note}</div>}
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 14 }}>
          <b style={{ color: C.green }}>Positive = they owe you</b> · <b style={{ color: C.red }}>negative = you owe them</b>. Weekly settlements from Fish Tank/All American/My Clubs and the settle checklist live in Bookkeeping. Everything else is logged in Ledger.
        </div>

        {/* ═══════════ BOOKKEEPING ═══════════ */}
        {view === "bookkeeping" && (
          <>
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 14 }}>
              <div style={{ color: C.mute, fontSize: 12.5 }}>
                Finished weeks appear below for review — edit a name or amount, or hit × to reject a row, then Accept folds the rest into the balances (once per week, no double-counting). Any line settled by Ak lands on the checklist beneath it — Fish Tank umbrellas, loose SAs, standalone players, and house-backed accounts assigned to Ak; AA SA/Agent/individual/Manager/Master/DL-umbrella lines that are Ak's personally or split-and-collected-by-Ak, and house-backed/staked accounts where Ak holds a share; and every My Clubs umbrella or unbundled player — so nobody gets missed. Ordinary mid-week ledger/staking/misc entries never touch the checklist.
              </div>
              <Btn tone="ghost" small onClick={refreshPending}>↻ Refresh</Btn>
            </div>
            <div style={{ color: C.mute, fontSize: 11.5, marginTop: -8, marginBottom: 14 }}>
              Unlocked, fixed, and relocked a week on Fish Tank/AA/My Clubs? Hit Refresh to pull the corrected numbers in before re-syncing below.
            </div>
            {(() => {
              const open = pending.filter((s) => !data.pushed[s.sourceKey] && s.items.some((it) => Math.abs(it.amount) > 0.005));
              if (!open.length) return <div style={{ color: C.mute, fontSize: 13, marginBottom: 16 }}>No finished weeks waiting for review.</div>;
              return (
                <div style={{ marginBottom: 16 }}>
                  <Card title="Weekly settlements ready to fold in" right={open.length > 1 ? <Btn tone="gold" small onClick={() => acceptSources(open)}>Accept all</Btn> : null}>
                    {open.map((s) => {
                      const rows = aggRows(s).map((r) => rowState(s, r));
                      const liveRows = rows.filter((r) => !r.skip);
                      const skipped = rows.length - liveRows.length;
                      const tot = liveRows.reduce((a, r2_) => a + r2_.amount, 0);
                      return (
                        <div key={s.sourceKey} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6, flexWrap: "wrap" }}>
                            <b style={{ fontSize: 13.5 }}>{s.label}</b>
                            <span style={{ color: C.mute, fontSize: 12 }}>
                              {liveRows.length} counterpart{liveRows.length === 1 ? "y" : "ies"} · net {fmt(tot)}
                              {skipped > 0 && <span style={{ color: C.red }}> · {skipped} rejected</span>}
                              {rows.some((r) => r.edited) && <span style={{ color: C.goldDark }}> · edited</span>}
                              {s.stakingItems?.length > 0 && <span> · {s.stakingItems.length} staking line{s.stakingItems.length > 1 ? "s" : ""} → Staking</span>}
                            </span>
                            <span style={{ marginLeft: "auto" }}>
                              <Btn tone="gold" small disabled={liveRows.length === 0} onClick={() => acceptSources([s])}>{skipped > 0 ? `Accept ${liveRows.length} — update tabs` : "Accept — update tabs"}</Btn>
                            </span>
                          </div>
                          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(310px, 1fr))", gap: "4px 18px" }}>
                            {rows.map((r) => (
                              <div key={r.key} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, padding: "2px 0", opacity: r.skip ? 0.45 : 1 }}>
                                <input value={tweaks[s.sourceKey]?.[r.key]?.name ?? r.orig.name} onChange={(e) => setTweak(s.sourceKey, r.key, { name: e.target.value })} disabled={r.skip}
                                  style={{ ...inputS, padding: "3px 6px", fontSize: 12.5, fontWeight: 600, flex: 1, minWidth: 90, textDecoration: r.skip ? "line-through" : "none", border: r.name !== r.orig.name ? `1px solid ${C.gold}` : `1px solid ${C.line}` }} />
                                <span style={{ color: C.mute, fontSize: 10.5, whiteSpace: "nowrap" }}>{r.kind}{r.orig.n > 1 ? ` ·${r.orig.n}` : ""}</span>
                                <NumInput width={86} value={r.amount} onChange={(v) => setTweak(s.sourceKey, r.key, { amount: v })} />
                                <button onClick={() => setTweak(s.sourceKey, r.key, { skip: !r.skip })} style={{ border: `1px solid ${C.line}`, background: C.surface, color: r.skip ? C.green : C.red, borderRadius: 5, cursor: "pointer", fontSize: 11, padding: "2px 7px", whiteSpace: "nowrap" }}>{r.skip ? "↩ keep" : "×"}</button>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </Card>
                </div>
              );
            })()}

            {(() => {
              const posted = pending.filter((s) => data.pushed[s.sourceKey]);
              if (!posted.length) return null;
              return (
                <div style={{ marginBottom: 16 }}>
                  <button onClick={() => setShowPosted(!showPosted)} style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontSize: 12.5, fontWeight: 700, padding: "4px 0" }}>
                    {showPosted ? "▼" : "►"} Posted weeks ({posted.length}) — re-sync after an unlock/edit/relock
                  </button>
                  {showPosted && (
                    <Card title="Already accepted">
                      <div style={{ fontSize: 12, color: C.mute, marginBottom: 8 }}>
                        These weeks are already in the tabs. If one got unlocked, corrected, and relocked on its site, re-sync it here — that clears the old posted amount first (and the staking rows it fed) so the corrected one replaces it cleanly, with no double-counting.
                      </div>
                      {posted.map((s) => {
                        const rows = aggRows(s).map((r) => rowState(s, r));
                        const liveRows = rows.filter((r) => !r.skip);
                        const tot = liveRows.reduce((a, r2_) => a + r2_.amount, 0);
                        return (
                          <div key={s.sourceKey} style={{ borderTop: `1px solid ${C.line}`, padding: "8px 0", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                            <b style={{ fontSize: 13 }}>{s.label}</b>
                            <span style={{ color: C.mute, fontSize: 12 }}>current numbers: {liveRows.length} counterpart{liveRows.length === 1 ? "y" : "ies"} · net {fmt(tot)}</span>
                            <span style={{ marginLeft: "auto" }}>
                              <Btn tone="ghost" small onClick={() => acceptSources([s], true)}>↻ Re-sync — update tabs</Btn>
                            </span>
                          </div>
                        );
                      })}
                    </Card>
                  )}
                </div>
              );
            })()}

            {(() => {
              const open = settleChecklist.filter((x) => !x.done);
              const done = settleChecklist.filter((x) => x.done);
              return (
                <Card title={`Settle checklist${open.length ? ` · ${open.length} to go` : ""}`} right={done.length > 0 ? <Btn tone="ghost" small onClick={clearDoneChecklist}>Clear checked ({done.length})</Btn> : null}>
                  <div style={{ fontSize: 12, color: C.mute, marginBottom: 8 }}>
                    Only the lines that are Ak's to settle from an accepted weekly settlement land here. Check them off once you've reached out, or × to drop them outright.
                  </div>
                  {settleChecklist.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>Nothing to settle right now — accept a weekly settlement above and names will show up here.</div>}
                  {[...open, ...done].map((x) => {
                    const cp = cpById(x.cpId);
                    return (
                      <div key={x.id} style={{ display: "flex", alignItems: "center", gap: 8, borderTop: `1px solid ${C.line}`, padding: "6px 0", fontSize: 13, opacity: x.done ? 0.5 : 1 }}>
                        <input type="checkbox" checked={!!x.done} onChange={() => toggleChecklistDone(x.id)} />
                        <span style={{ fontWeight: 600, textDecoration: x.done ? "line-through" : "none" }}>{x.name}</span>
                        <span style={{ color: C.mute, fontSize: 11 }}>{x.week}</span>
                        {cp && <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", color: C.mute }}>{money(balances[cp.id] || 0)}</span>}
                        {cp && <button onClick={() => { setFilterCp(cp.id); setView("ledger"); }} style={{ border: `1px solid ${C.line}`, background: C.surface, color: C.mute, borderRadius: 5, cursor: "pointer", fontSize: 10.5, padding: "2px 8px" }}>view</button>}
                        {iconBtn("×", () => removeChecklistItem(x.id), C.red, "Remove from checklist")}
                      </div>
                    );
                  })}
                </Card>
              );
            })()}
          </>
        )}

        {/* ═══════════ BALANCES ═══════════ */}
        {view === "balances" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14, marginBottom: 12 }}>
              {kindSection("player", "Player tabs")}
              {kindSection("club", "Club tabs")}
              {kindSection("other", "Other", otherTotals)}
            </div>
            <label style={{ fontSize: 12.5, color: C.mute, display: "flex", gap: 6, alignItems: "center" }}>
              <input type="checkbox" checked={showZero} onChange={(e) => setShowZero(e.target.checked)} /> show zero balances
            </label>
          </>
        )}

        {/* ═══════════ LEDGER ═══════════ */}
        {view === "ledger" && (
          <>
            <div style={{ marginBottom: 16 }}>
              <Card title="Add entry" right={
                <span style={{ display: "flex", gap: 2, background: C.surface, borderRadius: 7, padding: 2 }}>
                  {[["entry", "Entry"], ["swap", "Swap"]].map(([k, label]) => (
                    <button key={k} onClick={() => setAddMode(k)} style={{
                      border: "none", cursor: "pointer", padding: "5px 12px", fontSize: 12, fontWeight: 700,
                      background: addMode === k ? C.gold : "transparent", color: addMode === k ? "var(--onGold)" : C.mute,
                      borderRadius: 5 }}>{label}</button>
                  ))}
                </span>}>
                {addMode === "entry" ? (
                  <>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                      <select value={f.kind} onChange={(e) => setF({ ...f, kind: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
                        <option value="player">Player</option><option value="club">Club</option><option value="other">Other</option>
                      </select>
                      <input list="tabs-cp-names" placeholder="Who…" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} style={{ ...inputS, width: 170 }} />
                      <datalist id="tabs-cp-names">{cps.filter((c) => c.kind === f.kind).map((c) => <option key={c.id} value={c.name} />)}</datalist>
                      <NumInput width={100} value={f.amount} onChange={(v) => setF({ ...f, amount: v })} />
                      {methodInput(f.method, (m) => setF({ ...f, method: m }))}
                      {isCrypto(f.method) && <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, color: C.mute }}>vig <PctInput width={52} value={fRate} onChange={(v) => setF({ ...f, rate: v })} />
                        {fPreview && <span>→ tab {money(fPreview.amount)} · vig <b style={{ color: fPreview.vig >= 0 ? C.green : C.red }}>{fmt(fPreview.vig)}</b></span>}</span>}
                      <input placeholder="Note" value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} style={{ ...inputS, flex: 1, minWidth: 160 }} />
                      {dateInput(f.date, (v) => setF({ ...f, date: v }))}
                      <Btn tone="gold" small onClick={addEntry}>Add</Btn>
                    </div>
                    <div style={{ fontSize: 11.5, color: C.mute, marginTop: 6 }}>Amount = the money that moved (positive: you paid out / they owe you; negative: they paid you). Method "crypto" logs the vig automatically — you sending: they're credited amount × (1 + rate), vig gain = amount × rate. You receiving: you're only credited amount ÷ (1 + rate) since the vig eats into what lands, vig loss = amount owed − credited amount.</div>
                  </>
                ) : (
                  <>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                      <span style={{ fontSize: 11, color: C.mute, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>From</span>
                      <select value={sw.fromKind} onChange={(e) => setSw({ ...sw, fromKind: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
                        <option value="player">Player</option><option value="club">Club</option><option value="other">Other</option>
                      </select>
                      <input list="tabs-swap-from-names" placeholder="Who's sending…" value={sw.from} onChange={(e) => setSw({ ...sw, from: e.target.value })} style={{ ...inputS, width: 150 }} />
                      <datalist id="tabs-swap-from-names">{cps.filter((c) => c.kind === sw.fromKind).map((c) => <option key={c.id} value={c.name} />)}</datalist>
                      <span style={{ color: C.mute }}>→</span>
                      <span style={{ fontSize: 11, color: C.mute, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>To</span>
                      <select value={sw.toKind} onChange={(e) => setSw({ ...sw, toKind: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
                        <option value="player">Player</option><option value="club">Club</option><option value="other">Other</option>
                      </select>
                      <input list="tabs-swap-to-names" placeholder="Who's receiving…" value={sw.to} onChange={(e) => setSw({ ...sw, to: e.target.value })} style={{ ...inputS, width: 150 }} />
                      <datalist id="tabs-swap-to-names">{cps.filter((c) => c.kind === sw.toKind).map((c) => <option key={c.id} value={c.name} />)}</datalist>
                      <NumInput width={100} value={sw.amount} onChange={(v) => setSw({ ...sw, amount: v })} />
                      {methodInput(sw.method, (m) => setSw({ ...sw, method: m }))}
                      <input placeholder="Note" value={sw.note} onChange={(e) => setSw({ ...sw, note: e.target.value })} style={{ ...inputS, flex: 1, minWidth: 160 }} />
                      {dateInput(sw.date, (v) => setSw({ ...sw, date: v }))}
                      <Btn tone="gold" small onClick={addSwap} disabled={!sw.from.trim() || !sw.to.trim() || !sw.amount}>Swap</Btn>
                    </div>
                    <div style={{ fontSize: 11.5, color: C.mute, marginTop: 6 }}>Settle one tab against another with no cash moving — From and To can each be a player, club, or other counterparty. From's tab is reduced by the amount, To's tab is increased by the amount. E.g. a player owes you 5k and you owe a club 7k: swap 5k and the player is settled while the club drops to 2k owed. Logged as two linked lines (deleting one removes both).</div>
                  </>
                )}
              </Card>
            </div>
            <div style={{ background: C.card, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
              <div style={{ display: "flex", gap: 10, alignItems: "center", padding: "10px 16px", background: C.cream, flexWrap: "wrap" }}>
                <span style={{ fontWeight: 700 }}>Ledger</span>
                <select value={filterCp} onChange={(e) => setFilterCp(e.target.value)} style={{ ...inputS, fontSize: 12.5 }}>
                  <option value="">Everyone</option>
                  {cps.map((c) => <option key={c.id} value={c.id}>{c.name} ({c.kind})</option>)}
                </select>
                {filterCp && <span style={{ fontSize: 13 }}>balance {money(balances[filterCp] || 0)}{staking.makeupByCp[filterCp] > 0.005 && <span style={{ marginLeft: 8 }}><Pill tone="red">in makeup {fmt(staking.makeupByCp[filterCp])}</Pill></span>}</span>}
                {undoStack.length > 0 && (
                  <button onClick={undoLastDelete} title={`Restore: ${undoStack[undoStack.length - 1].label}`} style={{ border: `1px solid ${C.gold}`, background: C.surface, color: C.goldDark, borderRadius: 6, cursor: "pointer", fontSize: 11.5, fontWeight: 700, padding: "4px 10px" }}>
                    ↺ Undo delete{undoStack.length > 1 ? ` (${undoStack.length})` : ""}
                  </button>
                )}
                <span style={{ marginLeft: "auto", fontSize: 11.5, color: C.mute }}>newest first · click ✎ to edit any line</span>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", minWidth: 980, borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: C.cream }}>
                    <th style={{ ...th, textAlign: "left" }}>Date</th><th style={{ ...th, textAlign: "left" }}>Who</th>
                    <th style={th}>Amount</th><th style={{ ...th, textAlign: "left" }}>Method</th><th style={th}>Vig</th><th style={{ ...th, textAlign: "left" }}>Note</th>
                    <th style={{ ...th, textAlign: "left" }}>Source</th><th style={th}></th>
                  </tr></thead>
                  <tbody>
                    {ledgerRows.length === 0 && <tr><td colSpan={8} style={{ ...tdL, color: C.mute, padding: 16 }}>No entries yet.</td></tr>}
                    {ledgerRows.map((e, i) => {
                      const cp = cpById(e.cpId);
                      const editing = editId === e.id;
                      if (editing) return (
                        <tr key={e.id} style={{ background: C.banner, borderTop: `1px solid ${C.line}` }}>
                          <td style={tdL}>{dateInput(editDraft.date, (v) => setEditDraft({ ...editDraft, date: v }), 130)}</td>
                          <td style={tdL}>
                            <select value={editDraft.cpId} onChange={(ev) => setEditDraft({ ...editDraft, cpId: ev.target.value })} style={{ ...inputS, fontSize: 12 }}>
                              {cps.map((c) => <option key={c.id} value={c.id}>{c.name} ({c.kind})</option>)}
                            </select>
                          </td>
                          <td style={td}><NumInput width={96} value={editDraft.amount} onChange={(v) => setEditDraft({ ...editDraft, amount: v })} />{isCrypto(editDraft.method) && <div style={{ fontSize: 10, color: C.mute }}>base amount</div>}</td>
                          <td style={tdL}>{methodInput(editDraft.method, (m) => setEditDraft({ ...editDraft, method: m }), 100)}</td>
                          <td style={td}>{isCrypto(editDraft.method) ? <PctInput width={52} value={editDraft.rate} onChange={(v) => setEditDraft({ ...editDraft, rate: v })} /> : "—"}</td>
                          <td style={tdL}><input value={editDraft.note} onChange={(ev) => setEditDraft({ ...editDraft, note: ev.target.value })} style={{ ...inputS, width: "100%", minWidth: 200, boxSizing: "border-box" }} /></td>
                          <td style={{ ...tdL, color: C.mute, fontSize: 11.5 }}>{e.source === "week" ? e.week : e.source === "swap" ? `swap · ${cpById(e.swapWith)?.name || "?"}` : "manual"}</td>
                          <td style={{ ...td, whiteSpace: "nowrap" }}><Btn tone="gold" small onClick={commitEdit}>Save</Btn> {iconBtn("cancel", () => setEditId(null), C.mute)}</td>
                        </tr>
                      );
                      return (
                        <tr key={e.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                          <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{e.date}</td>
                          <td style={{ ...tdL, fontWeight: 600 }}>{cp?.name || "?"} <span style={{ color: C.mute, fontWeight: 400, fontSize: 11 }}>({cp?.kind})</span></td>
                          <td style={td}>{money(e.amount)}{e.baseAmount != null && Math.abs(e.baseAmount - e.amount) > 0.005 && <div style={{ fontSize: 10, color: C.mute }}>sent {fmt(e.baseAmount)}</div>}</td>
                          <td style={{ ...tdL, fontSize: 12 }}>{e.method ? <Pill tone={isCrypto(e.method) ? "gold" : "blue"}>{e.method}</Pill> : <span style={{ color: C.mute }}>—</span>}</td>
                          <td style={{ ...td, fontSize: 12 }}>{e.vig != null ? <span style={{ color: e.vig >= 0 ? C.green : C.red }}>{fmt(e.vig)} <span style={{ color: C.mute }}>@{e.vigRate}%</span></span> : ""}</td>
                          <td style={{ ...tdL, fontSize: 12.5, whiteSpace: "normal", minWidth: 220 }}>{e.note || ""}</td>
                          <td style={{ ...tdL, color: C.mute, fontSize: 11.5 }}>{e.source === "week" ? e.week : e.source === "staking" ? <span title="Produced by a staking result — edit it in Staking">staking</span> : e.source === "swap" ? <span title="Linked swap — deleting removes both sides">swap · {cpById(e.swapWith)?.name || "?"}</span> : "manual"}</td>
                          <td style={{ ...td, width: 60, whiteSpace: "nowrap" }}>
                            {e.source === "staking" ? iconBtn("→", () => { setView("staking"); }, C.goldDark, "Edit in Staking") : <>{iconBtn("✎", () => startEdit(e), C.goldDark, "Edit")}{iconBtn("×", () => delEntry(e.id), C.red, "Delete")}</>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {/* ═══════════ VIG ═══════════ */}
        {view === "vig" && (
          <>
            <div style={{ background: C.bar, borderRadius: 12, padding: "18px 24px", marginBottom: 16, color: "var(--barText)", display: "flex", gap: 30, alignItems: "center", flexWrap: "wrap" }}>
              <div><div style={{ fontSize: 11, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--barGold)" }}>Crypto vig — running total</div>
                <div style={{ fontFamily: "Georgia, serif", fontSize: 26, color: vigTotal >= 0 ? "var(--barGreen)" : "var(--barRed)" }}>{fmt(vigTotal)}</div></div>
              <div style={{ fontSize: 12.5 }}>Gains {fmt(vigEntries.filter((e) => e.vig > 0).reduce((a, e) => a + e.vig, 0))} · losses {fmt(vigEntries.filter((e) => e.vig < 0).reduce((a, e) => a + e.vig, 0))} · {vigEntries.length} transaction{vigEntries.length !== 1 ? "s" : ""}</div>
              <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>Default rate <PctInput width={56} value={data.vigRate ?? 5} onChange={(v) => v != null && save({ ...data, vigRate: v })} /> <span style={{ color: "var(--barSubtle)" }}>(per-transaction rate is adjustable below)</span></div>
            </div>
            <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>Every ledger entry with method <b>crypto</b> lands here. Positive amount (you → them) = vig gain; negative (them → you) = vig loss. Change a rate here and the tab amount recomputes; edit the base amount, note, or date from the Ledger.</div>
            <div style={{ background: C.card, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", minWidth: 820, borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: C.cream }}>
                    <th style={{ ...th, textAlign: "left" }}>Date</th><th style={{ ...th, textAlign: "left" }}>Who</th><th style={th}>Amount moved</th><th style={th}>Rate</th><th style={th}>Vig gain / loss</th><th style={th}>Tab effect</th><th style={{ ...th, textAlign: "left" }}>Note</th><th style={th}></th>
                  </tr></thead>
                  <tbody>
                    {vigEntries.length === 0 && <tr><td colSpan={8} style={{ ...tdL, color: C.mute, padding: 16 }}>No crypto transactions yet — add one in Ledger with method "crypto".</td></tr>}
                    {sortDateDesc(vigEntries, (e) => data.entries.indexOf(e)).map((e, i) => (
                      <tr key={e.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                        <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{e.date}</td>
                        <td style={{ ...tdL, fontWeight: 600 }}>{cpById(e.cpId)?.name || "?"}</td>
                        <td style={td}>{money(e.baseAmount ?? e.amount)}</td>
                        <td style={td}><PctInput width={52} value={e.vigRate ?? data.vigRate ?? 5} onChange={(v) => v != null && setEntryRate(e.id, v)} /></td>
                        <td style={td}><b style={{ color: (e.vig || 0) >= 0 ? C.green : C.red }}>{fmt(e.vig || 0)}</b> <span style={{ fontSize: 10.5, color: C.mute }}>{(e.vig || 0) >= 0 ? "gain" : "loss"}</span></td>
                        <td style={td}>{money(e.amount)}</td>
                        <td style={{ ...tdL, fontSize: 12.5, whiteSpace: "normal" }}>{e.note}</td>
                        <td style={{ ...td, whiteSpace: "nowrap" }}>{iconBtn("✎", () => { setView("ledger"); setFilterCp(e.cpId); startEdit(e); }, C.goldDark, "Edit in Ledger")}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {/* ═══════════ STAKING ═══════════ */}
        {view === "staking" && (
          <>
            <div style={{ background: C.bar, borderRadius: 12, padding: "18px 24px", marginBottom: 16, color: "var(--barText)", display: "flex", gap: 34, alignItems: "center", flexWrap: "wrap" }}>
              {[["Net action buy", staking.totals.netActionBuy], ["Total chopped on stake", staking.totals.chopped], ["Total makeup", staking.totals.makeup]].map(([l, v]) => (
                <div key={l}><div style={{ fontSize: 11, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--barGold)" }}>{l}</div>
                  <div style={{ fontFamily: "Georgia, serif", fontSize: 24, color: l === "Total makeup" ? "var(--barGold)" : v >= 0 ? "var(--barGreen)" : "var(--barRed)" }}>{fmt(v)}</div></div>
              ))}
              <div style={{ marginLeft: "auto", fontSize: 12, color: "var(--barSubtle)", maxWidth: 420 }}>Action buy: your share of P&L. Makeup: losses accrue to the player's makeup; wins clear it, the excess is chopped and your cut adds up here. Fish Tank / owner-club house-backed lines flow in automatically when a week is accepted (your share only) — unless a backed player is linked to a makeup deal below ("Unified deal" on their backed row), in which case their full weekly net feeds into that deal instead (dated to when it's accepted), combining every linked site — plus anything logged manually — into one shared, ongoing makeup pool + tab. Keep that deal on "settle per session" so each result updates the running makeup as it comes in.</div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 14, marginBottom: 16 }}>
              <Card title="New deal">
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                  <input list="tabs-player-names" placeholder="Player…" value={sd.name} onChange={(e) => setSd({ ...sd, name: e.target.value })} style={{ ...inputS, width: 150 }} />
                  <datalist id="tabs-player-names">{allSeenNames.map((n) => <option key={n} value={n} />)}</datalist>
                  <select value={sd.type} onChange={(e) => setSd({ ...sd, type: e.target.value, pct: e.target.value === "makeup" ? 100 : 50 })} style={{ ...inputS, fontSize: 12.5 }}>
                    <option value="action">Action buy</option><option value="makeup">Makeup deal</option>
                  </select>
                  <span style={{ fontSize: 12, color: C.mute }}>% staked</span><PctInput width={56} value={sd.pct} onChange={(v) => setSd({ ...sd, pct: v ?? 0 })} />
                  {sd.type === "makeup" && <>
                    <span style={{ fontSize: 12, color: C.mute }}>Ak chop %</span><PctInput width={56} value={sd.akChopPct} onChange={(v) => setSd({ ...sd, akChopPct: v ?? 0 })} />
                    <select value={sd.cadence} onChange={(e) => setSd({ ...sd, cadence: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
                      <option value="session">settle per session</option><option value="weekly">settle weekly (Mon–Sun)</option>
                    </select>
                    <span style={{ fontSize: 12, color: C.mute }}>makeup now</span><NumInput width={90} value={sd.makeupStart} onChange={(v) => setSd({ ...sd, makeupStart: v })} />
                  </>}
                  <input placeholder="Deal note" value={sd.note} onChange={(e) => setSd({ ...sd, note: e.target.value })} style={{ ...inputS, flex: 1, minWidth: 120 }} />
                  <Btn tone="gold" small onClick={addDeal}>Add deal</Btn>
                </div>
              </Card>
              <Card title="Log a result">
                {stakingDealsAll.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>Create a deal first.</div> : (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                    <select value={sr.dealId} onChange={(e) => { const d = stakingDealsAll.find((x) => x.deal.id === e.target.value); setSr({ ...sr, dealId: e.target.value, pct: "" }); }} style={{ ...inputS, fontSize: 12.5, maxWidth: 220 }}>
                      <option value="">— deal —</option>
                      {stakingDealsAll.map((d) => <option key={d.deal.id} value={d.deal.id}>{d.name} · {d.deal.type === "action" ? `action ${d.deal.pct}%` : `makeup ${d.deal.pct}%`}</option>)}
                    </select>
                    {dateInput(sr.date, (v) => setSr({ ...sr, date: v }))}
                    <input placeholder="Game (live 5/10, cGG, …)" value={sr.game} onChange={(e) => setSr({ ...sr, game: e.target.value })} style={{ ...inputS, width: 150 }} />
                    <span style={{ fontSize: 12, color: C.mute }}>P&L</span><NumInput width={96} value={sr.pnl} onChange={(v) => setSr({ ...sr, pnl: v })} />
                    <span style={{ fontSize: 12, color: C.mute }}>%</span><PctInput width={56} value={sr.pct === "" ? (stakingDealsAll.find((d) => d.deal.id === sr.dealId)?.deal.pct ?? "") : sr.pct} onChange={(v) => setSr({ ...sr, pct: v ?? "" })} />
                    <select value={sr.holder} onChange={(e) => setSr({ ...sr, holder: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
                      <option value="ak">held by AK</option><option value="player">held by {sr.dealId ? holderName(stakingDealsAll.find((d) => d.deal.id === sr.dealId)?.deal || {}) : "player"}</option>
                    </select>
                    <input placeholder="Note" value={sr.note} onChange={(e) => setSr({ ...sr, note: e.target.value })} style={{ ...inputS, flex: 1, minWidth: 100 }} />
                    <Btn tone="gold" small onClick={addResult} disabled={!sr.dealId}>Log</Btn>
                  </div>
                )}
                <div style={{ fontSize: 11.5, color: C.mute, marginTop: 6 }}>Net = P&L × %. "Held by" = who is holding the money after the session; the player's tab moves by what he holds minus what he's owed, so wins and losses land right either way.</div>
              </Card>
            </div>

            {stakingDealsAll.length === 0 && staking.imported.length === 0 && <Card title="No staking deals yet"><div style={{ color: C.mute, fontSize: 13 }}>Add a manual deal above, or accept a Fish Tank / owner-club week that has house-backed players.</div></Card>}

            {Object.entries(dealsByPlayer).map(([cpId, deals]) => (
              <div key={cpId} style={{ marginBottom: 14 }}>
                <Card title={cpById(cpId)?.name || "?"} right={
                  <span style={{ display: "flex", gap: 14, fontSize: 12.5, alignItems: "center" }}>
                    {deals.some((d) => d.deal.type === "action") && <span>Net action buy <b>{money(r2(deals.reduce((a, d) => a + d.netActionBuy, 0)))}</b></span>}
                    {deals.some((d) => d.deal.type === "makeup") && <><span>Total chopped <b>{money(r2(deals.reduce((a, d) => a + d.chopped, 0)))}</b></span><span>Makeup <b style={{ color: deals.some((d) => d.inMakeup) ? C.red : C.green }}>{fmt(r2(deals.reduce((a, d) => a + d.makeup, 0)))}</b></span></>}
                    <span>Tab <b>{money(balances[cpId] || 0)}</b></span>
                  </span>}>
                  {deals.map((d) => (
                    <div key={d.deal.id} style={{ borderTop: `1px solid ${C.line}`, padding: "8px 0" }}>
                      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", fontSize: 12.5 }}>
                        <button onClick={() => setOpenDeal({ ...openDeal, [d.deal.id]: !openDeal[d.deal.id] })} style={{ border: "none", background: "none", cursor: "pointer", color: C.goldDark, fontSize: 12, padding: 0 }}>{openDeal[d.deal.id] ? "▼" : "►"}</button>
                        <Pill tone={d.deal.type === "action" ? "blue" : "red"}>{d.deal.type === "action" ? "action buy" : "makeup"}</Pill>
                        <span>% staked <PctInput width={52} value={d.deal.pct} onChange={(v) => v != null && setDeal(d.deal.id, { pct: v })} /></span>
                        {d.deal.type === "makeup" && <>
                          <span>Ak chop <PctInput width={52} value={d.deal.akChopPct ?? 50} onChange={(v) => v != null && setDeal(d.deal.id, { akChopPct: v })} /></span>
                          <select value={d.deal.cadence || "session"} onChange={(e) => setDeal(d.deal.id, { cadence: e.target.value })} style={{ ...inputS, fontSize: 11.5, padding: "3px 6px" }}><option value="session">per session</option><option value="weekly">weekly</option></select>
                          <span>start makeup <NumInput width={80} value={d.deal.makeupStart || 0} onChange={(v) => setDeal(d.deal.id, { makeupStart: v })} /></span>
                          <span>now <b style={{ color: d.inMakeup ? C.red : C.green }}>{fmt(d.makeup)}</b></span>
                          <span>chopped <b>{money(d.chopped)}</b></span>
                        </>}
                        {d.deal.type === "action" && <span>net action buy <b>{money(d.netActionBuy)}</b></span>}
                        <input value={d.deal.note || ""} placeholder="note" onChange={(e) => setDeal(d.deal.id, { note: e.target.value })} style={{ ...inputS, fontSize: 11.5, width: 160 }} />
                        {d.deal.type === "makeup" && (
                          <button onClick={() => { try { navigator.clipboard.writeText(d.deal.id); } catch (e) {} flash("Deal ID copied — paste into a backed player's \"Unified deal\" field on Fish Tank / a club to fold their weekly net in here."); }}
                            title={d.deal.id} style={{ border: `1px solid ${C.line}`, borderRadius: 5, background: "none", color: C.mute, cursor: "pointer", fontSize: 10.5, fontFamily: "monospace", padding: "2px 6px" }}>
                            id …{d.deal.id.slice(-6)} ⧉
                          </button>
                        )}
                        <span style={{ marginLeft: "auto", color: C.mute }}>{d.rows.length} result{d.rows.length !== 1 ? "s" : ""}</span>
                        <button onClick={() => endStake(d.deal.id)} title="Close this deal out and log its lifetime result to Ended Stakes" style={{ border: `1px solid ${C.line}`, borderRadius: 5, background: "none", color: C.goldDark, cursor: "pointer", fontSize: 11, padding: "2px 7px" }}>End stake</button>
                        {iconBtn("× delete deal", () => delDeal(d.deal.id), C.red)}
                      </div>
                      {openDeal[d.deal.id] && (
                        <div style={{ overflowX: "auto", marginTop: 6 }}>
                          <table style={{ width: "100%", minWidth: 900, borderCollapse: "collapse" }}>
                            <thead><tr style={{ background: C.cream }}>
                              <th style={{ ...th, textAlign: "left" }}>Date</th><th style={{ ...th, textAlign: "left" }}>Game</th><th style={th}>P&L</th><th style={th}>%</th><th style={th}>Net</th><th style={{ ...th, textAlign: "left" }}>Holder</th>
                              {d.deal.type === "makeup" && <><th style={th}>Makeup after</th><th style={th}>Chop (Ak)</th><th style={th}>Chop (player)</th></>}
                              <th style={th}>Tab effect</th><th style={{ ...th, textAlign: "left" }}>Note</th><th style={th}></th>
                            </tr></thead>
                            <tbody>
                              {d.rows.length === 0 && <tr><td colSpan={12} style={{ ...tdL, color: C.mute, padding: 10 }}>No results yet.</td></tr>}
                              {[...d.rows].reverse().map((r, i) => editRes === r.id ? (
                                <tr key={r.id} style={{ background: C.banner, borderTop: `1px solid ${C.line}` }}>
                                  <td style={tdL}>
                                    {dateInput(resDraft.date, (v) => setResDraft({ ...resDraft, date: v }), 130)}
                                    {d.deal.type === "makeup" && (d.deal.cadence || "session") === "weekly" && (
                                      <div style={{ marginTop: 4 }}>
                                        {dateInput(resDraft.groupWeek || "", (v) => setResDraft({ ...resDraft, groupWeek: v }), 130)}
                                        <div style={{ fontSize: 9.5, color: C.mute, marginTop: 1 }}>group under wk of… (blank = auto from date)</div>
                                      </div>
                                    )}
                                  </td>
                                  <td style={tdL}><input value={resDraft.game} onChange={(e) => setResDraft({ ...resDraft, game: e.target.value })} style={{ ...inputS, width: 130 }} /></td>
                                  <td style={td}><NumInput width={90} value={resDraft.pnl} onChange={(v) => setResDraft({ ...resDraft, pnl: v })} /></td>
                                  <td style={td}><PctInput width={52} value={resDraft.pct} onChange={(v) => setResDraft({ ...resDraft, pct: v })} /></td>
                                  <td style={td}>—</td>
                                  <td style={tdL}><select value={resDraft.holder} onChange={(e) => setResDraft({ ...resDraft, holder: e.target.value })} style={{ ...inputS, fontSize: 12 }}><option value="ak">AK</option><option value="player">{d.name}</option></select></td>
                                  {d.deal.type === "makeup" && <td colSpan={3} style={td}>—</td>}
                                  <td style={td}>—</td>
                                  <td style={tdL}><input value={resDraft.note} onChange={(e) => setResDraft({ ...resDraft, note: e.target.value })} style={{ ...inputS, width: 140 }} /></td>
                                  <td style={{ ...td, whiteSpace: "nowrap" }}><Btn tone="gold" small onClick={commitRes}>Save</Btn> {iconBtn("cancel", () => setEditRes(null), C.mute)}</td>
                                </tr>
                              ) : (
                                <tr key={r.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                                  <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{r.date}{r.weekKey && <div style={{ fontSize: 10 }}>wk {r.weekKey}{r.groupWeek && <span title="Manually grouped — not derived from this result's own date" style={{ color: C.goldDark, fontWeight: 700 }}> · pinned</span>}</div>}</td>
                                  <td style={tdL}>{r.game || "—"}</td>
                                  <td style={td}>{money(+r.pnl || 0)}</td>
                                  <td style={td}>{r.pct}%</td>
                                  <td style={td}><b>{money(r.net)}</b>{r.weekKey && r.weekNet != null && r.tab != null && <div style={{ fontSize: 10, color: C.mute }}>week {fmt(r.weekNet)}</div>}</td>
                                  <td style={{ ...tdL, fontSize: 12 }}>{r.holder === "player" ? d.name : "AK"}</td>
                                  {d.deal.type === "makeup" && <>
                                    <td style={td}>{r.makeupAfter != null ? fmt(r.makeupAfter) : <span style={{ color: C.mute }}>·</span>}</td>
                                    <td style={td}>{r.akChop != null ? (r.akChop ? money(r.akChop) : "—") : <span style={{ color: C.mute }}>·</span>}</td>
                                    <td style={td}>{r.playerChop != null ? (r.playerChop ? fmt(r.playerChop) : "—") : <span style={{ color: C.mute }}>·</span>}</td>
                                  </>}
                                  <td style={td}>{r.tab != null ? (Math.abs(r.tab) > 0.005 ? money(r.tab) : "—") : <span style={{ color: C.mute }}>·</span>}</td>
                                  <td style={{ ...tdL, fontSize: 12, whiteSpace: "normal" }}>{r.note}</td>
                                  <td style={{ ...td, whiteSpace: "nowrap" }}>{iconBtn("✎", () => { setEditRes(r.id); setResDraft({ date: r.date, groupWeek: r.groupWeek || "", game: r.game || "", pnl: r.pnl, pct: r.pct, holder: r.holder, note: r.note || "" }); }, C.goldDark)}{iconBtn("×", () => delResult(r.id), C.red)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                          {d.deal.type === "makeup" && (d.deal.cadence || "session") === "weekly" && <div style={{ fontSize: 11, color: C.mute, marginTop: 4 }}>Weekly cadence: makeup / chop and the tab effect are computed on each Mon–Sun week's net and shown on the week's last row. Editing a result (✎) lets you pin it to a specific week regardless of its own date — handy when correcting an old week's numbers, since a re-synced auto result is dated the day of the fix, not the day the game happened.</div>}
                        </div>
                      )}
                    </div>
                  ))}
                </Card>
              </div>
            ))}

            {staking.endedDeals.length > 0 && (
              <Card title="Ended stakes" right={
                <span style={{ display: "flex", gap: 10, alignItems: "center", fontSize: 12.5 }}>
                  <span>Lifetime total <b style={{ color: (endedIncludeAction ? staking.endedTotals.all : staking.endedTotals.makeupOnly) >= 0 ? C.green : C.red }}>{fmt(endedIncludeAction ? staking.endedTotals.all : staking.endedTotals.makeupOnly)}</b></span>
                  <button onClick={() => setEndedIncludeAction((v) => !v)} title="Action buys rarely get formally ended — they just run on. This adds the running net action buy off every currently-open manual action-buy deal above (not gated on that player also having ended a makeup stake), on top of the ended-makeup total. House-backed imports are never included." style={{ border: `1px solid ${C.line}`, borderRadius: 5, background: endedIncludeAction ? C.cream : "none", color: C.mute, cursor: "pointer", fontSize: 11, padding: "2px 7px" }}>
                    {endedIncludeAction ? "+ live action buy" : "makeup only"}
                  </button>
                </span>}>
                <div style={{ fontSize: 11.5, color: C.mute, marginBottom: 8 }}>Closed-out deals, logged separately from the ongoing totals above. A makeup deal's result is total chopped minus outstanding makeup at the time it ended (the write-off, if it ended in the red); an action-buy deal's result is its lifetime net action buy.</div>
                {endedDealsSorted.map((d) => (
                  <div key={d.deal.id} style={{ borderTop: `1px solid ${C.line}`, padding: "8px 0", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", fontSize: 12.5 }}>
                    <Pill tone={d.deal.type === "action" ? "blue" : "red"}>{d.deal.type === "action" ? "action buy" : "makeup"}</Pill>
                    <b>{d.name}</b>
                    <span style={{ color: C.mute }}>ended {d.deal.endedAt}</span>
                    {d.deal.type === "makeup" && <span style={{ color: C.mute }}>chopped {money(d.deal.endChopped)} − makeup {money(d.deal.endMakeup)}</span>}
                    <Pill tone={d.deal.endOutcome === "win" ? "green" : "red"}>{d.deal.endOutcome} {money(d.deal.endAmount)}</Pill>
                    <span style={{ marginLeft: "auto" }} />
                    <button onClick={() => reopenStake(d.deal.id)} title="Move this deal back to ongoing and remove its ended-stakes entry" style={{ border: `1px solid ${C.line}`, borderRadius: 5, background: "none", color: C.mute, cursor: "pointer", fontSize: 11, padding: "2px 7px" }}>Reopen</button>
                    {iconBtn("× delete deal", () => delDeal(d.deal.id), C.red)}
                  </div>
                ))}
                {endedIncludeAction && staking.liveActionBuyRows.map((d) => (
                  <div key={"live-" + d.deal.id} style={{ borderTop: `1px dashed ${C.line}`, padding: "8px 0", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", fontSize: 12.5 }}>
                    <Pill tone="blue">live action buy</Pill>
                    <b>{d.name}</b>
                    <span style={{ color: C.mute }}>still ongoing — % staked {d.deal.pct}, manage it above</span>
                    <Pill tone={d.netActionBuy >= 0 ? "green" : "red"}>{money(d.netActionBuy)}</Pill>
                  </div>
                ))}
              </Card>
            )}

            {staking.imported.length > 0 && (
              <Card title="House-backed on the sites (your share) — imported weekly" right={<Pill tone="gold">Fish Tank · owner clubs</Pill>}>
                <div style={{ fontSize: 12, color: C.mute, marginBottom: 8 }}>Settlement math happens on each site; these are your book shares folded in when the week is accepted. Makeup shown is the live balance from the site's deal (your share).</div>
                {staking.imported.map((x) => (
                  <div key={x.name} style={{ borderTop: `1px solid ${C.line}`, padding: "8px 0" }}>
                    <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap", fontSize: 13 }}>
                      <b>{x.name}</b>
                      {x.rows.some((r) => r.kind === "action") && <span>Net action buy <b>{money(x.netActionBuy)}</b></span>}
                      {(x.rows.some((r) => r.kind === "makeup") || (x.live || []).length > 0) && <><span>Total chopped <b>{money(x.chopped)}</b></span><span>Makeup <b style={{ color: x.makeup > 0.005 ? C.red : C.green }}>{fmt(x.makeup)}</b></span></>}
                      {(x.live || []).map((l) => <span key={l.site + l.rawName} style={{ color: C.mute, fontSize: 11.5 }}>{l.site} · {l.rawName} · {l.rb}{l.share !== 1 ? ` · ${Math.round(l.share * 100)}% share` : ""}</span>)}
                    </div>
                    {x.rows.length > 0 && (
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: "2px 16px", marginTop: 4 }}>
                        {[...x.rows].reverse().map((r) => (
                          <div key={r.id} style={{ fontSize: 12, color: C.mute, display: "flex", gap: 6, alignItems: "center" }}>
                            <span>{r.sourceKey.split(":").slice(1).join(":") || r.date}</span>
                            <span>{r.site}</span>
                            <Pill tone={r.kind === "action" ? "blue" : "red"}>{r.kind}</Pill>
                            <span style={{ marginLeft: "auto", color: C.ink }}>{r.kind === "action" ? money(r.net) : <>chop {money(r.chopped || 0)}{r.makeupAfter != null ? ` · mk ${fmt(r.makeupAfter)}` : ""}</>}</span>
                            {iconBtn("×", () => delImported(r.id), C.red, "Remove")}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </Card>
            )}
          </>
        )}

        {/* ═══════════ MISC ═══════════ */}
        {view === "misc" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 14, marginBottom: 16, alignItems: "start" }}>
              <Card title="Log misc. P&L">
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                  {dateInput(mf.date, (v) => setMf({ ...mf, date: v }))}
                  <NumInput width={100} value={mf.amount} onChange={(v) => setMf({ ...mf, amount: v })} />
                  <input list="tabs-misc-cats" placeholder="Category" value={mf.category} onChange={(e) => setMf({ ...mf, category: e.target.value })} style={{ ...inputS, width: 150 }} />
                  <datalist id="tabs-misc-cats">{miscCats.map((c) => <option key={c} value={c} />)}</datalist>
                  <input placeholder="Note" value={mf.note} onChange={(e) => setMf({ ...mf, note: e.target.value })} style={{ ...inputS, flex: 1, minWidth: 160 }} />
                  <Btn tone="gold" small onClick={addMisc}>Add</Btn>
                </div>
                <div style={{ fontSize: 11.5, color: C.mute, marginTop: 6 }}>Anything that isn't a tab, crypto, or stake transaction — sales, promo, fees. Positive = income, negative = cost. Category totals show in Balances → Other.</div>
              </Card>
              <Card title={`By category · ${fmt(miscTotal)}`}>
                {Object.entries(miscByCat).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>Nothing yet.</div>}
                {Object.entries(miscByCat).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1])).map(([cat, v]) => (
                  <div key={cat} style={{ display: "flex", padding: "5px 0", borderTop: `1px solid ${C.line}`, fontSize: 13 }}>
                    <button onClick={() => setMiscCatFilter(miscCatFilter === cat ? "" : cat)} style={{ border: "none", background: "none", cursor: "pointer", color: miscCatFilter === cat ? C.goldDark : C.ink, fontWeight: 600, padding: 0 }}>{cat}</button>
                    <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(v)}</span>
                  </div>
                ))}
              </Card>
            </div>
            <div style={{ background: C.card, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
              <div style={{ display: "flex", gap: 10, alignItems: "center", padding: "10px 16px", background: C.cream, flexWrap: "wrap" }}>
                <span style={{ fontWeight: 700 }}>Misc. P&L</span>
                <span style={{ fontSize: 12.5, color: C.mute }}>sort</span>
                <select value={miscSort} onChange={(e) => setMiscSort(e.target.value)} style={{ ...inputS, fontSize: 12.5 }}><option value="date">by date (newest first)</option><option value="category">by category</option></select>
                <select value={miscCatFilter} onChange={(e) => setMiscCatFilter(e.target.value)} style={{ ...inputS, fontSize: 12.5 }}><option value="">all categories</option>{miscCats.map((c) => <option key={c} value={c}>{c}</option>)}</select>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", minWidth: 700, borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: C.cream }}><th style={{ ...th, textAlign: "left" }}>Date</th><th style={{ ...th, textAlign: "left" }}>Category</th><th style={th}>Amount</th><th style={{ ...th, textAlign: "left" }}>Note</th><th style={th}></th></tr></thead>
                  <tbody>
                    {miscRows.length === 0 && <tr><td colSpan={5} style={{ ...tdL, color: C.mute, padding: 16 }}>Nothing logged.</td></tr>}
                    {miscRows.map((e, i) => editMisc === e.id ? (
                      <tr key={e.id} style={{ background: C.banner, borderTop: `1px solid ${C.line}` }}>
                        <td style={tdL}>{dateInput(miscDraft.date, (v) => setMiscDraft({ ...miscDraft, date: v }), 130)}</td>
                        <td style={tdL}><input list="tabs-misc-cats" value={miscDraft.category} onChange={(ev) => setMiscDraft({ ...miscDraft, category: ev.target.value })} style={{ ...inputS, width: 140 }} /></td>
                        <td style={td}><NumInput width={96} value={miscDraft.amount} onChange={(v) => setMiscDraft({ ...miscDraft, amount: v })} /></td>
                        <td style={tdL}><input value={miscDraft.note} onChange={(ev) => setMiscDraft({ ...miscDraft, note: ev.target.value })} style={{ ...inputS, width: "100%", minWidth: 200, boxSizing: "border-box" }} /></td>
                        <td style={{ ...td, whiteSpace: "nowrap" }}><Btn tone="gold" small onClick={commitMisc}>Save</Btn> {iconBtn("cancel", () => setEditMisc(null), C.mute)}</td>
                      </tr>
                    ) : (
                      <tr key={e.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                        <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{e.date}</td>
                        <td style={tdL}><Pill tone="green">{e.category || "uncategorized"}</Pill></td>
                        <td style={td}>{money(e.amount)}</td>
                        <td style={{ ...tdL, fontSize: 12.5, whiteSpace: "normal" }}>{e.note}</td>
                        <td style={{ ...td, whiteSpace: "nowrap" }}>{iconBtn("✎", () => { setEditMisc(e.id); setMiscDraft({ date: e.date, amount: e.amount, note: e.note || "", category: e.category || "" }); }, C.goldDark)}{iconBtn("×", () => delMisc(e.id), C.red)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {/* ═══════════ PLAYER DATA ═══════════ */}
        {view === "players" && (
          <>
            <Card title="Player data — one person, many usernames" right={
              <span style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input placeholder="New person (e.g. Melon)…" value={newPerson} onChange={(e) => setNewPerson(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addPerson()} style={{ ...inputS, width: 190 }} />
                <Btn tone="gold" small onClick={addPerson}>+ Person</Btn>
              </span>}>
              <div style={{ fontSize: 12.5, color: C.mute, marginBottom: 8 }}>
                One person, many usernames across sites — scaled up from the same idea as DL umbrellas in Fish Tank/All American, but site-wide. Type a new username and where it plays (Fish Tank, {(clubs || []).map((c) => c.name).join(", ")}, My Clubs, or any) for a name that hasn't shown up on an export yet, or use <b>pick from known usernames</b> to chip-select from everyone already seen on the last imported week per site. When a week is accepted, settlements for every alias land on the person's single tab. <b>Tab shows as</b> lets someone who both plays under a username and owns/settles a club (e.g. a My Clubs club owner) fold their player tab and club tab into one — set it to Club, add their player alias(es), and both sides land on the same balance going forward. Notes are for accounting — payment methods, Discord vs Telegram, whatever helps. <b>Merge existing</b> pulls tab entries (of any kind — player, club, or other) already sitting under the alias names into the person's name.
              </div>
              {persons.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No people yet. Create "Melon", then add fewtire (Fish Tank), adbank (All American), P7713 (My Clubs).</div>}
              {persons.map((per) => (
                <div key={per.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
                  <div style={{ display: "flex", gap: 8, alignItems: "flex-start", flexWrap: "wrap" }}>
                    <div style={{ minWidth: 220 }}>
                      <input value={per.name} onChange={(e) => patchPerson(per.id, { name: e.target.value })} style={{ ...inputS, width: 200, fontWeight: 700, fontSize: 13.5 }} />
                      <div style={{ marginTop: 6, display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                        <span style={{ fontSize: 11, color: C.mute }}>Tab shows as</span>
                        <select value={per.kind || "player"} onChange={(e) => patchPerson(per.id, { kind: e.target.value })} style={{ ...inputS, padding: "2px 6px", fontSize: 11.5 }}>
                          <option value="player">Player</option><option value="club">Club</option><option value="other">Other</option>
                        </select>
                      </div>
                      <div style={{ marginTop: 6, display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                        <Btn tone="ghost" small onClick={() => mergePerson(per)}>Merge existing tab entries</Btn>
                        {(() => { const cp = cps.find((c) => c.kind === (per.kind || "player") && c.name.toLowerCase() === per.name.toLowerCase()); return cp ? <span style={{ fontSize: 12, color: C.mute }}>tab {money(balances[cp.id] || 0)}</span> : <span style={{ fontSize: 12, color: C.mute }}>no tab yet</span>; })()}
                      </div>
                    </div>
                    <div style={{ flex: 1, minWidth: 260 }}>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                        {per.aliases.map((a, i) => (
                          <span key={a.name + i} style={{ background: C.surface, border: `1px solid ${C.gold}`, borderRadius: 12, padding: "2px 4px 2px 9px", fontSize: 11.5, display: "inline-flex", gap: 4, alignItems: "center" }}>
                            <b>{a.name}</b>
                            <select value={a.site || ""} onChange={(e) => patchPerson(per.id, { aliases: per.aliases.map((x, j) => (j === i ? { ...x, site: e.target.value } : x)) })} style={{ ...inputS, padding: "1px 4px", fontSize: 10.5, border: "none", background: "transparent", color: C.mute }}>
                              {SITES.map((s) => <option key={s} value={s}>{s || "any site"}</option>)}
                            </select>
                            <button onClick={() => patchPerson(per.id, { aliases: per.aliases.filter((_, j) => j !== i) })} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 11 }}>×</button>
                          </span>
                        ))}
                        <input list={"alias-names-" + per.id} placeholder="add username…" value={addAlias[per.id]?.name || ""} onChange={(e) => setAddAlias({ ...addAlias, [per.id]: { ...(addAlias[per.id] || {}), name: e.target.value } })}
                          onKeyDown={(e) => { if (e.key === "Enter") { const v = (addAlias[per.id]?.name || "").trim(); if (v && !per.aliases.some((x) => x.name.toLowerCase() === v.toLowerCase())) patchPerson(per.id, { aliases: [...per.aliases, { name: v, site: addAlias[per.id]?.site || "" }] }); setAddAlias({ ...addAlias, [per.id]: { ...(addAlias[per.id] || {}), name: "" } }); } }}
                          style={{ ...inputS, width: 140, fontSize: 11.5 }} />
                        <select value={addAlias[per.id]?.site || ""} onChange={(e) => setAddAlias({ ...addAlias, [per.id]: { ...(addAlias[per.id] || {}), site: e.target.value } })} style={{ ...inputS, fontSize: 11.5, padding: "3px 6px" }}>
                          {SITES.map((s) => <option key={s} value={s}>{s || "any site"}</option>)}
                        </select>
                        <datalist id={"alias-names-" + per.id}>{allSeenNames.map((n) => <option key={n} value={n} />)}</datalist>
                        {usernameGroups.length > 0 && (
                          <button onClick={() => setPickerOpen({ ...pickerOpen, [per.id]: !pickerOpen[per.id] })} style={{ border: `1px solid ${C.line}`, background: C.surface, color: C.goldDark, borderRadius: 12, cursor: "pointer", fontSize: 11, padding: "2px 10px" }}>
                            {pickerOpen[per.id] ? "hide known usernames ▴" : "pick from known usernames ▾"}
                          </button>
                        )}
                      </div>
                      {pickerOpen[per.id] && (
                        <div style={{ marginTop: 8, padding: "8px 10px", background: C.surface, borderRadius: 6 }}>
                          <div style={{ fontSize: 10.5, color: C.mute, marginBottom: 6 }}>Every username seen on the last imported week across every site. Click to add as an alias for {per.name || "this person"}; dimmed names already belong to someone else.</div>
                          {usernameGroups.map(([site, list]) => (
                            <div key={site} style={{ marginBottom: 8 }}>
                              <div style={{ fontSize: 10.5, color: C.mute, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{site}</div>
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                                {list.map((u) => {
                                  const inThis = personHasUsername(per, u.name, u.site);
                                  const inOther = !inThis && persons.some((p2) => p2.id !== per.id && personHasUsername(p2, u.name, u.site));
                                  return (
                                    <button key={u.site + "|" + u.name} onClick={() => !inOther && toggleUsername(per, u.name, u.site)} style={{
                                      padding: "3px 10px", borderRadius: 12, fontSize: 11.5, fontWeight: 600, cursor: inOther ? "default" : "pointer",
                                      border: `1px solid ${inThis ? C.goldDark : C.line}`,
                                      background: inThis ? C.gold : C.card, color: inThis ? "var(--onGold)" : inOther ? "var(--chipOff)" : C.mute, opacity: inOther ? 0.6 : 1 }}>
                                      {u.name}{u.hint && <span style={{ opacity: 0.75, fontWeight: 400 }}> · {u.hint}</span>}
                                    </button>
                                  );
                                })}
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                      <textarea value={per.notes || ""} onChange={(e) => patchPerson(per.id, { notes: e.target.value })} placeholder="Notes — payment methods, Telegram / Discord, deal reminders…" rows={2} style={{ ...inputS, width: "100%", boxSizing: "border-box", marginTop: 6, fontFamily: "inherit", resize: "vertical" }} />
                    </div>
                    <button onClick={() => { if (window.confirm(`Delete person "${per.name}"? Tab entries stay under whatever name they have now.`)) setPersons(persons.filter((x) => x.id !== per.id)); }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 13 }}>× delete</button>
                  </div>
                </div>
              ))}
            </Card>
          </>
        )}
      </div>
    </div>
  );
}

async function downloadTabsExcel(data, cps, balances, allEntries, staking, vigEntries, vigTotal, miscByCat, miscTotal, otherTotals) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Balances");
  [30, 10, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `Tabs — balances as of ${today()}`);
  [["player", "Player tabs"], ["club", "Club tabs"], ["other", "Other"]].forEach(([kind, title]) => {
    const hr = ws.addRow([title]);
    hr.getCell(1).font = { name: "Arial", size: 11, bold: true, color: { argb: XLC.white } };
    for (let j = 1; j <= 3; j++) hr.getCell(j).fill = fillOf(XLC.gold);
    const list = cps.filter((c) => c.kind === kind && Math.abs(balances[c.id] || 0) > 0.005);
    list.forEach((c, i) => {
      const r = ws.addRow([]);
      xText(r.getCell(1), c.name + (staking.makeupByCp[c.id] > 0.005 ? ` (in makeup ${fmt(staking.makeupByCp[c.id])})` : ""), { bold: true });
      xMoney(r.getCell(3), balances[c.id] || 0);
      if (i % 2 === 1) for (let j = 1; j <= 3; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
    });
    const sub = ws.addRow([]);
    xText(sub.getCell(1), "Subtotal", { bold: true });
    xMoney(sub.getCell(3), list.reduce((a, c) => a + (balances[c.id] || 0), 0), { bold: true });
    for (let j = 1; j <= 3; j++) sub.getCell(j).fill = fillOf(XLC.cream);
    if (kind === "other") {
      const nr = ws.addRow(["Running totals (read-only, not in net position)"]); nr.getCell(1).font = { name: "Arial", size: 9, italic: true, color: { argb: XLC.mute } };
      otherTotals.forEach(([l, v]) => { const r = ws.addRow([]); xText(r.getCell(1), l, { mute: true }); xMoney(r.getCell(3), v); });
    }
    ws.addRow([]);
  });
  const tot = ws.addRow([]);
  xText(tot.getCell(1), "NET POSITION", { bold: true, white: true });
  xMoney(tot.getCell(3), Object.values(balances).reduce((a, v) => a + v, 0), { bold: true, white: true, colorSign: false });
  for (let j = 1; j <= 3; j++) tot.getCell(j).fill = fillOf(XLC.bar);

  const wl = wb.addWorksheet("Ledger");
  [11, 20, 9, 13, 11, 10, 40, 18].forEach((w, i) => (wl.getColumn(i + 1).width = w));
  xTitle(wl, "Ledger — full history (newest first)");
  xHeader(wl, ["Date", "Counterparty", "Kind", "Amount", "Method", "Vig", "Note", "Source"], 3);
  sortDateDesc(allEntries, (e) => allEntries.indexOf(e)).forEach((e, i) => {
    const cp = cps.find((c) => c.id === e.cpId);
    const r = wl.addRow([]);
    xText(r.getCell(1), e.date, { mute: true });
    xText(r.getCell(2), cp?.name || "?", { bold: true });
    xText(r.getCell(3), cp?.kind || "", { mute: true });
    xMoney(r.getCell(4), e.amount);
    xText(r.getCell(5), e.method || "", { mute: true });
    if (e.vig != null) xMoney(r.getCell(6), e.vig);
    xText(r.getCell(7), e.note || "", { mute: true });
    xText(r.getCell(8), e.source === "week" ? e.week || "week" : e.source || "manual", { mute: true });
    if (i % 2 === 1) for (let j = 1; j <= 8; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });

  const wv = wb.addWorksheet("Vig");
  [11, 20, 14, 8, 14, 14, 40].forEach((w, i) => (wv.getColumn(i + 1).width = w));
  xTitle(wv, `Crypto vig — total ${fmt(vigTotal)}`);
  xHeader(wv, ["Date", "Who", "Amount moved", "Rate", "Vig", "Tab effect", "Note"], 2);
  sortDateDesc(vigEntries, (e) => data.entries.indexOf(e)).forEach((e, i) => {
    const r = wv.addRow([]);
    xText(r.getCell(1), e.date, { mute: true }); xText(r.getCell(2), cps.find((c) => c.id === e.cpId)?.name || "?", { bold: true });
    xMoney(r.getCell(3), e.baseAmount ?? e.amount); xNum(r.getCell(4), e.vigRate ?? 0, '0.##"%"'); xMoney(r.getCell(5), e.vig || 0); xMoney(r.getCell(6), e.amount); xText(r.getCell(7), e.note || "", { mute: true });
    if (i % 2 === 1) for (let j = 1; j <= 7; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });

  const wsk = wb.addWorksheet("Staking");
  [20, 12, 11, 18, 12, 8, 12, 10, 12, 12, 12, 30].forEach((w, i) => (wsk.getColumn(i + 1).width = w));
  xTitle(wsk, `Staking — net action buy ${fmt(staking.totals.netActionBuy)} · chopped ${fmt(staking.totals.chopped)} · makeup ${fmt(staking.totals.makeup)}`);
  xHeader(wsk, ["Player", "Deal", "Date", "Game", "P&L", "%", "Net", "Holder", "Makeup after", "Ak chop", "Tab effect", "Note"], 4);
  staking.deals.forEach((d) => d.rows.forEach((r, i) => {
    const row = wsk.addRow([]);
    xText(row.getCell(1), d.name, { bold: true }); xText(row.getCell(2), d.deal.type, { mute: true }); xText(row.getCell(3), r.date, { mute: true }); xText(row.getCell(4), r.game || "", { mute: true });
    xMoney(row.getCell(5), +r.pnl || 0); xNum(row.getCell(6), r.pct, '0.##"%"'); xMoney(row.getCell(7), r.net); xText(row.getCell(8), r.holder === "player" ? d.name : "AK", { mute: true });
    if (r.makeupAfter != null) xMoney(row.getCell(9), r.makeupAfter, { colorSign: false }); if (r.akChop != null) xMoney(row.getCell(10), r.akChop); if (r.tab != null) xMoney(row.getCell(11), r.tab); xText(row.getCell(12), r.note || "", { mute: true });
    if (i % 2 === 1) for (let j = 1; j <= 12; j++) row.getCell(j).fill = fillOf(XLC.rowAlt);
  }));
  if (staking.imported.length) {
    wsk.addRow([]);
    xHeader(wsk, ["Imported (your share)", "Site", "Week", "Kind", "Net", "", "Chopped", "", "Makeup after"], 4);
    staking.imported.forEach((x) => x.rows.forEach((r) => { const row = wsk.addRow([]); xText(row.getCell(1), x.name, { bold: true }); xText(row.getCell(2), r.site, { mute: true }); xText(row.getCell(3), r.sourceKey.split(":").slice(1).join(":"), { mute: true }); xText(row.getCell(4), r.kind, { mute: true }); xMoney(row.getCell(5), r.net); xMoney(row.getCell(7), r.chopped || 0); if (r.makeupAfter != null) xMoney(row.getCell(9), r.makeupAfter, { colorSign: false }); }));
  }

  const wm = wb.addWorksheet("Misc PnL");
  [11, 18, 14, 40].forEach((w, i) => (wm.getColumn(i + 1).width = w));
  xTitle(wm, `Misc. P&L — total ${fmt(miscTotal)}`);
  xHeader(wm, ["Date", "Category", "Amount", "Note"], 2);
  [...data.misc].sort((a, b) => (a.category || "").localeCompare(b.category || "") || (b.date || "").localeCompare(a.date || "")).forEach((e, i) => {
    const r = wm.addRow([]); xText(r.getCell(1), e.date, { mute: true }); xText(r.getCell(2), e.category || "uncategorized", { bold: true }); xMoney(r.getCell(3), e.amount); xText(r.getCell(4), e.note || "", { mute: true });
    if (i % 2 === 1) for (let j = 1; j <= 4; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  wm.addRow([]);
  Object.entries(miscByCat).forEach(([cat, v]) => { const r = wm.addRow([]); xText(r.getCell(2), cat, { bold: true }); xMoney(r.getCell(3), v, { bold: true }); });
  await saveWb(wb, `Tabs_${today()}.xlsx`);
}
