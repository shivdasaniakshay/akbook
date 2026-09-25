import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

// Persistence lives in sync.js: localStorage is the cache, Supabase is the shared (encrypted) truth.
import { store, blobPut as idbPut, blobGet as idbGet, blobDel as idbDel } from "./sync.js";

// ———————————————— Seed / snapshot ————————————————
// Everything the app knows lives under these localStorage keys.
const STATE_KEYS_STATIC = ["fishtank-config-v4", "fishtank-lastweek-v4", "agentclubs-v3", "tabs-v1", "allamerican-v1", "allamerican-lastweek-v1", "ownerclubs-v1", "archive-index-v1", "book-weeks-v1", "book-checklist-v1", "fishtank-club-v1", "fishtank-club-week-v1", "ui-v1", "bankroll-v1"];
// Owner clubs added later live under oc-cfg-* / oc-week-* keys — pick those up too.
const allStateKeys = () => {
  const all = [...STATE_KEYS_STATIC];
  try { for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); if (k && (k.startsWith("oc-cfg-") || k.startsWith("oc-week-") || k.startsWith("wk:")) && !all.includes(k)) all.push(k); } } catch (e) {}
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
  for (const k of new Set([...STATE_KEYS_STATIC, ...Object.keys(seed).filter((k) => k.startsWith("oc-cfg-") || k.startsWith("oc-week-") || k.startsWith("wk:"))])) {
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
  const keys = [...new Set([...STATE_KEYS_STATIC, ...Object.keys(seed).filter((k) => k.startsWith("oc-cfg-") || k.startsWith("oc-week-") || k.startsWith("wk:"))])];
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
    "--s1": "#2a78d6", "--s2": "#eb6834", "--s3": "#4a3aa7",
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
    "--s1": "#3987e5", "--s2": "#d95926", "--s3": "#9085e9",
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
// Each week's underlying data (players + deals at the time) lives in the synced store,
// so any signed-in browser can reopen the numbers and rebuild the Excel.
const wkKey = (site, period) => `wk:${site}:${period}`;
async function archiveWeekData(club, cfg, players, period, jackpotFromExport) {
  if (!period || !players?.length) return;
  await store.set(wkKey(club.id, period), JSON.stringify({ site: club.id, siteName: club.name, period, players, cfg, jackpotFromExport: jackpotFromExport ?? null, savedAt: new Date().toISOString() }));
  const idx = await loadArchiveIndex();
  const old = idx.weeks.find((w) => w.site === club.id && w.period === period);
  const rec = { ...(old || { site: club.id, siteName: club.name, period, archivedAt: new Date().toISOString(), hasRaw: false, rawName: "", gen: [] }), siteName: club.name, hasData: true };
  idx.weeks = [...idx.weeks.filter((w) => w !== old), rec];
  await saveArchiveIndex(idx);
}
async function loadArchivedWeek(site, period) { try { const c = await store.get(wkKey(site, period)); return c?.value ? JSON.parse(c.value) : null; } catch (e) { return null; } }
async function deleteArchivedWeek(site, period) {
  try { await store.del(wkKey(site, period)); } catch (e) {}
  await idbDel(rawKey(site, period)).catch(() => {});
  await idbDel(genKey(site, period)).catch(() => {});
  const idx = await loadArchiveIndex();
  idx.weeks = idx.weeks.filter((w) => !(w.site === site && w.period === period));
  await saveArchiveIndex(idx);
}

function ArchiveView({ clubs }) {
  const [idx, setIdx] = useState(null);
  const [open, setOpen] = useState(null); // `${site}|${period}`
  const [detail, setDetail] = useState(null);
  const pastRef = useRef(null); const [pastClub, setPastClub] = useState(null);
  const reload = async () => setIdx(await loadArchiveIndex());
  useEffect(() => { (async () => {
    // Make sure every club's current week is in the archive too.
    for (const c of clubs) {
      try { const d = await store.get(c.weekKey), cf = await store.get(c.cfgKey);
        if (d?.value && cf?.value) { const w = JSON.parse(d.value); const has = await store.get(wkKey(c.id, w.period)); if (!has?.value && w.players?.length) await archiveWeekData(c, { ...AA_DEFAULT_CFG, ...JSON.parse(cf.value) }, w.players, w.period, w.jackpotFromExport); } } catch (e) {}
    }
    reload();
  })(); }, []);
  if (!idx) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;
  const clubOf = (w) => clubs.find((c) => c.id === w.site) || { id: w.site, name: w.siteName || w.site, owners: [], meId: null };
  const toggle = async (w) => {
    const k = w.site + "|" + w.period;
    if (open === k) { setOpen(null); return; }
    setOpen(k); setDetail(null);
    const d = await loadArchivedWeek(w.site, w.period);
    if (d) { const club = clubOf(w); setDetail({ club, d, model: buildAAModel(d.players, { ...AA_DEFAULT_CFG, ...d.cfg }, d.period, club) }); }
    else setDetail({ missing: true });
  };
  const dlRaw = async (w) => { const r = await idbGet(rawKey(w.site, w.period)).catch(() => null); if (r?.buf) downloadBytes(r.buf, r.name || `${w.siteName}_${w.period}_raw.xlsx`); else window.alert("The raw upload wasn't stored for this week."); };
  const dlGen = async (w, i) => { const g = await idbGet(genKey(w.site, w.period)).catch(() => null); const f = g?.files?.[i]; if (f) downloadBytes(f.buf, f.name); else if (i === 0) window.alert("Those files weren't carried over from the old site. Use \"Add a past week\" with the raw export to bring the numbers back."); };
  // Re-upload an old export: its numbers are saved with the club's current deals.
  const addPast = async (file) => {
    const club = clubs.find((c) => c.id === pastClub); if (!club) return;
    try {
      const buf = await file.arrayBuffer();
      const { players, period, jackpot, jackpotFound } = parseWorkbook(buf);
      const cf = await store.get(club.cfgKey); const cfg = { ...AA_DEFAULT_CFG, ...(cf?.value ? JSON.parse(cf.value) : {}) };
      await archiveWeekData(club, cfg, players, period, jackpotFound ? jackpot : null);
      try { await idbPut(rawKey(club.id, period), { site: club.id, siteName: club.name, period, name: file.name, buf }); const idx2 = await loadArchiveIndex(); idx2.weeks = idx2.weeks.map((w) => (w.site === club.id && w.period === period ? { ...w, hasRaw: true, rawName: file.name } : w)); await saveArchiveIndex(idx2); } catch (e) {}
      reload();
    } catch (e) { window.alert("Couldn't read that file: " + (e.message || e)); }
  };
  const sortKey = (p) => (p || "").replace(/[^\d]/g, "");
  const sites = [...new Set(idx.weeks.map((w) => w.site))].sort((x, y) => clubs.findIndex((c) => c.id === x) - clubs.findIndex((c) => c.id === y));
  return (
    <div style={{ padding: "18px clamp(10px, 2vw, 26px) 60px", maxWidth: 1400, margin: "0 auto", display: "grid", gap: 14 }}>
      <input ref={pastRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) addPast(f); e.target.value = ""; }} />
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>Archive</div>
        <span style={{ marginLeft: "auto", display: "flex", gap: 6, alignItems: "center" }}>
          <span style={{ fontSize: 12, color: C.mute }}>Add a past week:</span>
          {clubs.map((c) => <Btn key={c.id} tone="ghost" small onClick={() => { setPastClub(c.id); setTimeout(() => pastRef.current?.click(), 0); }}>{c.name}</Btn>)}
        </span>
      </div>
      {idx.weeks.length === 0 && <Card title="Nothing archived yet"><div style={{ color: C.mute, fontSize: 13 }}>Each uploaded week lands here automatically.</div></Card>}
      {sites.map((site) => {
        const ws = idx.weeks.filter((w) => w.site === site).sort((x, y) => sortKey(y.period).localeCompare(sortKey(x.period)));
        const club = clubOf(ws[0]);
        return (
          <Card key={site} title={`${club.name} · ${ws.length} week${ws.length !== 1 ? "s" : ""}`}>
            {ws.map((w) => {
              const k = w.site + "|" + w.period, isOpen = open === k;
              return (
                <div key={w.period} style={{ borderTop: `1px solid ${C.line}` }}>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", padding: "7px 0", fontSize: 13, flexWrap: "wrap" }}>
                    <button onClick={() => toggle(w)} style={{ border: "none", background: "none", color: C.ink, cursor: "pointer", fontWeight: 700, padding: 0, fontSize: 13 }}>{isOpen ? "▾" : "▸"} {bookLabel(w.period) || w.period}</button>
                    {w.hasData ? <Pill tone="green">numbers saved</Pill> : <Pill>files only</Pill>}
                    <span style={{ marginLeft: "auto", display: "flex", gap: 6, flexWrap: "wrap" }}>
                      {w.hasRaw && <Btn tone="ghost" small onClick={() => dlRaw(w)}>Raw upload</Btn>}
                      {(w.gen || []).length > 0 && <Btn tone="ghost" small onClick={() => (w.gen || []).forEach((_, j) => dlGen(w, j))}>Excel files ({w.gen.length})</Btn>}
                      <button onClick={async () => { if (window.confirm(`Delete archived week ${w.period}?`)) { await deleteArchivedWeek(w.site, w.period); reload(); } }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
                    </span>
                  </div>
                  {isOpen && (
                    <div style={{ padding: "4px 0 12px 18px" }}>
                      {!detail ? <div style={{ color: C.mute, fontSize: 12.5 }}>Loading…</div> : detail.missing ? <div style={{ color: C.mute, fontSize: 12.5 }}>Only files were kept for this week (it was archived before numbers were saved).</div> : (() => {
                        const { model: m, club: cl, d } = detail; const H = m.H;
                        return (
                          <div style={{ display: "grid", gap: 8 }}>
                            <div style={{ display: "flex", gap: 18, flexWrap: "wrap", fontSize: 13 }}>
                              <span>Rake <b>{fmt(m.clubRevenue)}</b></span><span>Pool <b>{fmt(m.pool)}</b></span>
                              {m.ownerIds.map((o) => <span key={o}>{H.lbl(o)} profit {money(m.profit[o])}</span>)}
                              <span style={{ color: C.mute }}>{m.transfers.length ? m.transfers.map((t) => `${H.lbl(t.from)} pays ${H.lbl(t.to)} ${fmt(t.amount)}`).join(" · ") : m.ready ? "even" : "settle-up incomplete"}</span>
                            </div>
                            <table style={{ width: "100%", borderCollapse: "collapse" }}>
                              <thead><tr><th style={{ ...th, textAlign: "left" }}>Line</th><th style={{ ...th, textAlign: "left" }}>Collected by</th><th style={th}>Winnings</th><th style={th}>Tips</th><th style={th}>Rakeback</th><th style={th}>Settlement</th></tr></thead>
                              <tbody>{[...m.entities, ...m.backedEntities].map((e) => <tr key={e.key} style={{ borderTop: `1px solid ${C.line}` }}><td style={tdL}>{e.name}</td><td style={tdL}>{e.collector ? H.lbl(e.collector) : e.backer ? H.lbl(e.backer) : "—"}</td><td style={td}>{fmt(e.pnl)}</td><td style={td}>{fmt(e.fee)}</td><td style={td}>{fmt(e.tipback ?? e.rbCredit ?? 0)}</td><td style={td}>{money(-e.settlement)}</td></tr>)}</tbody>
                            </table>
                            <div style={{ display: "flex", gap: 8 }}>
                              <Btn tone="gold" small onClick={() => downloadAAWorkbook(m, d.period, cl)}>Rebuild Excel</Btn>
                              {m.ownerIds.map((o) => <Btn key={o} tone="ghost" small onClick={() => downloadAAOwnerExcel(m, o, d.period, cl)}>{H.lbl(o)} report</Btn>)}
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  )}
                </div>
              );
            })}
          </Card>
        );
      })}
      <Notes><div>Every week you upload is saved here with its numbers and the deals in place at the time, so anyone with the site link and passphrase can reopen it and rebuild its Excel. Raw uploads and generated files are kept too when available. Settlement is shown as + they owe the club / − the club owes them.</div></Notes>
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
// Collapsible "How this works" block — definitions live at the bottom of each screen.
function Notes({ children, title = "How this works" }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ marginTop: 22, borderTop: `1px solid ${C.line}`, paddingTop: 10 }}>
      <button onClick={() => setOpen(!open)} style={{ border: "none", background: "none", color: C.mute, cursor: "pointer", fontSize: 12.5, fontWeight: 700, padding: 0 }}>{open ? "▾" : "▸"} {title}</button>
      {open && <div style={{ color: C.mute, fontSize: 12.5, lineHeight: 1.55, marginTop: 8, display: "grid", gap: 8 }}>{children}</div>}
    </div>
  );
}
// Tables inside .fit shrink their text/padding with the window so every column fits without sideways scroll.
const FIT_CSS = `.app table{width:100%;table-layout:auto}
.app th{white-space:normal!important;padding:6px 5px!important;font-size:clamp(8.5px,0.62vw,10.5px)!important;letter-spacing:.04em!important}
.app td{padding:5px 5px!important;font-size:clamp(10px,0.78vw,13.5px)!important}
.app td{white-space:normal!important;overflow-wrap:anywhere}
.app td select{max-width:100%;font-size:clamp(10px,0.74vw,12px)!important;padding:3px 4px!important}
.app td input{font-size:clamp(10px,0.74vw,12.5px)!important;padding:3px 4px!important;max-width:clamp(44px,5.5vw,90px)}
.fit{overflow:hidden!important}`;

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
const UI_KEY = "ui-v1";
export default function App() {
  const [ui, setUi] = useState({ theme: "dark", mode: "" });
  const [err, setErr] = useState("");
  const [loaded, setLoaded] = useState(false);
  const [clubs, setClubs] = useState([]);
  const seedRef = useRef(null);
  const saveClubs = async (next) => { setClubs(next); await saveOwnerClubs(next); };
  const up = (patch) => { const next = { ...ui, ...patch }; setUi(next); store.set(UI_KEY, JSON.stringify(next)).catch(() => {}); };
  const addClub = async () => {
    const name = window.prompt("Name of the new club:");
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

  useEffect(() => { (async () => {
    await applySeedOnce();
    let u = null;
    try { const c = await store.get(UI_KEY); if (c?.value) u = JSON.parse(c.value); } catch (e) {}
    // Older builds kept theme + current tab inside the Fish Tank config.
    if (!u) { try { const c = await store.get("fishtank-config-v4"); if (c?.value) { const s = JSON.parse(c.value); u = { theme: s.theme || "dark", mode: s.mode || "" }; } } catch (e) {} }
    if (u) setUi({ theme: "dark", ...u });
    try { setClubs(await loadOwnerClubs()); } catch (e) {}
    setLoaded(true);
  })(); }, []);

  if (!loaded) return <div style={{ fontFamily: "Georgia, serif", padding: 40, color: "#8A7E6C" }}>Loading…</div>;

  const theme = ui.theme === "light" ? "light" : "dark";
  let mode = ui.mode || (clubs[0] ? "oc:" + clubs[0].id : "tabs");
  if (mode === "fishtank") mode = "oc:fishtank";
  if (mode === "allamerican") mode = "oc:allamerican";
  const activeClub = mode.startsWith("oc:") ? clubs.find((c) => c.id === mode.slice(3)) : null;
  const navBtn = (k, label) => (
    <button key={k} onClick={() => up({ mode: k })} style={{
      border: "none", cursor: "pointer", borderRadius: 5, padding: "5px 12px", fontSize: 12.5, fontWeight: 700, whiteSpace: "nowrap",
      background: mode === k ? "var(--gold)" : "transparent", color: mode === k ? "var(--onGold)" : "var(--barMute)" }}>{label}</button>
  );
  const barBtn = { background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 12 };

  return (
    <div className="app" style={{ ...PALETTES[theme], minHeight: "100vh", background: C.paper, color: C.ink, fontFamily: "'Avenir Next', 'Segoe UI', system-ui, sans-serif", colorScheme: theme }}>
      <style>{FIT_CSS}</style>
      <div style={{ background: C.bar, padding: "12px clamp(10px, 2vw, 26px)", display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 21, color: "var(--barText)" }}>AK's Book</div>
        <div style={{ display: "flex", gap: 2, background: "rgba(255,255,255,0.08)", borderRadius: 7, padding: 3, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: 10, color: "var(--barSubtle)", padding: "0 6px", letterSpacing: ".08em", textTransform: "uppercase" }}>Clubs</span>
          {clubs.map((c) => navBtn("oc:" + c.id, c.name))}
          <button onClick={addClub} title="Add a club" style={{ border: "none", cursor: "pointer", borderRadius: 5, padding: "5px 9px", fontSize: 13, fontWeight: 700, background: "transparent", color: "var(--barMute)" }}>+</button>
        </div>
        <div style={{ display: "flex", gap: 2, background: "rgba(255,255,255,0.08)", borderRadius: 7, padding: 3, flexWrap: "wrap" }}>
          {[["agent", "My Clubs"], ["book", "Book"], ["tabs", "Tabs"], ["archive", "Archive"], ["bankroll", "Bankroll"]].map(([k, l]) => navBtn(k, l))}
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
          <input ref={seedRef} type="file" accept=".json,application/json" style={{ display: "none" }}
            onChange={(e) => { const f = e.target.files?.[0]; if (f) onSeedFile(f); e.target.value = ""; }} />
          <button onClick={exportSnapshot} title="Download all data as a backup file" style={barBtn}>Export</button>
          <button onClick={() => seedRef.current?.click()} title="Load a backup file" style={barBtn}>Import</button>
          <button onClick={() => up({ theme: theme === "dark" ? "light" : "dark" })} title="Toggle dark mode" style={{ ...barBtn, fontSize: 14 }}>{theme === "dark" ? "☀" : "☾"}</button>
        </div>
      </div>
      {err && <div style={{ background: "var(--errBg)", color: C.red, padding: "10px 14px", fontSize: 13.5 }}>{err}</div>}

      {mode === "agent" ? (
        <AgentClubs theme={theme} />
      ) : mode === "book" ? (
        <BookSection />
      ) : activeClub ? (
        <AllAmerican key={activeClub.id} club={activeClub} clubs={clubs} saveClubs={saveClubs} onDeleteClub={deleteClub} />
      ) : mode === "archive" ? (
        <ArchiveView clubs={clubs} />
      ) : mode === "bankroll" ? (
        <Bankroll clubs={clubs} />
      ) : (
        <TabsLedger clubs={clubs} />
      )}
    </div>
  );
}

// ———— Cross-module helpers ————
// Fish Tank now runs as an ordinary club (see FT_CLUB); the old standalone model is retired.
async function loadFishTankModel() { return null; }
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
// rows for a set of usernames out of an owner club week (All American, Honey
// Pot, or any future owner club) — same shape as ftRowsForNames, so it merges
// straight into the same report list. `oc` is one entry from
// loadAllOwnerClubModels(): { model, period, cfg, club }.
function aaRowsForNames(oc, nameSet) {
  if (!oc || !oc.model) return [];
  const rows = [];
  const label = `${oc.club?.name || "Owner club"}${oc.period ? ` (${oc.period})` : ""}`;
  const clubKey = oc.club?.id || label;
  oc.model.entities.forEach((e) => {
    (e.members || []).forEach((mm) => {
      if (nameSet.has(mm.name.trim().toLowerCase())) {
        // Members don't carry a pre-computed settlement (only the entity
        // aggregate does) — but it's additive: entity settlement = net − trCut
        // = (pnl + tipback) − trCut summed across members, so each member's
        // own slice is exactly pnl + tipback − trCut.
        rows.push({ id: "oc-" + clubKey + "-" + mm.memberId, clubName: label, name: mm.name, customDeal: `TB ${mm.tbPct}%${mm.tr ? ` · TR ${mm.tr}%` : ""}`, pnl: mm.pnl, tips: mm.fee, tipback: mm.tipback, settlement: r2(mm.pnl + mm.tipback - mm.trCut), margin: 0, played: true });
      }
    });
  });
  (oc.model.backedEntities || []).forEach((e) => {
    if (nameSet.has(e.name.trim().toLowerCase())) {
      rows.push({ id: "oc-" + clubKey + "-" + e.key, clubName: label, name: e.name, customDeal: e.dealType === "action" ? `action buy ${e.actionPct}%` : `makeup · RB ${e.rb}%`, pnl: e.pnl, tips: e.fee, tipback: e.rbCredit, settlement: e.settlement, margin: 0, played: true });
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
  clubTB: 80, clubAction: 0, hasBBJ: false,
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

// bbj (bad beat jackpot deduction, club-currency, defaults to 0) is kept
// completely out of tipback/net/actionCut — it never feeds the TR base (or a
// custom formula's inputs) — and is only taken off the settlement itself,
// after everything else including TR/formula, then converted like the rest.
function settleLine(pnl, tips, d, club, conv, bbj) {
  bbj = bbj || 0;
  const tipback = (tips * (d.tb || 0)) / 100;
  const net = pnl + tipback;
  const gross = pnl + tips;
  const tr = d.tr || 0;
  if (club._fn) {
    let s;
    try { s = club._fn({ pnl, tips, tb: d.tb || 0, tr, rebate: 0, tipback, net, gross }); }
    catch (e) { s = net; }
    if (!isFinite(s)) s = net;
    return { tipback, net, actionCut: net - s, settlement: (s - bbj) * conv };
  }
  const trBase = club.actionBase === "pnl" ? pnl : club.actionBase === "gross" ? gross : net;
  const actionCut = (trBase * tr) / 100;
  return { tipback, net, actionCut, settlement: (net - actionCut - bbj) * conv };
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
      // BBJ only applies where the club is flagged for it — a stray value left
      // over from before the toggle was on (or after it's switched off) is
      // ignored rather than silently still being deducted.
      const bbj = club.hasBBJ ? (+e.bbj || 0) : 0;
      const played = pnl !== 0 || tips !== 0 || bbj !== 0;
      const isMine = myAccSet.has(p.name.trim().toLowerCase());
      const clubDeal = { tb: club.clubTB || 0, tr: club.clubAction || 0 };
      const eff = isMine ? clubDeal : p; // your own accounts ride the club's deal automatically
      // BBJ is charged to the player and passed straight through to the club
      // (you're just the middleman on it), so it comes off both sides of the
      // settlement equally — your margin on this player is unaffected by it.
      const mine = settleLine(pnl, tips, eff, ctx, conv, bbj);
      // what the club pays you for this player's action
      const clubSide = settleLine(pnl, tips, clubDeal, ctx, conv, bbj);
      return { ...p, tb: eff.tb, tr: eff.tr, isMine, clubId: club.id, clubName: club.name, pnl, tips, bbj, played,
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
  const [ownerClubs, setOwnerClubs] = useState([]);
  useEffect(() => { (async () => { setFt(await loadFishTankModel()); setPersons(await loadPersons()); setOwnerClubs(await loadAllOwnerClubModels()); })(); }, []);

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
        {week && tab === "summary" && <AgentSummary model={model} wk={wk} setExportData={setExportData} acfg={acfg} up={up} ft={ft} persons={persons} ownerClubs={ownerClubs} />}
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
        Enter each player's <b>P&L</b> and <b>Tips</b> in club currency; blank = no play. <span style={{ color: C.green, fontWeight: 700 }}>Green</span> = you pay them · <span style={{ color: C.red, fontWeight: 700 }}>red</span> = they pay you. <b>Margin</b> is what you keep after the club pays you for that player. Clubs with the <b>BBJ deduction</b> toggle on (Clubs & deals) get a <b>BBJ</b> column — a bad-beat-jackpot amount charged to the player and passed through to the club, kept separate from P&L/tips and untouched by TB/TR, taken off the settlement last.
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
              {c.hasBBJ && <th style={th} title="Bad beat jackpot — charged to the player, passed through to the club, taken off the settlement last.">BBJ</th>}
              <th style={th}>Tipback</th><th style={th}>Settlement</th><th style={th}>Your margin</th>
            </tr></thead>
            <tbody>
              {c.playersC.map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}{model.inUmbrella.has(p.id) && <span style={{ marginLeft: 6 }}><Pill tone="blue">{(acfg.umbrellas.find((u) => u.playerIds.includes(p.id)) || {}).name}</Pill></span>}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 11.5 }}>{dealLabel(p)}</td>
                  <td style={td}><NumInput width={92} value={(week.entries[p.id] || {}).pnl ?? ""} onChange={(v) => setEntry(p.id, "pnl", v)} disabled={locked} /></td>
                  <td style={td}><NumInput width={82} value={(week.entries[p.id] || {}).tips ?? ""} onChange={(v) => setEntry(p.id, "tips", v)} disabled={locked} /></td>
                  {c.hasBBJ && <td style={td}><NumInput width={82} value={(week.entries[p.id] || {}).bbj ?? ""} onChange={(v) => setEntry(p.id, "bbj", v)} disabled={locked} /></td>}
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

function AgentSummary({ model, wk, setExportData, acfg, up, ft, persons, ownerClubs }) {
  const t = model.totals;
  const active = model.clubs.filter((c) => c.active);
  const [reportSel, setReportSel] = useState("");
  const bundled = new Set(persons.flatMap((p) => p.usernames.map((u) => u.trim().toLowerCase())));
  const ftNames = ft ? [...new Set(ft.model.entities.flatMap((e) => e.type === "backed" ? [e.name] : e.members.map((m) => m.name)))] : [];
  const ocNames = (ownerClubs || []).flatMap((oc) => (oc.model.entities || []).flatMap((e) => (e.members || []).map((m) => m.name)).concat((oc.model.backedEntities || []).map((e) => e.name)));
  const rawNames = [...new Set([...model.allPlayers.map((p) => p.name.trim()), ...ftNames.map((n) => n.trim()), ...ocNames.map((n) => n.trim())])]
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
  const aaRows = (ownerClubs || []).flatMap((oc) => aaRowsForNames(oc, nameSet));
  // clubs owned by this person also fold into their report (sign flipped: positive = you pay them)
  const clubRows = model.clubs.filter((c) => c.active && (c.owner || "").trim() && (nameSet.has(c.owner.trim().toLowerCase()) || c.owner.trim().toLowerCase() === reportLabel.toLowerCase()))
    .map((c) => ({ id: "club-" + c.id, clubName: c.name, name: c.owner, customDeal: "club settlement", pnl: c.pnl, tips: c.tips, tipback: 0, settlement: -c.clubSettlement, margin: 0, played: true }));
  const reportRows = [...ftRows, ...aaRows, ...mcRows, ...clubRows];
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
  const copyRows = (title, rows) => setExportData({ title, text: toTSV(["Club", "Player", "Deal", "P&L", "Tips", "BBJ", "Tipback", "Settlement", "Your margin"], rows.map((p) => [p.clubName, p.name, p.customDeal || dealLabel(p), p.pnl.toFixed(2), p.tips.toFixed(2), (p.bbj || 0).toFixed(2), p.tipback.toFixed(2), p.settlement.toFixed(2), (p.margin || 0).toFixed(2)])) });

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
                    <td style={td}>tips {fmt(p.tips)}{p.bbj ? ` · bbj ${fmt(p.bbj)}` : ""}</td>
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
                  <td style={td}>tips {fmt(p.tips)}{p.bbj ? ` · bbj ${fmt(p.bbj)}` : ""}</td>
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
                  <td style={td}>tips {fmt(p.tips)}{p.bbj ? ` · bbj ${fmt(p.bbj)}` : ""}</td>
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
const Toggle = ({ on, onClick, label, title }) => (
  <button onClick={onClick} title={title} style={{
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
              <Toggle on={!!c.hasBBJ} onClick={() => setClub(c.id, { hasBBJ: !c.hasBBJ })} label="BBJ deduction" title="Bad beat jackpot: a per-player amount entered on Weekly entry, charged to the player and passed straight through to the club — separate from P&L/tips, untouched by TB/TR/formula, and taken off the settlement last." />
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
    xMoney(r.getCell(6), p.bbj || 0, { colorSign: false });
    xMoney(r.getCell(7), p.tipback, { colorSign: false });
    xMoney(r.getCell(8), p.settlement);
    xMoney(r.getCell(9), p.margin || 0);
    if (i % 2 === 1) for (let j = 1; j <= 9; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const tr = ws.addRow([]);
  xText(tr.getCell(1), "TOTAL", { bold: true });
  xMoney(tr.getCell(8), rows.reduce((a, p) => a + p.settlement, 0), { bold: true });
  xMoney(tr.getCell(9), rows.reduce((a, p) => a + (p.margin || 0), 0), { bold: true });
  for (let j = 1; j <= 9; j++) tr.getCell(j).fill = fillOf(XLC.cream);
}
const AG_HEAD = ["Club", "Player", "Deal", "P&L", "Tips", "BBJ", "Tipback", "Settlement", "Your margin"];
async function downloadPlayerExcel(name, rows, wk) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(safeSheetName(name, wb));
  [18, 18, 22, 12, 12, 12, 12, 13, 13].forEach((w, i) => (ws.getColumn(i + 1).width = w));
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
    [18, 18, 22, 12, 12, 12, 12, 13, 13].forEach((w, i) => (w2.getColumn(i + 1).width = w));
    xTitle(w2, `${u.name} (umbrella) — ${wk}`);
    xHeader(w2, AG_HEAD, 3);
    playerSheetRows(w2, u.played, true);
  });
  active.forEach((c) => {
    const w2 = wb.addWorksheet(safeSheetName(c.name, wb));
    [18, 22, 12, 12, 12, 12, 12, 13, 13].forEach((w, i) => (w2.getColumn(i + 1).width = w));
    xTitle(w2, `${c.name} — ${wk}`);
    xHeader(w2, ["Player", "Deal", "", "P&L", "Tips", "BBJ", "Tipback", "Settlement", "Your margin"], 3);
    playerSheetRows(w2, c.playersC.filter((p) => p.played), false);
    if (c.clubAdj !== 0) { const ar = w2.addRow([]); xText(ar.getCell(1), "Club adjustments", { mute: true }); xMoney(ar.getCell(8), c.clubAdj); }
    const cr = w2.addRow([]);
    xText(cr.getCell(1), "CLUB SETTLEMENT (you ↔ club)", { bold: true });
    xMoney(cr.getCell(8), c.clubSettlement, { bold: true });
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
// Fish Tank used to run on its own engine; it's now an ordinary club (Ak/Jon 50/50, fees, stake rake to pool).
const FT_CLUB = { id: "fishtank", name: "Fish Tank", cfgKey: "fishtank-club-v1", weekKey: "fishtank-club-week-v1", meId: "ak",
  owners: [{ id: "ak", label: "Ak", poolPct: 50, jpPct: 50 }, { id: "jon", label: "Jon", poolPct: 50, jpPct: 50 }] };
// Old Fish Tank config/week → owner-club config/week. The old keys are left untouched.
function ftToClub(f, wk) {
  const collectors = {};
  Object.entries(f.assignments || {}).forEach(([k, v]) => { collectors[k.startsWith("u:") ? "dl:" + k.slice(2) : k] = v; });
  // The club engine rolls agent-only players (no SA) up to their agent line; carry the collector over.
  (wk?.players || []).forEach((p) => {
    if (p.saId === "-" && p.agentId !== "-" && collectors[`p:${p.memberId}`] && !collectors[`ag:${p.agentId}`]) collectors[`ag:${p.agentId}`] = collectors[`p:${p.memberId}`];
  });
  const cfg = { ...AA_DEFAULT_CFG,
    defaultTB: f.defaultTB ?? 80, deals: { ...(f.saDeals || {}) }, ownDeals: { ...(f.playerDeals || {}) },
    playerTr: Object.fromEntries(Object.entries(f.actionTax || {}).map(([k, v]) => [k, v.pct || 0])),
    collectors, backed: { ...(f.backed || {}) }, ownAccounts: { ak: f.ownAccounts?.ak || [], jon: f.ownAccounts?.jon || [] },
    finalizedPeriods: { ...(f.finalizedPeriods || {}) }, names: { ...(f.names || {}) },
    dlUmbrellas: (f.umbrellas || []).map((u) => ({ id: u.id, name: u.name, memberKeys: (u.saIds || []).map((id) => `sa:${id}`) })),
    fees: f.fees || [], feeBase: f.feeBase || "net", defaultClass: "agent", stakeMarginToPool: true };
  const week = wk ? { players: wk.players || null, period: wk.period || "", jackpotFromExport: null } : null;
  return { cfg, week };
}
async function loadOwnerClubs() {
  let clubs = null;
  try {
    const c = await store.get(OC_LIST_KEY);
    if (c?.value) { const v = JSON.parse(c.value); if (Array.isArray(v.clubs)) clubs = v.clubs; }
  } catch (e) {}
  // First run: seed from the legacy All American keys (present or not — it's the default club).
  let dirty = false;
  if (!clubs) { clubs = [OC_LEGACY_CLUB]; dirty = true; }
  // One-time: move Fish Tank onto the shared club engine.
  if (!clubs.some((c) => c.id === FT_CLUB.id)) {
    try {
      const fc = await store.get("fishtank-config-v4");
      if (fc?.value) {
        const fw = await store.get("fishtank-lastweek-v4");
        const { cfg, week } = ftToClub(JSON.parse(fc.value), fw?.value ? JSON.parse(fw.value) : null);
        await store.set(FT_CLUB.cfgKey, JSON.stringify(cfg));
        if (week) await store.set(FT_CLUB.weekKey, JSON.stringify(week));
        clubs = [FT_CLUB, ...clubs]; dirty = true;
      }
    } catch (e) {}
  }
  if (dirty) { try { await store.set(OC_LIST_KEY, JSON.stringify({ clubs })); } catch (e) {} }
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
      const tag = (cfg.owners || {})[e.key] || (e.type === "sa" || e.type === "agent" || e.type === "dlUmbrella" ? "agent" : cfg.defaultClass || soleOwner);
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
  // Stake rake margin goes to the backer by default, or into the pool when the club says so (Fish Tank).
  const stakeToPool = !!cfg.stakeMarginToPool;
  const stakePoolMargin = stakeToPool ? backedEntities.reduce((a, e) => a + e.margin, 0) : 0;
  const grossPool = entities.filter((e) => e.tag === "agent").reduce((a, e) => a + e.margin, 0) + stakePoolMargin;
  const poolTR = entities.filter((e) => e.tag === "agent").reduce((a, e) => a + e.trCut, 0);
  const personalMargin = zero();
  entities.forEach((e) => { if (e.tag && e.tag !== "agent" && personalMargin[e.tag] != null) personalMargin[e.tag] += e.margin; });
  const dealMargin = zero();
  if (!stakeToPool) backedEntities.forEach((e) => ownerIds.forEach((o) => (dealMargin[o] += e.margin * shareOf(e.backer, o))));
  // Club fees (accountant etc.) come off the pool's rake profit (TR excluded) before it's split.
  const allFee = entities.reduce((a, e) => a + e.fee, 0) + backedEntities.reduce((a, e) => a + e.fee, 0) + ownRows.reduce((a, p) => a + p.fee, 0);
  const feeBase = cfg.feeBase === "gross" ? allFee : grossPool - poolTR;
  const feeRows = (cfg.fees || []).map((f) => ({ ...f, amount: f.kind === "fixed" ? (+f.amount || 0) : (feeBase * (+f.pct || 0)) / 100 }));
  const totalFees = feeRows.reduce((a, f) => a + f.amount, 0);
  const pool = grossPool - totalFees;
  const feeToOwner = zero();
  feeRows.forEach((f) => { if (feeToOwner[f.recipient] != null) feeToOwner[f.recipient] += f.amount; });
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
  const profit = Object.fromEntries(ownerIds.map((o) => [o, poolShares[o] + personalMargin[o] + dealMargin[o] + ownPosition[o] + jpShares[o] + feeToOwner[o]]));

  // ——— Collections (actuals) ———
  // The jackpot cash sits in the jackpot pool (held by no one), so the
  // settle-up compares collections against profit MINUS the jackpot shares.
  // Own-account P&L stays IN the transfer: an owner's winnings are funded by
  // the week's collections, same as the Fish Tank Ak/Jon recon.
  const actual = zero();
  entities.forEach((e) => { if (e.collector && actual[e.collector] != null) actual[e.collector] += e.unionCash; });
  backedEntities.forEach((e) => ownerIds.forEach((o) => (actual[o] += e.unionCash * shareOf(e.backer, o))));
  // External fees are cash paid out of collections by whoever pays them.
  feeRows.forEach((f) => { if (f.recipient !== "external") return; ownerIds.forEach((o) => (actual[o] -= f.paidBy === "split" ? f.amount * H.poolShare(o) : f.paidBy === o ? f.amount : 0)); });

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
    pool, grossPool, feeRows, totalFees, feeToOwner, poolShares, personalMargin, dealMargin, totalPersonalMargin, totalDealMargin, jackpot, jpShare, jpShares, jpSharesRaw, jpHoldMoves, jpInCollections, profit, cashProfit, actual, delta, transfers, imbalance, balanceOk,
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

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Stake deals (makeup)</div>
      <div className="fit" style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Makeup in</th>
            <th style={th}>RB % normal / in makeup</th><th style={th}>Player %</th>
            <th style={th}>P&L</th><th style={th}>Net (P&L+RB)</th><th style={th}>Player gets</th><th style={th}>Backer book</th><th style={th}>Makeup after</th><th style={th}>Unified</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {stakePlayers.length === 0 && <tr><td colSpan={12} style={{ ...tdL, color: C.mute, padding: 14 }}>No stake deals yet — add one below.</td></tr>}
            {stakePlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{b.name}{!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}</td>
                  <td style={tdL}>{backerSel(k, b)}</td>
                  <td style={{ ...td, color: (b.makeup || 0) > 0.005 ? C.red : C.ink }}><NumInput width={80} value={b.makeup} onChange={(v) => setB(k, { makeup: v })} /></td>
                  <td style={td}><span style={{ display: "inline-flex", gap: 4, alignItems: "center" }}><PctInput width={44} max={999} value={b.rbNormal} onChange={(v) => v != null && setB(k, { rbNormal: v })} /><PctInput width={44} max={999} value={b.rbMakeup} onChange={(v) => v != null && setB(k, { rbMakeup: v })} /></span></td>
                  <td style={td}><PctInput width={46} value={b.playerProfitPct ?? 50} onChange={(v) => v != null && setB(k, { playerProfitPct: v })} /></td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td} title={e ? `RB ${fmt(e.rbCredit)} @${e.rb}% on ${fmt(e.fee)} tips` : ""}>{e ? money(e.net) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? money(e.backerBook) : "—"}</td>
                  <td style={td}>{e ? fmt(e.makeupAfter) : fmt(b.makeup || 0)}</td>
                  <td style={td}>
                    <input placeholder="deal id" value={b.unifiedDealId || ""} onChange={(ev) => setB(k, { unifiedDealId: ev.target.value.trim() })} style={{ ...inputS, width: 70, fontSize: 11 }} title="Paste a Tabs → Staking makeup deal's ID to fold this player's weekly net into that shared, ongoing pool instead of tracking makeup locally here. Use a 'settle per session' deal so it updates alongside any mid-week manual results in the order they happened." />
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
      <div className="fit" style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
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
      <Notes><div>Matched by exact username — a staked player inside an agent tree is pulled out of the tree onto their deal automatically.
        {" "}<b>Stake (makeup)</b>: week net = P&L + RB credit; above makeup the player is paid their % of the excess, below it the net accrues to makeup. RB rate follows makeup <b>entering</b> the week; Lock once to roll it forward. Unlock to fix something, then lock again to re-snapshot — if this week's already in Tabs, re-sync it from Tabs → Bookkeeping afterward.
        {" "}<b>Action buy</b>: the backer owns their % of (P&L + rakeback); the player settles the remainder.
        {" "}The rake margin on these lines goes to the <b>backer personally</b> (never the pool), and the deal P&L below is <b>between backer and player only</b> — it never enters the owner settle-up.
        {" "}Paste a Tabs → Staking makeup deal's ID into <b>Unified deal</b> on a row to fold that player's weekly net into one shared, ongoing pool across sites (and manual/external games) instead of tracking makeup locally here — the local balance freezes and the deal's own % staked / chop handle the tab. Use a "settle per session" deal (the default) so this week's report and any mid-week manual results each update the running makeup as they're logged, in the order they happened, rather than getting batched into one lump sum.</div></Notes>
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
        Click a row to expand. Owner reports show only that owner's own play, personal players, and shared agents.
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
function OwnerClubSetup({ club, clubs, saveClubs, onDeleteClub, cfg, up }) {
  const [name, setName] = useState(club.name);
  useEffect(() => setName(club.name), [club.id, club.name]);
  const patch = (p) => saveClubs(clubs.map((c) => (c.id === club.id ? { ...c, ...p } : c)));
  const setOwner = (id, p) => patch({ owners: club.owners.map((o) => (o.id === id ? { ...o, ...p } : o)) });
  const addOwner = () => { const id = "o-" + uid(); patch({ owners: [...club.owners, { id, label: "New owner", poolPct: 0, jpPct: 0 }] }); };
  const delOwner = (id) => { if (club.owners.length <= 1) return; if (!window.confirm("Remove this owner? Lines tagged to them will need re-tagging.")) return; patch({ owners: club.owners.filter((o) => o.id !== id), meId: club.meId === id ? club.owners.find((o) => o.id !== id)?.id : club.meId }); };
  const H = ocHelpers(club);
  return (
    <div style={{ marginTop: 6 }}>
      <Card title="Club">
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <span style={{ fontSize: 13, color: C.mute }}>Name</span>
          <input value={name} onChange={(e) => setName(e.target.value)} onBlur={() => name.trim() && name.trim() !== club.name && patch({ name: name.trim() })}
            onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()} style={{ ...inputS, width: 220, fontWeight: 700 }} />
          <span style={{ marginLeft: "auto" }}>
            <Btn tone="ghost" small onClick={() => { if (window.confirm(`Delete club "${club.name}" and all its saved weeks, deals, and tags? Archived files stay in Archive.`)) onDeleteClub(club.id); }}>Delete club</Btn>
          </span>
        </div>
      </Card>
      <div style={{ height: 14 }} />
      <Card title="Owners" right={<Btn tone="ghost" small onClick={addOwner}>+ Owner</Btn>}>
        <div style={{ fontSize: 12.5, color: C.mute, marginBottom: 10 }}><b>Pool %</b> splits the shared pool · <b>JP %</b> splits the bad beat jackpot · <b>Me</b> marks your seat.</div>
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
      {cfg && up && <><div style={{ height: 14 }} /><ClubRules club={club} cfg={cfg} up={up} /></>}
    </div>
  );
}

// Club-wide rules that used to be Fish Tank-only: defaults, fees off the pool, where stake rake goes.
function ClubRules({ club, cfg, up }) {
  const H = ocHelpers(club);
  const fees = cfg.fees || [];
  const setFee = (id, p) => up({ fees: fees.map((f) => (f.id === id ? { ...f, ...p } : f)) });
  const row = (label, ctl) => <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderTop: `1px solid ${C.line}`, fontSize: 13 }}><span style={{ color: C.mute }}>{label}</span><span style={{ marginLeft: "auto" }}>{ctl}</span></div>;
  return (
    <Card title="Club rules">
      {row("Default rakeback (TB %) for new lines", <PctInput value={cfg.defaultTB ?? 80} onChange={(v) => v != null && up({ defaultTB: v })} width={60} max={999} />)}
      {row("Default take rate (TR %)", <PctInput value={cfg.defaultTR ?? 0} onChange={(v) => v != null && up({ defaultTR: v })} width={60} />)}
      {row("Class for new player/manager lines", <select value={cfg.defaultClass || ""} onChange={(e) => up({ defaultClass: e.target.value || undefined })} style={inputS}>
        <option value="">Ask me each time</option>{H.tags.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
      </select>)}
      {row("Rake margin on staked/action players goes to", <select value={cfg.stakeMarginToPool ? "pool" : "backer"} onChange={(e) => up({ stakeMarginToPool: e.target.value === "pool" })} style={inputS}>
        <option value="backer">The backer</option><option value="pool">The shared pool</option>
      </select>)}
      <div style={{ borderTop: `1px solid ${C.line}`, paddingTop: 10, marginTop: 4 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
          <b style={{ fontSize: 13 }}>Fees off the pool</b>
          <select value={cfg.feeBase || "net"} onChange={(e) => up({ feeBase: e.target.value })} style={{ ...inputS, fontSize: 12 }}>
            <option value="net">% of pool rake profit</option><option value="gross">% of total rake</option>
          </select>
          <span style={{ marginLeft: "auto" }}><Btn tone="ghost" small onClick={() => up({ fees: [...fees, { id: "f" + uid(), label: "New fee", pct: 0, recipient: "external", paidBy: "split" }] })}>+ Fee</Btn></span>
        </div>
        {fees.length === 0 && <div style={{ color: C.mute, fontSize: 12.5 }}>No fees.</div>}
        {fees.map((f) => (
          <div key={f.id} style={{ display: "flex", gap: 8, alignItems: "center", padding: "4px 0", flexWrap: "wrap" }}>
            <input value={f.label} onChange={(e) => setFee(f.id, { label: e.target.value })} style={{ ...inputS, width: 180 }} />
            <PctInput value={f.pct ?? 0} onChange={(v) => setFee(f.id, { pct: v ?? 0 })} width={54} />
            <span style={{ fontSize: 12, color: C.mute }}>to</span>
            <select value={f.recipient} onChange={(e) => setFee(f.id, { recipient: e.target.value })} style={{ ...inputS, fontSize: 12 }}>
              <option value="external">someone outside</option>{H.owners.map((o) => <option key={o.id} value={o.id}>{o.label}</option>)}
            </select>
            {f.recipient === "external" && <><span style={{ fontSize: 12, color: C.mute }}>paid by</span>
              <select value={f.paidBy || "split"} onChange={(e) => setFee(f.id, { paidBy: e.target.value })} style={{ ...inputS, fontSize: 12 }}>
                <option value="split">pool split</option>{H.owners.map((o) => <option key={o.id} value={o.id}>{o.label}</option>)}
              </select></>}
            <button onClick={() => up({ fees: fees.filter((x) => x.id !== f.id) })} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
          </div>
        ))}
      </div>
    </Card>
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

  const [editing, setEditing] = useState(null);
  return (
    <div>
      <div style={{ display: "none" }}>
        Some super agents, agents, unlinked players, managers, and masters — including a club owner's own Manager/Master line — are all actually settled through the same person. Group those lines into one DL umbrella and they combine into a single line on Lines & ownership — one class, one collector — while each member keeps its own rate; expand "player overrides" on the merged line to override any subgroup or player underneath.
      </div>
      <Card title="DL umbrellas — lines settled through one person" right={
        <span style={{ display: "flex", gap: 8 }}>
          <input placeholder="New umbrella name…" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUmbrella()} style={{ ...inputS, width: 200 }} />
          <Btn tone="ghost" small onClick={addUmbrella}>+ Create</Btn>
        </span>
      }>
        {umbrellas.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No DL umbrellas yet.</div>}
        {umbrellas.map((u) => (
          <div key={u.id} style={{ borderTop: `1px solid ${C.line}`, padding: "6px 0" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: editing === u.id ? 10 : 0 }}>
              <input value={u.name} onChange={(e) => renameUmbrella(u.id, e.target.value)} style={{ ...inputS, fontWeight: 700, width: 200 }} />
              <Pill tone="blue">{u.memberKeys.length} member{u.memberKeys.length !== 1 ? "s" : ""}</Pill>
              <span style={{ fontSize: 12, color: C.mute, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 480 }}>{u.memberKeys.map((k) => (assignable.find((o) => o.key === k) || {}).name || cfg.names?.[k.split(":")[1]] || k).join(", ")}</span>
              <button onClick={() => setEditing(editing === u.id ? null : u.id)} style={{ marginLeft: "auto", border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontSize: 12, fontWeight: 700 }}>{editing === u.id ? "done" : "edit members"}</button>
              <button onClick={() => deleteUmbrella(u.id)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
            </div>
            {editing === u.id && groups.map(([label, list]) => (
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
            {editing === u.id && groups.length === 0 && <div style={{ color: C.mute, fontSize: 12.5 }}>No super agents, agents, or unlinked players on this week's export yet.</div>}
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
  const [showJp, setShowJp] = useState(false);
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
        try { await archiveWeekData(club, cfg, players, period, jpFromExport); } catch (e) {}
        try { await archiveWeek(club.id, club.name, period, async () => { await downloadAAWorkbook(model, period, club); for (const o of ownerIds) await downloadAAOwnerExcel(model, o, period, club); }); } catch (e) {}
      }
      const names = { ...cfg.names };
      p.forEach((x) => { names[x.memberId] = x.name; if (x.saId !== "-") names[x.saId] = x.saName; });
      const patch = { names };
      if (jackpotFound) patch.jackpots = { ...cfg.jackpots, [per]: Math.round(jp * 100) / 100 };
      await up(patch);
      setPlayers(p); setPeriod(per); setJpFromExport(jackpotFound ? Math.round(jp * 100) / 100 : null); setTab("settle");
      try { await store.set(club.weekKey, JSON.stringify({ players: p, period: per, jackpotFromExport: jackpotFound ? jp : null })); } catch (e) {}
      try { await archiveWeekData(club, { ...cfg, ...patch }, p, per, jackpotFound ? jp : null); } catch (e) {}
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

  const nzRow = (label, val, opts) => (Math.abs(val || 0) > 0.005 ? ownerRow(label, val, opts) : null);
  const collectorSel = (e) => (
    <select value={e.collector || ""} onChange={(ev) => setCollector(e.key, ev.target.value)} style={{ ...inputS, padding: "4px 6px", fontSize: 12, borderColor: e.collector ? C.line : C.red }}>
      {!e.collector && <option value="">— who collects? —</option>}
      {ownerIds.map((o) => <option key={o} value={o}>{lbl(o)}</option>)}
    </select>
  );

  return (
    <div style={{ padding: "18px clamp(10px, 2vw, 26px) 60px", maxWidth: 1600, margin: "0 auto" }}>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />

      {err && <div style={{ background: "var(--errBg)", color: C.red, padding: "10px 14px", borderRadius: 6, marginBottom: 14, fontSize: 13.5 }}>{err}</div>}
      {saveNote && <div style={{ background: C.banner, color: C.goldDark, padding: "8px 14px", borderRadius: 6, marginBottom: 14, fontSize: 12.5 }}>{saveNote}</div>}

      {!players && (
        <div style={{ background: C.card, border: `1px dashed ${C.gold}`, borderRadius: 10, padding: "50px 30px", textAlign: "center" }}>
          <div style={{ fontFamily: "Georgia, serif", fontSize: 20, marginBottom: 8 }}>{club.name}</div>
          <div style={{ color: C.mute, fontSize: 14, marginBottom: 18 }}>Upload this week's ClubGG export to start. Set owners and club rules in Setup.</div>
          <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
            <Btn onClick={() => fileRef.current?.click()}>Choose file</Btn>
            <Btn tone="ghost" onClick={() => setTab("setup")}>Owners & setup</Btn>
          </div>
          {tab === "setup" && <div style={{ textAlign: "left", marginTop: 20 }}><OwnerClubSetup club={club} clubs={clubs} saveClubs={saveClubs} onDeleteClub={onDeleteClub} cfg={cfg} up={up} /></div>}
        </div>
      )}

      {players && model && (
        <>
          <div style={{ display: "flex", gap: 4, marginBottom: 16, borderBottom: `2px solid ${C.line}`, flexWrap: "wrap" }}>
            {[["settle", "Lines & deals"], ["backed", "Stakes & action"], ["reports", "Reports"], ["owners", "Settle-up"], ["setup", "Setup"]].map(([k, label]) => (
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
              <Btn tone="ghost" small onClick={async () => { await archiveWeekData(club, cfg, players, period, jpFromExport); await archiveWeek(club.id, club.name, period, async () => { await downloadAAWorkbook(model, period, club); for (const o of ownerIds) await downloadAAOwnerExcel(model, o, period, club); }); setSaveNote(`Archived ${period} — see the Archive tab.`); }}>Archive this week</Btn>
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
              <div style={{ display: "flex", alignItems: "center", marginBottom: 10, gap: 8, flexWrap: "wrap" }}>
                <div style={{ fontFamily: "Georgia, serif", fontSize: 18 }}>Lines & deals</div>
                <span style={{ color: C.mute, fontSize: 12.5 }}>{model.entities.length} lines · set rate, class, and who collects</span>
                <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
                  <Btn tone="gold" small onClick={() => {
                    if (!model.ready) { window.alert(`Finish review first — ${model.untagged.length} untagged line(s), ${model.uncollected.length} without a collector.`); return; }
                    downloadAAWorkbook(model, period, club);
                  }}>Download Excel</Btn>
                  <Btn tone="ghost" small onClick={exportCsv}>Copy table</Btn>
                </div>
              </div>
              <div className="fit" style={{ background: C.card, borderRadius: 10, boxShadow: "0 1px 6px rgba(0,0,0,0.15)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: C.cream }}>
                    <th style={{ ...th, textAlign: "left" }}>Line</th>
                    <th style={{ ...th, textAlign: "left" }}>Class</th>
                    <th style={{ ...th, textAlign: "left" }}>Collected by</th>
                    <th style={th}>TB %</th>
                    <th style={th}>TR %</th>
                    <th style={th}>Winnings</th><th style={th}>Tips</th>
                    <th style={th}>Rakeback</th>
                    <th style={{ ...th }}>Margin → to</th><th style={th}>Club cash</th>
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
                              {open ? "hide" : "players"}
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
                        <td style={td}><PctInput value={cfg.deals[e.id] ?? cfg.defaultTB} onChange={(v) => setDeal(e.id, v)} width={46} max={999} /></td>
                        <td style={td}><PctInput value={cfg.tr?.[e.id] ?? cfg.defaultTR ?? 0} onChange={(v) => v != null && setTr(e.id, v)} width={42} /></td>
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
                  {model.totalFees > 0.005 && <span>Fees <b style={{ color: "var(--barRed)" }}>{fmt(model.totalFees)}</b></span>}
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

              <div style={{ height: 14 }} />
              <AADLUmbrellasTab model={model} cfg={cfg} up={up} />
              <AAOwnAccounts model={model} cfg={cfg} up={up} />
              <Notes>
                <div><b>Margin</b> = tips − rakeback = the profit on the line, plus any TR cut. The rakeback itself is always paid out to the agent or player at their TB%. <b>TR</b> is an extra cut of that line's net (P&L + rakeback) on top of the rake margin — e.g. a "70/10" deal (70% TB, 10% TR) — off by default; set "Default TR %" above for a club where every line runs on one (skips needing a separate stake per player), or override it per line. <b>Agent</b> margins feed the owner pool ({H.poolLabel}); <b>personal</b> margins go 100% to that owner (his player, his spread), and he collects the line himself. Owners' own accounts are handled below; staked/action players sit on the Stake & action tab. Every player rolls up onto their super agent or, lacking one, their agent — "player overrides" on a line lets you pin one player (or a whole agent) to their own rate, e.g. a VIP deal that shouldn't follow their agent's rate. <b>Managers</b> and <b>Masters</b> sit outside that hierarchy and show up as their own lines.</div>
                <div><b>DL umbrellas</b> fold several super agents / agents / players who all settle through one person into a single line. Each member keeps its own rate underneath (open "players" on the merged line).</div>
                <div><b>Class</b>: agent lines feed the shared pool ({H.poolLabel}); personal lines route 100% to their owner. <b>Collected by</b> is who actually takes the cash for that line.</div>
              </Notes>
            </>
          )}

          {tab === "backed" && <AABackedTab model={model} cfg={cfg} up={up} period={period} />}

          {tab === "reports" && <AAReportsTab model={model} cfg={cfg} period={period} expanded={repExpanded} setExpanded={setRepExpanded} club={club} />}
          {tab === "setup" && (
            <>
              <OwnerClubSetup club={club} clubs={clubs} saveClubs={saveClubs} onDeleteClub={onDeleteClub} cfg={cfg} up={up} />
              {((cfg.jpHolds || []).length > 0 || Object.values(cfg.jackpots || {}).some((v) => Math.abs(+v || 0) > 0.005)) && <><div style={{ height: 14 }} /><AABBJHoldsCard model={model} cfg={cfg} up={up} period={period} /></>}
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
                    {model.ready && (model.balanceOk ? " Balance check: ✓ books tie out." : ` ⚠ Books off by ${fmt(model.imbalance)} — the export's P&L doesn't net to its rake (promos, uncollected pots, or jackpot drop inside P&L).`)}
                  </div>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(310px, 1fr))", gap: 14, marginBottom: 18 }}>
                {(Math.abs(jackpot) > 0.005 || showJp || Object.values(cfg.jackpots || {}).some((v) => Math.abs(+v || 0) > 0.005)) ? <Card title="Bad beat jackpot">
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
                </Card> : <div><button onClick={() => setShowJp(true)} style={{ border: `1px dashed ${C.line}`, background: "none", color: C.mute, borderRadius: 10, padding: "14px 18px", cursor: "pointer", fontSize: 13, width: "100%" }}>+ Enter a bad beat jackpot for this week</button></div>}

                <Card title="Club economics">
                  {ownerRow("Rake collected", model.clubRevenue)}
                  {ownerRow("Rakeback → pool lines", -model.agentRB)}
                  {nzRow("Rakeback → players on personal lines", -model.personalRB)}
                  {nzRow("RB credits → stake/action deals", -model.backedRB)}
                  {nzRow("Feeback → owners' own accounts", -model.ownFeeback)}
                  {model.feeRows.filter((x) => x.amount).map((x) => ownerRow(`${x.label} fee (${x.pct}%)`, -x.amount))}
                  {ownerRow(`Pool after fees (${H.poolLabel})`, model.pool, { bold: true })}
                  {nzRow("Personal margins (owner-routed)", model.totalPersonalMargin, { bold: true })}
                  {nzRow("Stake/action rake margins → backer", model.totalDealMargin, { bold: true })}
                </Card>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(310px, 1fr))", gap: 14 }}>
                {ownerIds.map((o) => (
                  <Card key={o} title={`${lbl(o)} — profit`}>
                    {H.poolShare(o) > 0 && ownerRow(`Pool share (${H.pctS(H.poolShare(o))})`, model.poolShares[o])}
                    {nzRow("Personal-line margin · 100%", model.personalMargin[o])}
                    {nzRow("Stake/action rake margin · his deals", model.dealMargin[o])}
                    {nzRow("Own accounts P&L", model.ownPosition[o])}
                    {Math.abs(model.feeToOwner[o]) > 0.005 && ownerRow("Fees paid to him", model.feeToOwner[o])}
                    {(() => {
                      const heldAway = model.jpHoldMoves.find((m) => m.forOwnerId === o);
                      const heldFor = model.jpHoldMoves.filter((m) => m.holderId === o);
                      const label = heldAway ? `Jackpot share — held by ${lbl(heldAway.holderId)}` : heldFor.length ? `Jackpot share (${H.pctS(H.jpShareOf(o))}) + holding ${heldFor.map((m) => lbl(m.forOwnerId)).join(", ")}'s BBJ` : `Jackpot share (${H.pctS(H.jpShareOf(o))})`;
                      return Math.abs(model.jpShares[o]) > 0.005 ? ownerRow(label, model.jpShares[o], { rule: true }) : <div style={{ borderBottom: `1px solid ${C.line}` }} />;
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
              <Notes>
                <div><b>Transfers</b> compare what each owner collected against his collectable share. Stake/action books settle separately. Own-account P&L is in the transfer — an owner's winnings are paid out of the week's collections.</div>
                <div><b>Pool</b> = agent-line margins (plus stake rake if the club routes it there), minus club fees, split by pool %. Personal-line margins go 100% to their owner.</div>
                <div><b>Bad beat jackpot</b> splits by JP %. If the cash sits in the jackpot pool it's excluded from the transfer; if it came out of players' P&L the collectors hold it and it settles through the transfer.</div>
              </Notes>
            </>
          )}
        </>
      )}
    </div>
  );
}

// ———————————————— Bankroll tracker ————————————————
// Your own sessions (live/online, cash/tournament) plus your weekly personal play from the clubs.
const BR_KEY = "bankroll-v1";
const BR_GAMES = ["NLH", "PLO", "PLO5", "PLO Hi-Lo", "Big O", "Mixed", "Other"];
const BR_EMPTY = { start: 0, adjustments: [], sessions: [], hideClub: false };
async function loadBankroll() { try { const c = await store.get(BR_KEY); if (c?.value) return { ...BR_EMPTY, ...JSON.parse(c.value) }; } catch (e) {} return { ...BR_EMPTY }; }
async function saveBankroll(b) { try { await store.set(BR_KEY, JSON.stringify(b)); } catch (e) {} }
const periodEnd = (period) => { const m = String(period || "").match(/~\s*(\d{4}-\d\d-\d\d)/); return m ? m[1] : today(); };
// "9/14 - 9/20" style My Clubs week key → end date this year (or last year if that lands in the future).
const mcWeekEnd = (wk) => { const m = String(wk).match(/(\d{1,2})\/(\d{1,2})\s*$/); if (!m) return today(); const y = new Date().getFullYear(); let d = `${y}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`; if (d > today() && +m[1] > new Date().getMonth() + 2) d = `${y - 1}${d.slice(4)}`; return d; };
// Weekly personal play from every club week we have data for (current + archived) and every My Clubs week.
async function collectClubPlay(clubs) {
  const out = [];
  const idx = await loadArchiveIndex();
  for (const club of clubs) {
    const weeks = {};
    try { const d = await store.get(club.weekKey), cf = await store.get(club.cfgKey); if (d?.value) { const w = JSON.parse(d.value); if (w.players?.length) weeks[w.period] = { players: w.players, cfg: { ...AA_DEFAULT_CFG, ...(cf?.value ? JSON.parse(cf.value) : {}) } }; } } catch (e) {}
    for (const w of idx.weeks.filter((x) => x.site === club.id && x.hasData && !weeks[x.period])) { const a = await loadArchivedWeek(club.id, w.period); if (a?.players?.length) weeks[w.period] = { players: a.players, cfg: { ...AA_DEFAULT_CFG, ...a.cfg } }; }
    Object.entries(weeks).forEach(([period, w]) => {
      const m = buildAAModel(w.players, w.cfg, period, club);
      const mine = m.own.filter((p) => p.owner === club.meId);
      if (!mine.length) return;
      out.push({ key: `${club.id}|${period}`, date: periodEnd(period), venue: club.name, result: r2(mine.reduce((a, p) => a + p.position, 0)), hands: mine.reduce((a, p) => a + (p.hands || 0), 0), note: `${[...new Set(mine.map((p) => p.name))].join(", ")} · ${bookLabel(period)}` });
    });
  }
  try {
    const c = await store.get("agentclubs-v3");
    if (c?.value) {
      const acfg = normalizeAgent({ ...AGENT_DEFAULT, ...JSON.parse(c.value) });
      const myAcc = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
      Object.keys(acfg.weeks || {}).forEach((wk) => {
        const m = computeAgent(acfg, acfg.weeks[wk]);
        const rows = m.allPlayers.filter((p) => p.played && myAcc.has(p.name.trim().toLowerCase()));
        if (!rows.length) return;
        out.push({ key: `mc|${wk}`, date: mcWeekEnd(wk), venue: "My Clubs", result: r2(rows.reduce((a, p) => a + p.settlement, 0)), hands: 0, note: `${[...new Set(rows.map((p) => p.name))].join(", ")} · ${wk}` });
      });
    }
  } catch (e) {}
  return out;
}

function Bankroll({ clubs }) {
  const [b, setB] = useState(null);
  const [club, setClub] = useState([]);
  const [range, setRange] = useState("all");
  const [mode, setMode] = useState("bankroll");
  const [form, setForm] = useState(null);
  const [filter, setFilter] = useState("all");
  useEffect(() => { (async () => { setB(await loadBankroll()); setClub(await collectClubPlay(clubs)); })(); }, []);
  if (!b) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;
  const save = async (next) => { setB(next); await saveBankroll(next); };

  // Every result as one list: manual sessions + club weeks (unless hidden).
  const all = [
    ...b.sessions.map((x) => ({ ...x, result: r2((+x.cashOut || 0) - (+x.buyIn || 0)), manual: true })),
    ...(b.hideClub ? [] : club.map((x) => ({ ...x, id: x.key, type: "club", game: "NLH", start: x.date + "T23:59" }))),
  ].sort((x, y) => (x.start || x.date).localeCompare(y.start || y.date));
  const shown = all.filter((x) => filter === "all" || (filter === "club" ? x.type === "club" : filter === x.type));
  const dayOf = (x) => (x.start || x.date || "").slice(0, 10);
  const now = today();
  const cut = range === "month" ? now.slice(0, 8) + "01" : range === "30d" ? new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10) : range === "ytd" ? now.slice(0, 4) + "-01-01" : "";
  const inRange = shown.filter((x) => !cut || dayOf(x) >= cut);
  const adjTotal = r2(b.adjustments.reduce((a, x) => a + (+x.amount || 0), 0));
  const results = r2(all.reduce((a, x) => a + x.result, 0));
  const bankroll = r2((+b.start || 0) + adjTotal + results);
  const monthRes = r2(all.filter((x) => dayOf(x) >= now.slice(0, 8) + "01").reduce((a, x) => a + x.result, 0));
  const hrs = inRange.filter((x) => x.durationMin).reduce((a, x) => a + x.durationMin / 60, 0);
  const rangeRes = r2(inRange.reduce((a, x) => a + x.result, 0));
  const wins = inRange.filter((x) => x.result > 0).length;
  // Curve: bankroll or cumulative profit over time.
  let run = mode === "bankroll" ? (+b.start || 0) : 0;
  const events = [...all.map((x) => ({ d: dayOf(x), v: x.result })), ...(mode === "bankroll" ? b.adjustments.map((x) => ({ d: x.date, v: +x.amount || 0 })) : [])].sort((x, y) => x.d.localeCompare(y.d));
  const pts = events.map((e) => { run = r2(run + e.v); return { label: e.d.slice(5).replace("-", "/"), v: run, d: e.d }; }).filter((p) => !cut || p.d >= cut);
  const byKey = (f) => { const m = {}; inRange.forEach((x) => { const k = f(x) || "—"; m[k] = m[k] || { k, n: 0, res: 0, hrs: 0 }; m[k].n++; m[k].res = r2(m[k].res + x.result); m[k].hrs += (x.durationMin || 0) / 60; }); return Object.values(m).sort((a, c) => c.res - a.res); };
  const months = (() => { const m = {}; inRange.forEach((x) => { const k = dayOf(x).slice(0, 7); m[k] = r2((m[k] || 0) + x.result); }); return Object.entries(m).sort().map(([k, v]) => ({ label: new Date(k + "-15").toLocaleString("en-US", { month: "short", year: "2-digit" }), values: { v } })); })();

  const blank = { type: "cash", game: "NLH", stakes: "", venue: "", start: new Date(Date.now() - new Date().getTimezoneOffset() * 6e4).toISOString().slice(0, 16), durationMin: 120, buyIn: "", cashOut: "", notes: "" };
  const commit = async () => {
    const x = { ...form, id: form.id || uid(), buyIn: +form.buyIn || 0, cashOut: +form.cashOut || 0, durationMin: +form.durationMin || 0 };
    await save({ ...b, sessions: [...b.sessions.filter((s2) => s2.id !== x.id), x] }); setForm(null);
  };
  const fld = (label, ctl) => <label style={{ display: "grid", gap: 4, fontSize: 12, color: C.mute }}>{label}{ctl}</label>;
  const chip = (on, label, onClick) => <button key={label} onClick={onClick} style={{ border: `1px solid ${on ? C.gold : C.line}`, background: on ? C.gold : "transparent", color: on ? "var(--onGold)" : C.ink, borderRadius: 16, padding: "5px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>{label}</button>;
  const venues = [...new Set(all.map((x) => x.venue).filter(Boolean))];
  const stakesList = [...new Set(b.sessions.map((x) => x.stakes).filter(Boolean))];

  return (
    <div style={{ padding: "18px clamp(10px, 2vw, 26px) 60px", maxWidth: 1400, margin: "0 auto", display: "grid", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 2fr) minmax(280px, 1fr)", gap: 14, alignItems: "start" }}>
        <div style={{ background: C.card, borderRadius: 12, padding: "16px 18px", boxShadow: "0 1px 6px rgba(0,0,0,0.15)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {chip(mode === "bankroll", "Bankroll", () => setMode("bankroll"))}{chip(mode === "profit", "Profit", () => setMode("profit"))}
            <span style={{ marginLeft: "auto", fontSize: 13, color: monthRes >= 0 ? C.green : C.red, fontWeight: 700 }}>{monthRes >= 0 ? "↗ +" : "↘ "}{fmt(monthRes)} this month</span>
          </div>
          <div style={{ fontFamily: "Georgia, serif", fontSize: 40, margin: "8px 0 2px" }}>{fmt(mode === "bankroll" ? bankroll : results)}</div>
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 8 }}>{fmt(+b.start || 0)} start · {adjTotal >= 0 ? "+" : ""}{fmt(adjTotal)} adjustments · {results >= 0 ? "+" : ""}{fmt(results)} results</div>
          {pts.length > 1 ? <LineChart points={pts} height={200} label={mode === "bankroll" ? "bankroll" : "profit"} /> : <div style={{ color: C.mute, fontSize: 13, padding: 20 }}>Log a session to start the chart.</div>}
        </div>
        <div style={{ display: "grid", gap: 10 }}>
          <Btn tone="gold" onClick={() => setForm({ ...blank })}>+ Log session</Btn>
          <Card title="Bankroll settings">
            <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, marginBottom: 8 }}><span style={{ color: C.mute }}>Starting bankroll</span><span style={{ marginLeft: "auto" }}><NumInput value={b.start} onChange={(v) => save({ ...b, start: v || 0 })} width={100} /></span></div>
            <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12.5, color: C.mute }}><input type="checkbox" checked={!b.hideClub} onChange={(e) => save({ ...b, hideClub: !e.target.checked })} /> Include weekly club personal play ({club.length} weeks)</label>
            <div style={{ fontWeight: 700, fontSize: 12.5, margin: "10px 0 4px" }}>Deposits / withdrawals</div>
            {b.adjustments.map((a) => <div key={a.id} style={{ display: "flex", gap: 8, fontSize: 12.5, padding: "3px 0" }}><span style={{ color: C.mute }}>{a.date}</span><span>{a.note}</span><span style={{ marginLeft: "auto" }}>{money(+a.amount || 0)}</span><button onClick={() => save({ ...b, adjustments: b.adjustments.filter((x) => x.id !== a.id) })} style={{ border: "none", background: "none", color: C.red, cursor: "pointer" }}>×</button></div>)}
            <Btn tone="ghost" small onClick={() => { const v = window.prompt("Amount (negative for a withdrawal):"); if (v && !isNaN(+v)) save({ ...b, adjustments: [...b.adjustments, { id: uid(), date: today(), amount: +v, note: window.prompt("Note (optional):") || "" }] }); }}>+ Adjustment</Btn>
          </Card>
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        {[["month", "Month"], ["30d", "30d"], ["ytd", "YTD"], ["all", "All"]].map(([k, l]) => chip(range === k, l, () => setRange(k)))}
        <span style={{ width: 12 }} />
        {[["all", "All"], ["cash", "Cash"], ["tournament", "Tournaments"], ["club", "Club weeks"]].map(([k, l]) => chip(filter === k, l, () => setFilter(k)))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12 }}>
        {[["Profit", money(rangeRes)], ["Sessions", inRange.length], ["Win %", inRange.length ? Math.round((wins / inRange.length) * 100) + "%" : "—"], ["Hours", hrs ? hrs.toFixed(0) : "—"], ["$ / hour", hrs ? money(r2(inRange.filter((x) => x.durationMin).reduce((a, x) => a + x.result, 0) / hrs)) : "—"], ["Avg session", inRange.length ? money(r2(rangeRes / inRange.length)) : "—"]].map(([l, v]) => (
          <div key={l} style={{ background: C.card, borderRadius: 10, padding: "12px 14px", boxShadow: "0 1px 6px rgba(0,0,0,0.15)" }}>
            <div style={{ fontSize: 11.5, color: C.mute, textTransform: "uppercase", letterSpacing: ".06em" }}>{l}</div>
            <div style={{ fontFamily: "Georgia, serif", fontSize: 24, marginTop: 4 }}>{v}</div>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 14 }}>
        <Card title="By month">{months.length ? <BarChart groups={months} series={[{ key: "v", label: "Result", color: "var(--gold)" }]} colorBySign height={180} /> : <div style={{ color: C.mute, fontSize: 13 }}>No results in this range.</div>}</Card>
        {[["By venue", (x) => x.venue], ["By stakes", (x) => (x.type === "club" ? "club week" : x.stakes ? `${x.game} ${x.stakes}` : x.game)]].map(([title, f]) => (
          <Card key={title} title={title}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={{ ...th, textAlign: "left" }}></th><th style={th}>Sessions</th><th style={th}>Profit</th><th style={th}>$ / hr</th></tr></thead>
              <tbody>{byKey(f).slice(0, 8).map((r) => <tr key={r.k} style={{ borderTop: `1px solid ${C.line}` }}><td style={tdL}>{r.k}</td><td style={td}>{r.n}</td><td style={td}>{money(r.res)}</td><td style={td}>{r.hrs ? fmt(r.res / r.hrs, 0) : "—"}</td></tr>)}</tbody>
            </table>
          </Card>
        ))}
      </div>

      <Card title={`Sessions · ${inRange.length}`}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><th style={{ ...th, textAlign: "left" }}>Date</th><th style={{ ...th, textAlign: "left" }}>Type</th><th style={{ ...th, textAlign: "left" }}>Game · stakes</th><th style={{ ...th, textAlign: "left" }}>Venue</th><th style={th}>Hours</th><th style={th}>Buy-in</th><th style={th}>Cash out</th><th style={th}>Result</th><th style={{ ...th, textAlign: "left" }}>Notes</th><th style={th}></th></tr></thead>
          <tbody>{[...inRange].reverse().map((x, i) => (
            <tr key={x.id} style={{ borderTop: `1px solid ${C.line}`, background: i % 2 ? C.rowAlt : "transparent" }}>
              <td style={tdL}>{dayOf(x)}</td>
              <td style={tdL}>{x.type === "club" ? <Pill tone="blue">club week</Pill> : x.type === "tournament" ? <Pill>tournament</Pill> : <Pill tone="green">cash</Pill>}</td>
              <td style={tdL}>{x.game}{x.stakes ? ` ${x.stakes}` : ""}</td>
              <td style={tdL}>{x.venue}</td>
              <td style={td}>{x.durationMin ? (x.durationMin / 60).toFixed(1) : "—"}</td>
              <td style={td}>{x.manual ? fmt(+x.buyIn || 0) : "—"}</td>
              <td style={td}>{x.manual ? fmt(+x.cashOut || 0) : "—"}</td>
              <td style={{ ...td, fontWeight: 700 }}>{money(x.result)}</td>
              <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{x.notes || x.note || ""}</td>
              <td style={td}>{x.manual && <><button onClick={() => setForm({ ...x })} style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer" }}>✎</button><button onClick={() => window.confirm("Delete this session?") && save({ ...b, sessions: b.sessions.filter((s2) => s2.id !== x.id) })} style={{ border: "none", background: "none", color: C.red, cursor: "pointer" }}>×</button></>}</td>
            </tr>
          ))}</tbody>
        </table>
      </Card>
      <Notes><div>Bankroll = starting bankroll + deposits/withdrawals + every result. Club weeks are your own-account results from each club (current week plus every archived week with saved numbers) and your My Clubs accounts; they update on their own. $/hour only counts sessions with a duration.</div></Notes>

      {form && (
        <div onClick={() => setForm(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.55)", display: "grid", placeItems: "center", zIndex: 50 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.card, color: C.ink, borderRadius: 14, padding: 22, width: "min(460px, 92vw)", display: "grid", gap: 12, boxShadow: "0 12px 40px rgba(0,0,0,.4)" }}>
            <div style={{ display: "flex", alignItems: "center" }}><div style={{ fontFamily: "Georgia, serif", fontSize: 20 }}>{form.id ? "Edit session" : "Log session"}</div><button onClick={() => setForm(null)} style={{ marginLeft: "auto", border: "none", background: "none", color: C.mute, fontSize: 20, cursor: "pointer" }}>×</button></div>
            <div style={{ display: "flex", gap: 6 }}>{chip(form.type === "cash", "Cash", () => setForm({ ...form, type: "cash" }))}{chip(form.type === "tournament", "Tournament", () => setForm({ ...form, type: "tournament" }))}</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>{BR_GAMES.map((g) => chip(form.game === g, g, () => setForm({ ...form, game: g })))}</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              {fld(form.type === "cash" ? "Stakes" : "Event / buy-in level", <input list="br-stakes" value={form.stakes} onChange={(e) => setForm({ ...form, stakes: e.target.value })} placeholder={form.type === "cash" ? "10/20" : "Sunday Major"} style={inputS} />)}
              {fld("Venue", <input list="br-venues" value={form.venue} onChange={(e) => setForm({ ...form, venue: e.target.value })} placeholder="Rivers Casino" style={inputS} />)}
              {fld("Start", <input type="datetime-local" value={form.start} onChange={(e) => setForm({ ...form, start: e.target.value })} style={inputS} />)}
              {fld("Duration (hours)", <input type="number" step="0.25" value={form.durationMin ? form.durationMin / 60 : ""} onChange={(e) => setForm({ ...form, durationMin: Math.round((+e.target.value || 0) * 60) })} style={inputS} />)}
              {fld(form.type === "cash" ? "Buy-in (total)" : "Buy-in + rebuys", <input type="number" value={form.buyIn} onChange={(e) => setForm({ ...form, buyIn: e.target.value })} style={inputS} />)}
              {fld(form.type === "cash" ? "Cash out" : "Prize", <input type="number" value={form.cashOut} onChange={(e) => setForm({ ...form, cashOut: e.target.value })} style={inputS} />)}
            </div>
            {fld("Notes", <input value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} style={inputS} />)}
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontSize: 13, color: C.mute }}>Result {money(r2((+form.cashOut || 0) - (+form.buyIn || 0)))}</span>
              <span style={{ marginLeft: "auto" }}><Btn tone="gold" onClick={commit}>{form.id ? "Save" : "Log session"}</Btn></span>
            </div>
            <datalist id="br-venues">{venues.map((v) => <option key={v} value={v} />)}</datalist>
            <datalist id="br-stakes">{stakesList.map((v) => <option key={v} value={v} />)}</datalist>
          </div>
        </div>
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
  // One name = one tab: fold a person's separate player/club tabs together.
  if (!d.nameMerged) {
    const keep = {}, remap = {};
    d.counterparties.forEach((c) => { const k = c.name.trim().toLowerCase(); if (!keep[k]) keep[k] = c; else remap[c.id] = keep[k].id; });
    if (Object.keys(remap).length) {
      const m = (id) => remap[id] || id;
      const kept = new Set(Object.values(remap));
      d.entries = d.entries.map((e) => ({ ...e, cpId: m(e.cpId) }));
      d.staking = { ...d.staking, deals: d.staking.deals.map((x) => ({ ...x, cpId: m(x.cpId) })) };
      d.settleChecklist = d.settleChecklist.map((x) => ({ ...x, cpId: m(x.cpId) }));
      d.counterparties = d.counterparties.filter((c) => !remap[c.id]).map((c) => (kept.has(c.id) ? { ...c, kind: "player" } : c));
    }
    d.nameMerged = true; changed = true;
  }
  if (changed) { try { await store.set(TABS_KEY, JSON.stringify(d)); } catch (e) {} }
  return d;
}
async function saveTabs(data) { try { await store.set(TABS_KEY, JSON.stringify(data)); } catch (e) {} }

function applyWeekToTabsData(data, sourceKey, weekLabel, items) {
  if (data.pushed[sourceKey]) return data;
  const byKey = {}; data.counterparties.forEach((cp) => (byKey[cp.name.toLowerCase()] = cp));
  const d = today();
  const checklist = data.settleChecklist || (data.settleChecklist = []);
  const onList = new Set(checklist.map((x) => x.cpId));
  items.filter((it) => Math.abs(it.amount) > 0.005).forEach((it) => {
    const key = it.name.toLowerCase();
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
  const byKey = {}; data.counterparties.forEach((cp) => (byKey[cp.name.toLowerCase()] = cp));
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
      const cp = byKey[f.cp.name.toLowerCase()];
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
    pending.push({ sourceKey: club.id === "fishtank" ? `ft:${aa.period}` : `aa:${club.id === "allamerican" ? "" : club.id + ":"}${aa.period}`, label: `${site} · ${aa.period}`, items, stakingItems, unifiedFeeds });
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
      const makeupBefore = makeup;
      if (deal.type === "makeup") {
        if (net < 0) makeup = r2(makeup - net);
        else { recovered = Math.min(net, makeup); makeup = r2(makeup - recovered); excess = net - recovered; akChop = r2(excess * akPct / 100); playerChop = r2(excess - akChop); }
        chopped = r2(chopped + akChop);
      } else netAB = r2(netAB + net);
      const playerEnt = unstaked + playerChop;
      const tab = r2(held - playerEnt);
      if (Math.abs(tab) > 0.005) derived.push({ id: "stk-" + refId, date: when, cpId: deal.cpId, amount: tab, note: label, source: "staking", dealId: deal.id, refId });
      return { recovered, excess, akChop, playerChop, tab, makeupBefore, makeupAfter: makeup };
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
        partial.forEach((p, i) => rows.push({ ...p, weekKey: w, weekNet: r2(net), ...(i === partial.length - 1 ? x : { recovered: null, excess: null, akChop: null, playerChop: null, tab: null, makeupBefore: null, makeupAfter: null }) }));
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

// A "Unified deal" (Tabs → Staking) can be fed by several sites at once —
// Fish Tank, an owner club, a My Clubs sub-club — folding what would
// otherwise be several separate local makeup balances into one real backing
// relationship. Book Summary wants just that deal's OWN most-recent week,
// not its lifetime chopped/makeup totals, plus which sites actually fed it
// that week (for the "(Fish Tank, Honey Pot, Betflix)" label). Returns null
// for a deal with no dated rows at all (nothing to report).
function unifiedDealWeekly(d) {
  const rows = (d.rows || []).filter((r) => r.date);
  if (!rows.length) return null;
  const thisWeek = rows.reduce((mx, r) => { const w = weekOf(r.date); return !mx || w > mx ? w : mx; }, null);
  const weekRows = rows.filter((r) => weekOf(r.date) === thisWeek);
  const sites = [...new Set(weekRows.map((r) => r.game).filter(Boolean))];
  if (d.deal.type === "action") {
    return { name: d.name, kind: "action", net: r2(weekRows.reduce((a, r) => a + (r.net || 0), 0)), sites, thisWeek };
  }
  // Weekly-cadence rows only carry real akChop/makeupBefore/makeupAfter on
  // the last row of their week group (others are nulled out above); session-
  // cadence rows each carry their own. Either way: sum akChop across the
  // week for chop, and take the balance at the start of the week vs. the end
  // for accrual — not just the latest row's makeupAfter, so multiple
  // sessions within the same week net against each other correctly.
  const settled = weekRows.filter((r) => r.makeupAfter != null).sort((a, b) => a.date.localeCompare(b.date));
  if (!settled.length) return null;
  const chop = r2(weekRows.reduce((a, r) => a + (r.akChop || 0), 0));
  const accrued = Math.max(0, r2(settled[settled.length - 1].makeupAfter - settled[0].makeupBefore));
  return { name: d.name, kind: "makeup", chop, accrued, sites, thisWeek };
}

const TABS_VIEWS = [["balances", "Balances"], ["bookkeeping", "Weekly imports"], ["ledger", "Ledger"], ["staking", "Staking"], ["vig", "Vig"], ["misc", "Misc. P&L"], ["players", "People"]];
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

// Weekly totals history — same pattern as My Clubs' `weeks` object: a plain
// label → snapshot map, persisted so a past week's numbers stay lookupable
// after that week's Fish Tank/owner-club/Tabs data has been overwritten by
// the next one. Unlike My Clubs there's no "current week" to type into —
// everything here is either the live, freshly-computed totals, or a
// snapshot someone explicitly saved — so this only ever stores `weeks`.
const BOOK_WEEKS_KEY = "book-weeks-v1";
async function loadBookWeeks() {
  try { const c = await store.get(BOOK_WEEKS_KEY); if (c?.value) { const v = JSON.parse(c.value); if (v && typeof v.weeks === "object") return v.weeks; } } catch (e) {}
  return {};
}
async function saveBookWeeks(weeks) { try { await store.set(BOOK_WEEKS_KEY, JSON.stringify({ weeks })); } catch (e) {} }

// Book: this week's numbers with charts, plus a history of saved weeks.
const bookLabel = (period) => { const m = String(period || "").match(/(\d{4})-(\d\d)-(\d\d)\s*~\s*(\d{4})-(\d\d)-(\d\d)/); return m ? `${m[2]}/${m[3]} - ${m[5]}/${m[6]}` : ""; };
const shortLabel = (l) => bookLabel(l) || l;
const bookSortKey = (w) => (String(w?.period || "").match(/\d{4}-\d\d-\d\d/) || [w?.priorWeekStart || w?.savedAt || ""])[0];
// Older snapshots kept Fish Tank separate from the club list; fold it in so every week reads the same.
const bookClubs = (t) => [...((Math.abs(t.ftPersonal || 0) + Math.abs(t.ftFee || 0)) > 0.005 && !(t.ocRows || []).some((o) => o.name === "Fish Tank") ? [{ name: "Fish Tank", personal: t.ftPersonal || 0, fee: t.ftFee || 0 }] : []), ...(t.ocRows || []),
  ...((Math.abs(t.myPlayTotal || 0) + Math.abs(t.mcMargin || 0)) > 0.005 ? [{ name: "My Clubs", personal: t.myPlayTotal || 0, fee: t.mcMargin || 0 }] : [])];
const bookGrand = (t) => t.grandTotal ?? r2((t.clubTotal || 0) + (t.stakingVigMiscTotal || 0));

function BookSection() {
  const [subtab, setSubtab] = useState("week");
  const [loaded, setLoaded] = useState(false);
  const [ownerClubs, setOwnerClubs] = useState([]);
  const [agent, setAgent] = useState(null);
  const [persons, setPersons] = useState([]);
  const [tabs, setTabs] = useState(null);
  const [bookWeeks, setBookWeeks] = useState({});
  const [viewWeek, setViewWeek] = useState(""); // "" = live; else a saved week's label
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => { (async () => {
    setLoaded(false);
    const [ocM, agM, ppl, tb, bw] = await Promise.all([loadAllOwnerClubModels(), loadAgentModel(), loadPersons(), loadTabs(), loadBookWeeks()]);
    setOwnerClubs(ocM); setAgent(agM); setPersons(ppl); setTabs(tb); setBookWeeks(bw);
    setLoaded(true);
  })(); }, [refreshKey]);

  const liveTotals = useMemo(() => computeBookTotals({ ft: null, ownerClubs, agent, tabs, persons }), [ownerClubs, agent, tabs, persons]);
  const labels = Object.keys(bookWeeks).sort((x, y) => bookSortKey(bookWeeks[x]).localeCompare(bookSortKey(bookWeeks[y])));
  const liveLabel = bookLabel(liveTotals.period);
  const liveKey = labels.find((l) => shortLabel(l) === liveLabel) || liveLabel;
  const shown = viewWeek && bookWeeks[viewWeek] ? bookWeeks[viewWeek] : liveTotals;
  const shownLabel = viewWeek ? shortLabel(viewWeek) : liveLabel;
  const prevLabel = labels.filter((l) => bookSortKey(bookWeeks[l]) < bookSortKey(shown) && shortLabel(l) !== shownLabel).pop();

  const saveThisWeek = async () => {
    const label = liveKey || window.prompt("Label for this week (e.g. 09/01 - 09/07):", "");
    if (!label) return;
    if (bookWeeks[label] && !window.confirm(`"${shortLabel(label)}" is already saved — overwrite it with today's numbers?`)) return;
    const next = { ...bookWeeks, [label]: { ...liveTotals, savedAt: new Date().toISOString() } };
    setBookWeeks(next); await saveBookWeeks(next);
  };
  const deleteWeek = async (label) => {
    if (!window.confirm(`Delete the saved week "${shortLabel(label)}"? This can't be undone.`)) return;
    const next = { ...bookWeeks }; delete next[label];
    setBookWeeks(next); if (viewWeek === label) setViewWeek(""); await saveBookWeeks(next);
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 4, padding: "10px clamp(10px, 2vw, 26px) 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap", alignItems: "center" }}>
        {[["week", "Week"], ["history", `History${labels.length ? ` · ${labels.length}` : ""}`]].map(([k, label]) => (
          <button key={k} onClick={() => setSubtab(k)} style={{ border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700, background: subtab === k ? C.card : "transparent", color: subtab === k ? C.ink : C.mute, borderRadius: "8px 8px 0 0", marginBottom: -2 }}>{label}</button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", paddingBottom: 6 }}>
          {subtab === "week" && <select value={viewWeek} onChange={(e) => setViewWeek(e.target.value)} style={{ ...inputS, fontSize: 12.5 }}>
            <option value="">This week (live){liveLabel ? ` · ${liveLabel}` : ""}</option>
            {[...labels].reverse().map((k) => <option key={k} value={k}>{shortLabel(k)}</option>)}
          </select>}
          {subtab === "week" && !viewWeek && (bookWeeks[liveKey] ? <Pill tone="green">saved</Pill> : null)}
          <Btn tone="gold" small onClick={saveThisWeek}>{bookWeeks[liveKey] ? "Re-save this week" : "Save this week"}</Btn>
          <button onClick={() => setRefreshKey((k) => k + 1)} title="Reload club data" style={{ border: `1px solid ${C.line}`, background: "none", color: C.mute, cursor: "pointer", borderRadius: 6, padding: "4px 10px", fontSize: 12 }}>↻</button>
        </div>
      </div>
      <div style={{ padding: "18px clamp(10px, 2vw, 26px) 60px", maxWidth: 1400, margin: "0 auto" }}>
        {!loaded ? <div style={{ color: C.mute, padding: 20 }}>Loading…</div> : subtab === "week"
          ? <BookWeek t={shown} label={shownLabel} prev={prevLabel ? bookWeeks[prevLabel] : null} prevLabel={prevLabel && shortLabel(prevLabel)} saved={!!viewWeek} />
          : <BookHistory weeks={bookWeeks} labels={labels} open={(l) => { setViewWeek(l); setSubtab("week"); }} del={deleteWeek} />}
      </div>
    </div>
  );
}

// ——— small SVG charts (one y-axis, zero baseline, hover tooltip) ———
const shortMoney = (v) => { const a = Math.abs(v); return (v < 0 ? "-" : "") + (a >= 1000 ? (a / 1000).toFixed(a >= 10000 ? 0 : 1) + "k" : a.toFixed(0)); };
function niceTicks(lo, hi) {
  const span = hi - lo || 1, step0 = span / 4, mag = Math.pow(10, Math.floor(Math.log10(step0)));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((x) => x >= step0);
  const out = []; let v = Math.floor(lo / step) * step; out.push(v);
  while (v < hi - 1e-9) { v += step; out.push(v); }
  return out;
}
// Chart width follows its container so text stays a normal size.
function useWidth(init = 800) {
  const ref = useRef(null); const [w, setW] = useState(init);
  useEffect(() => { if (!ref.current) return; const ro = new ResizeObserver(([e]) => setW(Math.max(280, e.contentRect.width))); ro.observe(ref.current); return () => ro.disconnect(); }, []);
  return [ref, w];
}
// groups: [{ label, values: { key: number } }], series: [{ key, label, color }]; colorBySign colors one series green/red.
function BarChart({ groups, series, height = 220, colorBySign = false }) {
  const [hover, setHover] = useState(null);
  const [ref, W] = useWidth();
  const H = height, padL = 46, padB = 26, padT = 10;
  const vals = groups.flatMap((g) => series.map((s) => g.values[s.key] || 0));
  const ticks = niceTicks(Math.min(0, ...vals), Math.max(0, ...vals));
  const lo = ticks[0], hi = ticks[ticks.length - 1];
  const y = (v) => padT + ((hi - v) / (hi - lo || 1)) * (H - padT - padB);
  const gw = (W - padL) / Math.max(1, groups.length), bw = Math.max(3, Math.min(34, (gw * 0.6) / series.length));
  return (
    <div ref={ref} style={{ position: "relative" }}>
      {series.length > 1 && <div style={{ display: "flex", gap: 14, fontSize: 12, color: C.mute, marginBottom: 6 }}>{series.map((s) => <span key={s.key} style={{ display: "inline-flex", alignItems: "center", gap: 5 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: s.color }} />{s.label}</span>)}</div>}
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }} onMouseLeave={() => setHover(null)}>
        {ticks.map((t) => <g key={t}><line x1={padL} x2={W} y1={y(t)} y2={y(t)} stroke="var(--line)" strokeWidth={t === 0 ? 1.5 : 1} /><text x={padL - 6} y={y(t) + 4} textAnchor="end" fontSize="11" fill="var(--mute)">{shortMoney(t)}</text></g>)}
        {groups.map((g, i) => {
          const x0 = padL + i * gw + (gw - bw * series.length - 2 * (series.length - 1)) / 2;
          return (
            <g key={g.label} onMouseEnter={() => setHover(i)}>
              <rect x={padL + i * gw} y={padT} width={gw} height={H - padT - padB} fill={hover === i ? "var(--rowAlt)" : "transparent"} />
              {series.map((s, j) => { const v = g.values[s.key] || 0, top = Math.min(y(v), y(0)), h = Math.max(1, Math.abs(y(v) - y(0)));
                return <rect key={s.key} x={x0 + j * (bw + 2)} y={top} width={bw} height={h} rx={Math.min(4, bw / 2)} fill={colorBySign ? (v >= 0 ? "var(--green)" : "var(--red)") : s.color} />; })}
              {(groups.length <= 16 || i % Math.ceil(groups.length / 16) === 0) && <text x={padL + i * gw + gw / 2} y={H - 8} textAnchor="middle" fontSize="11" fill="var(--mute)">{g.label}</text>}
            </g>
          );
        })}
      </svg>
      {hover != null && groups[hover] && (
        <div style={{ position: "absolute", top: 24, left: `${Math.min(80, ((padL + hover * gw + gw) / W) * 100)}%`, background: C.card, border: `1px solid ${C.line}`, borderRadius: 6, padding: "6px 10px", fontSize: 12, pointerEvents: "none", boxShadow: "0 2px 8px rgba(0,0,0,.25)", whiteSpace: "nowrap" }}>
          <b>{groups[hover].label}</b>
          {series.map((s) => <div key={s.key} style={{ display: "flex", gap: 10 }}><span style={{ color: C.mute }}>{s.label}</span><span style={{ marginLeft: "auto" }}>{money(groups[hover].values[s.key] || 0)}</span></div>)}
        </div>
      )}
    </div>
  );
}
function LineChart({ points, height = 180, label }) {
  const [hover, setHover] = useState(null);
  const [ref, W] = useWidth();
  const H = height, padL = 46, padB = 26, padT = 10;
  const ticks = niceTicks(Math.min(0, ...points.map((p) => p.v)), Math.max(0, ...points.map((p) => p.v)));
  const lo = ticks[0], hi = ticks[ticks.length - 1];
  const x = (i) => padL + 14 + (points.length <= 1 ? (W - padL - 28) / 2 : (i / (points.length - 1)) * (W - padL - 42));
  const y = (v) => padT + ((hi - v) / (hi - lo || 1)) * (H - padT - padB);
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }} onMouseLeave={() => setHover(null)}
        onMouseMove={(e) => { const r = e.currentTarget.getBoundingClientRect(); const px = ((e.clientX - r.left) / r.width) * W; let best = 0; points.forEach((_, i) => { if (Math.abs(x(i) - px) < Math.abs(x(best) - px)) best = i; }); setHover(best); }}>
        {ticks.map((t) => <g key={t}><line x1={padL} x2={W} y1={y(t)} y2={y(t)} stroke="var(--line)" strokeWidth={t === 0 ? 1.5 : 1} /><text x={padL - 6} y={y(t) + 4} textAnchor="end" fontSize="11" fill="var(--mute)">{shortMoney(t)}</text></g>)}
        <polyline fill="none" stroke="var(--s1)" strokeWidth="2" points={points.map((p, i) => `${x(i)},${y(p.v)}`).join(" ")} />
        {points.map((p, i) => <circle key={i} cx={x(i)} cy={y(p.v)} r={hover === i ? 5 : 3.5} fill="var(--s1)" stroke="var(--card)" strokeWidth="2" />)}
        {hover != null && <line x1={x(hover)} x2={x(hover)} y1={padT} y2={H - padB} stroke="var(--mute)" strokeDasharray="3 3" />}
        {points.map((p, i) => (points.length <= 16 || i % Math.ceil(points.length / 16) === 0) && <text key={"l" + i} x={x(i)} y={H - 8} textAnchor="middle" fontSize="11" fill="var(--mute)">{p.label}</text>)}
      </svg>
      {hover != null && <div style={{ position: "absolute", top: 8, left: `${Math.min(78, (x(hover) / W) * 100)}%`, background: C.card, border: `1px solid ${C.line}`, borderRadius: 6, padding: "6px 10px", fontSize: 12, pointerEvents: "none", whiteSpace: "nowrap" }}><b>{points[hover].label}</b> · {label} {money(points[hover].v)}</div>}
    </div>
  );
}

const statTile = (label, v, prev, color) => (
  <div style={{ background: C.card, borderRadius: 10, padding: "14px 16px", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", borderTop: `3px solid ${color}` }}>
    <div style={{ fontSize: 11.5, color: C.mute, textTransform: "uppercase", letterSpacing: ".06em" }}>{label}</div>
    <div style={{ fontFamily: "Georgia, serif", fontSize: 26, marginTop: 4, color: v > 0.005 ? C.green : v < -0.005 ? C.red : C.ink }}>{fmt(v)}</div>
    {prev != null && <div style={{ fontSize: 11.5, color: C.mute, marginTop: 2 }}>{v - prev >= 0 ? "▲" : "▼"} {fmt(Math.abs(v - prev))} vs last saved week</div>}
  </div>
);

function BookWeek({ t, label, prev, prevLabel, saved }) {
  const clubs = bookClubs(t);
  const line = (l, v, opts = {}) => <div style={{ display: "flex", padding: "5px 0", fontSize: 13, borderTop: opts.rule ? `1px solid ${C.line}` : "none", fontWeight: opts.bold ? 700 : 400 }}><span style={{ color: opts.bold ? C.ink : C.mute }}>{l}</span><span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(v || 0)}</span></div>;
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 20 }}>{label ? `Week ${label}` : "This week"}</div>
        <span style={{ color: C.mute, fontSize: 12.5 }}>{saved ? `saved ${t.savedAt ? new Date(t.savedAt).toLocaleDateString() : ""}` : "live — from the clubs' current weeks"}</span>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
        {statTile("Total", bookGrand(t), prev ? bookGrand(prev) : null, "var(--gold)")}
        {statTile("Personal play", t.personalTotal || 0, prev ? prev.personalTotal : null, "var(--s1)")}
        {statTile("Rake profit", t.rakeProfitTotal || 0, prev ? prev.rakeProfitTotal : null, "var(--s2)")}
        {statTile("Staking · vig · misc", t.stakingVigMiscTotal || 0, prev ? prev.stakingVigMiscTotal : null, "var(--s3)")}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 3fr) minmax(260px, 2fr)", gap: 14, alignItems: "start" }}>
        <Card title="By club">
          {clubs.length ? <BarChart groups={clubs.map((c) => ({ label: c.name, values: { personal: c.personal, fee: c.fee } }))} series={[{ key: "personal", label: "Personal play", color: "var(--s1)" }, { key: "fee", label: "Rake profit", color: "var(--s2)" }]} /> : <div style={{ color: C.mute, fontSize: 13 }}>No club weeks loaded.</div>}
          <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 8 }}>
            <thead><tr><th style={{ ...th, textAlign: "left" }}>Club</th><th style={th}>Personal play</th><th style={th}>Rake profit</th><th style={th}>Total</th></tr></thead>
            <tbody>{clubs.map((c) => <tr key={c.name} style={{ borderTop: `1px solid ${C.line}` }}><td style={tdL}><b>{c.name}</b> <span style={{ color: C.mute, fontSize: 11 }}>{bookLabel(c.period)}</span></td><td style={td}>{money(c.personal)}</td><td style={td}>{money(c.fee)}</td><td style={td}>{money(r2(c.personal + c.fee))}</td></tr>)}</tbody>
          </table>
        </Card>
        <Card title="Staking, vig & misc">
          {line("Makeup chopped profit", t.makeupChopTotal)}
          {line("Action buys net", t.actionNetTotal)}
          {line(`Crypto vig${t.priorWeekStart ? ` (${t.priorWeekStart.slice(5)} – ${(t.priorWeekEnd || "").slice(5)})` : ""}`, t.vigWeekTotal)}
          {line("Misc P&L", t.miscWeekTotal)}
          {line("Total", t.stakingVigMiscTotal, { rule: true, bold: true })}
          <div style={{ fontSize: 12, color: C.mute, marginTop: 6 }}>Makeup still owed (open, not counted): <b style={{ color: C.goldDark }}>{fmt(t.makeupAccruedTotal || 0)}</b></div>
          {(t.makeupRows || []).length + (t.actionRows || []).length > 0 && (
            <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 10 }}>
              <thead><tr><th style={{ ...th, textAlign: "left" }}>Player</th><th style={th}>Chop / net</th><th style={th}>Makeup added</th></tr></thead>
              <tbody>
                {(t.makeupRows || []).map((r) => <tr key={"m" + r.name} style={{ borderTop: `1px solid ${C.line}` }}><td style={tdL}>{r.name} <span style={{ color: C.mute, fontSize: 11 }}>stake</span></td><td style={td}>{money(r.chop)}</td><td style={td}>{fmt(r.accrued)}</td></tr>)}
                {(t.actionRows || []).map((r) => <tr key={"a" + r.name} style={{ borderTop: `1px solid ${C.line}` }}><td style={tdL}>{r.name} <span style={{ color: C.mute, fontSize: 11 }}>action</span></td><td style={td}>{money(r.net)}</td><td style={td}>—</td></tr>)}
              </tbody>
            </table>
          )}
          {((t.vigRows || []).length + (t.miscRows || []).length) > 0 && <div style={{ marginTop: 10 }}>
            {(t.vigRows || []).filter((r) => Math.abs(r.vig) > 0.005).map((r) => line(`Vig · ${r.name}`, r.vig))}
            {(t.miscRows || []).map((r) => line(`Misc · ${r.category}`, r.amount))}
          </div>}
        </Card>
      </div>
      <Notes>
        <div><b>Total</b> = personal play + rake profit + staking/vig/misc. <b>Personal play</b> is your own accounts across every club. <b>Rake profit</b> is your share of each club's margin (pool share, personal lines, fees paid to you). Staking counts only realized money (chop, action-buy net); makeup still owed is shown but not counted.</div>
        <div>Vig and misc use the last completed Mon–Sun week. Save each week once the clubs are settled so it lands in History{prevLabel ? ` — compared here against ${prevLabel}` : ""}.</div>
      </Notes>
    </div>
  );
}

function BookHistory({ weeks, labels, open, del }) {
  if (!labels.length) return <Card title="No saved weeks yet"><div style={{ color: C.mute, fontSize: 13 }}>Use <b>Save this week</b> once a week is settled — it lands here with charts.</div></Card>;
  let cum = 0;
  const rows = labels.map((key) => { const w = weeks[key]; const g = bookGrand(w); cum = r2(cum + g); return { key, l: shortLabel(key), w, g, cum }; });
  const best = rows.reduce((a, r) => (r.g > a.g ? r : a), rows[0]), worst = rows.reduce((a, r) => (r.g < a.g ? r : a), rows[0]);
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
        {statTile(`Running total · ${rows.length} weeks`, cum, null, "var(--gold)")}
        {statTile("Average week", r2(cum / rows.length), null, "var(--s1)")}
        {statTile(`Best week · ${best.l}`, best.g, null, "var(--green)")}
        {statTile(`Worst week · ${worst.l}`, worst.g, null, "var(--red)")}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))", gap: 14 }}>
        <Card title="Weekly total"><BarChart groups={rows.map((r) => ({ label: r.l.slice(0, 5), values: { g: r.g } }))} series={[{ key: "g", label: "Total", color: "var(--gold)" }]} colorBySign /></Card>
        <Card title="Running total"><LineChart points={rows.map((r) => ({ label: r.l.slice(0, 5), v: r.cum }))} label="running" /></Card>
      </div>
      <Card title="Where each week came from">
        <BarChart groups={rows.map((r) => ({ label: r.l.slice(0, 5), values: { p: r.w.personalTotal || 0, f: r.w.rakeProfitTotal || 0, s: r.w.stakingVigMiscTotal || 0 } }))}
          series={[{ key: "p", label: "Personal play", color: "var(--s1)" }, { key: "f", label: "Rake profit", color: "var(--s2)" }, { key: "s", label: "Staking · vig · misc", color: "var(--s3)" }]} />
      </Card>
      <Card title="Weeks">
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><th style={{ ...th, textAlign: "left" }}>Week</th><th style={th}>Personal play</th><th style={th}>Rake profit</th><th style={th}>Staking · vig · misc</th><th style={th}>Total</th><th style={th}>Running</th><th style={th}></th></tr></thead>
          <tbody>{[...rows].reverse().map((r, i) => (
            <tr key={r.key} style={{ borderTop: `1px solid ${C.line}`, background: i % 2 ? C.rowAlt : "transparent" }}>
              <td style={tdL}><button onClick={() => open(r.key)} style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontWeight: 700, padding: 0, fontSize: "inherit" }}>{r.l}</button></td>
              <td style={td}>{money(r.w.personalTotal || 0)}</td><td style={td}>{money(r.w.rakeProfitTotal || 0)}</td><td style={td}>{money(r.w.stakingVigMiscTotal || 0)}</td>
              <td style={{ ...td, fontWeight: 700 }}>{money(r.g)}</td><td style={td}>{money(r.cum)}</td>
              <td style={td}><button onClick={() => del(r.key)} title="Delete week" style={{ border: "none", background: "none", color: C.red, cursor: "pointer" }}>×</button></td>
            </tr>
          ))}</tbody>
        </table>
      </Card>
    </div>
  );
}

// Every number Book Summary shows, computed once from whatever's currently
// loaded (or, for a saved week, exactly this same shape read back out of
// storage — BookSummary itself doesn't know or care which). Kept a plain,
// JSON-serializable object (no JSX) on purpose: it's exactly what "+ Save
// this week" persists verbatim, so what you save is exactly what you saw.
function computeBookTotals({ ft, ownerClubs, agent, tabs, persons }) {
  // Fish Tank: entitle.ak bundles personal play + ½ profit + staking together —
  // split it apart and drop the staking piece (backed books).
  const ftPersonal = ft ? ft.model.ownPosition.ak : 0;
  const ftStaking = ft ? ft.model.backedBook.ak : 0;
  const ftFee = ft ? r2(ft.model.entitle.ak - ftPersonal - ftStaking) : 0;

  // Owner clubs (Midnight Bazaar + any others added) — Ak's share on each.
  // Unlike Fish Tank, "profit" here already excludes the deal/backed book, so
  // the fee margin is just profit minus own-account play.
  const ocRows = (ownerClubs || []).map((oc) => {
    const meId = oc.club.meId;
    const personal = oc.model.ownPosition[meId] || 0;
    const staking = oc.model.backedBook[meId] || 0;
    const fee = r2((oc.model.profit[meId] || 0) - personal);
    return { name: oc.club.name, period: oc.period, personal, staking, fee };
  });
  const ocPersonalTotal = r2(ocRows.reduce((a, o) => a + o.personal, 0));
  const ocFeeTotal = r2(ocRows.reduce((a, o) => a + o.fee, 0));
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
  const myPlayNames = [...new Set(myPlayRows.map((p) => p.name))];
  const mcMargin = agent ? agent.model.totals.margin : 0;
  const mcAdj = agent ? agent.model.totals.globalAdjTotal : 0;

  // Headline #1 — total P&L playing on personal accounts, everywhere.
  const personalTotal = r2(ftPersonal + ocPersonalTotal + myPlayTotal);
  // Headline #2 — total ClubGG rake profit made (fee/rakeback margins earned
  // running the clubs, no personal play mixed in).
  const rakeProfitTotal = r2(ftFee + ocFeeTotal + mcMargin);
  const clubTotal = r2(personalTotal + rakeProfitTotal);

  // ——— Staking: Makeup deals + Action buys, one row per PLAYER (not per
  // club) across every site that fed them this week. A backed makeup player
  // linked to a "Unified deal" (Tabs → Staking) has their local per-club
  // number excluded here — the moment a unified deal exists, that site's own
  // local balance is frozen/vestigial (see the unifiedDealId wiring
  // elsewhere) — and their real, combined-across-sites weekly number comes
  // from the shared deal instead (unifiedDealWeekly), which is also the only
  // place a My Clubs contribution shows up, since My Clubs has no local
  // staking system of its own — only the option to link a player into a
  // unified deal. Everyone's grouped by their canonical Player Data name so
  // the same person under different site usernames still lands on one line.
  const mapName = makeNameMapper(persons || []);
  const unifiedIdOf = (backedCfg, key) => (backedCfg?.[key.slice(2)]?.unifiedDealId || "").trim();
  const localShareRows = [];
  if (ft) {
    const local = ft.model.backedEntities.filter((e) => !(e.dealType === "makeup" && unifiedIdOf(ft.cfg?.backed, e.key)));
    localShareRows.push(...backedShareRows(local, (backer) => (backer === "split" ? 0.5 : backer === "ak" ? 1 : 0), "Fish Tank", mapName));
  }
  (ownerClubs || []).forEach((oc) => {
    const meId = oc.club.meId;
    const local = oc.model.backedEntities.filter((e) => !(e.dealType === "makeup" && unifiedIdOf(oc.cfg?.backed, e.key)));
    localShareRows.push(...backedShareRows(local, (backer) => oc.model.H.shareOf(backer, meId), oc.club.name, mapName));
  });
  const staking = computeStaking(tabs || TABS_EMPTY, tabs?.counterparties || [], [], persons || []);
  const unifiedWeekly = staking.deals.filter((d) => !d.deal.ended).map(unifiedDealWeekly).filter(Boolean);

  const makeupGroups = new Map(), actionGroups = new Map();
  const bump = (map, name, patch) => { const g = map.get(name) || { name, sites: new Set(), chop: 0, accrued: 0, net: 0 }; patch(g); map.set(name, g); };
  localShareRows.forEach((r) => {
    if (r.kind === "makeup") bump(makeupGroups, r.name, (g) => { g.sites.add(r.site); g.chop = r2(g.chop + r.chopped); g.accrued = r2(g.accrued + Math.max(0, r2(r.makeupAfter - r.makeupBefore))); });
    else bump(actionGroups, r.name, (g) => { g.sites.add(r.site); g.net = r2(g.net + r.net); });
  });
  unifiedWeekly.forEach((u) => {
    if (u.kind === "makeup") bump(makeupGroups, u.name, (g) => { u.sites.forEach((s) => g.sites.add(s)); g.chop = r2(g.chop + u.chop); g.accrued = r2(g.accrued + u.accrued); });
    else bump(actionGroups, u.name, (g) => { u.sites.forEach((s) => g.sites.add(s)); g.net = r2(g.net + u.net); });
  });
  const makeupRows = [...makeupGroups.values()].filter((g) => Math.abs(g.chop) > 0.005 || g.accrued > 0.005).map((g) => ({ ...g, sites: [...g.sites] })).sort((a, b) => b.accrued - a.accrued || b.chop - a.chop);
  const actionRows = [...actionGroups.values()].filter((g) => Math.abs(g.net) > 0.005).map((g) => ({ ...g, sites: [...g.sites] })).sort((a, b) => b.net - a.net);
  const makeupChopTotal = r2(makeupRows.reduce((a, g) => a + g.chop, 0));
  const makeupAccruedTotal = r2(makeupRows.reduce((a, g) => a + g.accrued, 0));
  const actionNetTotal = r2(actionRows.reduce((a, g) => a + g.net, 0));

  // Crypto vig + Misc P&L, prior week only (Tabs → Vig / Misc P&L) — the
  // most recently completed Mon–Sun week, not whatever week is still in
  // progress today.
  const priorWeekStart = (() => { const d = new Date(weekOf(today()) + "T00:00:00"); d.setDate(d.getDate() - 7); return d.toISOString().slice(0, 10); })();
  const priorWeekEnd = (() => { const d = new Date(priorWeekStart + "T00:00:00"); d.setDate(d.getDate() + 6); return d.toISOString().slice(0, 10); })();
  const vigWeekEntries = tabs ? tabs.entries.filter((e) => isCrypto(e.method) && weekOf(e.date) === priorWeekStart) : [];
  const vigWeekTotal = r2(vigWeekEntries.reduce((a, e) => a + (+e.vig || 0), 0));
  const vigGains = r2(vigWeekEntries.filter((e) => e.vig > 0).reduce((a, e) => a + e.vig, 0));
  const vigLosses = r2(vigWeekEntries.filter((e) => e.vig < 0).reduce((a, e) => a + e.vig, 0));
  const cpName = (id) => tabs?.counterparties.find((c) => c.id === id)?.name || "?";
  const vigByCp = {};
  vigWeekEntries.forEach((e) => { vigByCp[e.cpId] = r2((vigByCp[e.cpId] || 0) + (+e.vig || 0)); });
  const vigRows = Object.entries(vigByCp).map(([id, v]) => ({ name: cpName(id), vig: v })).sort((a, b) => b.vig - a.vig);

  const miscWeekEntries = tabs ? (tabs.misc || []).filter((e) => weekOf(e.date) === priorWeekStart) : [];
  const miscWeekTotal = r2(miscWeekEntries.reduce((a, e) => a + (+e.amount || 0), 0));
  const miscByCat = {};
  miscWeekEntries.forEach((e) => { const k = e.category || "uncategorized"; miscByCat[k] = r2((miscByCat[k] || 0) + (+e.amount || 0)); });
  const miscRows = Object.entries(miscByCat).map(([category, amount]) => ({ category, amount })).sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));

  // Headline #3 — total P&L on staking + vig + misc, combined. Only the
  // REALIZED pieces count as P&L here (chop, action-buy net, vig, misc) —
  // makeup still just accruing is deliberately left out, same "not a loss
  // yet" reasoning as the breakdown card below; it's shown there for
  // reference but doesn't feed this total.
  const stakingVigMiscTotal = r2(makeupChopTotal + actionNetTotal + vigWeekTotal + miscWeekTotal);

  return {
    period: ft?.period || (ocRows.find((o) => o.name === "Fish Tank") || ocRows[0] || {}).period || "",
    grandTotal: r2(clubTotal + stakingVigMiscTotal),
    ftPersonal, ftFee, ocRows, ocPersonalTotal, ocFeeTotal, ocLabel, ocFeeLabel,
    myPlayTotal, myPlayNames, mcMargin, mcAdj,
    personalTotal, rakeProfitTotal, clubTotal,
    makeupRows, actionRows, makeupChopTotal, makeupAccruedTotal, actionNetTotal,
    priorWeekStart, priorWeekEnd, vigRows, vigWeekTotal, vigGains, vigLosses,
    miscRows, miscWeekTotal,
    stakingVigMiscTotal,
    hasFt: !!ft, hasAgent: !!agent,
  };
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
  const [tabSearch, setTabSearch] = useState("");
  const [tabSort, setTabSort] = useState("amount");

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
    const key = name.trim().toLowerCase();
    let cp = list.find((c) => c.name.toLowerCase() === key);
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
    const dup = cps.find((c) => c.id !== cp.id && c.name.toLowerCase() === name.trim().toLowerCase());
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
      const k = it.name.toLowerCase();
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
  const SITES = ["", ...new Set(["Fish Tank", ...(clubs || []).map((c) => c.name)]), "My Clubs"];
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
    const matches = cps.filter((c) => names.has(c.name.toLowerCase()) && c.name.toLowerCase() !== per.name.toLowerCase());
    if (!matches.length) { flash(`No separate tab entries found under ${per.name}'s aliases.`); return; }
    if (!window.confirm(`Merge ${matches.map((m) => m.name).join(", ")} into "${per.name}"? Their ledger entries and staking deals move under the one name.`)) return;
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
  const findPersonForCp = (cp) => persons.find((per) => (per.name.toLowerCase() === cp.name.toLowerCase() || per.aliases.some((a) => a.name.toLowerCase() === cp.name.toLowerCase())));
  const setCpNotes = (id, notes) => save({ ...data, counterparties: data.counterparties.map((c) => (c.id === id ? { ...c, notes } : c)) });
  const bundleIntoPlayerData = async (cp) => {
    if (persons.some((per) => per.name.toLowerCase() === cp.name.toLowerCase())) { flash(`${cp.name} is already in People.`); return; }
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


  const exportAll = () => setExportData({ title: "Tabs ledger", text: toTSV(["Date", "Counterparty", "Kind", "Amount", "Method", "Vig", "Note", "Source"],
    sortDateDesc(allEntries, (e) => allEntries.indexOf(e)).map((e) => [e.date, cpById(e.cpId)?.name || "?", cpById(e.cpId)?.kind || "", e.amount.toFixed(2), e.method || "", e.vig != null ? e.vig.toFixed(2) : "", e.note || "", e.source === "week" ? e.week : e.source])) });

  // ——— People: duplicate suggestions, club settle-with names ———
  const normName = (n) => n.toLowerCase().replace(/[^a-z0-9]/g, "").replace(/\d+$/, "");
  const dupPairs = useMemo(() => {
    const out = [], dismissed = new Set(data.dupDismissed || []);
    const samePerson = (a, b) => { const pa = findPersonForCp(a), pb = findPersonForCp(b); return pa && pb && pa.id === pb.id; };
    for (let i = 0; i < cps.length; i++) for (let j = i + 1; j < cps.length; j++) {
      const a = cps[i], b = cps[j], na = normName(a.name), nb = normName(b.name);
      if (na.length < 3 || nb.length < 3) continue;
      const close = na === nb || ((na.startsWith(nb) || nb.startsWith(na)) && Math.min(na.length, nb.length) >= 4 && Math.abs(na.length - nb.length) <= 4);
      if (close && !dismissed.has(a.id + "|" + b.id) && !samePerson(a, b)) out.push([a, b]);
    }
    return out;
  }, [cps, data.dupDismissed, persons]);
  // Merge one tab into another and remember the old name as a username of that person.
  const mergeTabInto = async (from, into) => {
    let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries] };
    next = mergeCps(next, [from.id], into.id);
    let ps = [...persons];
    let per = ps.find((x) => x.name.toLowerCase() === into.name.toLowerCase()) || findPersonForCp(into);
    if (!per) { per = { id: uid(), name: into.name, kind: "player", aliases: [], notes: into.notes || "" }; ps.push(per); }
    const fromPer = ps.find((x) => x.id !== per.id && x.name.toLowerCase() === from.name.toLowerCase());
    const extra = [{ name: from.name, site: "" }, ...(fromPer ? fromPer.aliases : [])].filter((a) => !per.aliases.some((x) => x.name.toLowerCase() === a.name.toLowerCase()) && a.name.toLowerCase() !== per.name.toLowerCase());
    ps = ps.filter((x) => !fromPer || x.id !== fromPer.id).map((x) => (x.id === per.id ? { ...x, aliases: [...x.aliases, ...extra] } : x));
    next.persons = normPersons(ps);
    await save(next);
    flash(`Merged ${from.name} into ${into.name}.`);
  };
  const dismissDup = (a, b) => save({ ...data, dupDismissed: [...(data.dupDismissed || []), a.id + "|" + b.id] });
  const [mcClubs, setMcClubs] = useState([]);
  useEffect(() => { (async () => { try { const c = await store.get("agentclubs-v3"); if (c?.value) setMcClubs(JSON.parse(c.value).clubs || []); } catch (e) {} })(); }, []);
  const setClubOwner = async (clubId, owner) => {
    const c = await store.get("agentclubs-v3"); if (!c?.value) return;
    const a = JSON.parse(c.value);
    a.clubs = (a.clubs || []).map((x) => (x.id === clubId ? { ...x, owner } : x));
    await store.set("agentclubs-v3", JSON.stringify(a)); setMcClubs(a.clubs);
    // Existing tab under the club's own name moves onto the owner's tab.
    const club = a.clubs.find((x) => x.id === clubId);
    const old = cps.find((x) => x.name.toLowerCase() === (club?.name || "").toLowerCase());
    if (owner.trim() && old && old.name.toLowerCase() !== owner.trim().toLowerCase() && window.confirm(`Move the "${old.name}" club tab onto ${owner.trim()}'s tab?`)) {
      let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries] };
      const [list, target] = findOrCreateCp(next.counterparties, owner.trim(), "player");
      next.counterparties = list;
      await save(mergeCps(next, [old.id], target.id));
    }
    flash(`${club?.name} now settles with ${owner.trim() || "the club itself"}.`);
  };
  const [peopleSearch, setPeopleSearch] = useState("");
  // Old weeks logged under a club's own name before its owner was set.
  const strayClubTabs = mcClubs.filter((c) => (c.owner || "").trim() && c.owner.trim().toLowerCase() !== c.name.toLowerCase()).map((c) => [c, cps.find((x) => x.name.toLowerCase() === c.name.toLowerCase())]).filter(([, cp]) => cp);
  const moveStrayClubTabs = async () => {
    if (!window.confirm(`Move ${strayClubTabs.length} club tab(s) onto their owners' tabs? (${strayClubTabs.map(([c]) => `${c.name} → ${c.owner}`).join(", ")})`)) return;
    let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries] };
    strayClubTabs.forEach(([c, cp]) => { const [list, target] = findOrCreateCp(next.counterparties, c.owner.trim(), "player"); next.counterparties = list; next = mergeCps(next, [cp.id], target.id); });
    await save(next);
    flash(`Moved ${strayClubTabs.length} club tabs onto their owners.`);
  };

  const tabRow = (c) => {
          const open = expandedCp === c.id;
          const per = findPersonForCp(c);
          return (
            <div key={c.id} style={{ borderTop: `1px solid ${C.line}` }}>
              <div onClick={() => { setExpandedCp(open ? null : c.id); if (open) setConfirmDeleteCp((x) => (x === c.id ? null : x)); }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 0", fontSize: 13.5, cursor: "pointer" }}>
                <span style={{ color: C.mute, fontSize: 10, width: 10, display: "inline-block", flexShrink: 0 }}>{open ? "▾" : "▸"}</span>
                <span style={{ color: C.ink, fontWeight: 600, fontSize: 13.5 }}>{c.name}</span>
                {iconBtn("✎", (e) => { e.stopPropagation(); renameCp(c); }, C.mute, "Rename (renaming onto an existing name merges)")}
                {c.kind === "club" && <Pill tone="blue">club</Pill>}
                {per && per.aliases.length > 0 && <span style={{ fontSize: 11, color: C.mute, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 260 }}>{per.aliases.map((a) => a.name).join(", ")}</span>}
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

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;
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
      <div style={{ padding: "18px clamp(10px, 2vw, 26px) 60px", maxWidth: 1400, margin: "0 auto" }}>
        {note && <div style={{ background: C.banner, color: C.goldDark, padding: "8px 14px", borderRadius: 6, marginBottom: 14, fontSize: 12.5 }}>{note}</div>}
        <div style={{ color: C.mute, fontSize: 12, marginBottom: 12 }}><b style={{ color: C.green }}>+ they owe you</b> · <b style={{ color: C.red }}>− you owe them</b></div>

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
        {view === "balances" && (() => {
          const q = tabSearch.trim().toLowerCase();
          const rows = cps.filter((c) => (showZero || Math.abs(balances[c.id] || 0) > 0.005 || staking.makeupByCp[c.id]) && (!q || c.name.toLowerCase().includes(q) || (findPersonForCp(c)?.aliases || []).some((a) => a.name.toLowerCase().includes(q))))
            .sort((x, y) => (tabSort === "name" ? x.name.localeCompare(y.name) : Math.abs(balances[y.id] || 0) - Math.abs(balances[x.id] || 0)));
          const owed = r2(cps.reduce((acc, c) => acc + Math.max(0, balances[c.id] || 0), 0));
          const owe = r2(cps.reduce((acc, c) => acc + Math.min(0, balances[c.id] || 0), 0));
          return (
            <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 2fr) minmax(260px, 1fr)", gap: 14, alignItems: "start" }}>
              <Card title={`Tabs · ${rows.length}`} right={
                <span style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <input placeholder="Search name or username…" value={tabSearch} onChange={(e) => setTabSearch(e.target.value)} style={{ ...inputS, width: 190, fontSize: 12 }} />
                  <select value={tabSort} onChange={(e) => setTabSort(e.target.value)} style={{ ...inputS, fontSize: 12 }}><option value="amount">Biggest first</option><option value="name">A–Z</option></select>
                </span>}>
                {rows.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>Nothing open.</div>}
                {rows.map((c) => tabRow(c))}
                <label style={{ fontSize: 12, color: C.mute, display: "flex", gap: 6, alignItems: "center", marginTop: 10 }}>
                  <input type="checkbox" checked={showZero} onChange={(e) => setShowZero(e.target.checked)} /> show settled (zero) tabs
                </label>
              </Card>
              <Card title="Totals">
                {[["They owe you", owed], ["You owe", owe]].map(([l, v]) => <div key={l} style={{ display: "flex", padding: "6px 0", fontSize: 13.5 }}><span style={{ color: C.mute }}>{l}</span><span style={{ marginLeft: "auto" }}>{money(v)}</span></div>)}
                <div style={{ display: "flex", padding: "8px 0", fontSize: 14, borderTop: `1px solid ${C.line}`, fontWeight: 700 }}><span>Net position</span><span style={{ marginLeft: "auto" }}>{money(totalAll)}</span></div>
                <div style={{ fontSize: 10.5, letterSpacing: "0.08em", textTransform: "uppercase", color: C.goldDark, fontWeight: 700, margin: "12px 0 2px" }}>Running totals · not in net position</div>
                {otherTotals.map(([label, v]) => (
                  <div key={label} style={{ display: "flex", padding: "5px 0", borderTop: `1px solid ${C.line}`, fontSize: 12.5 }}>
                    <span style={{ color: C.mute }}>{label}</span><span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(v)}</span>
                  </div>
                ))}
              </Card>
            </div>
          );
        })()}

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
                      <datalist id="tabs-cp-names">{cps.map((c) => <option key={c.id} value={c.name} />)}</datalist>
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
                      <datalist id="tabs-swap-from-names">{cps.map((c) => <option key={c.id} value={c.name} />)}</datalist>
                      <span style={{ color: C.mute }}>→</span>
                      <span style={{ fontSize: 11, color: C.mute, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>To</span>
                      <select value={sw.toKind} onChange={(e) => setSw({ ...sw, toKind: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
                        <option value="player">Player</option><option value="club">Club</option><option value="other">Other</option>
                      </select>
                      <input list="tabs-swap-to-names" placeholder="Who's receiving…" value={sw.to} onChange={(e) => setSw({ ...sw, to: e.target.value })} style={{ ...inputS, width: 150 }} />
                      <datalist id="tabs-swap-to-names">{cps.map((c) => <option key={c.id} value={c.name} />)}</datalist>
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
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
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
                          <td style={{ ...tdL, fontWeight: 600 }}>{cp?.name || "?"}</td>
                          <td style={td}>{money(e.amount)}{e.baseAmount != null && Math.abs(e.baseAmount - e.amount) > 0.005 && <div style={{ fontSize: 10, color: C.mute }}>sent {fmt(e.baseAmount)}</div>}</td>
                          <td style={{ ...tdL, fontSize: 12 }}>{e.method ? <Pill tone={isCrypto(e.method) ? "gold" : "blue"}>{e.method}</Pill> : <span style={{ color: C.mute }}>—</span>}</td>
                          <td style={{ ...td, fontSize: 12 }}>{e.vig != null ? <span style={{ color: e.vig >= 0 ? C.green : C.red }}>{fmt(e.vig)} <span style={{ color: C.mute }}>@{e.vigRate}%</span></span> : ""}</td>
                          <td style={{ ...tdL, fontSize: 12.5, whiteSpace: "normal" }}>{e.note || ""}</td>
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
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
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
                          <table style={{ width: "100%", borderCollapse: "collapse" }}>
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
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
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
          <div style={{ display: "grid", gap: 14 }}>
            {dupPairs.length > 0 && (
              <Card title={`Possible duplicates · ${dupPairs.length}`}>
                {dupPairs.slice(0, 30).map(([x, y]) => (
                  <div key={x.id + y.id} style={{ display: "flex", gap: 8, alignItems: "center", padding: "6px 0", borderTop: `1px solid ${C.line}`, fontSize: 13, flexWrap: "wrap" }}>
                    <b>{x.name}</b> <span style={{ color: C.mute }}>{fmt(balances[x.id] || 0)}</span>
                    <span style={{ color: C.mute }}>and</span>
                    <b>{y.name}</b> <span style={{ color: C.mute }}>{fmt(balances[y.id] || 0)}</span>
                    <span style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
                      <Btn tone="ghost" small onClick={() => mergeTabInto(y, x)}>Merge into {x.name}</Btn>
                      <Btn tone="ghost" small onClick={() => mergeTabInto(x, y)}>Merge into {y.name}</Btn>
                      <button onClick={() => dismissDup(x, y)} style={{ border: "none", background: "none", color: C.mute, cursor: "pointer", fontSize: 12 }}>not the same</button>
                    </span>
                  </div>
                ))}
              </Card>
            )}

            <Card title="People" right={
              <span style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input placeholder="Search…" value={peopleSearch} onChange={(e) => setPeopleSearch(e.target.value)} style={{ ...inputS, width: 150, fontSize: 12 }} />
                <input placeholder="New person…" value={newPerson} onChange={(e) => setNewPerson(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addPerson()} style={{ ...inputS, width: 150, fontSize: 12 }} />
                <Btn tone="gold" small onClick={addPerson}>+ Person</Btn>
              </span>}>
              {persons.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No people yet.</div>}
              {persons.filter((per) => { const q = peopleSearch.trim().toLowerCase(); return !q || per.name.toLowerCase().includes(q) || per.aliases.some((x) => x.name.toLowerCase().includes(q)); })
                .sort((x, y) => x.name.localeCompare(y.name)).map((per) => {
                const cp = cps.find((c) => c.name.toLowerCase() === per.name.toLowerCase());
                const strays = cps.filter((c) => c.id !== cp?.id && per.aliases.some((x) => x.name.toLowerCase() === c.name.toLowerCase()));
                const draft = addAlias[per.id]?.name || "";
                const addUser = () => { const v = draft.trim(); if (v && !per.aliases.some((x) => x.name.toLowerCase() === v.toLowerCase())) patchPerson(per.id, { aliases: [...per.aliases, { name: v, site: addAlias[per.id]?.site || "" }] }); setAddAlias({ ...addAlias, [per.id]: { name: "", site: "" } }); };
                return (
                  <div key={per.id} style={{ display: "grid", gridTemplateColumns: "minmax(140px, 180px) 1fr minmax(160px, 220px) 90px 24px", gap: 10, alignItems: "center", padding: "8px 0", borderTop: `1px solid ${C.line}` }}>
                    <input value={per.name} onChange={(e) => patchPerson(per.id, { name: e.target.value })} style={{ ...inputS, fontWeight: 700, width: "100%", boxSizing: "border-box" }} />
                    <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center" }}>
                      {per.aliases.map((x, i) => (
                        <span key={x.name + i} title={x.site ? `only on ${x.site}` : "any site"} style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "1px 4px 1px 9px", fontSize: 11.5, display: "inline-flex", gap: 4, alignItems: "center" }}>
                          {x.name}{x.site && <span style={{ color: C.mute, fontSize: 10 }}>· {x.site}</span>}
                          <button onClick={() => patchPerson(per.id, { aliases: per.aliases.filter((_, j) => j !== i) })} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 12, padding: 0 }}>×</button>
                        </span>
                      ))}
                      <input list="people-names" placeholder="+ username" value={draft} onChange={(e) => setAddAlias({ ...addAlias, [per.id]: { ...(addAlias[per.id] || {}), name: e.target.value } })} onKeyDown={(e) => e.key === "Enter" && addUser()} onBlur={addUser} style={{ ...inputS, width: 110, fontSize: 11.5, padding: "2px 6px" }} />
                      {strays.length > 0 && <Btn tone="ghost" small onClick={() => mergePerson(per)}>merge {strays.length} tab{strays.length > 1 ? "s" : ""}</Btn>}
                    </div>
                    <input value={per.notes || ""} onChange={(e) => patchPerson(per.id, { notes: e.target.value })} placeholder="Notes (payment, Telegram…)" style={{ ...inputS, fontSize: 12, width: "100%", boxSizing: "border-box" }} />
                    <span style={{ textAlign: "right", fontVariantNumeric: "tabular-nums" }}>{cp ? money(balances[cp.id] || 0) : <span style={{ color: C.mute, fontSize: 12 }}>no tab</span>}</span>
                    <button onClick={() => { if (window.confirm(`Remove "${per.name}" from People? Their tab stays.`)) setPersons(persons.filter((x) => x.id !== per.id)); }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
                  </div>
                );
              })}
              <datalist id="people-names">{allSeenNames.map((n) => <option key={n} value={n} />)}</datalist>
            </Card>

            {mcClubs.length > 0 && (
              <Card title="My Clubs — who you settle with" right={strayClubTabs.length > 0 ? <Btn tone="gold" small onClick={moveStrayClubTabs}>Move {strayClubTabs.length} old club tabs onto owners</Btn> : null}>
                {mcClubs.map((c) => (
                  <div key={c.id} style={{ display: "flex", gap: 10, alignItems: "center", padding: "5px 0", borderTop: `1px solid ${C.line}`, fontSize: 13 }}>
                    <b style={{ minWidth: 160 }}>{c.name}</b>
                    <span style={{ color: C.mute, fontSize: 12 }}>settles with</span>
                    <input list="people-names" defaultValue={c.owner || ""} placeholder="the club itself" onBlur={(e) => e.target.value.trim() !== (c.owner || "") && setClubOwner(c.id, e.target.value.trim())} onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()} style={{ ...inputS, width: 200, fontSize: 12 }} />
                    {c.owner && <span style={{ color: C.mute, fontSize: 12 }}>→ lands on {c.owner}'s tab</span>}
                  </div>
                ))}
              </Card>
            )}
            <Notes>
              <div>Each person has one tab. Usernames listed under a person (from any club) all settle onto that one tab. A username can be limited to one site if the same name means different people on different sites.</div>
              <div><b>Possible duplicates</b> are tabs with near-identical names (e.g. Rlawns / Rlawnsgud). Merging moves all entries and staking deals onto one tab and keeps the old name as a username.</div>
              <div><b>Settles with</b>: when you settle a club with its owner directly, name them here — that club's weekly amount goes on the owner's tab. Several clubs can share one owner.</div>
            </Notes>
          </div>
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
