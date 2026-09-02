// Cloud sync layer: every localStorage key the app uses is mirrored (encrypted)
// into a Supabase `kv` table, and archive workbooks go (encrypted) to Supabase
// Storage. localStorage stays the fast local cache; Supabase Realtime pushes
// other people's edits into it and tells the UI to re-read.
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
export const configured = !!(SUPABASE_URL && SUPABASE_ANON_KEY);
export const supabase = configured ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

// One shared login for the co-owners; the passphrase doubles as its password
// and (through PBKDF2) as the AES key that encrypts every stored value.
const LOGIN_EMAIL = import.meta.env.VITE_LOGIN_EMAIL || "owners@akbook.app";
const KDF_SALT = "akbook-kv-v1";
const BUCKET = "archive";
const PASS_KEY = "akbook-pass"; // sessionStorage: survives refresh, gone when the tab closes

let aesKey = null;
const clientId = crypto.randomUUID(); // lets us ignore the echo of our own writes
const listeners = new Set();  // (changedKeys[]) => void
const statusListeners = new Set(); // ("synced" | "saving" | "offline" | "error", message) => void
let status = "synced";

// ———————————————— crypto helpers ————————————————
const enc = new TextEncoder(), dec = new TextDecoder();
function bytesToB64(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
function b64ToBytes(b64) {
  const s = atob(b64); const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
async function deriveKey(passphrase) {
  const base = await crypto.subtle.importKey("raw", enc.encode(passphrase), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey({ name: "PBKDF2", salt: enc.encode(KDF_SALT), iterations: 250000, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
}
async function encryptBytes(bytes) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, aesKey, bytes));
  const out = new Uint8Array(12 + ct.length); out.set(iv); out.set(ct, 12);
  return out;
}
async function decryptBytes(bytes) {
  return new Uint8Array(await crypto.subtle.decrypt({ name: "AES-GCM", iv: bytes.subarray(0, 12) }, aesKey, bytes.subarray(12)));
}
const encryptText = async (s) => bytesToB64(await encryptBytes(enc.encode(s)));
const decryptText = async (b64) => dec.decode(await decryptBytes(b64ToBytes(b64)));

// ———————————————— status + change notifications ————————————————
function setStatus(next, msg = "") { status = next; statusListeners.forEach((fn) => fn(next, msg)); }
export const getStatus = () => status;
export function onStatus(fn) { statusListeners.add(fn); return () => statusListeners.delete(fn); }
export function onRemoteChange(fn) { listeners.add(fn); return () => listeners.delete(fn); }

let pendingKeys = new Set(), notifyTimer = null;
function notify(key) {
  pendingKeys.add(key);
  clearTimeout(notifyTimer);
  notifyTimer = setTimeout(() => { const keys = [...pendingKeys]; pendingKeys = new Set(); listeners.forEach((fn) => fn(keys)); }, 250);
}

// ———————————————— unlock / lock ————————————————
export const savedPassphrase = () => { try { return sessionStorage.getItem(PASS_KEY) || ""; } catch (e) { return ""; } };

export async function unlock(passphrase) {
  if (!configured) throw new Error("Supabase isn't configured — set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY.");
  const { error } = await supabase.auth.signInWithPassword({ email: LOGIN_EMAIL, password: passphrase });
  if (error) throw new Error("That passphrase didn't work.");
  aesKey = await deriveKey(passphrase);
  try { sessionStorage.setItem(PASS_KEY, passphrase); } catch (e) {}
  await hydrate();
  subscribe();
}

export async function lock() {
  try { sessionStorage.removeItem(PASS_KEY); } catch (e) {}
  aesKey = null;
  if (channel) { await supabase.removeChannel(channel); channel = null; }
  await supabase.auth.signOut();
  // Wipe the local cache so the next person at this browser sees nothing without the passphrase.
  Object.keys(localStorage).filter((k) => !k.startsWith("sb-")).forEach((k) => localStorage.removeItem(k));
}

// Pull every row into localStorage (cloud wins over whatever this browser had).
async function hydrate() {
  const { data, error } = await supabase.from("kv").select("key,value");
  if (error) throw new Error("Couldn't load data: " + error.message);
  const cloud = {};
  for (const row of data) {
    try { cloud[row.key] = await decryptText(row.value); }
    catch (e) { throw new Error("Data couldn't be decrypted with this passphrase."); }
  }
  Object.keys(localStorage).filter((k) => !k.startsWith("sb-")).forEach((k) => localStorage.removeItem(k));
  Object.entries(cloud).forEach(([k, v]) => localStorage.setItem(k, v));
}

// ———————————————— realtime ————————————————
let channel = null;
function subscribe() {
  if (channel) return;
  channel = supabase.channel("kv-live")
    .on("postgres_changes", { event: "*", schema: "public", table: "kv" }, async (p) => {
      if (p.eventType === "DELETE") {
        const k = p.old?.key;
        if (k && localStorage.getItem(k) != null) { localStorage.removeItem(k); notify(k); }
        return;
      }
      if (p.new.by === clientId) return; // our own write coming back
      try {
        const v = await decryptText(p.new.value);
        if (localStorage.getItem(p.new.key) === v) return;
        localStorage.setItem(p.new.key, v);
        notify(p.new.key);
      } catch (e) {}
    })
    .subscribe((s) => { if (s === "SUBSCRIBED") setStatus("synced"); else if (s === "CHANNEL_ERROR" || s === "TIMED_OUT") setStatus("offline", "Live updates paused — reconnecting"); });
}

// ———————————————— key/value store (drop-in for the old localStorage `store`) ————————————————
const pushTimers = new Map();
function push(k, v) {
  // Debounce rapid edits to the same key into one write.
  clearTimeout(pushTimers.get(k));
  setStatus("saving");
  pushTimers.set(k, setTimeout(async () => {
    pushTimers.delete(k);
    try {
      const { error } = await supabase.from("kv").upsert({ key: k, value: await encryptText(v), by: clientId, updated_at: new Date().toISOString() });
      if (error) throw error;
      if (pushTimers.size === 0) setStatus("synced");
    } catch (e) { setStatus("error", "Couldn't save to the cloud: " + (e.message || e)); }
  }, 400));
}

export const store = {
  get: async (k) => { const v = localStorage.getItem(k); return v ? { value: v } : null; },
  set: async (k, v) => { localStorage.setItem(k, v); push(k, v); return { value: v }; },
  del: async (k) => { localStorage.removeItem(k); await supabase.from("kv").delete().eq("key", k); },
};

// ———————————————— archive files (Supabase Storage, encrypted) ————————————————
// Values look like { ..., buf: ArrayBuffer } or { files: [{ name, buf }] }; buffers travel as base64.
const isBytes = (v) => v instanceof ArrayBuffer || ArrayBuffer.isView(v);
const toBytes = (v) => (v instanceof ArrayBuffer ? new Uint8Array(v) : new Uint8Array(v.buffer, v.byteOffset, v.byteLength));
const replacer = (_, v) => (isBytes(v) ? { __bytes: bytesToB64(toBytes(v)) } : v);
const reviver = (_, v) => (v && typeof v === "object" && typeof v.__bytes === "string" ? b64ToBytes(v.__bytes) : v);
// Storage object names are restrictive, so file keys travel as base64url.
const objectPath = (key) => btoa(unescape(encodeURIComponent(key))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "") + ".bin";

export async function blobPut(key, val) {
  const bytes = await encryptBytes(enc.encode(JSON.stringify(val, replacer)));
  const { error } = await supabase.storage.from(BUCKET).upload(objectPath(key), new Blob([bytes]), { upsert: true, contentType: "application/octet-stream" });
  if (error) throw error;
}
export async function blobGet(key) {
  const { data, error } = await supabase.storage.from(BUCKET).download(objectPath(key));
  if (error || !data) return null;
  return JSON.parse(dec.decode(await decryptBytes(new Uint8Array(await data.arrayBuffer()))), reviver);
}
export async function blobDel(key) {
  await supabase.storage.from(BUCKET).remove([objectPath(key)]);
}
