// Shared helpers for the Vercel functions: verify the caller's Supabase session,
// then talk to GitHub on their behalf with the repo token.
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY;
export const REPO = process.env.GITHUB_REPO; // "owner/name"
const TOKEN = process.env.GITHUB_TOKEN;

export async function requireUser(req) {
  const token = (req.headers.authorization || "").replace(/^Bearer /, "");
  if (!token) throw Object.assign(new Error("Not signed in"), { status: 401 });
  const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  const { data, error } = await sb.auth.getUser(token);
  if (error || !data?.user) throw Object.assign(new Error("Not signed in"), { status: 401 });
  return data.user;
}

export async function gh(path, opts = {}) {
  const r = await fetch("https://api.github.com" + path, {
    ...opts,
    headers: { Authorization: "Bearer " + TOKEN, Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28", "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  const body = await r.json().catch(() => ({}));
  if (!r.ok) throw Object.assign(new Error(body.message || "GitHub error " + r.status), { status: 502 });
  return body;
}

export function handler(fn) {
  return async (req, res) => {
    try {
      if (!REPO || !TOKEN) throw Object.assign(new Error("GITHUB_REPO / GITHUB_TOKEN aren't set on Vercel"), { status: 500 });
      await requireUser(req);
      res.status(200).json(await fn(req));
    } catch (e) {
      res.status(e.status || 500).json({ error: e.message || String(e) });
    }
  };
}
