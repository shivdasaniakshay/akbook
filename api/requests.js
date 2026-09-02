// GET  → recent change requests (GitHub issues labelled claude-request) + their PRs
// POST → { text } opens a new request; the Claude Code action picks it up
import { gh, REPO, handler } from "./_lib.js";

export const LABEL = "claude-request";

async function list() {
  const [issues, prs] = await Promise.all([
    gh(`/repos/${REPO}/issues?labels=${LABEL}&state=all&per_page=25&sort=created&direction=desc`),
    gh(`/repos/${REPO}/pulls?state=all&per_page=50&sort=created&direction=desc`),
  ]);
  const out = [];
  for (const is of issues.filter((i) => !i.pull_request)) {
    // Claude's branches are named claude/issue-<n>-…; match the PR by head branch or a "#<n>" in its body.
    const pr = prs.find((p) => new RegExp(`issue-${is.number}(-|$)`).test(p.head.ref) || new RegExp(`#${is.number}\\b`).test(p.body || ""));
    let prInfo = null;
    if (pr) {
      let preview_url = null;
      try {
        const comments = await gh(`/repos/${REPO}/issues/${pr.number}/comments?per_page=30`);
        for (const c of comments) { const m = (c.body || "").match(/https:\/\/[\w.-]+\.vercel\.app[^\s)\]]*/); if (m) preview_url = m[0]; }
      } catch (e) {}
      prInfo = { number: pr.number, html_url: pr.html_url, state: pr.state, merged: !!pr.merged_at, preview_url };
    }
    out.push({ number: is.number, title: is.title, state: is.state, created_at: is.created_at, html_url: is.html_url, pr: prInfo });
  }
  return out;
}

async function create(text) {
  const firstLine = text.trim().split("\n")[0].trim();
  const title = firstLine.length > 80 ? firstLine.slice(0, 77) + "…" : firstLine;
  const body = `@claude ${text.trim()}\n\n---\n_Requested from the AK's Book site. Implement it, run \`npm run build\`, and open a pull request._`;
  return gh(`/repos/${REPO}/issues`, { method: "POST", body: JSON.stringify({ title, body, labels: [LABEL] }) });
}

export default handler(async (req) => {
  if (req.method === "POST") {
    const { text } = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
    if (!text || !text.trim()) throw Object.assign(new Error("Say what you'd like changed."), { status: 400 });
    const is = await create(text);
    return { number: is.number, html_url: is.html_url };
  }
  return list();
});
