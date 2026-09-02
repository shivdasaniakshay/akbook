// POST { pr } → squash-merge Claude's pull request; Vercel deploys main automatically.
import { gh, REPO, handler } from "./_lib.js";

export default handler(async (req) => {
  const { pr } = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
  if (!pr) throw Object.assign(new Error("Missing pr"), { status: 400 });
  const r = await gh(`/repos/${REPO}/pulls/${pr}/merge`, { method: "PUT", body: JSON.stringify({ merge_method: "squash" }) });
  return { merged: !!r.merged, message: r.message };
});
