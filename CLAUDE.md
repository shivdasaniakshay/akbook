# AK's Book — notes for Claude

This is a weekly poker-club accounting app (React + Vite) with an encrypted realtime
backend (Supabase) and a Vercel deployment. Requests arrive as GitHub issues opened
from the "Ask Claude" panel inside the site.

## How to work here
- The whole app UI lives in `src/App.jsx` (large, single file — keep it that way).
  `src/sync.js` is the persistence layer (Supabase kv table + storage, AES-GCM
  encryption, realtime). `src/main.jsx` has the passphrase gate and the floating dock.
  `src/ClaudePanel.jsx` is the request panel. `api/` are Vercel serverless functions.
- Persist state through `store.get/set/del` (never touch `localStorage` directly for
  app data) so edits sync to everyone. Big binary files go through `blobPut/blobGet`.
- Keep the code style: compact, few defensive guards, short one-sentence comments.
- Run `npm install` and `npm run build` before finishing; the build must pass.
- Don't change `KDF_SALT`, `LOGIN_EMAIL`, the kv schema, or anything that would make
  existing encrypted data unreadable.
- Leave the workflow, `api/`, and `SETUP.md` alone unless the request is about them.
