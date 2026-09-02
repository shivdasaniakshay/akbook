# AK's Book — setup

Four accounts, about 20 minutes: **GitHub** (code), **Supabase** (shared database + realtime),
**Vercel** (hosting), and an **Anthropic API key** (so Claude can build changes you request
from inside the site). Everything below is on free tiers except the API usage.

## 1. Put the code on GitHub

1. Create an empty repo (e.g. `akbook`, private is fine).
2. In this folder:
   ```
   git init && git add -A && git commit -m "AK's Book"
   git branch -M main
   git remote add origin https://github.com/YOUR-USER/akbook.git
   git push -u origin main
   ```

## 2. Supabase (database, realtime, file storage)

1. https://supabase.com → New project. Pick a name and a database password (you won't need it in the app).
2. **SQL Editor → New query**: paste the contents of `supabase/schema.sql` and Run.
3. **Authentication → Users → Add user → Create new user**:
   - Email: `owners@akbook.app` (the app's shared login — any email works if you also set `VITE_LOGIN_EMAIL`)
   - Password: **this is the site passphrase.** Choose something long; it also encrypts the data.
   - Tick *Auto Confirm User*.
4. **Authentication → Sign In / Providers → Email**: turn **off** "Allow new users to sign up" (only you create logins).
5. **Project Settings → API**: copy the *Project URL* and the *anon public* key.

> Changing the passphrase later means re-encrypting the data — export data first (header → Export data),
> set the new password on the user in Supabase, sign in with it, then Import data.

## 3. GitHub token + Anthropic key (for "Ask Claude")

1. GitHub → Settings → Developer settings → **Fine-grained personal access tokens** → Generate new token.
   Repository access: just this repo. Permissions: **Issues: Read and write**, **Pull requests: Read and write**,
   **Contents: Read and write**. Copy it (`github_pat_…`).
2. https://console.anthropic.com → API keys → create one.
3. In the GitHub repo → Settings → Secrets and variables → Actions → **New repository secret**:
   `ANTHROPIC_API_KEY` = your key. (Or use `CLAUDE_CODE_OAUTH_TOKEN` from `claude setup-token`
   to bill a Claude Pro/Max subscription instead — and swap the commented line in `.github/workflows/claude.yml`.)
4. Repo → Settings → Actions → General → Workflow permissions: **Read and write**, and tick
   **Allow GitHub Actions to create and approve pull requests**.
5. Install the Claude GitHub app on the repo: https://github.com/apps/claude (Install → pick the repo).

## 4. Vercel (hosting)

1. https://vercel.com → Add New → Project → import the GitHub repo. Framework: Vite (auto-detected).
2. Environment variables (all environments):
   | Name | Value |
   |---|---|
   | `VITE_SUPABASE_URL` | Project URL from step 2 |
   | `VITE_SUPABASE_ANON_KEY` | anon key from step 2 |
   | `GITHUB_REPO` | `YOUR-USER/akbook` |
   | `GITHUB_TOKEN` | the fine-grained token from step 3 |
3. Deploy. Your site is `https://akbook-….vercel.app` (add a custom domain under Settings → Domains if you like).
4. Vercel → Project → Settings → Git: make sure **Preview deployments** are on so every Claude PR gets a preview URL.

## 5. First run

Open the site, enter the passphrase, then header → **Import data** and load a `fishtank-seed-*.json`
exported from the old app (or start fresh). From now on every edit is saved to Supabase and appears
on every open browser within a second. Send the URL + passphrase to your co-owners.

## Asking Claude for changes

Bottom-right → **Ask Claude** → describe the change → *Send to Claude*. Within a few minutes the
request shows **Ready to review** with an *Open preview* link (a full copy of the site running the
new code) and a **Deploy** button; Deploy merges it and the live site updates for everyone.
To ask for tweaks before deploying, comment on the pull request with `@claude …`.

## Local development

```
cp .env.example .env.local   # fill in the Supabase values
npm install
npm run dev
```
(The Ask Claude panel needs the API functions, which only run on Vercel — use `npx vercel dev` for that.)

## How it's built

- `src/App.jsx` — the app, unchanged apart from persistence.
- `src/sync.js` — passphrase → PBKDF2 → AES-GCM key; every value is encrypted before it reaches
  Supabase, so the database and its backups are unreadable without the passphrase. Supabase Realtime
  pushes other people's edits into the local cache and the app re-renders.
- Concurrency is last-writer-wins per saved key (e.g. the whole Tabs ledger is one key), so two people
  editing the same section in the same second can overwrite each other — fine for a couple of owners.
