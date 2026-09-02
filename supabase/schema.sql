-- Run this once in the Supabase SQL editor (Dashboard → SQL Editor → New query → Run).

-- One row per app key. `value` is AES-GCM ciphertext (base64) — Supabase never sees plaintext.
create table if not exists public.kv (
  key text primary key,
  value text not null,
  by text,                                  -- writer's session id (lets a browser ignore its own echo)
  updated_at timestamptz not null default now()
);

-- Only signed-in owners (the shared login) can read or write.
alter table public.kv enable row level security;
drop policy if exists "owners read" on public.kv;
drop policy if exists "owners write" on public.kv;
create policy "owners read"  on public.kv for select to authenticated using (true);
create policy "owners write" on public.kv for all    to authenticated using (true) with check (true);

-- Realtime: broadcast row changes to connected browsers.
alter publication supabase_realtime add table public.kv;
alter table public.kv replica identity full;

-- Archive workbooks (encrypted blobs) live in a private storage bucket.
insert into storage.buckets (id, name, public) values ('archive', 'archive', false)
  on conflict (id) do nothing;
drop policy if exists "owners archive read"   on storage.objects;
drop policy if exists "owners archive write"  on storage.objects;
drop policy if exists "owners archive update" on storage.objects;
drop policy if exists "owners archive delete" on storage.objects;
create policy "owners archive read"   on storage.objects for select to authenticated using (bucket_id = 'archive');
create policy "owners archive write"  on storage.objects for insert to authenticated with check (bucket_id = 'archive');
create policy "owners archive update" on storage.objects for update to authenticated using (bucket_id = 'archive');
create policy "owners archive delete" on storage.objects for delete to authenticated using (bucket_id = 'archive');
