# GWA Champions Platform — Agent Working Rules

> **Single-owner policy / سياسة وكيل واحد**
> This repository is maintained by **one** agent only ("Agent A") together with the owner
> **MohdGems2050**. Any other agent, bot or automation must **not** push to `main`.
> يجب ألا يرفع أي وكيل أو بوت آخر إلى فرع `main`. الوكيل المسؤول الوحيد هو "Agent A".

---

## 1. Repository facts

| Item | Value |
|------|-------|
| App type | Vite + React + TypeScript single-page app (`src/App.tsx`) |
| Hosting | **Vercel** (production alias: `gwa-champions-platform.vercel.app`) |
| Backend | **Supabase** project `cpinsianfugimhbgzkkx` |
| Build | `npm run build` → output `dist` |
| Legacy hosting | Netlify project `stirring-unicorn-f3175f` (build credits exhausted, do not rely on it) |

## 2. Required environment variables (Vercel)

```
VITE_SUPABASE_URL=https://cpinsianfugimhbgzkkx.supabase.co
VITE_SUPABASE_ANON_KEY=<publishable key from Supabase dashboard>
```

Never commit real keys. The anon/publishable key is browser-safe only because every
table has RLS enabled with explicit policies.

## 3. Cloud data model (Supabase)

| Table | Purpose | Notes |
|-------|---------|-------|
| `accounts` | **Source of truth for logins** | `id, full_name, role, grade, section, teacher_id, username (unique), password, active` |
| `works_cloud` | Student/teacher submissions | `id, competition_id, student_id, title, desc, video_url, thumb, at` |
| `evaluations` | Teacher scorecards | `scores` is `jsonb` (10 numbers) |
| `certificates` | Issued certificates | |
| `users_data` | **LEGACY** — auth-era table, no `username`/`password` columns | never use for login |
| `works` | **LEGACY** — has a `NOT NULL uuid owner_id` column | writes fail, never use |
| storage bucket `videos` | Public video files | max 50 MB on the free plan |

## 4. Known pitfalls (do not regress)

1. `desc` is a reserved SQL keyword → always quote it (`"desc"`) in raw SQL.
2. The legacy `public.works` table requires `owner_id uuid` → use `works_cloud`.
3. Login must read credentials from `accounts` (case-insensitive username); **never**
   gate login behind `supabase.auth` (no auth users exist).
4. Upserts must be deduplicated by `username` and chunked; one conflicting row must not
   abort the whole sync.
5. Videos must be uploaded to Supabase Storage bucket `videos` and the resulting public
   URL stored in `works_cloud.video_url`.

## 5. Change workflow

1. Work on a single branch: `agent-a/<short-topic>`.
2. Keep `main` deployable — Vercel deploys every push to `main` automatically.
3. Verify with `npx esbuild src/App.tsx --outfile=/tmp/check.js` before pushing.
4. Write clear commit messages describing the user-visible change.
5. Never force-push `main`, never rewrite history.
6. If a rollback is required, use `git revert` / restore a known commit and push a new
   commit — do not delete commits.

## 6. Current stable checkpoint

```
8aa12e4  revert: roll back the competition delete/edit + teacher-login feature set
```

Cloud login (`accounts`) and submissions (`works_cloud` + `videos` bucket) are verified
working against the live Supabase project.
