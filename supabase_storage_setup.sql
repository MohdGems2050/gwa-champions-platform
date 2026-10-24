-- ============================================================
-- GWA Champions - Supabase Storage for student videos
-- Run once in Supabase Dashboard -> SQL Editor
-- ============================================================

-- 1) Public bucket that holds submission videos
insert into storage.buckets (id, name, public)
values ('videos', 'videos', true)
on conflict (id) do update set public = true;

-- 2) Anyone may read the videos (needed for playback on any device)
drop policy if exists "videos public read" on storage.objects;
create policy "videos public read"
on storage.objects for select
using (bucket_id = 'videos');

-- 3) Browser clients may upload videos into this bucket
drop policy if exists "videos anon upload" on storage.objects;
create policy "videos anon upload"
on storage.objects for insert
with check (bucket_id = 'videos');

-- 4) Any visitor may remove their own deleted submission file
drop policy if exists "videos anon delete" on storage.objects;
create policy "videos anon delete"
on storage.objects for delete
using (bucket_id = 'videos');
