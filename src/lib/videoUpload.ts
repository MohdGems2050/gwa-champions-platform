/**
 * Uploads a submission video to Supabase Storage (the shared cloud bucket),
 * so the file is playable from any device.
 *
 * If Supabase Storage is unavailable, it falls back to the legacy Netlify Blob
 * endpoint so nothing breaks for existing deployments.
 */
import { supabase, isSupabaseConfigured } from "./supabase";

const BUCKET = "videos";
const CHUNK_SIZE = 3 * 1024 * 1024;

function safeName(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9._-]+/g, "-").slice(-80);
}

async function uploadToSupabase(file: File, onProgress?: (pct: number) => void): Promise<string> {
  const path = `submissions/${Date.now()}-${crypto.randomUUID()}-${safeName(file.name || "video.mp4")}`;
  onProgress?.(15);

  const { error } = await supabase.storage.from(BUCKET).upload(path, file, {
    cacheControl: "3600",
    upsert: false,
    contentType: file.type || "video/mp4",
  });
  if (error) throw new Error(error.message);

  onProgress?.(85);
  const { data } = supabase.storage.from(BUCKET).getPublicUrl(path);
  if (!data?.publicUrl) throw new Error("Could not create a public video link");
  onProgress?.(100);
  return data.publicUrl;
}

async function uploadToLegacyBlobs(file: File, onProgress?: (pct: number) => void): Promise<string> {
  const id = crypto.randomUUID();
  const chunks = Math.ceil(file.size / CHUNK_SIZE);
  if (chunks === 0) throw new Error("The selected video is empty");

  for (let i = 0; i < chunks; i++) {
    const part = file.slice(i * CHUNK_SIZE, Math.min(file.size, (i + 1) * CHUNK_SIZE));
    let res: Response | null = null;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        res = await fetch(`/api/videos/${id}/chunk?index=${i}`, {
          method: "PUT",
          headers: { "Content-Type": "application/octet-stream" },
          body: part,
        });
        if (res.ok) break;
      } catch {
        res = null;
      }
    }
    if (!res?.ok) throw new Error(`Failed to upload part ${i + 1} of ${chunks}`);
    onProgress?.(Math.round(((i + 1) / chunks) * 95));
  }

  const res = await fetch(`/api/videos/${id}/complete`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chunks,
      chunkSize: CHUNK_SIZE,
      size: file.size,
      type: file.type || "video/mp4",
      name: file.name,
    }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.url) throw new Error(data.error || "Failed to finalize upload");
  onProgress?.(100);
  return data.url as string;
}

export async function uploadVideo(file: File, onProgress?: (pct: number) => void): Promise<string> {
  if (!file || file.size === 0) throw new Error("The selected video is empty");

  let cloudError: unknown = null;
  if (isSupabaseConfigured) {
    try {
      return await uploadToSupabase(file, onProgress);
    } catch (err) {
      cloudError = err;
      console.warn("Supabase Storage upload failed, trying legacy endpoint:", err);
    }
  }

  try {
    return await uploadToLegacyBlobs(file, onProgress);
  } catch (legacyError) {
    if (cloudError) {
      const message = cloudError instanceof Error ? cloudError.message : String(cloudError);
      throw new Error(
        `تعذر رفع الفيديو إلى التخزين السحابي (${message}). تأكد من إنشاء Bucket باسم "videos" في Supabase Storage وجعله عاماً.`
      );
    }
    throw legacyError;
  }
}

export async function deleteVideo(url: string): Promise<void> {
  if (!url) return;

  const marker = `/storage/v1/object/public/${BUCKET}/`;
  if (url.includes(marker)) {
    const path = decodeURIComponent(url.split(marker)[1] || "");
    if (path) await supabase.storage.from(BUCKET).remove([path]);
    return;
  }

  if (!url.startsWith("/api/videos/")) return;
  const res = await fetch(url, { method: "DELETE" });
  if (!res.ok) throw new Error("Failed to delete video");
}
