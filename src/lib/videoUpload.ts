/**
 * Uploads a video to Netlify Blobs through the /api/videos function.
 * The file is sent in chunks so each request stays under the function payload limit.
 */
const CHUNK_SIZE = 3 * 1024 * 1024;

export async function uploadVideo(file: File, onProgress?: (pct: number) => void): Promise<string> {
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
    body: JSON.stringify({ chunks, chunkSize: CHUNK_SIZE, size: file.size, type: file.type || "video/mp4", name: file.name }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.url) throw new Error(data.error || "Failed to finalize upload");
  onProgress?.(100);
  return data.url as string;
}

export async function deleteVideo(url: string): Promise<void> {
  if (!url.startsWith("/api/videos/")) return;
  const res = await fetch(url, { method: "DELETE" });
  if (!res.ok) throw new Error("Failed to delete video");
}
