import type { Config, Context } from "@netlify/functions";
import { getStore } from "@netlify/blobs";

// Videos are uploaded in chunks (each under the 6 MB request limit) and stored
// as separate blobs, plus a manifest describing the whole file.
type Manifest = { chunks: number; chunkSize: number; size: number; type: string; name: string };

const MAX_CHUNK = 5 * 1024 * 1024;
const ID_RE = /^[a-z0-9-]{8,64}$/;

const store = () => getStore({ name: "videos", consistency: "strong" });
const chunkKey = (id: string, i: number) => `${id}/chunk-${String(i).padStart(5, "0")}`;
const manifestKey = (id: string) => `${id}/manifest`;

export default async (req: Request, context: Context) => {
  const { id, action } = context.params;
  if (!id || !ID_RE.test(id)) return Response.json({ error: "Invalid video id" }, { status: 400 });
  const videos = store();

  if (req.method === "PUT" && action === "chunk") {
    const index = Number(new URL(req.url).searchParams.get("index"));
    if (!Number.isInteger(index) || index < 0) return Response.json({ error: "Invalid chunk index" }, { status: 400 });
    const data = await req.arrayBuffer();
    if (data.byteLength === 0 || data.byteLength > MAX_CHUNK) {
      return Response.json({ error: "Invalid chunk size" }, { status: 400 });
    }
    await videos.set(chunkKey(id, index), data);
    return Response.json({ ok: true });
  }

  if (req.method === "POST" && action === "complete") {
    const m = (await req.json()) as Manifest;
    if (!Number.isInteger(m.chunks) || m.chunks < 1 || !Number.isInteger(m.size) || m.size < 1) {
      return Response.json({ error: "Invalid manifest" }, { status: 400 });
    }
    for (let i = 0; i < m.chunks; i++) {
      if (!(await videos.getMetadata(chunkKey(id, i)))) {
        return Response.json({ error: `Missing chunk ${i}` }, { status: 400 });
      }
    }
    await videos.setJSON(manifestKey(id), {
      chunks: m.chunks,
      chunkSize: m.chunkSize,
      size: m.size,
      type: String(m.type || "video/mp4"),
      name: String(m.name || "video"),
    });
    return Response.json({ url: `/api/videos/${id}` });
  }

  if (req.method === "DELETE" && !action) {
    const { blobs } = await videos.list({ prefix: `${id}/` });
    await Promise.all(blobs.map((b) => videos.delete(b.key)));
    return Response.json({ ok: true });
  }

  if ((req.method === "GET" || req.method === "HEAD") && !action) {
    const m = (await videos.get(manifestKey(id), { type: "json" })) as Manifest | null;
    if (!m) return new Response("Not found", { status: 404 });

    // Serve at most one stored chunk per response; video players request the rest via Range.
    let start = 0;
    const range = req.headers.get("range");
    const match = range?.match(/bytes=(\d*)-(\d*)/);
    if (match) {
      if (match[1]) start = Number(match[1]);
      else if (match[2]) start = Math.max(0, m.size - Number(match[2]));
    }
    if (start >= m.size) {
      return new Response(null, { status: 416, headers: { "Content-Range": `bytes */${m.size}` } });
    }
    const index = Math.floor(start / m.chunkSize);
    const chunk = (await videos.get(chunkKey(id, index), { type: "arrayBuffer" })) as ArrayBuffer | null;
    if (!chunk) return new Response("Not found", { status: 404 });

    const chunkStart = index * m.chunkSize;
    let end = chunkStart + chunk.byteLength - 1;
    if (match?.[1] && match[2]) end = Math.min(end, Number(match[2]));
    const body = chunk.slice(start - chunkStart, end - chunkStart + 1);

    const headers = {
      "Content-Type": m.type,
      "Accept-Ranges": "bytes",
      "Content-Length": String(body.byteLength),
      "Content-Range": `bytes ${start}-${end}/${m.size}`,
      "Cache-Control": "public, max-age=31536000, immutable",
    };
    return new Response(req.method === "HEAD" ? null : body, { status: 206, headers });
  }

  return new Response("Method not allowed", { status: 405 });
};

export const config: Config = {
  path: ["/api/videos/:id", "/api/videos/:id/:action"],
};
