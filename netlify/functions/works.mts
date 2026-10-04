import type { Config, Context } from "@netlify/functions";
import { desc, eq } from "drizzle-orm";
import { db } from "../../db/index.js";
import { works } from "../../db/schema.js";

export default async (req: Request, context: Context) => {
  if (req.method === "GET") {
    const rows = await db.select().from(works).orderBy(desc(works.at));
    return Response.json(rows);
  }

  if (req.method === "POST") {
    const w = await req.json();
    if (!w?.id || !w.competitionId || !w.studentId || !w.title || !w.videoUrl) {
      return Response.json({ error: "Missing required fields" }, { status: 400 });
    }
    const [row] = await db
      .insert(works)
      .values({
        id: String(w.id),
        competitionId: String(w.competitionId),
        studentId: String(w.studentId),
        title: String(w.title),
        desc: String(w.desc ?? ""),
        videoUrl: String(w.videoUrl),
        thumb: String(w.thumb ?? ""),
        at: w.at ? String(w.at) : new Date().toISOString(),
      })
      .returning();
    return Response.json(row, { status: 201 });
  }

  if (req.method === "DELETE" && context.params.id) {
    await db.delete(works).where(eq(works.id, context.params.id));
    return Response.json({ ok: true });
  }

  return new Response("Method not allowed", { status: 405 });
};

export const config: Config = {
  path: ["/api/works", "/api/works/:id"],
};
