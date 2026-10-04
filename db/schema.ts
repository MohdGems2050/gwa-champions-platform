import { pgTable, text, timestamp } from "drizzle-orm/pg-core";

export const works = pgTable("works", {
  id: text().primaryKey(),
  competitionId: text("competition_id").notNull(),
  studentId: text("student_id").notNull(),
  title: text().notNull(),
  desc: text().notNull().default(""),
  videoUrl: text("video_url").notNull(),
  thumb: text().notNull().default(""),
  at: timestamp({ mode: "string", withTimezone: true }).notNull().defaultNow(),
});
