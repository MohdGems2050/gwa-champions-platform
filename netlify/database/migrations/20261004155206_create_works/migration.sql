CREATE TABLE "works" (
	"id" text PRIMARY KEY,
	"competition_id" text NOT NULL,
	"student_id" text NOT NULL,
	"title" text NOT NULL,
	"desc" text DEFAULT '' NOT NULL,
	"video_url" text NOT NULL,
	"thumb" text DEFAULT '' NOT NULL,
	"at" timestamp with time zone DEFAULT now() NOT NULL
);
