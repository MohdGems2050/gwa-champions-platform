import type { Role, User } from "../types";

export function cn(...classes: Array<string | false | null | undefined>): string {
  return classes.filter(Boolean).join(" ");
}

export function makeId(prefix: string): string {
  return `${prefix}-${Math.random().toString(36).slice(2, 8)}`;
}

export function totalScore(scores: number[]): number {
  return scores.reduce((sum, score) => sum + score, 0);
}

export function slugify(name: string): string {
  return name
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, ".")
    .replace(/^\.+|\.+$/g, "");
}

export function createUser(
  name: string,
  role: Role,
  grade: string,
  section: string,
  teacherId: string | undefined,
  existingUsers: User[]
): User {
  let username = slugify(name) || "gwa.user";
  let counter = 1;
  while (existingUsers.some((u) => u.username === username)) {
    username = slugify(name) + counter;
    counter++;
  }
  const num = String(existingUsers.length + 2050).slice(-4);
  return {
    id: makeId(role.charAt(0)),
    fullName: name.trim(),
    role,
    grade,
    section: role === "student" ? section : "",
    teacherId,
    username,
    password: (role === "teacher" ? "Teach" : "GWA") + num,
    createdAt: new Date().toISOString(),
    active: true,
  };
}

export function getClassColor(grade: string) {
  const colors = [
    { bg: "rgba(22,151,232,0.35)", border: "#1697E8", glow: "rgba(22,151,232,0.15)" },
    { bg: "rgba(253,181,31,0.35)", border: "#FDB51F", glow: "rgba(253,181,31,0.15)" },
    { bg: "rgba(16,176,155,0.35)", border: "#10B09B", glow: "rgba(16,176,155,0.15)" },
    { bg: "rgba(226,20,85,0.35)", border: "#E21455", glow: "rgba(226,20,85,0.15)" },
    { bg: "rgba(139,92,246,0.35)", border: "#8B5CF6", glow: "rgba(139,92,246,0.15)" },
    { bg: "rgba(236,72,153,0.35)", border: "#EC4899", glow: "rgba(236,72,153,0.15)" },
  ];
  let hash = 0;
  for (let i = 0; i < grade.length; i++) {
    hash = grade.charCodeAt(i) + ((hash << 5) - hash);
  }
  return colors[Math.abs(hash) % colors.length];
}
