export type Role = "teacher" | "student";

export type View =
  | "home" | "showcase" | "classBoard" | "allWorks" | "winners"
  | "adminLogin" | "admin" | "teacherLogin" | "teacher"
  | "studentLogin" | "student";

export type User = {
  id: string; fullName: string; role: Role; grade: string; section: string;
  teacherId?: string; username: string; password: string; createdAt: string; active: boolean;
};

export type Comp = {
  id: string; nameAr: string; nameEn: string; logoUrl: string; grades: string[];
  rulesAr: string; rulesEn: string; closingDate: string; certDate: string;
  started: boolean; archived: boolean; criteria: string[]; evalCode: string;
};

export type Work = {
  id: string; competitionId: string; studentId: string; title: string;
  desc: string; videoUrl: string; thumb: string; at: string;
};

export type Ev = { id: string; workId: string; scores: number[]; at: string; approved?: boolean };

export type Like = { workId: string; type: "friend" | "parent" | "teacher"; name: string; at: string };

export type Cert = {
  id: string; workId: string; studentId: string; compId: string; compName: string;
  score: number; rank: number; grade: string; section: string; issuedAt: string;
};

export type ClassInfo = {
  grade: string; section: string; teacherId: string; teacherName: string; count: number;
};

export type ArchiveEntry = {
  id: string; compId: string; compName: string; compNameAr: string; archivedAt: string;
  winners: Array<{ rank: number; studentName: string; grade: string; section: string; score: number }>;
};

export type CompPreset = { nameEn: string; nameAr: string; criteria: string[] };
