-- Phase 2A: Authentication Foundation Schema
-- DRAFT ONLY — DO NOT EXECUTE YET
-- Purpose: secure profiles foundation for Supabase Auth migration

CREATE TABLE profiles (
  id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  role TEXT NOT NULL CHECK (role IN ('student', 'teacher', 'admin', 'coordinator')),
  full_name TEXT NOT NULL,
  grade TEXT,
  section TEXT,
  teacher_id UUID REFERENCES profiles(id),
  active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;

-- =====================================================
-- SECURITY DEFINER HELPERS (avoid RLS recursion safely)
-- =====================================================
-- These helpers run with elevated privileges, so we lock them down by:
-- 1) setting a safe search_path
-- 2) returning only booleans/UUIDs that are strictly required
-- 3) not exposing role-changing or arbitrary row access

CREATE OR REPLACE FUNCTION public.get_my_role()
RETURNS TEXT
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = ''
AS $$
  SELECT p.role
  FROM public.profiles AS p
  WHERE p.id = auth.uid()
  LIMIT 1;
$$;

CREATE OR REPLACE FUNCTION public.get_my_teacher_id()
RETURNS UUID
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = ''
AS $$
  SELECT p.teacher_id
  FROM public.profiles AS p
  WHERE p.id = auth.uid()
  LIMIT 1;
$$;

CREATE OR REPLACE FUNCTION public.is_my_student(target_profile_id UUID)
RETURNS BOOLEAN
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = ''
AS $$
  SELECT EXISTS (
    SELECT 1
    FROM public.profiles AS student_profile
    WHERE student_profile.id = target_profile_id
      AND student_profile.teacher_id = auth.uid()
  );
$$;

CREATE OR REPLACE FUNCTION public.is_admin_or_coordinator()
RETURNS BOOLEAN
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = ''
AS $$
  SELECT EXISTS (
    SELECT 1
    FROM public.profiles AS p
    WHERE p.id = auth.uid()
      AND p.role IN ('admin', 'coordinator')
  );
$$;

-- ==============================
-- SELECT policies
-- ==============================

-- 1) User can read their own profile only
CREATE POLICY "profiles_select_own"
ON public.profiles
FOR SELECT
TO authenticated
USING (auth.uid() = id);

-- 2) Student can read assigned teacher profile only
CREATE POLICY "profiles_select_assigned_teacher"
ON public.profiles
FOR SELECT
TO authenticated
USING (id = public.get_my_teacher_id());

-- 3) Teacher can read profiles of students assigned to them
CREATE POLICY "profiles_select_teacher_students"
ON public.profiles
FOR SELECT
TO authenticated
USING (
  public.get_my_role() = 'teacher'
  AND public.is_my_student(id)
);

-- 4) Admin / Coordinator can read all profiles
CREATE POLICY "profiles_select_admin_all"
ON public.profiles
FOR SELECT
TO authenticated
USING (public.is_admin_or_coordinator());

-- ==============================
-- NO browser-side INSERT / UPDATE / DELETE
-- ==============================
-- We intentionally do NOT create INSERT / UPDATE / DELETE policies for browser clients.
-- Profile creation and modification must happen only server-side via Edge Functions
-- using the service_role key in Supabase server secrets.
-- Anonymous users have NO access to profiles.
-- Public showcase must NOT query profiles directly.
-- Users cannot modify role from the browser because no client-side write policies exist.

