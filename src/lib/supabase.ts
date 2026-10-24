import { createClient } from "@supabase/supabase-js";

// Uses the browser-safe publishable anon key
const supabaseUrl = (import.meta as any).env.VITE_SUPABASE_URL || "https://cpinsianfugimhbgzkkx.supabase.co";
const supabaseAnonKey = (import.meta as any).env.VITE_SUPABASE_ANON_KEY || "sb_publishable_ZvcDP1VymahSV-qIjWBotQ_Qdpmq8fn";

export const isSupabaseConfigured = Boolean(supabaseUrl && supabaseAnonKey);

if (!isSupabaseConfigured) {
  console.warn(
    "Supabase is not configured: set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY. Supabase features are disabled."
  );
}

// createClient throws when the URL is empty, which would crash the whole app (white page).
// We use a placeholder URL so the app can still mount and run localStorage features.
export const supabase = createClient(
  supabaseUrl || "https://placeholder-project.supabase.co",
  supabaseAnonKey || "placeholder-anon-key"
);
