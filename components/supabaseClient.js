import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // Ini cuma warning di console, bukan error yang menghentikan app,
  // supaya build tetap jalan walau env belum diisi (misal saat development awal).
  console.warn(
    "[supabaseClient] NEXT_PUBLIC_SUPABASE_URL atau NEXT_PUBLIC_SUPABASE_ANON_KEY belum diset di .env.local"
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);