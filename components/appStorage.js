import { supabase } from "./supabaseClient";

// =====================================================
// Modul ini menggantikan window.storage (yang cuma jalan
// di lingkungan Claude Artifacts) dengan penyimpanan
// permanen beneran lewat Supabase, tapi bentuk API-nya
// dibikin semirip mungkin: getItem/setItem/deleteItem
// berdasarkan "key", sama seperti window.storage.get/set/delete.
//
// Tabelnya cuma key-value sederhana (lihat SQL di bawah),
// jadi hampir semua logic di KartuStokPFA.jsx yang sudah
// ada (JSON.stringify/JSON.parse per productId) tetap sama,
// cuma sumber datanya pindah dari browser storage ke database.
// =====================================================

const TABLE = "app_kv_store";

// Ambil value berdasarkan key. Return null kalau belum ada datanya.
export async function getItem(key) {
  const { data, error } = await supabase
    .from(TABLE)
    .select("value")
    .eq("key", key)
    .maybeSingle();

  if (error) throw error;
  if (!data) return null;

  return { key, value: data.value };
}

// Simpan/update value untuk sebuah key.
export async function setItem(key, value) {
  const { error } = await supabase
    .from(TABLE)
    .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: "key" });

  if (error) throw error;

  return { key, value };
}

// Hapus sebuah key.
export async function deleteItem(key) {
  const { error } = await supabase.from(TABLE).delete().eq("key", key);

  if (error) throw error;

  return { key, deleted: true };
}