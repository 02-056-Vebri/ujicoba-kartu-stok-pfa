"use client";

import { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { exportExcelProfessional } from "./exportExcel";

// ---------- Data produk (diambil dari STOCK_CARD_PFA_DES_2025.xlsx) ----------
const RAW_PRODUCTS = [
  "PFA 86%", "PFA 88%",
  "PFA 92% JB250", "PFA 92% JB300", "PFA 92% JB800", "PFA 92% ZAK25",
  "PFA 92% ZAK25 KITE", "PFA 92% JB500", "PFA 92% JB500 KITE",
  "92% JB300", "92% JB300 KITE", "PFA 92% JB1000", "PFA 92% JB1000 KITE",
  "PFA 96% ZAK 25 KITE", "PFA 96% ZAK 25", "PFA 96% JB 450", "PFA 96% JB 450 KITE",
  "PFA 96% JB500", "PFA 96% JB500 KITE", "PFA 96% JB1000", "PFA 96% JB1000 KITE",
  "PFA 97% ZAK 20 KITE", "PFA 97% ZAK 500 KITE",
  "PFA Others @20", "PFA Others @25", "PFA Transisi", "PFA Others @jb 500",
];

function categorize(name) {
  if (/86%/.test(name)) return "86%";
  if (/88%/.test(name)) return "88%";
  if (/92%/.test(name)) return "92%";
  if (/96%/.test(name)) return "96%";
  if (/97%/.test(name)) return "97%";
  return "Lainnya";
}
function slugify(name) {
  return name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
}
const CATEGORY_ORDER = ["86%", "88%", "92%", "96%", "97%", "Lainnya"];
const PRODUCTS = RAW_PRODUCTS.map((name) => {
  const trimmed = name.trim();
  return { id: slugify(trimmed), name: trimmed, category: categorize(trimmed) };
});

const MONTHS_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
function fmtDate(iso) {
  if (!iso) return "-";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return iso;
  return `${String(d.getDate()).padStart(2, "0")} ${MONTHS_ID[d.getMonth()]} ${d.getFullYear()}`;
}
function monthKey(iso) {
  return iso ? iso.slice(0, 7) : "";
}
function monthLabel(key) {
  const [y, m] = key.split("-");
  return `${MONTHS_ID[parseInt(m, 10) - 1]} ${y}`;
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function numFmt(n) {
  return new Intl.NumberFormat("id-ID").format(n || 0);
}

function getPackagingUnit(name) {
  const upper = name.toUpperCase();
  let m = upper.match(/JB\s?(\d+)/);
  if (m) return `JB ${m[1]}`;
  m = upper.match(/ZAK\s?(\d+)/);
  if (m) return `ZAK ${m[1]}`;
  m = upper.match(/@\s?(\d+)/);
  if (m) return `ZAK ${m[1]}`;
  return "Zak";
}

const SEED_TXNS = {
  "pfa-86": [
    { id: "seed-1", date: "2025-06-16", ref: "EXT25-00207", type: "in", zak: 41, kg: 1025, lot: "", tujuan: "", lokasi: "", ket: "Penerimaan bahan baku" },
    { id: "seed-2", date: "2025-09-01", ref: "PA25-08160", type: "out", zak: 120, kg: 3000, lot: "", tujuan: "", lokasi: "", ket: "Pemakaian produksi" },
  ],
};

const EMPTY_FORM = { date: todayISO(), ref: "", type: "in", zak: "", kg: "", lot: "", tujuan: "", lokasi: "", ket: "" };

function computeBalance(txns) {
  const sorted = [...txns].sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  let zak = 0, kg = 0;
  return sorted.map((t) => {
    if (t.type === "in") { zak += Number(t.zak) || 0; kg += Number(t.kg) || 0; }
    else { zak -= Number(t.zak) || 0; kg -= Number(t.kg) || 0; }
    return { ...t, sisaZak: zak, sisaKg: kg };
  });
}

function buildSheetHeader(unit) {
  return ["Tanggal", "Referensi", `Masuk (${unit})`, "Masuk (Kg)", `Keluar (${unit})`, "Keluar (Kg)", `Sisa (${unit})`, "Total (Kg)", "Lokasi", "Lot No", "Pallet Kayu", "Keterangan"];
}

function buildSheetRows(productName, balancedTxns, unit) {

  const rows = [];

  // Header tabel saja
  rows.push(buildSheetHeader(unit));

  // Isi data
  for (const t of balancedTxns) {
    rows.push([
      fmtDate(t.date),
      t.ref || "",
      t.type === "in" ? t.zak : "",
      t.type === "in" ? t.kg : "",
      t.type === "out" ? t.zak : "",
      t.type === "out" ? t.kg : "",
      t.sisaZak,
      t.sisaKg,
      t.lokasi || "",
      t.lot || "",
      t.tujuan || "",
      t.ket || "",
    ]);
  }

  if (balancedTxns.length === 0) {
    rows.push(["Belum ada transaksi"]);
  }

  return rows;
}

export default function KartuStokPFA() {
  const [role, setRole] = useState("editor"); // 'editor' | 'viewer'
  const [selectedId, setSelectedId] = useState(PRODUCTS[0].id);
  const [search, setSearch] = useState("");
  const [txnsByProduct, setTxnsByProduct] = useState({});
  const [loadingProduct, setLoadingProduct] = useState(false);
  const [formOpen, setFormOpen] = useState(false);
  const [form, setForm] = useState(EMPTY_FORM);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [toast, setToast] = useState(null);
  const [formError, setFormError] = useState("");
  const [monthlyOpen, setMonthlyOpen] = useState(false);
  const [filterMode, setFilterMode] = useState("all"); // 'all' | 'day' | 'month' | 'year' | 'range'
  const [filterValue, setFilterValue] = useState("");
  const [rangeFrom, setRangeFrom] = useState("");
  const [rangeTo, setRangeTo] = useState("");
  const [calendarOpen, setCalendarOpen] = useState(false);
  const [exportScope, setExportScope] = useState("current"); // 'current' | 'all'
  const now = new Date();
  const [viewYear, setViewYear] = useState(now.getFullYear());
  const [viewMonth, setViewMonth] = useState(now.getMonth()); // 0-indexed

  const selectedProduct = PRODUCTS.find((p) => p.id === selectedId);
  const unit = useMemo(() => getPackagingUnit(selectedProduct.name), [selectedProduct]);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2200);
  }, []);

  const loadProduct = useCallback(async (productId) => {
    if (txnsByProduct[productId]) return;
    setLoadingProduct(true);
    try {
      const res = await window.storage.get(`txns:${productId}`, true);
      const parsed = res ? JSON.parse(res.value) : [];
      setTxnsByProduct((prev) => ({ ...prev, [productId]: parsed }));
    } catch (e) {
      const seeded = SEED_TXNS[productId] || [];
      setTxnsByProduct((prev) => ({ ...prev, [productId]: seeded }));
      if (seeded.length) {
        try { await window.storage.set(`txns:${productId}`, JSON.stringify(seeded), true); } catch (_) {}
      }
    } finally {
      setLoadingProduct(false);
    }
  }, [txnsByProduct]);

  useEffect(() => { loadProduct(selectedId); }, [selectedId, loadProduct]);

  const persist = useCallback((productId, next) => {
    setTxnsByProduct((prev) => ({ ...prev, [productId]: next }));
    if (typeof window === "undefined" || !window.storage) {
      showToast("Tersimpan di sesi ini (penyimpanan permanen tidak tersedia).");
      return;
    }
    window.storage.set(`txns:${productId}`, JSON.stringify(next), true)
      .then((res) => { if (!res) showToast("Gagal menyimpan ke penyimpanan permanen."); })
      .catch(() => showToast("Gagal menyimpan ke penyimpanan permanen."));
  }, [showToast]);

  const rawTxns = txnsByProduct[selectedId] || [];

  const sortedWithBalance = useMemo(() => computeBalance(rawTxns), [rawTxns]);

  const txnDatesSet = useMemo(() => new Set(sortedWithBalance.map((t) => t.date)), [sortedWithBalance]);

  const matchesFilter = useCallback((t) => {
    if (filterMode === "day") return t.date === filterValue;
    if (filterMode === "month") return t.date.slice(0, 7) === filterValue;
    if (filterMode === "year") return t.date.slice(0, 4) === filterValue;
    if (filterMode === "range") return (!rangeFrom || t.date >= rangeFrom) && (!rangeTo || t.date <= rangeTo);
    return true;
  }, [filterMode, filterValue, rangeFrom, rangeTo]);

  const displayedTxns = useMemo(() => sortedWithBalance.filter(matchesFilter), [sortedWithBalance, matchesFilter]);

  const periodLabel = useMemo(() => {
    if (filterMode === "day") return fmtDate(filterValue);
    if (filterMode === "month") return monthLabel(filterValue);
    if (filterMode === "year") return filterValue;
    if (filterMode === "range") return `${fmtDate(rangeFrom)} – ${fmtDate(rangeTo)}`;
    return "Semua Periode";
  }, [filterMode, filterValue, rangeFrom, rangeTo]);

  const filenameSlug = useMemo(() => {
    if (filterMode === "day") return filterValue;
    if (filterMode === "month") return filterValue;
    if (filterMode === "year") return filterValue;
    if (filterMode === "range") return `${rangeFrom}_sd_${rangeTo}`;
    return "semua";
  }, [filterMode, filterValue, rangeFrom, rangeTo]);

  function pad2(n) { return String(n).padStart(2, "0"); }

  function goPrevMonth() {
    if (viewMonth === 0) { setViewMonth(11); setViewYear((y) => y - 1); } else setViewMonth((m) => m - 1);
  }
  function goNextMonth() {
    if (viewMonth === 11) { setViewMonth(0); setViewYear((y) => y + 1); } else setViewMonth((m) => m + 1);
  }
  function goPrevYear() { setViewYear((y) => y - 1); }
  function goNextYear() { setViewYear((y) => y + 1); }

  function pickDay(dateStr) {
    setFilterMode("day"); setFilterValue(dateStr); setCalendarOpen(false);
  }
  function pickToday() {
    setFilterMode("day"); setFilterValue(todayISO()); setCalendarOpen(false);
  }
  function pickWholeMonth() {
    setFilterMode("month"); setFilterValue(`${viewYear}-${pad2(viewMonth + 1)}`); setCalendarOpen(false);
  }
  function pickWholeYear() {
    setFilterMode("year"); setFilterValue(`${viewYear}`); setCalendarOpen(false);
  }
  function applyRange() {
    if (!rangeFrom || !rangeTo) { showToast("Isi tanggal awal dan akhir dulu."); return; }
    if (rangeFrom > rangeTo) { showToast("Tanggal awal harus sebelum tanggal akhir."); return; }
    setFilterMode("range"); setCalendarOpen(false);
  }
  function resetFilter() { setFilterMode("all"); setFilterValue(""); setRangeFrom(""); setRangeTo(""); setCalendarOpen(false); }

  const calendarCells = useMemo(() => {
    const firstWeekday = new Date(viewYear, viewMonth, 1).getDay();
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
    const cells = [];
    for (let i = 0; i < firstWeekday; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) {
      cells.push(`${viewYear}-${pad2(viewMonth + 1)}-${pad2(d)}`);
    }
    return cells;
  }, [viewYear, viewMonth]);

  const [exporting, setExporting] = useState(false);

  const handleExport = useCallback(async () => {
    if (exportScope === "current") {
      if (displayedTxns.length === 0) {
        showToast(`Tidak ada transaksi pada periode ${periodLabel} untuk produk ini.`);
        return;
      }
      const rows = buildSheetRows(
        `${selectedProduct.name} - ${periodLabel}`,
        displayedTxns,
        unit
      );
      await exportExcelProfessional(
        rows,
        selectedProduct.name.slice(0, 31),
        `Kartu-Stok-${selectedProduct.id}-${filenameSlug}.xlsx`
      );
      showToast(`File Excel (${selectedProduct.name}, ${periodLabel}) berhasil dibuat.`);
      return;
    }

    setExporting(true);
    try {
      const wb = XLSX.utils.book_new();
      const usedNames = new Set();
      let anyData = false;
      for (const p of PRODUCTS) {
        let txns = txnsByProduct[p.id];
        if (!txns) {
          try {
            const res = await window.storage.get(`txns:${p.id}`, true);
            txns = res ? JSON.parse(res.value) : [];
          } catch (e) {
            txns = SEED_TXNS[p.id] || [];
          }
        }
        const balanced = computeBalance(txns).filter(matchesFilter);
        if (balanced.length) anyData = true;
        const productUnit = getPackagingUnit(p.name);
        const ws = XLSX.utils.aoa_to_sheet(buildSheetRows(p.name, balanced, productUnit));
        let sheetName = p.name.slice(0, 31) || p.id;
        let i = 1;
        while (usedNames.has(sheetName)) { sheetName = `${p.name.slice(0, 27)}_${i++}`; }
        usedNames.add(sheetName);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }
      if (!anyData) {
        showToast(`Tidak ada transaksi pada periode ${periodLabel} di semua produk.`);
        return;
      }
      XLSX.writeFile(wb, `Kartu-Stok-Semua-Produk-${filenameSlug}.xlsx`);
      showToast(`File Excel (Semua Produk, ${periodLabel}) berhasil dibuat.`);
    } catch (e) {
      showToast("Gagal membuat file Excel.");
    } finally {
      setExporting(false);
    }
  }, [exportScope, displayedTxns, periodLabel, selectedProduct, unit, filenameSlug, txnsByProduct, matchesFilter, showToast]);

  const monthlyRecap = useMemo(() => {
    const map = new Map();
    for (const t of sortedWithBalance) {
      map.set(monthKey(t.date), { sisaZak: t.sisaZak, sisaKg: t.sisaKg });
    }
    return Array.from(map.entries()).sort(([a], [b]) => (a < b ? -1 : 1));
  }, [sortedWithBalance]);

  const currentStock = sortedWithBalance.length
    ? sortedWithBalance[sortedWithBalance.length - 1]
    : { sisaZak: 0, sisaKg: 0 };

  const filteredProducts = PRODUCTS.filter((p) =>
    p.name.toLowerCase().includes(search.toLowerCase())
  );
  const grouped = CATEGORY_ORDER.map((cat) => ({
    cat,
    items: filteredProducts.filter((p) => p.category === cat),
  })).filter((g) => g.items.length);

  function openAddForm() {
    setForm(EMPTY_FORM);
    setFormError("");
    setFormOpen(true);
  }

  async function submitForm(e) {
    e.preventDefault();
    const zakNum = Number(form.zak) || 0;
    const kgNum = Number(form.kg) || 0;
    if (!form.date) {
      setFormError("Tanggal wajib diisi.");
      return;
    }
    if (zakNum <= 0 && kgNum <= 0) {
      setFormError(`Isi jumlah (${unit}) atau berat (Kg) — minimal salah satu harus lebih dari 0.`);
      return;
    }
    setFormError("");
    const entry = {
      id: `t-${Date.now()}`,
      date: form.date,
      ref: form.ref.trim(),
      type: form.type,
      zak: zakNum,
      kg: kgNum,
      lot: form.lot.trim(),
      tujuan: form.tujuan,
      lokasi: form.lokasi.trim(),
      ket: form.ket.trim(),
    };
    const next = [...rawTxns, entry];
    await persist(selectedId, next);
    setFormOpen(false);
    showToast("Transaksi tersimpan.");
  }

  function deleteTxn(id) {
    const next = rawTxns.filter((t) => t.id !== id);
    persist(selectedId, next);
    setConfirmDeleteId(null);
    showToast("Transaksi dihapus.");
  }

  return (
    <div className="ks-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

        .ks-root {
          --bg: #14181c;
          --panel: #1b2126;
          --panel-alt: #212830;
          --border: #2b333b;
          --text: #e9edf0;
          --muted: #8b96a1;
          --accent: #e8a33d;
          --accent-dim: rgba(232,163,61,0.14);
          --positive: #4fa98a;
          --positive-dim: rgba(79,169,138,0.14);
          --negative: #d9694f;
          --negative-dim: rgba(217,105,79,0.14);
          font-family: 'Inter', sans-serif;
          background: var(--bg);
          color: var(--text);
          border-radius: 10px;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          min-height: 640px;
          max-width: 100%;
          border: 1px solid var(--border);
        }
        .ks-root * { box-sizing: border-box; }
        .ks-topbar {
          display: flex; align-items: center; justify-content: space-between;
          gap: 12px; padding: 14px 18px; border-bottom: 1px solid var(--border);
          background: var(--panel); flex-wrap: wrap;
        }
        .ks-title { font-family: 'Oswald', sans-serif; font-size: 19px; font-weight: 600;
          letter-spacing: 0.03em; text-transform: uppercase; display: flex; align-items: center; gap: 10px; }
        .ks-title small { font-family: 'Inter', sans-serif; text-transform: none; font-weight: 400;
          font-size: 12px; color: var(--muted); letter-spacing: 0; }
        .ks-role-select {
          background: var(--panel-alt); color: var(--text); border: 1px solid var(--border);
          border-radius: 6px; padding: 7px 10px; font-size: 13px; font-family: 'Inter', sans-serif;
        }
        .ks-body { display: flex; flex: 1; min-height: 0; }
        .ks-sidebar { width: 250px; border-right: 1px solid var(--border); background: var(--panel);
          display: flex; flex-direction: column; }
        .ks-search { padding: 12px; border-bottom: 1px solid var(--border); }
        .ks-search input {
          width: 100%; background: var(--panel-alt); border: 1px solid var(--border); color: var(--text);
          padding: 8px 10px; border-radius: 6px; font-size: 13px; font-family: 'Inter', sans-serif;
        }
        .ks-search input:focus { outline: 1px solid var(--accent); }
        .ks-list { overflow-y: auto; flex: 1; padding: 6px 0; }
        .ks-cat-label { padding: 10px 16px 4px; font-size: 11px; color: var(--muted); letter-spacing: 0.08em;
          text-transform: uppercase; font-weight: 600; }
        .ks-item {
          width: 100%; text-align: left; background: transparent; border: none; color: var(--text);
          padding: 8px 16px; font-size: 13.5px; cursor: pointer; border-left: 3px solid transparent;
          font-family: 'Inter', sans-serif;
        }
        .ks-item:hover { background: var(--panel-alt); }
        .ks-item.active { background: var(--accent-dim); border-left-color: var(--accent); font-weight: 600; }
        .ks-main { flex: 1; padding: 20px 24px; overflow-y: auto; min-width: 0; }
        .ks-headrow { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 16px; }
        .ks-product-name { font-family: 'Oswald', sans-serif; font-size: 24px; font-weight: 600; letter-spacing: 0.01em; }
        .ks-btn {
          background: var(--accent); color: #17140d; border: none; border-radius: 6px; padding: 9px 16px;
          font-size: 13.5px; font-weight: 600; cursor: pointer; font-family: 'Inter', sans-serif;
        }
        .ks-btn:hover { filter: brightness(1.08); }
        .ks-btn.ghost { background: transparent; color: var(--text); border: 1px solid var(--border); }
        .ks-btn:disabled { opacity: 0.55; cursor: not-allowed; }
        .ks-exportbar { display: flex; align-items: center; justify-content: space-between; gap: 10px;
          flex-wrap: wrap; margin-bottom: 14px; padding: 10px 12px; background: var(--panel);
          border: 1px solid var(--border); border-radius: 8px; }
        .ks-scope-toggle { display: flex; border: 1px solid var(--border); border-radius: 6px; overflow: hidden; }
        .ks-scope-btn { background: var(--panel-alt); color: var(--muted); border: none; padding: 8px 14px;
          font-size: 13px; cursor: pointer; font-family: 'Inter', sans-serif; }
        .ks-scope-btn.active { background: var(--accent-dim); color: var(--accent); font-weight: 600; }
        .ks-cal-range { margin-top: 10px; padding-top: 10px; border-top: 1px solid var(--border); }
        .ks-cal-range-inputs { display: flex; align-items: center; gap: 6px; margin: 6px 0; }
        .ks-cal-range-inputs input {
          flex: 1; background: var(--panel-alt); border: 1px solid var(--border); color: var(--text);
          padding: 6px 7px; border-radius: 6px; font-size: 12px; font-family: 'Inter', sans-serif; min-width: 0;
        }
        .ks-cal-range-inputs span { color: var(--muted); font-size: 12px; }
        .ks-filterbar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
        .ks-filter-label { font-size: 12px; color: var(--muted); }
        .ks-filter-count { font-size: 12px; color: var(--muted); margin-left: 4px; }
        .ks-calendar-wrap { position: relative; display: inline-block; }
        .ks-calendar-trigger { font-family: 'JetBrains Mono', monospace; font-size: 12.5px; }
        .ks-calendar-popover {
          position: absolute; top: calc(100% + 6px); left: 0; z-index: 40; width: 268px;
          background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 12px;
          box-shadow: 0 12px 28px rgba(0,0,0,0.45);
        }
        .ks-cal-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
        .ks-cal-title { font-family: 'Oswald', sans-serif; font-size: 13.5px; letter-spacing: 0.03em; text-transform: uppercase; }
        .ks-cal-nav-btn {
          background: var(--panel-alt); border: 1px solid var(--border); color: var(--text);
          border-radius: 5px; width: 24px; height: 24px; font-size: 12px; cursor: pointer;
        }
        .ks-cal-nav-btn:hover { background: var(--accent-dim); }
        .ks-cal-weekdays { display: grid; grid-template-columns: repeat(7, 1fr); margin-bottom: 4px; }
        .ks-cal-weekday { text-align: center; font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.04em; }
        .ks-cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }
        .ks-cal-day {
          position: relative; aspect-ratio: 1; background: transparent; border: none; color: var(--text);
          font-size: 12px; font-family: 'JetBrains Mono', monospace; border-radius: 6px; cursor: pointer;
        }
        .ks-cal-day.empty { cursor: default; }
        .ks-cal-day:not(.empty):hover { background: var(--panel-alt); }
        .ks-cal-day.today { outline: 1px solid var(--border); }
        .ks-cal-day.selected { background: var(--accent); color: #17140d; font-weight: 700; }
        .ks-cal-day.has-data::after {
          content: ""; position: absolute; bottom: 3px; left: 50%; transform: translateX(-50%);
          width: 4px; height: 4px; border-radius: 50%; background: var(--positive);
        }
        .ks-cal-day.selected.has-data::after { background: #17140d; }
        .ks-cal-actions { display: flex; gap: 8px; margin-top: 10px; }
        .ks-btn.small { padding: 6px 10px; font-size: 12px; }
        .ks-table-wrap { border: 1px solid var(--border); border-radius: 8px; overflow: hidden; }
        table.ks-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        table.ks-table thead th {
          background: var(--panel-alt); text-align: left; padding: 9px 10px; font-size: 11px;
          text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); border-bottom: 1px solid var(--border);
          white-space: nowrap;
        }
        table.ks-table tbody td { padding: 9px 10px; border-bottom: 1px solid var(--border); vertical-align: top; }
        table.ks-table tbody tr:last-child td { border-bottom: none; }
        table.ks-table tbody tr:hover { background: rgba(255,255,255,0.02); }
        .ks-mono { font-family: 'JetBrains Mono', monospace; }
        .ks-in { color: var(--positive); }
        .ks-out { color: var(--negative); }
        .ks-sisa-badge { font-family: 'JetBrains Mono', monospace; background: var(--accent-dim); color: var(--accent);
          padding: 3px 8px; border-radius: 5px; font-weight: 600; }
        .ks-empty { padding: 40px 20px; text-align: center; color: var(--muted); font-size: 13.5px; }
        .ks-del-btn { background: transparent; border: none; color: var(--negative); cursor: pointer; font-size: 12px; }
        .ks-modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.55); display: flex;
          align-items: center; justify-content: center; z-index: 50; }
        .ks-modal { background: var(--panel); border: 1px solid var(--border); border-radius: 10px; width: 460px;
          max-width: 92vw; padding: 20px; max-height: 88vh; overflow-y: auto; }
        .ks-modal h3 { font-family: 'Oswald', sans-serif; margin: 0 0 14px; font-size: 17px; letter-spacing: 0.02em; text-transform: uppercase; }
        .ks-field { margin-bottom: 11px; }
        .ks-field label { display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }
        .ks-field input, .ks-field select, .ks-field textarea {
          width: 100%; background: var(--panel-alt); border: 1px solid var(--border); color: var(--text);
          padding: 8px 10px; border-radius: 6px; font-size: 13px; font-family: 'Inter', sans-serif;
        }
        .ks-row2 { display: flex; gap: 10px; }
        .ks-row2 > div { flex: 1; }
        .ks-modal-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
        .ks-form-error { background: var(--negative-dim); color: var(--negative); border: 1px solid var(--negative);
          border-radius: 6px; padding: 9px 12px; font-size: 12.5px; margin-top: 6px; }
        .ks-toast { position: fixed; bottom: 18px; left: 50%; transform: translateX(-50%);
          background: var(--panel-alt); border: 1px solid var(--accent); color: var(--text);
          padding: 9px 16px; border-radius: 8px; font-size: 13px; z-index: 45; pointer-events: none; }
        .ks-monthly { margin-top: 18px; border: 1px solid var(--border); border-radius: 8px; }
        .ks-monthly-head { padding: 11px 14px; cursor: pointer; display: flex; justify-content: space-between;
          align-items: center; font-size: 13px; font-weight: 600; background: var(--panel); }
        .ks-monthly-body { padding: 0 14px 12px; }
        .ks-monthly-row { display: flex; justify-content: space-between; padding: 6px 0; font-size: 12.5px;
          border-bottom: 1px solid var(--border); }
        .ks-monthly-row:last-child { border-bottom: none; }
        .ks-loading { padding: 30px; text-align: center; color: var(--muted); font-size: 13px; }
        @media (max-width: 640px) {
          .ks-body { flex-direction: column; }
          .ks-sidebar { width: 100%; max-height: 220px; }
        }
      `}</style>

      <div className="ks-topbar">
        <div className="ks-title">
          Kartu Stok — Paraformaldehyde
          <small>Prototipe (demo)</small>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 12, color: "var(--muted)" }}>Login sebagai</span>
          <select className="ks-role-select" value={role} onChange={(e) => setRole(e.target.value)}>
            <option value="editor">Staff / Atasan (bisa edit)</option>
            <option value="viewer">Manager / Lainnya (lihat saja)</option>
          </select>
        </div>
      </div>

      <div className="ks-body">
        <div className="ks-sidebar">
          <div className="ks-search">
            <input placeholder="Cari produk..." value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <div className="ks-list">
            {grouped.map((g) => (
              <div key={g.cat}>
                <div className="ks-cat-label">PFA {g.cat}</div>
                {g.items.map((p) => (
                  <button
                    key={p.id}
                    className={`ks-item ${p.id === selectedId ? "active" : ""}`}
                    onClick={() => setSelectedId(p.id)}
                  >
                    {p.name}
                  </button>
                ))}
              </div>
            ))}
            {grouped.length === 0 && (
              <div style={{ padding: 16, fontSize: 13, color: "var(--muted)" }}>Tidak ditemukan.</div>
            )}
          </div>
        </div>

        <div className="ks-main">
          <div className="ks-headrow">
            <div>
              <div className="ks-product-name">{selectedProduct.name}</div>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              {role === "editor" && (
                <button className="ks-btn" onClick={openAddForm}>+ Tambah Transaksi</button>
              )}
            </div>
          </div>

          <div className="ks-exportbar">
            <div className="ks-scope-toggle">
              <button className={`ks-scope-btn ${exportScope === "current" ? "active" : ""}`} onClick={() => setExportScope("current")}>
                Produk ini
              </button>
              <button className={`ks-scope-btn ${exportScope === "all" ? "active" : ""}`} onClick={() => setExportScope("all")}>
                Semua produk
              </button>
            </div>
            <button className="ks-btn" onClick={handleExport} disabled={exporting}>
              {exporting ? "Menyiapkan..." : `⬇ Export ke Excel (${periodLabel})`}
            </button>
          </div>

          <div className="ks-filterbar">
            <span className="ks-filter-label">Cari periode:</span>
            <div className="ks-calendar-wrap">
              <button className="ks-btn ghost ks-calendar-trigger" onClick={() => setCalendarOpen((o) => !o)}>
                📅 {periodLabel}
              </button>
              {calendarOpen && (
                <div className="ks-calendar-popover">
                  <div className="ks-cal-header">
                    <button className="ks-cal-nav-btn" onClick={goPrevYear} title="Tahun sebelumnya">«</button>
                    <button className="ks-cal-nav-btn" onClick={goPrevMonth} title="Bulan sebelumnya">‹</button>
                    <span className="ks-cal-title">{MONTHS_ID[viewMonth]} {viewYear}</span>
                    <button className="ks-cal-nav-btn" onClick={goNextMonth} title="Bulan berikutnya">›</button>
                    <button className="ks-cal-nav-btn" onClick={goNextYear} title="Tahun berikutnya">»</button>
                  </div>
                  <div className="ks-cal-weekdays">
                    {["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"].map((w) => (
                      <div key={w} className="ks-cal-weekday">{w}</div>
                    ))}
                  </div>
                  <div className="ks-cal-grid">
                    {calendarCells.map((dateStr, idx) => {
                      if (!dateStr) return <div key={idx} className="ks-cal-day empty" />;
                      const day = parseInt(dateStr.slice(8, 10), 10);
                      const hasData = txnDatesSet.has(dateStr);
                      const isSelected = filterMode === "day" && filterValue === dateStr;
                      const isToday = dateStr === todayISO();
                      return (
                        <button
                          key={dateStr}
                          className={`ks-cal-day ${hasData ? "has-data" : ""} ${isSelected ? "selected" : ""} ${isToday ? "today" : ""}`}
                          onClick={() => pickDay(dateStr)}
                          title={hasData ? "Ada transaksi" : ""}
                        >
                          {day}
                        </button>
                      );
                    })}
                  </div>
                  <div className="ks-cal-actions">
                    <button className="ks-btn ghost small" onClick={pickToday}>Hari Ini</button>
                    <button className="ks-btn ghost small" onClick={pickWholeMonth}>Bulan Ini</button>
                    <button className="ks-btn ghost small" onClick={pickWholeYear}>Tahun Ini</button>
                  </div>
                  <div className="ks-cal-range">
                    <span className="ks-filter-label">Atau rentang tanggal custom:</span>
                    <div className="ks-cal-range-inputs">
                      <input type="date" value={rangeFrom} onChange={(e) => setRangeFrom(e.target.value)} />
                      <span>–</span>
                      <input type="date" value={rangeTo} onChange={(e) => setRangeTo(e.target.value)} />
                    </div>
                    <button className="ks-btn ghost small" onClick={applyRange} style={{ width: "100%" }}>Terapkan Rentang</button>
                  </div>
                </div>
              )}
            </div>
            {filterMode !== "all" && (
              <button className="ks-btn ghost" onClick={resetFilter}>Reset</button>
            )}
            <span className="ks-filter-count">{displayedTxns.length} transaksi ditemukan</span>
          </div>

          {loadingProduct ? (
            <div className="ks-loading">Memuat data...</div>
          ) : (
            <>
              <div className="ks-table-wrap">
                <table className="ks-table">
                  <thead>
                    <tr>
                      <th>Tanggal</th>
                      <th>Referensi</th>
                      <th>Masuk</th>
                      <th>Keluar</th>
                      <th>Sisa</th>
                      <th>Total</th>
                      <th>Lokasi</th>
                      <th>Lot No</th>
                      <th>Pallet Kayu</th>
                      <th>Keterangan</th>
                      {role === "editor" && <th></th>}
                    </tr>
                  </thead>
                  <tbody>
                    {displayedTxns.length === 0 && (
                      <tr>
                        <td colSpan={role === "editor" ? 11 : 10}>
                          <div className="ks-empty">
                            {sortedWithBalance.length === 0
                              ? <>Belum ada transaksi tercatat untuk produk ini.{role === "editor" && " Klik \u201c+ Tambah Transaksi\u201d untuk mulai mencatat."}</>
                              : `Tidak ada transaksi pada periode ${periodLabel}.`}
                          </div>
                        </td>
                      </tr>
                    )}
                    {displayedTxns.slice().reverse().map((t) => (
                      <tr key={t.id}>
                        <td className="ks-mono">{fmtDate(t.date)}</td>
                        <td className="ks-mono">{t.ref || "-"}</td>
                        <td className="ks-mono ks-in">{t.type === "in" ? `${numFmt(t.zak)} ${unit} / ${numFmt(t.kg)} kg` : "-"}</td>
                        <td className="ks-mono ks-out">{t.type === "out" ? `${numFmt(t.zak)} ${unit} / ${numFmt(t.kg)} kg` : "-"}</td>
                        <td><span className="ks-sisa-badge">{numFmt(t.sisaZak)} {unit}</span></td>
                        <td className="ks-mono">{numFmt(t.sisaKg)} kg</td>
                        <td>{t.lokasi || "-"}</td>
                        <td className="ks-mono">{t.lot || "-"}</td>
                        <td>{t.tujuan || "-"}</td>
                        <td>{t.ket || "-"}</td>
                        {role === "editor" && (
                          <td>
                            {confirmDeleteId === t.id ? (
                              <div style={{ display: "flex", gap: 6 }}>
                                <button className="ks-del-btn" onClick={() => deleteTxn(t.id)}>Yakin?</button>
                                <button className="ks-del-btn" style={{ color: "var(--muted)" }} onClick={() => setConfirmDeleteId(null)}>Batal</button>
                              </div>
                            ) : (
                              <button className="ks-del-btn" onClick={() => setConfirmDeleteId(t.id)}>Hapus</button>
                            )}
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="ks-monthly">
                <div className="ks-monthly-head" onClick={() => setMonthlyOpen(!monthlyOpen)}>
                  <span>Rekap Bulanan (otomatis)</span>
                  <span>{monthlyOpen ? "−" : "+"}</span>
                </div>
                {monthlyOpen && (
                  <div className="ks-monthly-body">
                    {monthlyRecap.length === 0 && (
                      <div style={{ padding: "8px 0", color: "var(--muted)", fontSize: 12.5 }}>Belum ada data.</div>
                    )}
                    {monthlyRecap.map(([key, v]) => (
                      <div className="ks-monthly-row" key={key}>
                        <span>{monthLabel(key)}</span>
                        <span className="ks-mono">{numFmt(v.sisaZak)} {unit} / {numFmt(v.sisaKg)} kg</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>

      {formOpen && (
        <div className="ks-modal-overlay" onClick={(e) => e.target === e.currentTarget && setFormOpen(false)}>
          <div className="ks-modal">
            <h3>Tambah Transaksi</h3>
            <form onSubmit={submitForm}>
              <div className="ks-row2">
                <div className="ks-field">
                  <label>Tanggal</label>
                  <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
                </div>
                <div className="ks-field">
                  <label>Jenis</label>
                  <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                    <option value="in">Masuk</option>
                    <option value="out">Keluar</option>
                  </select>
                </div>
              </div>
              <div className="ks-field">
                <label>No. Referensi</label>
                <input value={form.ref} onChange={(e) => setForm({ ...form, ref: e.target.value })} placeholder="mis. EXT25-00207" />
              </div>
              <div className="ks-row2">
                <div className="ks-field">
                  <label>Jumlah ({unit})</label>
                  <input type="number" value={form.zak} onChange={(e) => setForm({ ...form, zak: e.target.value })} placeholder="0" />
                </div>
                <div className="ks-field">
                  <label>Berat Total (Kg)</label>
                  <input type="number" value={form.kg} onChange={(e) => setForm({ ...form, kg: e.target.value })} placeholder="0" />
                </div>
              </div>
              <div className="ks-row2">
                <div className="ks-field">
                  <label>Lot No.</label>
                  <input value={form.lot} onChange={(e) => setForm({ ...form, lot: e.target.value })} />
                </div>
                <div className="ks-field">
                  <label>Pallet Kayu</label>
                  <select value={form.tujuan} onChange={(e) => setForm({ ...form, tujuan: e.target.value })}>
                    <option value="">-</option>
                    <option value="Ekspor">Ekspor</option>
                    <option value="Lokal">Lokal</option>
                  </select>
                </div>
              </div>
              <div className="ks-field">
                <label>Lokasi</label>
                <input value={form.lokasi} onChange={(e) => setForm({ ...form, lokasi: e.target.value })} placeholder="mis. Gudang A - Rak 3" />
              </div>
              <div className="ks-field">
                <label>Keterangan</label>
                <textarea rows={2} value={form.ket} onChange={(e) => setForm({ ...form, ket: e.target.value })} />
              </div>
              {formError && <div className="ks-form-error">⚠ {formError}</div>}
              <div className="ks-modal-actions">
                <button type="button" className="ks-btn ghost" onClick={() => setFormOpen(false)}>Batal</button>
                <button type="submit" className="ks-btn">Simpan</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {toast && <div className="ks-toast">{toast}</div>}
    </div>
  );
}