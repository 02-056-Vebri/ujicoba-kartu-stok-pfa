"use client";

import { useState, useEffect, useMemo, useCallback } from "react";
import { exportExcelProfessional, exportExcelAllProducts, exportExcelResume } from "./exportExcel";
import { getItem, setItem, deleteItem } from "./appStorage";

// ---------- Data produk (diambil dari STOCK_CARD_PFA_DES_2025.xlsx) ----------
const RAW_PRODUCTS = [
  "PFA 86% ZAK25", "PFA 88% ZAK25",
  "PFA 92% ZAK25", "PFA 92% ZAK25 KITE", "PFA 92% JB500", "PFA 92% JB500 KITE",
  "PFA 92% JB300", "PFA 92% JB300 KITE", "PFA 92% JB1000", "PFA 92% JB1000 KITE",
  "PFA 96% ZAK25", "PFA 96% ZAK25 KITE", "PFA 96% JB450", "PFA 96% JB450 KITE",
  "PFA 96% JB500", "PFA 96% JB500 KITE", "PFA 96% JB1000", "PFA 96% JB1000 KITE",
  "PFA 97% ZAK20 KITE", "PFA 97% ZAK500 KITE",
  "PFA Others @20KG", "PFA Others @25KG", "PFA Transisi @25KG", "PFA Others @JB500",
];

function categorize(name) {
  const m = name.match(/(\d{1,3})%/);
  if (m) return `${m[1]}%`;
  return "Lainnya";
}
function slugify(name) {
  return name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
}
// Urutan kategori dihitung otomatis di dalam komponen (lihat variabel productCategories),
// supaya kalau ada persentase baru (misal 99%) otomatis kesortir sesuai angka, bukan hardcode.
// Menjaga id produk tetap sama seperti sebelumnya walau nama tampilannya berubah,
// supaya data transaksi yang sudah tersimpan (dikunci ke id lama) tidak hilang/terpisah.
const ID_OVERRIDES = {
  "PFA 86% ZAK25": "pfa-86",
  "PFA 88% ZAK25": "pfa-88",
  "PFA Others @20KG": "pfa-others-20",
  "PFA Others @25KG": "pfa-others-25",
  "PFA Transisi @25KG": "pfa-transisi",
  "PFA Others @JB500": "pfa-others-jb500",
};
function buildProduct(name) {
  const trimmed = name.trim();
  return { id: ID_OVERRIDES[trimmed] || slugify(trimmed), name: trimmed, category: categorize(trimmed) };
}
const BASE_PRODUCTS = RAW_PRODUCTS.map(buildProduct);

const MONTHS_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
const MONTHS_ID_FULL = [
  "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
  "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER",
];
function monthLabelFull(key) {
  const [y, m] = key.split("-");
  return `${MONTHS_ID_FULL[parseInt(m, 10) - 1]} ${y}`;
}
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
// Prefix Lot No otomatis dari tanggal: 2 digit tahun + huruf bulan (A=Jan ... L=Des) + 2 digit tanggal
// Contoh: 2023-12-22 -> "23L22"
function getLotPrefix(dateStr) {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-");
  if (!y || !m || !d) return "";
  const yy = y.slice(-2);
  const monthNum = parseInt(m, 10);
  if (!monthNum || monthNum < 1 || monthNum > 12) return "";
  const monthLetter = String.fromCharCode(64 + monthNum); // 1->A, 2->B ... 12->L
  return `${yy}${monthLetter}${d}`;
}
function numFmt(n) {
  return new Intl.NumberFormat("id-ID").format(n || 0);
}

function getPackagingUnit(name) {
  const trimmed = name.trim();
  if (KNOWN_UNIT_WEIGHTS[trimmed] != null) return `ZAK ${KNOWN_UNIT_WEIGHTS[trimmed]}`;
  const upper = trimmed.toUpperCase();
  let m = upper.match(/JB\s?(\d+)/);
  if (m) return `JB ${m[1]}`;
  m = upper.match(/ZAK\s?(\d+)/);
  if (m) return `ZAK ${m[1]}`;
  m = upper.match(/@\s?(\d+)/);
  if (m) return `ZAK ${m[1]}`;
  return "Zak";
}
// Produk yang beratnya nggak tercantum eksplisit di nama, tapi berat per satuannya sudah diketahui
const KNOWN_UNIT_WEIGHTS = {};

function getUnitWeight(name) {
  const trimmed = name.trim();
  if (KNOWN_UNIT_WEIGHTS[trimmed] != null) return KNOWN_UNIT_WEIGHTS[trimmed];
  const upper = trimmed.toUpperCase();
  let m = upper.match(/JB\s?(\d+)/);
  if (m) return Number(m[1]);
  m = upper.match(/ZAK\s?(\d+)/);
  if (m) return Number(m[1]);
  m = upper.match(/@\s?(\d+)/);
  if (m) return Number(m[1]);
  return null;
}

// Nomor referensi data contoh dari awal pengembangan prototype — otomatis dibersihkan kalau masih ketemu
const LEGACY_DEMO_REFS = ["EXT25-00207", "PA25-08160"];

// ---------- Helper untuk fitur Resume Bulanan (semua produk sekaligus) ----------
function lastDayOfMonthStr(year, monthIndex) {
  const days = new Date(year, monthIndex + 1, 0).getDate();
  return `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(days).padStart(2, "0")}`;
}
function buildResumeSheetHeader() {
  return ["Produk", "IN", "IN (Kg)", "OUT", "OUT (Kg)", "Sisa (Satuan)", "Total (Kg)"];
}
function buildResumeSheetRows(monthLabelText, rows) {
  const out = [[`Resume Bulanan - ${monthLabelText}`], [], buildResumeSheetHeader()];
  for (const r of rows) {
    out.push([r.name, r.inZak, r.inKg, r.outZak, r.outKg, r.sisaZak, r.sisaKg]);
  }
  if (rows.length === 0) {
    out.push(["(Tidak ada produk)"]);
  } else {
    const t = rows.reduce(
      (acc, r) => ({
        inZak: acc.inZak + (Number(r.inZak) || 0),
        inKg: acc.inKg + (Number(r.inKg) || 0),
        outZak: acc.outZak + (Number(r.outZak) || 0),
        outKg: acc.outKg + (Number(r.outKg) || 0),
        sisaZak: acc.sisaZak + (Number(r.sisaZak) || 0),
        sisaKg: acc.sisaKg + (Number(r.sisaKg) || 0),
      }),
      { inZak: 0, inKg: 0, outZak: 0, outKg: 0, sisaZak: 0, sisaKg: 0 }
    );
    out.push(["TOTAL", t.inZak, t.inKg, t.outZak, t.outKg, t.sisaZak, t.sisaKg]);
  }
  return out;
}

const EMPTY_FORM = { date: todayISO(), ref: "", type: "in", zak: "", kg: "", lot: "", palletEkspor: "", palletLokal: "", lokasi: "", ket: "" };

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
  const unitBase = unit.replace(/\s*\d+$/, "").trim(); // "ZAK 25" -> "ZAK", "JB 1000" -> "JB"
  return ["Tanggal", "Referensi", `IN (${unitBase})`, "Masuk (Kg)", `OUT (${unitBase})`, "Keluar (Kg)", `Sisa (${unitBase})`, "Total (Kg)", "Lokasi", "Lot No", "Pallet Kayu (Eksport)", "Pallet Kayu (Lokal)", "Keterangan"];
}

function buildSheetRows(productName, balancedTxns, unit) {
  const rows = [[productName], [], buildSheetHeader(unit)];

  // Sisipkan baris "SUB TOTAL <BULAN> <TAHUN>" setiap kali bulan transaksi berganti,
  // persis seperti kartu stok manual (lihat contoh) — baris ini akan diwarnai hijau di Excel.
  let currentMonthKey = null;
  let monthAgg = null;

  function pushMonthSubtotal() {
    if (!monthAgg) return;
    rows.push([
      `SUB TOTAL ${monthLabelFull(currentMonthKey)}`,
      "",
      monthAgg.inZak, monthAgg.inKg,
      monthAgg.outZak, monthAgg.outKg,
      monthAgg.lastSisaZak, monthAgg.lastSisaKg,
      "", "", "", "", "",
    ]);
  }

  for (const t of balancedTxns) {
    const mk = monthKey(t.date);
    if (mk !== currentMonthKey) {
      pushMonthSubtotal();
      currentMonthKey = mk;
      monthAgg = { inZak: 0, inKg: 0, outZak: 0, outKg: 0, lastSisaZak: 0, lastSisaKg: 0 };
    }

    rows.push([
      fmtDate(t.date), t.ref || "", t.type === "in" ? t.zak : "", t.type === "in" ? t.kg : "",
      t.type === "out" ? t.zak : "", t.type === "out" ? t.kg : "", t.sisaZak, t.sisaKg,
      t.lokasi || "", t.lot || "", t.palletEkspor === "" || t.palletEkspor == null ? "" : t.palletEkspor,
      t.palletLokal === "" || t.palletLokal == null ? "" : t.palletLokal, t.ket || "",
    ]);

    monthAgg.inZak += t.type === "in" ? Number(t.zak) || 0 : 0;
    monthAgg.inKg += t.type === "in" ? Number(t.kg) || 0 : 0;
    monthAgg.outZak += t.type === "out" ? Number(t.zak) || 0 : 0;
    monthAgg.outKg += t.type === "out" ? Number(t.kg) || 0 : 0;
    monthAgg.lastSisaZak = t.sisaZak;
    monthAgg.lastSisaKg = t.sisaKg;
  }
  pushMonthSubtotal(); // subtotal untuk bulan terakhir

  if (balancedTxns.length === 0) {
    rows.push(["(Belum ada transaksi)"]);
  }
  return rows;
}

// Ganti username/password di bawah ini dengan yang rahasia, hanya dikasih tau ke staff & kepala bagian yang bersangkutan
const ACCOUNTS = [
  { username: "margaretta", password: "staff2026", name: "Margaretta (Staff Gudang)" },
  { username: "kepalabagian", password: "kabag2026", name: "Kepala Bagian" },
];

export default function KartuStokPFA() {
  const [role, setRole] = useState(null); // null | 'editor' | 'viewer'
  const [loginStep, setLoginStep] = useState("choose"); // 'choose' | 'password'
  const [userInput, setUserInput] = useState("");
  const [pwInput, setPwInput] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [loggedInName, setLoggedInName] = useState("");
  const [customProductNames, setCustomProductNames] = useState([]);
  const PRODUCTS = useMemo(() => {
    const merged = [...BASE_PRODUCTS];
    const seen = new Set(merged.map((p) => p.id));
    for (const name of customProductNames) {
      const p = buildProduct(name);
      if (!seen.has(p.id)) { seen.add(p.id); merged.push(p); }
    }
    return merged;
  }, [customProductNames]);
  const [selectedId, setSelectedId] = useState(PRODUCTS[0].id);
  const [search, setSearch] = useState("");
  const [txnsByProduct, setTxnsByProduct] = useState({});
  const [loadingProduct, setLoadingProduct] = useState(false);
  const [formOpen, setFormOpen] = useState(false);
  const [form, setForm] = useState(EMPTY_FORM);
  const [editingTxnId, setEditingTxnId] = useState(null);
  const [addProductOpen, setAddProductOpen] = useState(false);
  const [newProductName, setNewProductName] = useState("");
  const [addProductError, setAddProductError] = useState("");
  const [addingProduct, setAddingProduct] = useState(false);
  const [confirmDeleteProduct, setConfirmDeleteProduct] = useState(false);
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
  const [resumeOpen, setResumeOpen] = useState(false);
  const [resumeLoading, setResumeLoading] = useState(false);
  const [resumeExporting, setResumeExporting] = useState(false);
  const [resumeYear, setResumeYear] = useState(now.getFullYear());
  const [resumeMonth, setResumeMonth] = useState(now.getMonth());
  const [resumeRows, setResumeRows] = useState([]);
  const [hasSelectedProduct, setHasSelectedProduct] = useState(false);

  const selectedProduct = PRODUCTS.find((p) => p.id === selectedId);
  const [expandedCats, setExpandedCats] = useState(() => new Set([selectedProduct.category]));

  function toggleCat(cat) {
    setExpandedCats((prev) => {
      const next = new Set(prev);
      if (next.has(cat)) next.delete(cat); else next.add(cat);
      return next;
    });
  }

  function selectProduct(p) {
    setSelectedId(p.id);
    setExpandedCats((prev) => new Set(prev).add(p.category));
    setConfirmDeleteProduct(false);
    setHasSelectedProduct(true);
  }

  function openAddProductForm() {
    setNewProductName("");
    setAddProductError("");
    setAddProductOpen(true);
  }

  async function submitNewProduct(e) {
    if (e && e.preventDefault) e.preventDefault();
    const trimmed = newProductName.trim();
    if (!trimmed) {
      setAddProductError("Nama produk wajib diisi.");
      return;
    }
    const newP = buildProduct(trimmed);
    if (PRODUCTS.some((p) => p.id === newP.id)) {
      setAddProductError("Produk dengan nama ini sudah ada.");
      return;
    }
    setAddingProduct(true);
    try {
      const nextNames = [...customProductNames, trimmed];
      setCustomProductNames(nextNames);
      try {
        await setItem("custom-products", JSON.stringify(nextNames));
      } catch (e) {
        showToast("Produk ditambahkan, tapi gagal disimpan permanen. Coba tambahkan lagi kalau hilang.");
      }
      selectProduct(newP);
      setAddProductOpen(false);
      showToast(`Jenis produk "${trimmed}" ditambahkan.`);
    } finally {
      setAddingProduct(false);
    }
  }

  const unit = useMemo(() => getPackagingUnit(selectedProduct.name), [selectedProduct]);
  const unitWeight = useMemo(() => getUnitWeight(selectedProduct.name), [selectedProduct]);
  const isCustomProduct = useMemo(
    () => customProductNames.some((name) => buildProduct(name).id === selectedProduct.id),
    [customProductNames, selectedProduct]
  );

  async function deleteCustomProduct() {
    const deletedId = selectedProduct.id;
    const deletedName = selectedProduct.name;
    const nextNames = customProductNames.filter((name) => buildProduct(name).id !== deletedId);
    const remaining = [...BASE_PRODUCTS, ...nextNames.map(buildProduct)];
    setCustomProductNames(nextNames);
    setSelectedId(remaining[0].id);
    setConfirmDeleteProduct(false);
    try {
      await setItem("custom-products", JSON.stringify(nextNames));
      await deleteItem(`txns:${deletedId}`);
    } catch (e) {
      // penyimpanan mungkin gagal, tapi produk tetap dihapus dari tampilan
    }
    showToast(`Jenis produk "${deletedName}" dihapus.`);
  }

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2200);
  }, []);

  const loadProduct = useCallback(async (productId) => {
    if (txnsByProduct[productId]) return;
    setLoadingProduct(true);
    try {
      const res = await getItem(`txns:${productId}`);
      let parsed = res ? JSON.parse(res.value) : [];
      // Bersihkan otomatis data contoh dari awal pengembangan prototype (bukan data asli dari staff)
      if (LEGACY_DEMO_REFS.length) {
        const isLegacyDemo = parsed.length > 0 && parsed.every((t) => LEGACY_DEMO_REFS.includes(t.ref));
        if (isLegacyDemo) {
          parsed = [];
          try { await setItem(`txns:${productId}`, JSON.stringify([])); } catch (_) {}
        }
      }
      setTxnsByProduct((prev) => ({ ...prev, [productId]: parsed }));
    } catch (e) {
      setTxnsByProduct((prev) => ({ ...prev, [productId]: [] }));
    } finally {
      setLoadingProduct(false);
    }
  }, [txnsByProduct]);

  useEffect(() => { loadProduct(selectedId); }, [selectedId, loadProduct]);

  useEffect(() => {
    (async () => {
      try {
        const res = await getItem("custom-products");
        const parsed = res ? JSON.parse(res.value) : [];
        if (Array.isArray(parsed)) setCustomProductNames(parsed);
      } catch (e) {
        // belum ada produk custom tersimpan, biarkan kosong
      }
    })();
  }, []);

  const persist = useCallback(async (productId, next) => {
    setTxnsByProduct((prev) => ({ ...prev, [productId]: next }));
    try {
      const res = await setItem(`txns:${productId}`, JSON.stringify(next));
      if (!res) showToast("Gagal menyimpan ke penyimpanan permanen.");
    } catch (e) {
      showToast("Gagal menyimpan ke penyimpanan permanen.");
    }
  }, [showToast]);

  const loadResumeData = useCallback(async (year, monthIndex) => {
    setResumeLoading(true);
    try {
      const monthKeyStr = `${year}-${String(monthIndex + 1).padStart(2, "0")}`;
      const endOfMonth = lastDayOfMonthStr(year, monthIndex);

      // Ambil dulu produk mana aja yang datanya belum ada di cache (txnsByProduct),
      // lalu fetch semuanya SEKALIGUS (paralel), bukan satu-satu bergantian.
      const needFetch = PRODUCTS.filter((p) => !txnsByProduct[p.id]);
      const fetchResults = await Promise.all(
        needFetch.map(async (p) => {
          try {
            const res = await getItem(`txns:${p.id}`);
            let parsed = res ? JSON.parse(res.value) : [];
            if (LEGACY_DEMO_REFS.length && parsed.length > 0 && parsed.every((t) => LEGACY_DEMO_REFS.includes(t.ref))) {
              parsed = [];
            }
            return [p.id, parsed];
          } catch (e) {
            return [p.id, []];
          }
        })
      );
      const fetched = Object.fromEntries(fetchResults);

      const rows = [];
      for (const p of PRODUCTS) {
        const txns = txnsByProduct[p.id] || fetched[p.id] || [];
        const balanced = computeBalance(txns);
        let inZak = 0, inKg = 0, outZak = 0, outKg = 0;
        for (const t of balanced) {
          if (monthKey(t.date) === monthKeyStr) {
            if (t.type === "in") { inZak += Number(t.zak) || 0; inKg += Number(t.kg) || 0; }
            else { outZak += Number(t.zak) || 0; outKg += Number(t.kg) || 0; }
          }
        }
        const upToMonth = balanced.filter((t) => t.date <= endOfMonth);
        const last = upToMonth.length ? upToMonth[upToMonth.length - 1] : null;
        rows.push({
          id: p.id,
          name: p.name,
          category: p.category,
          unit: getPackagingUnit(p.name),
          inZak, inKg, outZak, outKg,
          sisaZak: last ? last.sisaZak : 0,
          sisaKg: last ? last.sisaKg : 0,
        });
      }
      if (Object.keys(fetched).length > 0) {
        setTxnsByProduct((prev) => ({ ...prev, ...fetched }));
      }
      setResumeRows(rows);
    } catch (e) {
      showToast("Gagal memuat data resume.");
    } finally {
      setResumeLoading(false);
    }
  }, [PRODUCTS, txnsByProduct, showToast]);

  function openResume() {
    setResumeOpen(true);
    loadResumeData(resumeYear, resumeMonth);
  }
  function resumePrevMonth() {
    let y = resumeYear, m = resumeMonth;
    if (m === 0) { m = 11; y -= 1; } else m -= 1;
    setResumeYear(y); setResumeMonth(m);
    loadResumeData(y, m);
  }
  function resumeNextMonth() {
    let y = resumeYear, m = resumeMonth;
    if (m === 11) { m = 0; y += 1; } else m += 1;
    setResumeYear(y); setResumeMonth(m);
    loadResumeData(y, m);
  }
  function resumeSetMonth(monthIndex) {
    setResumeMonth(monthIndex);
    loadResumeData(resumeYear, monthIndex);
  }
  function resumeSetYear(year) {
    setResumeYear(year);
    loadResumeData(year, resumeMonth);
  }
  const resumeYearOptions = useMemo(() => {
    const years = [];
    for (let y = 2050; y >= 2005; y--) years.push(y);
    if (!years.includes(resumeYear)) years.push(resumeYear);
    return years.sort((a, b) => b - a);
  }, [resumeYear]);
  const resumeMonthLabel = `${MONTHS_ID[resumeMonth]} ${resumeYear}`;
  const resumeTotals = useMemo(() => {
    return resumeRows.reduce(
      (acc, r) => ({
        inZak: acc.inZak + (Number(r.inZak) || 0),
        inKg: acc.inKg + (Number(r.inKg) || 0),
        outZak: acc.outZak + (Number(r.outZak) || 0),
        outKg: acc.outKg + (Number(r.outKg) || 0),
        sisaZak: acc.sisaZak + (Number(r.sisaZak) || 0),
        sisaKg: acc.sisaKg + (Number(r.sisaKg) || 0),
      }),
      { inZak: 0, inKg: 0, outZak: 0, outKg: 0, sisaZak: 0, sisaKg: 0 }
    );
  }, [resumeRows]);

  async function handleResumeExport() {
    setResumeExporting(true);
    try {
      // rows[0] = [judul], rows[1] = baris kosong, rows[2] = header, rows[3+] = data
      const rows = buildResumeSheetRows(resumeMonthLabel, resumeRows);
      await exportExcelResume(
        rows.slice(2),
        `Resume Bulanan - ${resumeMonthLabel}`,
        `Resume-Bulanan-PFA-${resumeYear}-${String(resumeMonth + 1).padStart(2, "0")}.xlsx`
      );
      showToast(`File Excel Resume Bulanan (${resumeMonthLabel}) berhasil dibuat.`);
    } catch (e) {
      showToast("Gagal membuat file Excel resume.");
    } finally {
      setResumeExporting(false);
    }
  }

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
      // rows[0] = [judul], rows[1] = baris kosong, rows[2] = header, rows[3+] = data
      const rows = buildSheetRows(`${selectedProduct.name} - ${periodLabel}`, displayedTxns, unit);
      await exportExcelProfessional(
        rows.slice(2),
        selectedProduct.name.slice(0, 31),
        `Kartu-Stok-${selectedProduct.id}-${filenameSlug}.xlsx`,
        `${selectedProduct.name} - ${periodLabel}`
      );
      showToast(`File Excel (${selectedProduct.name}, ${periodLabel}) berhasil dibuat.`);
      return;
    }

    setExporting(true);
    try {
      // Ambil dulu produk mana aja yang datanya belum ada di cache,
      // lalu fetch SEKALIGUS secara paralel (bukan satu-satu bergantian) biar lebih cepat.
      const needFetch = PRODUCTS.filter((p) => !txnsByProduct[p.id]);
      const fetchResults = await Promise.all(
        needFetch.map(async (p) => {
          try {
            const res = await getItem(`txns:${p.id}`);
            return [p.id, res ? JSON.parse(res.value) : []];
          } catch (e) {
            return [p.id, []];
          }
        })
      );
      const fetched = Object.fromEntries(fetchResults);
      if (Object.keys(fetched).length > 0) {
        setTxnsByProduct((prev) => ({ ...prev, ...fetched }));
      }

      const sheetsData = [];
      const usedNames = new Set();
      let anyData = false;
      for (const p of PRODUCTS) {
        let txns = txnsByProduct[p.id] || fetched[p.id] || [];
        if (txns.length > 0 && txns.every((t) => LEGACY_DEMO_REFS.includes(t.ref))) txns = [];
        const balanced = computeBalance(txns).filter(matchesFilter);
        if (balanced.length) anyData = true;
        const productUnit = getPackagingUnit(p.name);
        const rows = buildSheetRows(p.name, balanced, productUnit);
        let sheetName = p.name.slice(0, 31) || p.id;
        let i = 1;
        while (usedNames.has(sheetName)) { sheetName = `${p.name.slice(0, 27)}_${i++}`; }
        usedNames.add(sheetName);
        sheetsData.push({
          sheetName,
          rows: rows.slice(2),
          titleText: `${p.name} - ${periodLabel}`,
        });
      }
      if (!anyData) {
        showToast(`Tidak ada transaksi pada periode ${periodLabel} di semua produk.`);
        return;
      }
      await exportExcelAllProducts(sheetsData, `Kartu-Stok-Semua-Produk-${filenameSlug}.xlsx`);
      showToast(`File Excel (Semua Produk, ${periodLabel}) berhasil dibuat.`);
    } catch (e) {
      showToast("Gagal membuat file Excel.");
    } finally {
      setExporting(false);
    }
  }, [exportScope, displayedTxns, periodLabel, selectedProduct, unit, filenameSlug, txnsByProduct, matchesFilter, showToast, PRODUCTS]);

  const monthlyRecap = useMemo(() => {
    const map = new Map();
    for (const t of sortedWithBalance) {
      const key = monthKey(t.date);
      const prev = map.get(key) || { inZak: 0, inKg: 0, outZak: 0, outKg: 0, sisaZak: 0, sisaKg: 0 };
      if (t.type === "in") { prev.inZak += t.zak; prev.inKg += t.kg; }
      else { prev.outZak += t.zak; prev.outKg += t.kg; }
      prev.sisaZak = t.sisaZak;
      prev.sisaKg = t.sisaKg;
      map.set(key, prev);
    }
    return Array.from(map.entries()).sort(([a], [b]) => (a < b ? -1 : 1));
  }, [sortedWithBalance]);

  const currentStock = sortedWithBalance.length
    ? sortedWithBalance[sortedWithBalance.length - 1]
    : { sisaZak: 0, sisaKg: 0 };

  const filteredProducts = PRODUCTS.filter((p) =>
    p.name.toLowerCase().includes(search.toLowerCase())
  );
  const productCategories = Array.from(new Set(PRODUCTS.map((p) => p.category))).sort((a, b) => {
    if (a === "Lainnya") return 1;
    if (b === "Lainnya") return -1;
    return parseInt(a, 10) - parseInt(b, 10);
  });
  const grouped = productCategories.map((cat) => ({
    cat,
    items: filteredProducts.filter((p) => p.category === cat),
  })).filter((g) => g.items.length);

  function openAddForm() {
    const today = todayISO();
    setForm({ ...EMPTY_FORM, date: today, lot: getLotPrefix(today) });
    setEditingTxnId(null);
    setFormError("");
    setFormOpen(true);
  }

  function openEditForm(t) {
    setForm({
      date: t.date,
      ref: t.ref || "",
      type: t.type,
      zak: t.zak ? String(t.zak) : "",
      kg: t.kg ? String(t.kg) : "",
      lot: t.lot || "",
      palletEkspor: t.palletEkspor === "" || t.palletEkspor == null ? "" : String(t.palletEkspor),
      palletLokal: t.palletLokal === "" || t.palletLokal == null ? "" : String(t.palletLokal),
      lokasi: t.lokasi || "",
      ket: t.ket || "",
    });
    setEditingTxnId(t.id);
    setFormError("");
    setFormOpen(true);
  }

  function handleEditorLogin(e) {
    if (e && e.preventDefault) e.preventDefault();
    const found = ACCOUNTS.find(
      (a) => a.username.toLowerCase() === userInput.trim().toLowerCase() && a.password === pwInput.trim()
    );
    if (found) {
      setRole("editor");
      setLoggedInName(found.name);
      setLoginError("");
      setUserInput("");
      setPwInput("");
    } else {
      setLoginError("Username atau password salah. Coba lagi.");
    }
  }

  function loginAsViewer() {
    setRole("viewer");
    setLoggedInName("Manager / Lainnya");
  }

  function logout() {
    setRole(null);
    setLoginStep("choose");
    setUserInput("");
    setPwInput("");
    setLoginError("");
    setLoggedInName("");
    setHasSelectedProduct(false);
  }

  async function submitForm(e) {
    if (e && e.preventDefault) e.preventDefault();
    try {
      const zakNum = Number(form.zak) || 0;
      const kgNum = Number(form.kg) || 0;
      if (!form.date) {
        setFormError("Tanggal wajib diisi.");
        return;
      }
      if (zakNum <= 0 && kgNum <= 0) {
        setFormError(`Jumlah wajib diisi dengan nilai lebih dari 0.`);
        return;
      }
      setFormError("");
      const entry = {
        id: editingTxnId || `t-${Date.now()}`,
        date: form.date,
        ref: form.ref.trim(),
        type: form.type,
        zak: zakNum,
        kg: kgNum,
        lot: form.lot.trim(),
        palletEkspor: form.palletEkspor === "" ? "" : Number(form.palletEkspor),
        palletLokal: form.palletLokal === "" ? "" : Number(form.palletLokal),
        lokasi: form.lokasi.trim(),
        ket: form.ket.trim(),
      };
      const next = editingTxnId
        ? rawTxns.map((t) => (t.id === editingTxnId ? entry : t))
        : [...rawTxns, entry];
      await persist(selectedId, next);
      setFormOpen(false);
      setEditingTxnId(null);
      showToast(editingTxnId ? "Transaksi diperbarui." : "Transaksi tersimpan.");
    } catch (err) {
      setFormError(`Terjadi error teknis: ${err && err.message ? err.message : String(err)}`);
    }
  }

  async function deleteTxn(id) {
    const next = rawTxns.filter((t) => t.id !== id);
    await persist(selectedId, next);
    setConfirmDeleteId(null);
    showToast("Transaksi dihapus.");
  }

  return (
    <div className="ks-root ks-login-theme">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

        .ks-root {
          position: relative;
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
          display: flex;
          flex-direction: column;
          min-height: 100vh;
          width: 100%;
          max-width: 100%;
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
        .ks-search-box {
          display: flex; align-items: center; gap: 8px; border: 1.5px solid var(--accent);
          border-radius: 999px; padding: 0 12px; background: var(--panel-alt);
        }
        .ks-search-icon { display: flex; align-items: center; color: var(--accent); flex-shrink: 0; }
        .ks-search-box input {
          flex: 1; min-width: 0; background: transparent; border: none; color: var(--text);
          padding: 9px 0; font-size: 14px; font-family: 'Inter', sans-serif;
        }
        .ks-search-box input:focus { outline: none; }
        .ks-add-product-wrap { padding: 0 12px 10px; border-bottom: 1px solid var(--border); }
        .ks-add-product-btn { width: 100%; font-size: 12.5px; padding: 7px; }
        .ks-search-clear {
          background: transparent; color: var(--muted); border: none; padding: 4px; cursor: pointer;
          font-size: 13px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;
        }
        .ks-search-clear:hover { color: var(--text); }
        .ks-list { overflow-y: auto; flex: 1; padding: 6px 0; }
        .ks-cat-label { width: 100%; display: flex; align-items: center; justify-content: space-between;
          background: transparent; border: none; cursor: pointer; padding: 10px 16px 8px; font-size: 11px;
          color: var(--muted); letter-spacing: 0.08em; text-transform: uppercase; font-weight: 600; font-family: 'Inter', sans-serif; }
        .ks-cat-label:hover { color: var(--text); }
        .ks-cat-chevron { font-size: 10px; }
        .ks-item {
          width: 100%; text-align: left; background: transparent; border: none; color: var(--text);
          padding: 8px 16px; font-size: 13.5px; cursor: pointer; border-left: 3px solid transparent;
          font-family: 'Inter', sans-serif;
        }
        .ks-item:hover { background: var(--panel-alt); }
        .ks-item.active { background: var(--accent-dim); border-left-color: var(--accent); font-weight: 600; }
        .ks-main { flex: 1; padding: 20px 24px; overflow-y: auto; min-width: 0; }
        .ks-welcome {
          height: 100%; min-height: 340px; display: flex; flex-direction: column; align-items: center;
          justify-content: center; text-align: center; color: var(--muted); gap: 6px;
        }
        .ks-welcome-icon {
          width: 62px; height: 62px; border-radius: 16px; display: flex; align-items: center; justify-content: center;
          background: var(--accent-dim); color: var(--accent); margin-bottom: 10px;
        }
        .ks-welcome-title { font-family: 'Oswald', sans-serif; font-size: 19px; font-weight: 600; color: var(--text); letter-spacing: 0.01em; }
        .ks-welcome-sub { font-size: 13.5px; max-width: 320px; }
        .ks-headrow { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 16px; }
        .ks-product-name { font-family: 'Oswald', sans-serif; font-size: 24px; font-weight: 600; letter-spacing: 0.01em; }
        .ks-btn {
          background: var(--accent); color: #17140d; border: none; border-radius: 6px; padding: 9px 16px;
          font-size: 13.5px; font-weight: 600; cursor: pointer; font-family: 'Inter', sans-serif;
        }
        .ks-btn:hover { filter: brightness(1.08); }
        .ks-btn.ghost { background: transparent; color: var(--text); border: 1px solid var(--border); }
        .ks-btn.small { padding: 6px 10px; font-size: 12px; }
        .ks-btn-danger { background: var(--negative); color: #fff; }
        .ks-clearall-confirm { display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--muted);
          background: var(--negative-dim); border: 1px solid var(--negative); border-radius: 6px; padding: 6px 10px; }
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
        .ks-edit-btn { color: var(--accent); }
        .ks-modal-overlay { position: absolute; top: 0; left: 0; right: 0; min-height: 100%; background: rgba(0,0,0,0.55); display: flex;
          align-items: flex-start; justify-content: center; padding-top: 50px; z-index: 50; }
        .ks-modal { background: var(--panel); border: 1px solid var(--border); border-radius: 10px; width: 460px;
          max-width: 92vw; padding: 20px; max-height: 88vh; overflow-y: auto; }
        .ks-modal h3 { font-family: 'Oswald', sans-serif; margin: 0 0 14px; font-size: 17px; letter-spacing: 0.02em; text-transform: uppercase; }
        .ks-field { margin-bottom: 11px; }
        .ks-field label { display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }
        .ks-input-auto { opacity: 0.75; cursor: not-allowed; }
        .ks-hint { font-size: 11px; color: var(--muted); margin-top: 4px; }
        .ks-kg-display { background: var(--panel-alt); border: 1px solid var(--border); border-radius: 6px;
          padding: 8px 10px; }
        .ks-kg-value { display: block; font-family: 'JetBrains Mono', monospace; font-size: 15px;
          color: var(--accent); font-weight: 600; }
        .ks-kg-display .ks-hint { margin-top: 2px; }
        .ks-field input, .ks-field select, .ks-field textarea {
          width: 100%; background: var(--panel-alt); border: 1px solid var(--border); color: var(--text);
          padding: 8px 10px; border-radius: 6px; font-size: 13px; font-family: 'Inter', sans-serif;
        }
        .ks-row2 { display: flex; gap: 10px; }
        .ks-row2 > div { flex: 1; }
        .ks-modal-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
        .ks-form-error { background: var(--negative-dim); color: var(--negative); border: 1px solid var(--negative);
          border-radius: 6px; padding: 9px 12px; font-size: 12.5px; margin-top: 6px; }
        .ks-toast { position: absolute; bottom: 18px; left: 50%; transform: translateX(-50%);
          background: var(--panel-alt); border: 1px solid var(--accent); color: var(--text);
          padding: 9px 16px; border-radius: 8px; font-size: 13px; z-index: 45; pointer-events: none; }
        .ks-monthly { margin-top: 18px; border: 1px solid var(--border); border-radius: 8px; }
        .ks-monthly-head { padding: 11px 14px; cursor: pointer; display: flex; justify-content: space-between;
          align-items: center; font-size: 13px; font-weight: 600; background: var(--panel); }
        .ks-monthly-body { padding: 0 14px 12px; }
        .ks-monthly-row-full { padding: 10px 0; border-bottom: 1px solid var(--border); }
        .ks-monthly-row-full:last-child { border-bottom: none; }
        .ks-monthly-month { font-family: 'Oswald', sans-serif; font-size: 13px; letter-spacing: 0.02em;
          text-transform: uppercase; margin-bottom: 6px; color: var(--accent); }
        .ks-monthly-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 16px; }
        .ks-monthly-grid > div { display: flex; justify-content: space-between; font-size: 12px; gap: 8px; }
        .ks-monthly-label { color: var(--muted); }
        .ks-loading { padding: 30px; text-align: center; color: var(--muted); font-size: 13px; }
        @media (max-width: 640px) {
          .ks-body { flex-direction: column; }
          .ks-sidebar { width: 100%; max-height: 220px; }
        }
        .ks-login-wrap { flex: 1; min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 30px; }
        .ks-login-card { background: var(--panel); border: 1px solid var(--border); border-radius: 12px;
          padding: 28px; width: 360px; max-width: 100%; box-shadow: 0 16px 36px rgba(0,0,0,0.4); }
        .ks-login-title { font-family: 'Oswald', sans-serif; font-size: 19px; font-weight: 600;
          letter-spacing: 0.02em; text-transform: uppercase; text-align: center; }
        .ks-login-sub { color: var(--muted); font-size: 12.5px; text-align: center; margin-top: 4px; margin-bottom: 22px; }
        .ks-login-choices { display: flex; flex-direction: column; gap: 10px; }
        .ks-login-choice { background: var(--panel-alt); border: 1px solid var(--border); border-radius: 8px;
          padding: 14px 16px; text-align: left; cursor: pointer; color: var(--text); font-family: 'Inter', sans-serif; }
        .ks-login-choice:hover { border-color: var(--accent); background: var(--accent-dim); }
        .ks-login-choice-title { font-size: 14px; font-weight: 600; }
        .ks-login-choice-sub { font-size: 12px; color: var(--muted); margin-top: 3px; }
        .ks-login-form { display: flex; flex-direction: column; gap: 4px; }
        .ks-login-form label { font-size: 12px; color: var(--muted); margin-top: 10px; }
        .ks-login-form input {
          width: 100%; background: var(--panel-alt); border: 1px solid var(--border); color: var(--text);
          padding: 9px 11px; border-radius: 6px; font-size: 13px; font-family: 'Inter', sans-serif; margin-top: 4px;
        }
        .ks-login-form-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; }
        .ks-pw-wrap { position: relative; margin-top: 4px; }
        .ks-pw-wrap input { margin-top: 0; padding-right: 38px; }
        .ks-pw-toggle { position: absolute; right: 6px; top: 50%; transform: translateY(-50%);
          background: transparent; border: none; cursor: pointer; padding: 4px 6px; line-height: 1;
          color: var(--muted); display: flex; align-items: center; }
        .ks-pw-toggle:hover { color: var(--text); }
        .ks-role-badge { font-family: 'JetBrains Mono', monospace; font-size: 12px; background: var(--panel-alt);
          border: 1px solid var(--border); border-radius: 6px; padding: 6px 10px; }
        .ks-role-badge.editor { color: var(--accent); border-color: var(--accent); }
        .ks-role-badge.viewer { color: var(--muted); }
        .ks-modal-wide { width: 780px; }
        .ks-resume-nav { display: flex; align-items: center; justify-content: center; gap: 14px; margin-bottom: 14px; }
        .ks-resume-month { font-family: 'Oswald', sans-serif; font-size: 14px; letter-spacing: 0.02em;
          text-transform: uppercase; min-width: 140px; text-align: center; }
        .ks-resume-table-wrap { border: 1px solid var(--border); border-radius: 8px; overflow: auto; max-height: 55vh; }
        .ks-resume-total-row td {
          background: var(--panel-alt); font-weight: 700; border-top: 2px solid var(--border);
          position: sticky; bottom: 0; padding: 12px 10px; vertical-align: top;
        }
        .ks-resume-select {
          background: var(--panel-alt); color: var(--text); border: 1px solid var(--border);
          border-radius: 6px; padding: 7px 10px; font-size: 13px; font-family: 'Inter', sans-serif; cursor: pointer;
        }

        /* ---------- Tema halaman login (sesuai desain Figma PT. Dover Chemical) ---------- */
        .ks-login-theme {
          --bg: #eef1f7;
          --panel: #ffffff;
          --panel-alt: #f4f6fa;
          --border: #d7dbe3;
          --text: #101828;
          --muted: #6b7280;
          --accent: #2563eb;
          --accent-dim: rgba(37,99,235,0.10);
          --positive: #16a34a;
          --positive-dim: rgba(22,163,74,0.12);
          --negative: #dc2626;
          --negative-dim: rgba(220,38,38,0.12);
        }
        .ks-login-theme .ks-login-wrap {
          padding: 40px 20px;
          background-image: linear-gradient(rgba(0,0,0,0.32), rgba(0,0,0,0.32)), url('/bg-plant.jpg');
          background-size: cover;
          background-position: center;
        }
        .ks-login-theme .ks-login-card {
          width: 420px;
          border-radius: 18px;
          padding: 34px 38px 30px;
          border: none;
          text-align: center;
          box-shadow: 0 30px 60px -20px rgba(0,0,0,0.45), 0 4px 14px rgba(0,0,0,0.15);
        }
        .ks-login-logo-img {
          width: 130px;
          height: auto;
          display: block;
          margin: 0 auto 8px;
        }
        .ks-login-company {
          font-family: 'Inter', sans-serif;
          font-size: 15px;
          font-weight: 600;
          letter-spacing: 0.03em;
          color: #4b5563;
          text-transform: uppercase;
          margin-bottom: 18px;
        }
        .ks-login-theme .ks-login-title {
          text-align: center; font-size: 15px; letter-spacing: 0.04em; color: #111827;
          text-transform: uppercase; font-weight: 600; font-family: 'Inter', sans-serif;
        }
        .ks-login-system-name {
          font-family: 'Inter', sans-serif; font-size: 15px; font-weight: 600;
          color: #111827; margin-top: 2px;
        }
        .ks-login-theme .ks-login-sub { text-align: center; margin-top: 8px; margin-bottom: 26px; color: #9ca3af; }
        .ks-login-theme .ks-login-choice {
          border-radius: 12px; padding: 13px 16px; transition: all .15s ease;
          display: flex; align-items: center; justify-content: center; gap: 10px;
          font-size: 14.5px; font-weight: 500; border: 1px solid var(--border);
        }
        .ks-login-theme .ks-login-choice:hover {
          border-color: var(--accent); background: var(--accent); color: #fff;
        }
        .ks-login-theme .ks-login-choice-title { color: inherit; font-weight: 500; font-size: 14.5px; }
        .ks-login-theme .ks-login-form input {
          border-radius: 12px; border: 1px solid var(--border); background: #fff;
          padding: 12px 14px; color: #101828; transition: all .15s ease; font-size: 13.5px;
        }
        .ks-login-theme .ks-login-form input:focus {
          outline: none; border-color: var(--accent);
          box-shadow: 0 0 0 3px rgba(37,99,235,0.15);
        }
        .ks-login-theme .ks-login-form label { text-align: left; }
        .ks-login-theme .ks-btn { border-radius: 12px; }
        .ks-login-theme .ks-btn:not(.ghost) {
          background: var(--accent); border: 1px solid var(--accent); color: #ffffff; font-weight: 500;
        }
        .ks-login-theme .ks-btn:not(.ghost):hover { filter: brightness(1.1); }
        .ks-login-theme .ks-btn.ghost { background: #fff; color: #111827; border: 1px solid var(--border); font-weight: 500; }
        .ks-login-theme .ks-btn.ghost:hover { background: var(--accent); border-color: var(--accent); color: #fff; }
        .ks-login-theme .ks-btn.ks-add-product-btn { border-radius: 6px; }
        .ks-login-theme .ks-login-form-actions { gap: 12px; }
        .ks-login-theme .ks-login-form-actions .ks-btn { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; }
        .ks-login-theme .ks-pw-toggle { color: var(--muted); }
        .ks-login-theme .ks-form-error {
          color: var(--negative); font-size: 12px; text-align: left; margin-top: 8px;
          display: flex; align-items: center; gap: 5px;
        }
      `}</style>

      {!role ? (
        <div className="ks-login-wrap">
          <div className="ks-login-card">
            <img src="/logo-dover-chemical.png" alt="PT. Dover Chemical" className="ks-login-logo-img" />
            <div className="ks-login-company">PT. Dover Chemical</div>
            <div className="ks-login-title">Stock Card</div>
            <div className="ks-login-system-name">Paraformaldehyde Inventory System</div>
            <div className="ks-login-sub">Sign In To Continue</div>

            {loginStep === "choose" && (
              <div className="ks-login-choices">
                <button className="ks-login-choice" onClick={() => { setLoginStep("password"); setLoginError(""); }}>
                  <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z" />
                    <path d="M15 5 19 9" />
                  </svg>
                  <span className="ks-login-choice-title">Sign in as Editor</span>
                </button>
                <button className="ks-login-choice" onClick={loginAsViewer}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M1 12s4-7.5 11-7.5S23 12 23 12s-4 7.5-11 7.5S1 12 1 12Z" />
                    <circle cx="12" cy="12" r="3" />
                  </svg>
                  <span className="ks-login-choice-title">Sign in as Viewer</span>
                </button>
              </div>
            )}

            {loginStep === "password" && (
              <div className="ks-login-form">
                <input
                  type="text"
                  value={userInput}
                  onChange={(e) => setUserInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter") handleEditorLogin(e); }}
                  placeholder="Username"
                  autoFocus
                  style={{ marginTop: 0 }}
                />
                <div className="ks-pw-wrap" style={{ marginTop: 10 }}>
                  <input
                    type={showPw ? "text" : "password"}
                    value={pwInput}
                    onChange={(e) => setPwInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") handleEditorLogin(e); }}
                    placeholder="Password"
                  />
                  <button type="button" className="ks-pw-toggle" onClick={() => setShowPw((v) => !v)} tabIndex={-1}>
                    {showPw ? (
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7Z" />
                        <circle cx="12" cy="12" r="3" />
                        <line x1="3" y1="3" x2="21" y2="21" />
                      </svg>
                    ) : (
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7Z" />
                        <circle cx="12" cy="12" r="3" />
                      </svg>
                    )}
                  </button>
                </div>
                {loginError && (
                  <div className="ks-form-error">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z" />
                      <line x1="12" y1="9" x2="12" y2="13" />
                      <line x1="12" y1="17" x2="12.01" y2="17" />
                    </svg>
                    {loginError}
                  </div>
                )}
                <div className="ks-login-form-actions">
                  <button type="button" className="ks-btn ghost" onClick={() => { setLoginStep("choose"); setLoginError(""); setUserInput(""); setPwInput(""); }}>Kembali</button>
                  <button type="button" className="ks-btn" onClick={handleEditorLogin}>
                    Masuk
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M5 12h14" />
                      <path d="m13 6 6 6-6 6" />
                    </svg>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      ) : (
      <>
      <div className="ks-topbar">
        <div className="ks-title">
          Stock Card — Paraformaldehyde
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span className={`ks-role-badge ${role}`}>{role === "editor" ? "Editor" : "Viewer"}</span>
          <button className="ks-btn ghost" onClick={openResume}>📊 Resume Bulanan</button>
          <button className="ks-btn ghost" onClick={logout}>Keluar</button>
        </div>
      </div>

      <div className="ks-body">
        <div className="ks-sidebar">
          <div className="ks-search">
            <div className="ks-search-box">
              <span className="ks-search-icon">
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="11" cy="11" r="7" />
                  <path d="m21 21-4.3-4.3" />
                </svg>
              </span>
              <input placeholder="Cari produk..." value={search} onChange={(e) => setSearch(e.target.value)} />
              {search.trim() !== "" && (
                <button
                  type="button"
                  className="ks-search-clear"
                  onClick={() => setSearch("")}
                  title="Hapus pencarian"
                >
                  ✕
                </button>
              )}
            </div>
          </div>
          {role === "editor" && (
            <div className="ks-add-product-wrap">
              <button className="ks-btn ghost ks-add-product-btn" onClick={openAddProductForm}>+ Tambah Jenis Produk</button>
            </div>
          )}
          <div className="ks-list">
            {search.trim() !== "" ? (
              <>
                {filteredProducts.length === 0 && (
                  <div style={{ padding: 16, fontSize: 13, color: "var(--muted)" }}>Tidak ditemukan.</div>
                )}
                {filteredProducts.map((p) => (
                  <button
                    key={p.id}
                    className={`ks-item ${p.id === selectedId && hasSelectedProduct ? "active" : ""}`}
                    onClick={() => { selectProduct(p); setSearch(""); }}
                  >
                    {p.name}
                  </button>
                ))}
              </>
            ) : (
              grouped.map((g) => {
                if (g.items.length === 1) {
                  const p = g.items[0];
                  return (
                    <button
                      key={p.id}
                      className={`ks-item ks-item-single ${p.id === selectedId && hasSelectedProduct ? "active" : ""}`}
                      onClick={() => selectProduct(p)}
                    >
                      {p.name}
                    </button>
                  );
                }
                const isOpen = expandedCats.has(g.cat);
                return (
                  <div key={g.cat}>
                    <button className="ks-cat-label" onClick={() => toggleCat(g.cat)}>
                      <span>PFA {g.cat}</span>
                      <span className="ks-cat-chevron">{isOpen ? "▾" : "▸"}</span>
                    </button>
                    {isOpen && g.items.map((p) => (
                      <button
                        key={p.id}
                        className={`ks-item ${p.id === selectedId && hasSelectedProduct ? "active" : ""}`}
                        onClick={() => selectProduct(p)}
                      >
                        {p.name}
                      </button>
                    ))}
                  </div>
                );
              })
            )}
          </div>
        </div>

        <div className="ks-main">
          {!hasSelectedProduct ? (
            <div className="ks-welcome">
              <div className="ks-welcome-icon">
                <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M9 2h6" />
                  <path d="M10 2v6.34a2 2 0 0 1-.4 1.2L5.2 15.8A3 3 0 0 0 7.6 21h8.8a3 3 0 0 0 2.4-5.2l-4.4-6.26a2 2 0 0 1-.4-1.2V2" />
                  <path d="M6.5 15h11" />
                </svg>
              </div>
              <div className="ks-welcome-title">Pilih Jenis Produk</div>
              <div className="ks-welcome-sub">Klik salah satu jenis produk di sidebar kiri untuk melihat kartu stoknya.</div>
            </div>
          ) : (
          <>
          <div className="ks-headrow">
            <div>
              <div className="ks-product-name">{selectedProduct.name}</div>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              {role === "editor" && isCustomProduct && (
                confirmDeleteProduct ? (
                  <span className="ks-clearall-confirm">
                    Hapus jenis produk ini beserta semua transaksinya?
                    <button className="ks-btn ghost small" onClick={() => setConfirmDeleteProduct(false)}>Batal</button>
                    <button className="ks-btn small ks-btn-danger" onClick={deleteCustomProduct}>Ya, Hapus</button>
                  </span>
                ) : (
                  <button className="ks-btn ghost" onClick={() => setConfirmDeleteProduct(true)}>Hapus Jenis Produk</button>
                )
              )}
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
          </div>

          {loadingProduct ? (
            <div className="ks-loading">Memuat data...</div>
          ) : (
            <>
              <div className="ks-table-wrap">
                <table className="ks-table">
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Reff</th>
                      <th>In</th>
                      <th>Out</th>
                      <th>Sisa</th>
                      <th>Total</th>
                      <th>Lokasi</th>
                      <th>Lot No.</th>
                      <th>Pallet Kayu<br/>(Eksport)</th>
                      <th>Pallet Kayu<br/>(Lokal)</th>
                      <th>Keterangan</th>
                      {role === "editor" && <th></th>}
                    </tr>
                  </thead>
                  <tbody>
                    {displayedTxns.length === 0 && (
                      <tr>
                        <td colSpan={role === "editor" ? 12 : 11}>
                          <div className="ks-empty">
                            {sortedWithBalance.length === 0
                              ? "Belum ada transaksi tercatat untuk produk ini."
                              : `Tidak ada transaksi pada periode ${periodLabel}.`}
                          </div>
                        </td>
                      </tr>
                    )}
                    {displayedTxns.map((t) => (
                      <tr key={t.id}>
                        <td className="ks-mono">{fmtDate(t.date)}</td>
                        <td className="ks-mono">{t.ref || "-"}</td>
                        <td className="ks-mono ks-in">{t.type === "in" ? `${numFmt(t.zak)} / ${numFmt(t.kg)} kg` : "-"}</td>
                        <td className="ks-mono ks-out">{t.type === "out" ? `${numFmt(t.zak)} / ${numFmt(t.kg)} kg` : "-"}</td>
                        <td><span className="ks-sisa-badge">{numFmt(t.sisaZak)}</span></td>
                        <td className="ks-mono">{numFmt(t.sisaKg)} kg</td>
                        <td>{t.lokasi || "-"}</td>
                        <td className="ks-mono">{t.lot || "-"}</td>
                        <td className="ks-mono">{t.palletEkspor === "" || t.palletEkspor == null ? "-" : numFmt(t.palletEkspor)}</td>
                        <td className="ks-mono">{t.palletLokal === "" || t.palletLokal == null ? "-" : numFmt(t.palletLokal)}</td>
                        <td>{t.ket || "-"}</td>
                        {role === "editor" && (
                          <td>
                            {confirmDeleteId === t.id ? (
                              <div style={{ display: "flex", gap: 6 }}>
                                <button className="ks-del-btn" onClick={() => deleteTxn(t.id)}>Yakin?</button>
                                <button className="ks-del-btn" style={{ color: "var(--muted)" }} onClick={() => setConfirmDeleteId(null)}>Batal</button>
                              </div>
                            ) : (
                              <div style={{ display: "flex", gap: 10 }}>
                                <button className="ks-del-btn ks-edit-btn" onClick={() => openEditForm(t)}>Ubah</button>
                                <button className="ks-del-btn" onClick={() => setConfirmDeleteId(t.id)}>Hapus</button>
                              </div>
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
                  <span>Rekap Bulanan</span>
                  <span className="ks-cat-chevron">{monthlyOpen ? "▾" : "▸"}</span>
                </div>
                {monthlyOpen && (
                  <div className="ks-monthly-body">
                    {monthlyRecap.length === 0 && (
                      <div style={{ padding: "8px 0", color: "var(--muted)", fontSize: 12.5 }}>Belum ada data.</div>
                    )}
                    {monthlyRecap.map(([key, v]) => (
                      <div className="ks-monthly-row-full" key={key}>
                        <div className="ks-monthly-month">{monthLabel(key)}</div>
                        <div className="ks-monthly-grid">
                          <div><span className="ks-monthly-label">IN</span><span className="ks-mono ks-in">{numFmt(v.inZak)} / {numFmt(v.inKg)} kg</span></div>
                          <div><span className="ks-monthly-label">OUT</span><span className="ks-mono ks-out">{numFmt(v.outZak)} / {numFmt(v.outKg)} kg</span></div>
                          <div><span className="ks-monthly-label">SISA</span><span className="ks-mono">{numFmt(v.sisaZak)}</span></div>
                          <div><span className="ks-monthly-label">TOTAL</span><span className="ks-mono">{numFmt(v.sisaKg)} kg</span></div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </>
          )}
          </>
          )}
        </div>
      </div>

      {formOpen && (
        <div className="ks-modal-overlay" onClick={(e) => e.target === e.currentTarget && setFormOpen(false)}>
          <div className="ks-modal">
            <h3>{editingTxnId ? "Ubah Transaksi" : "Tambah Transaksi"}</h3>
            <div>
              <div className="ks-row2">
                <div className="ks-field">
                  <label>DATE</label>
                  <input
                    type="date"
                    value={form.date}
                    onChange={(e) => {
                      const newDate = e.target.value;
                      const oldPrefix = getLotPrefix(form.date);
                      const newPrefix = getLotPrefix(newDate);
                      let newLot = form.lot;
                      if (!form.lot || form.lot === oldPrefix) {
                        newLot = newPrefix;
                      } else if (oldPrefix && form.lot.startsWith(oldPrefix)) {
                        newLot = newPrefix + form.lot.slice(oldPrefix.length);
                      }
                      setForm({ ...form, date: newDate, lot: newLot });
                    }}
                  />
                </div>
                <div className="ks-field">
                  <label>TYPE</label>
                  <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                    <option value="in">IN</option>
                    <option value="out">OUT</option>
                  </select>
                </div>
              </div>
              <div className="ks-field">
                <label>REFF</label>
                <input value={form.ref} onChange={(e) => setForm({ ...form, ref: e.target.value })} />
              </div>
              <div className="ks-row2">
                <div className="ks-field">
                  <label>Jumlah</label>
                  <input
                    type="number"
                    value={form.zak}
                    onChange={(e) => {
                      const val = e.target.value;
                      const zakNum = Number(val) || 0;
                      setForm((f) => ({ ...f, zak: val, kg: unitWeight != null ? String(zakNum * unitWeight) : f.kg }));
                    }}
                    placeholder="0"
                  />
                </div>
                <div className="ks-field">
                  <label>Berat Total (Kg)</label>
                  {unitWeight != null ? (
                    <div className="ks-kg-display">
                      <span className="ks-kg-value">{numFmt((Number(form.zak) || 0) * unitWeight)} kg</span>
                    </div>
                  ) : (
                    <>
                      <input
                        type="number"
                        value={form.kg}
                        onChange={(e) => setForm({ ...form, kg: e.target.value })}
                        placeholder="0"
                      />
                      <div className="ks-hint">Isi manual (satuan produk ini tidak tercantum di nama)</div>
                    </>
                  )}
                </div>
              </div>
              <div className="ks-row2">
                <div className="ks-field">
                  <label>Lot No.</label>
                  <input value={form.lot} onChange={(e) => setForm({ ...form, lot: e.target.value })} />
                </div>
                <div className="ks-field">
                  <label>Pallet Kayu (Eksport)</label>
                  <input type="number" value={form.palletEkspor} onChange={(e) => setForm({ ...form, palletEkspor: e.target.value })} />
                </div>
              </div>
              <div className="ks-field">
                <label>Pallet Kayu (Lokal)</label>
                <input type="number" value={form.palletLokal} onChange={(e) => setForm({ ...form, palletLokal: e.target.value })} />
              </div>
              <div className="ks-field">
                <label>Lokasi</label>
                <input value={form.lokasi} onChange={(e) => setForm({ ...form, lokasi: e.target.value })} />
              </div>
              <div className="ks-field">
                <label>Keterangan</label>
                <textarea rows={2} value={form.ket} onChange={(e) => setForm({ ...form, ket: e.target.value })} />
              </div>
              {formError && <div className="ks-form-error">⚠ {formError}</div>}
              <div className="ks-modal-actions">
                <button type="button" className="ks-btn ghost" onClick={() => { setFormOpen(false); setEditingTxnId(null); }}>Batal</button>
                <button type="button" className="ks-btn" onClick={submitForm}>Simpan</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {addProductOpen && (
        <div className="ks-modal-overlay" onClick={(e) => e.target === e.currentTarget && setAddProductOpen(false)}>
          <div className="ks-modal">
            <h3>Tambah Jenis Produk</h3>
            <div className="ks-field">
              <label>Nama Produk</label>
              <input
                type="text"
                value={newProductName}
                onChange={(e) => setNewProductName(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") submitNewProduct(e); }}
                autoFocus
              />
            </div>
            {addProductError && <div className="ks-form-error">⚠ {addProductError}</div>}
            <div className="ks-modal-actions">
              <button type="button" className="ks-btn ghost" onClick={() => setAddProductOpen(false)}>Batal</button>
              <button type="button" className="ks-btn" onClick={submitNewProduct} disabled={addingProduct}>
                {addingProduct ? "Menyimpan..." : "Tambah"}
              </button>
            </div>
          </div>
        </div>
      )}

      {resumeOpen && (
        <div className="ks-modal-overlay" onClick={(e) => e.target === e.currentTarget && setResumeOpen(false)}>
          <div className="ks-modal ks-modal-wide">
            <h3>Resume Bulanan — Semua Produk</h3>
            <div className="ks-resume-nav">
              <button className="ks-btn ghost small" onClick={resumePrevMonth} title="Bulan sebelumnya">‹</button>
              <select
                className="ks-resume-select"
                value={resumeMonth}
                onChange={(e) => resumeSetMonth(Number(e.target.value))}
              >
                {MONTHS_ID.map((m, idx) => (
                  <option key={m} value={idx}>{m}</option>
                ))}
              </select>
              <select
                className="ks-resume-select"
                value={resumeYear}
                onChange={(e) => resumeSetYear(Number(e.target.value))}
              >
                {resumeYearOptions.map((y) => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
              <button className="ks-btn ghost small" onClick={resumeNextMonth} title="Bulan berikutnya">›</button>
            </div>
            {resumeLoading ? (
              <div className="ks-loading">Memuat data semua produk...</div>
            ) : (
              <>
                <div className="ks-resume-table-wrap">
                  <table className="ks-table">
                    <thead>
                      <tr>
                        <th>Produk</th>
                        <th>IN</th>
                        <th>OUT</th>
                        <th>SISA</th>
                        <th>TOTAL</th>
                      </tr>
                    </thead>
                    <tbody>
                      {resumeRows.length === 0 && (
                        <tr><td colSpan={5}><div className="ks-empty">Tidak ada produk.</div></td></tr>
                      )}
                      {resumeRows.map((r) => (
                        <tr key={r.id}>
                          <td>{r.name}</td>
                          <td className="ks-mono ks-in">{numFmt(r.inZak)} / {numFmt(r.inKg)} kg</td>
                          <td className="ks-mono ks-out">{numFmt(r.outZak)} / {numFmt(r.outKg)} kg</td>
                          <td><span className="ks-sisa-badge">{numFmt(r.sisaZak)}</span></td>
                          <td className="ks-mono">{numFmt(r.sisaKg)} kg</td>
                        </tr>
                      ))}
                    </tbody>
                    {resumeRows.length > 0 && (
                      <tfoot>
                        <tr className="ks-resume-total-row">
                          <td>TOTAL</td>
                          <td className="ks-mono ks-in">{numFmt(resumeTotals.inZak)} / {numFmt(resumeTotals.inKg)} kg</td>
                          <td className="ks-mono ks-out">{numFmt(resumeTotals.outZak)} / {numFmt(resumeTotals.outKg)} kg</td>
                          <td><span className="ks-sisa-badge">{numFmt(resumeTotals.sisaZak)}</span></td>
                          <td className="ks-mono">{numFmt(resumeTotals.sisaKg)} kg</td>
                        </tr>
                      </tfoot>
                    )}
                  </table>
                </div>
                <div className="ks-modal-actions">
                  <button type="button" className="ks-btn ghost" onClick={() => setResumeOpen(false)}>Tutup</button>
                  <button type="button" className="ks-btn" onClick={handleResumeExport} disabled={resumeExporting || resumeRows.length === 0}>
                    {resumeExporting ? "Menyiapkan..." : "⬇ Export Resume ke Excel"}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {toast && <div className="ks-toast">{toast}</div>}
      </>
      )}
    </div>
  );
}