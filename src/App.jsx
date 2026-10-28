// Pinned subject: simple RT finance app for a neighborhood treasurer. Plan: warm paper ledger meets modern utility, with terracotta as the one owned accent, a fast entry flow, and monthly recaps that make cash movement obvious.
import { useEffect, useMemo, useRef, useState } from "react";
import { ArrowDownLeft, ArrowUpRight, BookOpen, CalendarDays, ChevronDown, FileText, Home, Menu, Pencil, Plus, ReceiptText, Search, Settings2, Trash2, Users, WalletCards, X } from "lucide-react";
import * as XLSX from "xlsx";
import "./styles.css";

const formatTanggalIndonesia = (date = new Date()) =>
  new Intl.DateTimeFormat("id-ID", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
  }).format(date);

const namaBulan = [
  "Januari", "Februari", "Maret", "April",
  "Mei", "Juni", "Juli", "Agustus",
  "September", "Oktober", "November", "Desember"
];

const getPeriodeSekarang = () => {
  const sekarang = new Date();
  return `${namaBulan[sekarang.getMonth()]} ${sekarang.getFullYear()}`;
};

const getTanggalTransaksi = () => {
  return new Intl.DateTimeFormat("id-ID", {
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(new Date());
};

const getDaftarPeriode = () => {
  // Tampilkan hanya Januari–Desember pada tahun berjalan.
  const tahunSekarang = new Date().getFullYear();
  return namaBulan.map((bulan) => `${bulan} ${tahunSekarang}`);
};

const periodeDariTanggal = (tanggal = "") => {
  const match = String(tanggal).match(/\b(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})\b/);
  if (!match) return "";
  const bulanMap = { jan: "Januari", feb: "Februari", mar: "Maret", apr: "April", mei: "Mei", jun: "Juni", jul: "Juli", agu: "Agustus", sep: "September", okt: "Oktober", nov: "November", des: "Desember" };
  const singkatan = match[2].toLowerCase().slice(0, 3);
  const bulan = bulanMap[singkatan];
  return bulan ? `${bulan} ${match[3]}` : "";
};

const cocokPeriode = (item, periode) => {
  if (periode === "Semua periode") return true;
  if (item.period) return item.period === periode;

  // Untuk data lama yang belum menyimpan field periode, cocokkan bulan DAN tahun pada tanggal.
  const [bulanDipilih, tahunDipilih] = periode.split(" ");
  const indeksBulan = namaBulan.indexOf(bulanDipilih);
  const singkatanBulan = ["jan", "feb", "mar", "apr", "mei", "jun", "jul", "agu", "sep", "okt", "nov", "des"][indeksBulan];
  const tanggal = (item.date || "").toLocaleLowerCase("id-ID");
  return Boolean(singkatanBulan && tanggal.includes(singkatanBulan) && tanggal.includes(tahunDipilih));
};

const seedResidents = [
  { id: 1, block: "A", subBlock: "A1", name: "Budi Santoso", house: "A-01" },
  { id: 2, block: "A", subBlock: "A2", name: "Siti Aminah", house: "A-02" },
  { id: 3, block: "B", subBlock: "B1", name: "Andi Wijaya", house: "B-03" },
  { id: 4, block: "B", subBlock: "B2", name: "Rina Kartika", house: "B-04" },
  { id: 5, block: "C", subBlock: "C1", name: "Dedi Kurniawan", house: "C-07" },
];
const seedEntries = [
  { id: 1, date: "21 Okt 2026", resident: "Budi Santoso", block: "A", clean: 15000, kas: 10000, other: 0 },
  { id: 2, date: "20 Okt 2026", resident: "Siti Aminah", block: "A", clean: 15000, kas: 10000, other: 5000 },
  { id: 3, date: "18 Okt 2026", resident: "Andi Wijaya", block: "B", clean: 15000, kas: 10000, other: 0 },
  { id: 4, date: "16 Okt 2026", resident: "Rina Kartika", block: "B", clean: 15000, kas: 10000, other: 0 },
];
const STORAGE_KEY = "kas-warga-rt-data-v1";
const readStored = () => { try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "null"); } catch { return null; } };
const money = (n) => new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", maximumFractionDigits: 0 }).format(n);
const exportExcel = (entries, period = "Semua periode") => {
  const rows = entries.map((e, index) => ({
    No: index + 1,
    "Nama Warga": e.resident,
    Periode: e.period || period,
    "Tanggal Bayar": e.date || "",
    Blok: e.block,
    "Iuran Kebersihan": e.clean,
    "Iuran Kas RT": e.kas,
    "Iuran Lainnya": e.other || 0,
    Total: e.clean + e.kas + (e.other || 0),
    Status: "Lunas",
    Tanggal: e.date,
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet["!cols"] = [{ wch: 6 }, { wch: 22 }, { wch: 20 }, { wch: 18 }, { wch: 10 }, { wch: 19 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 12 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, `Rekap ${period}`.slice(0, 31));
  XLSX.writeFile(workbook, `rekap-kas-rt-${period.toLowerCase().replace(/\s+/g, "-")}.xlsx`);
};

export default function App({ user, onSignOut }) {
  const [active, setActive] = useState("overview");
  const signOutStarted = useRef(false);
  const [isOnline, setIsOnline] = useState(() => navigator.onLine);
  const [lastSync, setLastSync] = useState(() => localStorage.getItem("kas-warga-rt-last-sync") || "Belum disinkronkan");
  const stored = readStored();
  const [residents, setResidents] = useState(stored?.residents ?? []);
  const [entries, setEntries] = useState(stored?.entries ?? []);
  const defaultCashOut = [{ id: 1, date: "19 Okt 2026", title: "Pembelian alat kebersihan", amount: 850000 }, { id: 2, date: "12 Okt 2026", title: "Lampu jalan RT", amount: 400000 }];
  const [cashOut, setCashOut] = useState(stored?.cashOut ?? []);
  const [openingBalance, setOpeningBalance] = useState(stored?.openingBalance ?? 0);
  const [manualAdjustment, setManualAdjustment] = useState(stored?.manualAdjustment ?? 0);
  const [balanceHistory, setBalanceHistory] = useState(stored?.balanceHistory ?? []);
  const [expenseTitle, setExpenseTitle] = useState("");
  const [expenseAmount, setExpenseAmount] = useState("");
  const [expensePeriod, setExpensePeriod] = useState(getPeriodeSekarang());
  const [editingExpense, setEditingExpense] = useState(null);
  const [block, setBlock] = useState("");
  const [paymentSubBlock, setPaymentSubBlock] = useState("");
  const [resident, setResident] = useState("");
  const selectedResident = residents.find((r) => r.name === resident);
  const autoSubBlock = selectedResident?.subBlock || "";
  const [clean, setClean] = useState("");
  const [kas, setKas] = useState("");
  const [other, setOther] = useState("");
  const [paymentPeriod, setPaymentPeriod] = useState(getPeriodeSekarang());
  const [residentName, setResidentName] = useState("");
  const [residentBlock, setResidentBlock] = useState("A");
  const [residentSubBlock, setResidentSubBlock] = useState("A1");
  const [editingResident, setEditingResident] = useState(null);
  const [editingEntry, setEditingEntry] = useState(null);
  const [residentSearch, setResidentSearch] = useState("");
  const [recapSearch, setRecapSearch] = useState("");
  const [residentFilter, setResidentFilter] = useState("Semua blok");
  const [residentSubFilter, setResidentSubFilter] = useState("Semua sub blok");
  const [notice, setNotice] = useState("");
  const [summaryPeriod, setSummaryPeriod] = useState(getPeriodeSekarang());
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const fileInput = { current: null };

  useEffect(() => {
    const online = () => { setIsOnline(true); const now = new Date().toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" }); localStorage.setItem("kas-warga-rt-last-sync", now); setLastSync(now); setNotice("Sinkronisasi otomatis selesai"); setTimeout(() => setNotice(""), 2200); }; const offline = () => setIsOnline(false);
    window.addEventListener("online", online); window.addEventListener("offline", offline);
    return () => { window.removeEventListener("online", online); window.removeEventListener("offline", offline); };
  }, []);
  // Keluar otomatis: setelah 15 menit tanpa aktivitas, serta upaya saat tab/jendela ditutup.
  // Catatan: browser dapat menghentikan permintaan jaringan saat ditutup; timeout tidak bergantung pada pagehide.
  useEffect(() => {
    if (typeof onSignOut !== "function") return undefined;

    let idleTimer;
    const IDLE_TIMEOUT_MS = 15 * 60 * 1000;

    const jalankanLogout = (alasan = "otomatis") => {
      if (signOutStarted.current) return;
      signOutStarted.current = true;
      window.removeEventListener("pagehide", handlePageHide);
      try {
        const hasil = onSignOut();
        if (hasil && typeof hasil.catch === "function") {
          hasil.catch((error) => {
            console.error(`Gagal menjalankan logout ${alasan}:`, error);
            signOutStarted.current = false;
          });
        }
      } catch (error) {
        console.error(`Gagal menjalankan logout ${alasan}:`, error);
        signOutStarted.current = false;
      }
    };

    const handlePageHide = () => jalankanLogout("saat halaman ditutup");
    const aturUlangTimer = () => {
      window.clearTimeout(idleTimer);
      if (!signOutStarted.current) {
        idleTimer = window.setTimeout(() => jalankanLogout("karena tidak aktif selama 15 menit"), IDLE_TIMEOUT_MS);
      }
    };

    const aktivitas = ["pointerdown", "keydown", "mousemove", "touchstart", "scroll"];
    aktivitas.forEach((eventName) => window.addEventListener(eventName, aturUlangTimer, { passive: true }));
    window.addEventListener("pagehide", handlePageHide);
    aturUlangTimer();

    return () => {
      window.clearTimeout(idleTimer);
      aktivitas.forEach((eventName) => window.removeEventListener(eventName, aturUlangTimer));
      window.removeEventListener("pagehide", handlePageHide);
    };
  }, [onSignOut]);
  useEffect(() => {
    setEntries((current) => current.map((item) => item.resident ? item : { ...item, resident: residents.find((r) => r.block === item.block && r.subBlock === item.subBlock)?.name || "Nama belum diisi" }));
  }, []);
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ residents, entries, cashOut, openingBalance, manualAdjustment, balanceHistory }));
  }, [residents, entries, cashOut, openingBalance, manualAdjustment, balanceHistory]);

  const total = useMemo(() => entries.reduce((a, e) => a + e.clean + e.kas + (e.other || 0), 0), [entries]);
  const runningBalance = useMemo(() => openingBalance + total - cashOut.reduce((a, e) => a + e.amount, 0) + manualAdjustment, [openingBalance, total, cashOut, manualAdjustment]);
  const backupData = () => {
    const payload = { app: "Kas Warga RT", version: 1, exportedAt: new Date().toISOString(), residents, entries, cashOut, openingBalance, manualAdjustment, balanceHistory };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob); const link = document.createElement("a");
    link.href = url; link.download = `backup-kas-rt-${new Date().toISOString().slice(0, 10)}.json`; link.click(); URL.revokeObjectURL(url);
    setNotice("Backup data berhasil diunduh"); setTimeout(() => setNotice(""), 2200);
  };
  const resetData = () => {
    if (!window.confirm("Kosongkan semua data kas, warga, transaksi, dan riwayat? Data dummy akan dihapus.")) return;
    setResidents([]); setEntries([]); setCashOut([]); setOpeningBalance(0); setManualAdjustment(0); setBalanceHistory([]);
    setNotice("Semua data berhasil dikosongkan"); setTimeout(() => setNotice(""), 2600);
  };
  const restoreData = (event) => {
    const file = event.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = () => { try { const data = JSON.parse(reader.result); if (!Array.isArray(data.residents) || !Array.isArray(data.entries)) throw new Error(); setResidents(data.residents); setEntries(data.entries); if (Number.isFinite(Number(data.openingBalance))) setOpeningBalance(Number(data.openingBalance)); if (Number.isFinite(Number(data.manualAdjustment))) setManualAdjustment(Number(data.manualAdjustment)); if (Array.isArray(data.balanceHistory)) setBalanceHistory(data.balanceHistory); setNotice("Data berhasil dipulihkan"); } catch { setNotice("File backup tidak valid"); } setTimeout(() => setNotice(""), 2600); event.target.value = ""; };
    reader.readAsText(file);
  };
  const addExpense = (event) => {
    event.preventDefault();
    if (!expenseTitle.trim() || Number(expenseAmount) <= 0) return;
    
const record = {
  id: editingExpense ? editingExpense.id : Date.now(),
  date: editingExpense
    ? editingExpense.date
    : getTanggalTransaksi(),
  title: expenseTitle.trim(),
  period: expensePeriod || getPeriodeSekarang(),
  amount: Number(expenseAmount)
};

    setCashOut(editingExpense ? cashOut.map((item) => item.id === editingExpense.id ? record : item) : [record, ...cashOut]);
    setEditingExpense(null);
    setExpenseTitle("");
    setExpenseAmount("");
    setExpensePeriod(getPeriodeSekarang());
    setNotice(editingExpense ? "Pengeluaran diperbarui" : "Pengeluaran berhasil dicatat");
    setTimeout(() => setNotice(""), 2200);
  };
  const addEntry = (e) => {
  e.preventDefault();

  const record = {
    id: editingEntry ? editingEntry.id : Date.now(),

    // Tanggal transaksi mengikuti tanggal saat ini.
    // Jika mengedit transaksi lama, tanggalnya dipertahankan.
    date: editingEntry
      ? editingEntry.date
      : getTanggalTransaksi(),

    // Periode iuran dipilih terpisah dari tanggal transaksi.
    period: paymentPeriod || getPeriodeSekarang(),

    resident,
    block: block.trim().toUpperCase(),
    subBlock: paymentSubBlock || autoSubBlock,
    clean: Number(clean),
    kas: Number(kas),
    other: Number(other),
  };

  setEntries(
    editingEntry
      ? entries.map((item) =>
          item.id === editingEntry.id ? record : item
        )
      : [record, ...entries]
  );

  setEditingEntry(null);

  // Setelah penyimpanan, form kembali ke periode bulan ini.
  setPaymentPeriod(getPeriodeSekarang());
  setBlock("");
  setResident("");
  setClean("");
  setKas("");
  setOther("");
  setPaymentSubBlock("");

  setNotice(
    editingEntry
      ? "Pembayaran diperbarui"
      : "Pembayaran berhasil dicatat"
  );

  setTimeout(() => setNotice(""), 2200);
};
  
  const addResident = (e) => {
    e.preventDefault();
    const name = residentName.trim();
    const blockName = residentBlock.trim();
    const subBlockName = residentSubBlock.trim();
    if (!name || !blockName || !subBlockName) {
      setNotice("Nama, blok, dan sub blok wajib diisi");
      setTimeout(() => setNotice(""), 2600);
      return;
    }
    const duplicate = residents.some((r) => r.name.trim().toLowerCase() === name.toLowerCase() && r.id !== editingResident);
    if (duplicate) {
      setNotice("Nama warga sudah terdaftar");
      setTimeout(() => setNotice(""), 2600);
      return;
    }
    if (editingResident) {
      setResidents(residents.map((r) => r.id === editingResident ? { ...r, name, block: blockName, subBlock: subBlockName } : r));
      setEditingResident(null);
      setNotice("Data warga diperbarui");
    } else {
      setResidents([...residents, { id: Date.now(), block: blockName, subBlock: subBlockName, name, house: `${subBlockName}-${String(residents.length + 1).padStart(2, "0")}` }]);
      setNotice("Nama warga ditambahkan");
    }
    setResidentName("");
    setTimeout(() => setNotice(""), 2200);
  };
  const startEditExpense = (expense) => {
    setEditingExpense(expense);
    setExpenseTitle(expense.title);
    setExpenseAmount(String(expense.amount));
    setExpensePeriod(expense.period || periodeDariTanggal(expense.date) || getPeriodeSekarang());
    setActive("cashflow");
  };
  const deleteExpense = (id) => {
    const target = cashOut.find((item) => item.id === id);
    if (!target || !window.confirm(`Hapus pengeluaran ${target.title}?`)) return;
    setCashOut(cashOut.filter((item) => item.id !== id));
    setNotice("Pengeluaran dihapus");
    setTimeout(() => setNotice(""), 2200);
  };
  const startEditEntry = (entry) => {
    setEditingEntry(entry);
    setResident(entry.resident);
    setBlock(entry.block);
    setPaymentSubBlock(entry.subBlock || "");
    setClean(String(entry.clean));
    setKas(String(entry.kas));
    setOther(String(entry.other || 0));
    setActive("input");
  };
  const deleteEntry = (id) => {
    const target = entries.find((item) => item.id === id);
    if (!target || !window.confirm(`Hapus pembayaran ${target.resident}?`)) return;
    setEntries(entries.filter((item) => item.id !== id));
    setNotice("Pembayaran dihapus");
    setTimeout(() => setNotice(""), 2200);
  };
  const startEditResident = (r) => {
    setEditingResident(r.id);
    setResidentName(r.name);
    setResidentBlock(r.block);
    setResidentSubBlock(r.subBlock || `${r.block}1`);
  };
  const deleteResident = (id) => {
    const target = residents.find((r) => r.id === id);
    if (!target || !window.confirm(`Hapus data ${target.name}?`)) return;
    setResidents(residents.filter((r) => r.id !== id));
    setNotice("Data warga dihapus");
    setTimeout(() => setNotice(""), 2200);
  };
  return (
    <div className="app-shell">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        :root, body, .app-shell, .app-shell *, .login-screen, .login-screen * {
          font-family: 'Inter', sans-serif !important;
        }
        :root, body, .app-shell { font-size: 14px; }

        /* Tema Emerald Modern: hijau elegan, latar segar, dan kartu bersih */
        :root {
          --emerald-forest: #123B31;
          --emerald: #0F766E;
          --emerald-hover: #115E59;
          --mint: #34D399;
          --page-bg: #F3F8F5;
          --surface: #FFFFFF;
          --border-soft: #DCE9E2;
          --text-main: #20352E;
          --text-muted: #6B7F76;
        }
        body, .app-shell, .main {
          background: var(--page-bg) !important;
          color: var(--text-main);
        }
        .sidebar {
          background: var(--emerald-forest) !important;
          color: #F2FBF6 !important;
          border-color: #245347 !important;
        }
        .brand strong, .brand small, .side-label, .sidebar .nav-item,
        .sidebar .profile, .sidebar .profile small {
          color: #E5F4EC !important;
        }
        .brand-mark {
          color: var(--mint) !important;
        }
        .sidebar .nav-item {
          border-color: transparent !important;
        }
        .sidebar .nav-item:hover, .sidebar .nav-item.active {
          background: rgba(52, 211, 153, .15) !important;
          color: #FFFFFF !important;
        }
        .sidebar .nav-item.active {
          box-shadow: inset 3px 0 0 var(--mint);
        }
        .avatar, .resident-avatar {
          background: #D1FAE5 !important;
          color: var(--emerald-forest) !important;
        }
        .main, .content {
          color: var(--text-main);
        }
        .topbar, .section-head, .form-intro, .report-toolbar {
          color: var(--text-main);
        }
        .eyebrow, .muted, .subcopy, .field-hint, .stat-card small,
        .history-empty, .login-note {
          color: var(--text-muted);
        }
        .hero-card, .recap-banner, .recap-banner.cashflow-banner,
        .cashflow-banner, .report-heading {
          background: linear-gradient(125deg, #123B31 0%, #0F766E 100%) !important;
          color: #FFFFFF !important;
          border-color: #0F766E !important;
        }
        .hero-card .eyebrow, .hero-card .hero-note,
        .recap-banner .eyebrow, .recap-banner h2, .recap-banner p,
        .cashflow-banner .eyebrow, .cashflow-banner h2, .cashflow-banner p {
          color: #F0FDF4 !important;
        }
        .report-heading {
          padding: 28px 32px !important;
          border-radius: 12px !important;
          min-height: 150px;
          display: flex;
          align-items: center;
        }
        .report-heading .report-kicker,
        .report-heading h1,
        .report-heading p {
          color: #FFFFFF !important;
        }
        .report-heading h1 {
          margin: 8px 0 !important;
          line-height: 1.2 !important;
        }
        .report-period-control {
          display: flex;
          align-items: center;
          gap: 10px;
          flex-wrap: wrap;
          margin: 0 0 20px;
          color: var(--text-muted);
          font-size: 13px;
          font-weight: 600;
        }
        .report-period-control select {
          min-width: 190px;
          padding: 10px 12px;
          border: 1px solid var(--border-soft);
          border-radius: 10px;
          background: #FFFFFF;
          color: var(--text-main);
          font: inherit;
        }
        .report-period-control select:focus {
          outline: 2px solid rgba(15, 118, 110, .2);
          border-color: var(--emerald);
        }
        .hero-art .sun, .login-sun {
          background: var(--mint) !important;
        }
        .hero-art .hill, .login-orbit {
          border-color: rgba(167, 243, 208, .5) !important;
        }
        .stat-card, .big-card, .table-card, .form-card, .history-card,
        .backup-card, .balance-settings, .public-breakdown, .public-kpi,
        .resident-card, .report-paper, .expense-form, .login-card {
          background: var(--surface) !important;
          border-color: var(--border-soft) !important;
          box-shadow: 0 8px 24px rgba(18, 59, 49, .055) !important;
        }
        .stat-icon.green, .stat-icon.gold {
          background: #D1FAE5 !important;
          color: var(--emerald) !important;
        }
        .stat-icon.red {
          background: #FCE7E7 !important;
          color: #B45353 !important;
        }
        .primary-btn, .login-btn {
          background: var(--emerald) !important;
          color: #FFFFFF !important;
          border-color: var(--emerald) !important;
        }
        .primary-btn:hover, .login-btn:hover {
          background: var(--emerald-hover) !important;
          border-color: var(--emerald-hover) !important;
        }
        .text-btn, .print-btn, .restore-btn {
          color: var(--emerald) !important;
        }
        .outline-light {
          color: #FFFFFF !important;
          border-color: rgba(255,255,255,.65) !important;
          background: rgba(255,255,255,.08) !important;
        }
        .outline-light:hover {
          background: rgba(255,255,255,.18) !important;
        }
        input, select, textarea, .input-money {
          border-color: #C9DCD2 !important;
          background-color: #FFFFFF !important;
          color: var(--text-main) !important;
        }
        input:focus, select:focus, textarea:focus {
          border-color: var(--emerald) !important;
          outline-color: rgba(15, 118, 110, .18) !important;
          box-shadow: 0 0 0 3px rgba(15, 118, 110, .10) !important;
        }
        .input-money > span {
          color: var(--emerald) !important;
          background: #F0FDF4 !important;
        }
        .table-card table thead, .table-card thead tr {
          background: #EAF4EE !important;
        }
        .table-card th {
          color: #31594B !important;
        }
        .table-card tbody tr:hover {
          background: #F5FBF7 !important;
        }
        .block-tag, .status-paid, .flow-in {
          background: #D1FAE5 !important;
          color: #065F46 !important;
          border-color: #A7F3D0 !important;
        }
        .flow-out, .icon-btn.delete {
          background: #FCE7E7 !important;
          color: #B45353 !important;
        }
        .icon-btn.edit {
          background: #E6F4ED !important;
          color: var(--emerald) !important;
        }
        .bar, .bar i {
          background: #D1FAE5 !important;
        }
        .bar i {
          background: var(--emerald) !important;
        }
        .notebook-icon, .tip > span, .history-dot {
          background: #D1FAE5 !important;
          color: var(--emerald) !important;
        }
        .privacy-note {
          background: #ECFDF5 !important;
          border-color: #A7F3D0 !important;
          color: #065F46 !important;
        }
        .connection-status span, .nav-dot {
          background: var(--mint) !important;
        }
        .mobile-menu-backdrop {
          background: rgba(8, 32, 25, .48) !important;
        }
        .toast {
          background: #123B31 !important;
          color: #FFFFFF !important;
        }
        .login-screen {
          background: var(--page-bg) !important;
        }
        .login-art {
          background: linear-gradient(145deg, #123B31 0%, #0F766E 100%) !important;
        }
        .login-copy {
          color: var(--text-muted);
        }
        @media (max-width: 760px) {
          .sidebar { background: var(--emerald-forest) !important; }
        }

        
/* Cetak/PDF: A4 Portrait, ringkas, elegan */
@page {
  size: A4 portrait;
  margin: 10mm;
}

@media print {
  :root {
    color-scheme: light !important;
  }

  html,
  body,
  #root,
  .app-shell {
    width: auto !important;
    min-width: 0 !important;
    max-width: none !important;
    height: auto !important;
    min-height: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
    overflow: visible !important;
    background: #FFFFFF !important;
  }

  body {
    color: #20352E !important;
    font-family: Arial, Helvetica, sans-serif !important;
    font-size: 9pt !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

  *,
  *::before,
  *::after {
    animation: none !important;
    transition: none !important;
    box-shadow: none !important;
    text-shadow: none !important;
  }

  /* Sembunyikan elemen navigasi dan kontrol */
  .sidebar,
  .topbar,
  .connection-status,
  .mobile-menu-backdrop,
  .recap-actions,
  .public-controls,
  .recap-page .primary-btn,
  .search,
  .toast,
  .login-screen,
  .add-resident,
  .resident-actions,
  .resident-toolbar select,
  .resident-toolbar .print-btn,
  .report-period-control,
  button,
  input,
  select,
  textarea {
    display: none !important;
  }

  /* Area utama memenuhi halaman cetak */
  .main,
  .content,
  .content.narrow,
  .recap-page,
  .public-summary,
  .report-page,
  .cashflow-page,
  .overview-page {
    display: block !important;
    width: 100% !important;
    min-width: 0 !important;
    max-width: none !important;
    min-height: 0 !important;
    height: auto !important;
    margin: 0 !important;
    padding: 0 !important;
    overflow: visible !important;
    background: #FFFFFF !important;
    border: 0 !important;
    border-radius: 0 !important;
  }

  /* Judul laporan */
  .report-toolbar {
    margin: 0 0 8px !important;
    padding: 0 0 6px !important;
    border-bottom: 1.5px solid #0F766E !important;
  }

  .report-toolbar h2,
  .public-heading h2 {
    margin: 0 0 3px !important;
    color: #123B31 !important;
    font-size: 17pt !important;
    line-height: 1.2 !important;
  }

  /* Header dan ringkasan */
  .recap-banner,
  .hero-card,
  .report-heading,
  .public-heading {
    min-height: 0 !important;
    height: auto !important;
    margin: 0 0 8px !important;
    padding: 10px 12px !important;
    color: #123B31 !important;
    background: #F0F7F3 !important;
    border: 1px solid #D7E7DE !important;
    border-radius: 7px !important;
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }

  .recap-banner h2,
  .hero-card h2,
  .report-heading h1 {
    margin: 2px 0 4px !important;
    color: #123B31 !important;
    font-size: 17pt !important;
    line-height: 1.2 !important;
  }

  .recap-banner p,
  .hero-card p,
  .report-heading p,
  .report-toolbar .subcopy,
  .public-heading p {
    margin: 2px 0 0 !important;
    color: #52665D !important;
    font-size: 8.5pt !important;
  }

  .report-heading h1,
  .report-heading p {
    color: #123B31 !important;
  }

  .eyebrow,
  .report-kicker {
    color: #0F766E !important;
    font-size: 7pt !important;
    font-weight: 700 !important;
    letter-spacing: .08em !important;
  }

  /* Kartu ringkasan dibuat lebih hemat ruang */
  .recap-grid,
  .public-kpis,
  .report-stats {
    display: grid !important;
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
    gap: 7px !important;
    margin: 7px 0 10px !important;
  }

  .big-card,
  .public-kpi,
  .public-breakdown,
  .privacy-note,
  .report-paper,
  .table-card,
  .stat-card,
  .form-card,
  .history-card,
  .resident-card,
  .expense-form,
  .balance-settings {
    min-width: 0 !important;
    min-height: 0 !important;
    height: auto !important;
    margin: 0 0 7px !important;
    padding: 8px 9px !important;
    color: #20352E !important;
    background: #FFFFFF !important;
    border: 1px solid #DCE7E1 !important;
    border-radius: 6px !important;
    break-inside: auto !important;
    page-break-inside: auto !important;
  }

  .big-card strong,
  .public-kpi strong,
  .report-stats b {
    display: block !important;
    margin: 2px 0 !important;
    color: #123B31 !important;
    font-size: 13pt !important;
    line-height: 1.2 !important;
  }

  .big-card small,
  .public-kpi small,
  .public-kpi span {
    color: #52665D !important;
    font-size: 7.5pt !important;
  }

  .bar {
    height: 4px !important;
    margin: 4px 0 !important;
  }

  /* Judul setiap bagian */
  .section-head {
    display: flex !important;
    justify-content: space-between !important;
    align-items: end !important;
    gap: 6px !important;
    margin: 9px 0 5px !important;
    padding: 0 0 4px !important;
    border-bottom: 1px solid #DCE7E1 !important;
  }

  .section-head h2,
  .report-paper h3,
  .public-breakdown h2 {
    margin: 1px 0 3px !important;
    color: #123B31 !important;
    font-size: 10.5pt !important;
    line-height: 1.25 !important;
  }

  /* Tabel laporan */
  .table-card,
  .table-card table,
  .recap-page table,
  .report-paper table,
  .public-summary table,
  table {
    display: table !important;
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    margin: 0 !important;
    overflow: visible !important;
    border-collapse: collapse !important;
    table-layout: fixed !important;
  }

  table {
    font-size: 7pt !important;
    color: #263B33 !important;
  }

  thead {
    display: table-header-group !important;
  }

  tbody {
    display: table-row-group !important;
  }

  tr {
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }

  th,
  td,
  .table-card th,
  .table-card td {
    padding: 4px !important;
    line-height: 1.25 !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
    word-break: normal !important;
    vertical-align: top !important;
    border-bottom: 1px solid #E2EAE5 !important;
  }

  th,
  .table-card th {
    color: #214D3E !important;
    background: #EAF4EE !important;
    font-size: 6.5pt !important;
    font-weight: 700 !important;
  }

  tbody tr:nth-child(even) {
    background: #F8FBF9 !important;
  }

  /* Label status pembayaran */
  .status-paid,
  .block-tag {
    display: inline-block !important;
    padding: 2px 4px !important;
    color: #065F46 !important;
    background: #E6F4ED !important;
    border: 1px solid #C8E7D7 !important;
    border-radius: 4px !important;
    font-size: 6.5pt !important;
  }

  /* Laporan dan catatan kaki */
  .report-paper {
    padding: 10px !important;
  }

  .report-intro,
  .report-footer p {
    margin: 6px 0 !important;
    color: #52665D !important;
    font-size: 8pt !important;
  }

  .report-footer {
    display: flex !important;
    justify-content: space-between !important;
    gap: 12px !important;
    margin-top: 12px !important;
    padding-top: 7px !important;
    border-top: 1px solid #DCE7E1 !important;
  }

  .report-footer strong {
    color: #123B31 !important;
  }

  .public-breakdown-row {
    padding: 5px 0 !important;
    border-bottom: 1px solid #E2EAE5 !important;
  }

  .public-footer {
    margin-top: 7px !important;
    padding-top: 5px !important;
    color: #52665D !important;
    border-top: 1px solid #DCE7E1 !important;
    font-size: 7.5pt !important;
  }

  /* Daftar warga dua kolom */
  .resident-grid {
    display: grid !important;
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
    gap: 6px !important;
  }

  .resident-card {
    display: flex !important;
    align-items: center !important;
    gap: 7px !important;
    padding: 7px !important;
    break-inside: avoid !important;
  }

  .resident-card small {
    font-size: 7pt !important;
  }

  .resident-avatar {
    flex: 0 0 28px !important;
    width: 28px !important;
    height: 28px !important;
    background: #E6F4ED !important;
    color: #0F766E !important;
  }

  .recap-page .table-card {
    padding: 0 !important;
  }

  .recap-details-title {
    font-size: 10.5pt !important;
  }

  .report-paper h3 {
    margin-top: 10px !important;
    padding-bottom: 3px !important;
    border-bottom: 1px solid #DCE7E1 !important;
  }

  .empty-table,
  .empty-residents {
    color: #64756D !important;
  }

  .content > *:last-child,
  .report-paper > *:last-child {
    margin-bottom: 0 !important;
  }
}
      `}</style>
      {mobileMenuOpen && (
  <button
    type="button"
    className="mobile-menu-backdrop"
    aria-label="Tutup menu"
    onClick={() => setMobileMenuOpen(false)}
  />
)}
      <aside className={`sidebar ${mobileMenuOpen ? "mobile-open" : ""}`}>
        <div className="brand"><Logo /><div><strong>kas warga</strong><small>RT 26 / RW 04</small></div></div>
        <div className="side-label">Menu utama</div>
        <nav>
          {[["overview", Home, "Ringkasan"], ["input", ReceiptText, "Input pembayaran"], ["cashflow", WalletCards, "Arus keuangan"], ["recap", CalendarDays, "Rekap bulanan"], ["report", FileText, "Laporan warga"], ["public", Users, "Ringkasan publik"], ["residents", Users, "Data warga"]].map(([id, Icon, label]) => (
            <button className={`nav-item ${active === id ? "active" : ""}`} key={id} onClick={() => { setActive(id); setSettingsOpen(false); setMobileMenuOpen(false); if (id === "input") { setPaymentPeriod(getPeriodeSekarang()); setBlock(""); setPaymentSubBlock(""); setResident(""); setClean(""); setKas(""); setOther(""); setEditingEntry(null); } }}><Icon size={18} />{label}{id === "input" && <span className="nav-dot" />}</button>
          ))}
        </nav>
        <div className="side-bottom"><button className={`nav-item ${settingsOpen ? "active" : ""}`} onClick={() => {setSettingsOpen(!settingsOpen);if (!settingsOpen) setActive("settings");setMobileMenuOpen(false);}}><Settings2 size={18} />Pengaturan</button><div className="profile"><span className="avatar">{(user?.user_metadata?.full_name || user?.email || "W").slice(0,2).toUpperCase()}</span><div><b>{user?.user_metadata?.full_name || "Pengelola"}</b><small>{user?.email || "Akun terverifikasi"}</small></div><button type="button" className="signout-btn" onClick={() => { setMobileMenuOpen(false); onSignOut?.(); }}>Keluar</button></div></div>
      </aside>
      <main className="main">
        <div className="connection-status"><span /> {isOnline ? "Online" : "Mode offline"} · {isOnline ? `Sinkron terakhir: ${lastSync}` : "Data tersimpan di perangkat"}</div>
        <header className="topbar"><button
  type="button"
  className="mobile-menu-toggle"
  onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
  aria-label="Buka atau tutup menu"
>
  {mobileMenuOpen ? <X size={22} /> : <Menu size={22} />}
</button><div><span className="eyebrow">{formatTanggalIndonesia()}</span><h1>{active === "overview" ? `Selamat datang, ${user?.user_metadata?.full_name || "Pengelola"}` : active === "input" ? "Input pembayaran" : active === "cashflow" ? "Arus keuangan" : active === "recap" ? "Rekap bulanan" : active === "report" ? "Laporan warga" : active === "public" ? "Ringkasan publik" : active === "settings" ? "Pengaturan" : "Data warga"}</h1></div><button className="month-pill"><CalendarDays size={16} /> Oktober 2026 <ChevronDown size={15} /></button></header>
        {notice && <div className="toast"><span>✓</span>{notice}</div>}
        {active === "overview" && !settingsOpen && <Overview entries={entries} cashOut={cashOut} residents={residents} total={total} openingBalance={openingBalance} manualAdjustment={manualAdjustment} outgoingTotal={cashOut.reduce((a, e) => a + e.amount, 0)} runningBalance={runningBalance} summaryPeriod={summaryPeriod} setSummaryPeriod={setSummaryPeriod} setActive={setActive} startEditEntry={startEditEntry} deleteEntry={deleteEntry} />}
        {active === "input" && !settingsOpen && <InputView {...{ block, setBlock, autoSubBlock, paymentSubBlock, setPaymentSubBlock, resident, setResident, clean, setClean, kas, setKas, other, setOther, paymentPeriod, setPaymentPeriod, addEntry, residents, editingEntry, setEditingEntry }} />}
        {active === "cashflow" && !settingsOpen && <Cashflow entries={entries} cashOut={cashOut} runningBalance={runningBalance} addExpense={addExpense} expenseTitle={expenseTitle} setExpenseTitle={setExpenseTitle} expenseAmount={expenseAmount} setExpenseAmount={setExpenseAmount} expensePeriod={expensePeriod} setExpensePeriod={setExpensePeriod} editingExpense={editingExpense} setEditingExpense={setEditingExpense} startEditExpense={startEditExpense} deleteExpense={deleteExpense} />}
        {active === "recap" && !settingsOpen && <Recap entries={entries} total={total} residents={residents} recapSearch={recapSearch} setRecapSearch={setRecapSearch} />}
        {active === "report" && !settingsOpen && <ResidentReport entries={entries} cashOut={cashOut} runningBalance={runningBalance} openingBalance={openingBalance} manualAdjustment={manualAdjustment} residents={residents} />}
        {active === "public" && !settingsOpen && <PublicSummary entries={entries} cashOut={cashOut} openingBalance={openingBalance} manualAdjustment={manualAdjustment} />}
        {settingsOpen && <SettingsPanel backupData={backupData} restoreData={restoreData} resetData={resetData} openingBalance={openingBalance} setOpeningBalance={setOpeningBalance} manualAdjustment={manualAdjustment} setManualAdjustment={setManualAdjustment} balanceHistory={balanceHistory} setBalanceHistory={setBalanceHistory} setNotice={setNotice} />}
        {active === "residents" && !settingsOpen && <Residents {...{ residents, residentName, setResidentName, residentBlock, setResidentBlock, residentSubBlock, setResidentSubBlock, addResident, editingResident, setEditingResident, startEditResident, deleteResident, residentSearch, setResidentSearch, residentFilter, setResidentFilter, residentSubFilter, setResidentSubFilter }} />}
      </main>
    </div>
  );
}

function Logo({ large = false }) { return <span className={`brand-mark ${large ? "large" : ""}`} aria-label="Logo Kas Warga"><svg viewBox="0 0 48 48" aria-hidden="true"><path d="M12 10v28M12 24l20-14M12 24l21 14" fill="none" stroke="currentColor" strokeWidth="4.2" strokeLinecap="round" strokeLinejoin="round"/><path d="M29 10h7v7M30 38h6v-7" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round"/><path d="M8 42h32" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" opacity=".65"/></svg></span>; }

function LoginScreen({ onLogin }) { return <main className="login-screen"><div className="login-art"><div className="login-orbit orbit-one" /><div className="login-orbit orbit-two" /><div className="login-sun" /><div className="login-logo-lockup"><Logo large /></div></div><section className="login-card"><div className="brand login-brand"><Logo /><div><strong>kas warga</strong><small>RT 26 / RW 04</small></div></div><span className="eyebrow">Ruang bendahara</span><h1>Kelola kas warga,<br/><em>tanpa ribet.</em></h1><p className="login-copy">Masuk untuk mencatat iuran, memantau arus kas, dan membagikan laporan yang transparan.</p><button className="login-btn" type="button" onClick={onLogin}>Masuk ke aplikasi <span>→</span></button><small className="login-note">Akses lokal untuk pengelola kas RT</small></section></main>; }

function SettingsPanel({ backupData, restoreData, resetData, openingBalance, setOpeningBalance, manualAdjustment, setManualAdjustment, balanceHistory, setBalanceHistory, setNotice }) {
  const saveBalanceChange = (event) => { event.preventDefault(); setBalanceHistory([{ id: Date.now(), date: "23 Okt 2026, 13:26", opening: openingBalance, correction: manualAdjustment }, ...balanceHistory]); setNotice("Perubahan saldo dicatat"); setTimeout(() => setNotice(""), 2200); };
  return <section className="content settings-panel">
    <div className="form-intro"><div className="notebook-icon"><Settings2 size={25} /></div><div><h2>Pengaturan data</h2><p>Amankan data kas RT dengan backup berkala.</p></div></div>
    <form className="balance-settings" onSubmit={saveBalanceChange}><div><span className="eyebrow">Saldo awal kas</span><h3>Atur saldo awal</h3><p>Nilai ini menjadi titik awal perhitungan saldo berjalan.</p></div><div className="settings-inputs"><label>Saldo awal<div className="input-money"><span>Rp</span><input type="number" min="0" value={openingBalance} onChange={e => setOpeningBalance(Number(e.target.value || 0))} /></div></label><label>Koreksi manual<div className="input-money"><span>Rp</span><input type="number" value={manualAdjustment} onChange={e => setManualAdjustment(Number(e.target.value || 0))} /></div></label><button className="primary-btn" type="submit"><Plus size={18} /> Simpan koreksi</button></div></form>
    <div className="history-card"><div className="section-head"><div><span className="eyebrow">Audit saldo</span><h2>Riwayat perubahan</h2></div></div>{balanceHistory.length === 0 ? <p className="history-empty">Belum ada perubahan saldo atau koreksi.</p> : <div className="history-list">{balanceHistory.map(item => <div className="history-row" key={item.id}><span className="history-dot">↻</span><div><b>Saldo diperbarui</b><small>{item.date}</small></div><div className="history-values"><span>Awal: {money(item.opening)}</span><span>Koreksi: {item.correction >= 0 ? "+" : ""}{money(item.correction)}</span></div></div>)}</div>}</div>
    <div className="backup-card"><div><span className="eyebrow">Backup dan pemulihan</span><h3>Simpan salinan data kas</h3><p>Backup menyimpan daftar warga dan seluruh catatan pembayaran dalam file JSON.</p></div><div className="backup-actions"><button className="primary-btn" onClick={backupData}><FileText size={17} /> Backup data</button><label className="restore-btn"><UploadIcon /> Pulihkan data<input type="file" accept="application/json,.json" onChange={restoreData} /></label></div></div><div className="tip"><span>✦</span><div><b>Saran aman</b><p>Simpan file backup di tempat aman dan buat salinan baru setelah ada perubahan penting.</p></div></div>
  </section>;
}
function UploadIcon() { return <span className="upload-icon">↥</span>; }
function Overview({ entries, cashOut = [], total, openingBalance = 0, manualAdjustment = 0, outgoingTotal, runningBalance, summaryPeriod, setSummaryPeriod, setActive, startEditEntry, deleteEntry, residents = [] }) { const periodEntries = entries.filter((e) => cocokPeriode(e, summaryPeriod)); const periodExpenses = cashOut.filter((e) => cocokPeriode(e, summaryPeriod)); const periodTotal = periodEntries.reduce((a, e) => a + e.clean + e.kas + (e.other || 0), 0); const periodOutgoing = periodExpenses.reduce((a, e) => a + Number(e.amount || 0), 0); const displayedBalance = summaryPeriod === "Semua periode" ? runningBalance : openingBalance + periodTotal - periodOutgoing + manualAdjustment;
  return <section className="content">
    <div className="hero-card"><div><span className="eyebrow light">Saldo kas berjalan{summaryPeriod === "Semua periode" ? " · keseluruhan" : ` · ${summaryPeriod}`}</span><div className="hero-number">{money(displayedBalance)}</div><span className="hero-note"><ArrowUpRight size={15} /> {summaryPeriod === "Semua periode" ? "Saldo total seluruh periode" : `Perhitungan untuk ${summaryPeriod}`}</span></div><div className="hero-art"><div className="sun" /><div className="hill h1" /><div className="hill h2" /><WalletCards size={42} /></div></div>
    <div className="stat-grid"><Stat title={`Pemasukan ${summaryPeriod === "Semua periode" ? "keseluruhan" : "periode"}`} value={money(periodTotal)} icon={ArrowDownLeft} tone="green" meta={`Dari ${periodEntries.length} pembayaran`} /><Stat title={summaryPeriod === "Semua periode" ? "Pengeluaran keseluruhan" : `Pengeluaran ${summaryPeriod}`} value={money(periodOutgoing)} icon={ArrowUpRight} tone="red" meta={`${periodOutgoing ? "Transaksi tercatat" : "Belum ada transaksi"}`} /><Stat title="Warga sudah bayar" value={`${periodEntries.length} / ${residents.length}`} icon={Users} tone="gold" meta={residents.length ? `${Math.round((periodEntries.length / residents.length) * 100)}% tingkat pembayaran` : "Belum ada data warga"} /></div>
    <div className="section-head"><div><span className="eyebrow">Aktivitas terbaru</span><h2>Catatan pembayaran</h2></div><select className="period-filter" value={summaryPeriod} onChange={e => setSummaryPeriod(e.target.value)}><option>Semua periode</option>{getDaftarPeriode().map((periode) => <option key={periode} value={periode}>{periode}</option>)}</select><button className="text-btn" onClick={() => setActive("input")}>+ Input pembayaran</button></div>
    <div className="table-card"><table><thead><tr><th>Tanggal</th><th>Nama warga</th><th>Blok</th><th>Iuran kebersihan</th><th>Iuran kas RT</th><th>Iuran lainnya</th><th>Total</th><th>Aksi</th></tr></thead><tbody>{periodEntries.length === 0 ? <tr><td colSpan="8" className="empty-table">Belum ada data pembayaran untuk periode ini.</td></tr> : periodEntries.map((e) => <tr key={e.id}><td className="muted">{e.date}</td><td><b>{e.resident || residents.find((r) => r.block === e.block && r.subBlock === e.subBlock)?.name || "Nama belum diisi"}</b></td><td><span className="block-tag">{e.subBlock || "-"}</span></td><td>{money(e.clean)}</td><td>{money(e.kas)}</td><td>{money(e.other || 0)}</td><td><b>{money(e.clean + e.kas + (e.other || 0))}</b></td><td><div className="entry-actions"><button className="icon-btn edit" title="Edit pembayaran" onClick={() => startEditEntry(e)}><Pencil size={15} /></button><button className="icon-btn delete" title="Hapus pembayaran" onClick={() => deleteEntry(e.id)}><Trash2 size={15} /></button></div></td></tr>)}</tbody></table></div>
  </section>;
}
function Stat({ title, value, icon: Icon, tone, meta }) { return <div className="stat-card"><div className={`stat-icon ${tone}`}><Icon size={18} /></div><span className="eyebrow">{title}</span><strong>{value}</strong><small>{meta}</small></div>; }
function InputView({ block, setBlock, autoSubBlock, paymentSubBlock, setPaymentSubBlock, resident, setResident, clean, setClean, kas, setKas, other, setOther, paymentPeriod, setPaymentPeriod, addEntry, residents, editingEntry, setEditingEntry }) {
  const normalizedBlock = block.trim().toUpperCase();
  const subBlocksForBlock = [...new Set(residents.filter((r) => r.block.trim().toUpperCase() === normalizedBlock).map((r) => r.subBlock).filter(Boolean))].sort();
  const matchingResident = residents.find((r) => r.block.trim().toUpperCase() === normalizedBlock && r.subBlock === paymentSubBlock);
  const syncResident = (nextBlock, nextSubBlock) => {
    const match = residents.find((r) => r.block.trim().toUpperCase() === nextBlock.trim().toUpperCase() && r.subBlock === nextSubBlock);
    setResident(match?.name || "");
  };
  return <section className="content narrow"><div className="form-intro"><div className="notebook-icon"><BookOpen size={25} /></div><div><h2>Catat iuran warga</h2><p>Isi sekali, total dihitung otomatis. Praktis buat bendahara.</p></div></div><form className="form-card" onSubmit={addEntry}><div className="form-grid"><label>
  Periode pembayaran
  <select
    value={paymentPeriod}
    onChange={(e) => setPaymentPeriod(e.target.value)}
    required
  >
    {getDaftarPeriode().map((periode) => (
      <option key={periode} value={periode}>
        {periode}
      </option>
    ))}
  </select>
</label><label>Blok<input value={block} onChange={(e) => { const nextBlock = e.target.value.toUpperCase(); const firstSub = residents.find((r) => r.block.trim().toUpperCase() === nextBlock)?.subBlock || ""; setBlock(nextBlock); setPaymentSubBlock(firstSub); syncResident(nextBlock, firstSub); }} placeholder="Contoh: A" required /></label><label>Sub blok<select value={paymentSubBlock} onChange={(e) => { setPaymentSubBlock(e.target.value); syncResident(block, e.target.value); }} disabled={!normalizedBlock}><option value="">{normalizedBlock ? "Pilih sub blok" : "Isi blok dahulu"}</option>{subBlocksForBlock.map((sub) => (
  <option key={sub} value={sub}>
    {sub} - {
      residents.find(
        (r) =>
          r.block.trim().toUpperCase() === normalizedBlock &&
          (r.subBlock || "").trim() === sub.trim()
      )?.name || "Nama belum diisi"
    }
  </option>
))}</select><small className="field-hint">Pilihan diambil dari Data Warga</small></label><label>Nama warga<input value={matchingResident?.name || resident} readOnly placeholder="Otomatis dari blok dan sub blok" /><small className="field-hint">Terisi otomatis dari Data Warga</small></label><label>Iuran kebersihan<div className="input-money"><span>Rp</span><input type="number" min="0" value={clean} onChange={(e) => setClean(e.target.value)} /></div></label><label>Iuran kas RT<div className="input-money"><span>Rp</span><input type="number" min="0" value={kas} onChange={(e) => setKas(e.target.value)} /></div></label><label>Iuran lainnya <em>opsional</em><div className="input-money"><span>Rp</span><input type="number" min="0" value={other} onChange={(e) => setOther(e.target.value)} /></div></label></div><div className="total-row"><div><span>Total pembayaran</span><strong>{money(Number(clean || 0) + Number(kas || 0) + Number(other || 0))}</strong></div><button className="primary-btn" type="submit"><Plus size={18} /> {editingEntry ? "Simpan perubahan" : "Simpan pembayaran"}</button>{editingEntry && <button className="cancel-btn" type="button" onClick={() => setEditingEntry(null)}><X size={16} /> Batal</button>}</div></form><div className="tip"><span>✦</span><div><b>Nama warga otomatis</b><p>Pilih blok dan sub blok, nama akan diambil langsung dari Data Warga.</p></div></div></section>;
}
function Cashflow({ entries, cashOut, runningBalance, addExpense, expenseTitle, setExpenseTitle, expenseAmount, setExpenseAmount, expensePeriod, setExpensePeriod, editingExpense, setEditingExpense, startEditExpense, deleteExpense }) { const incoming = entries.reduce((a,e)=>a+e.clean+e.kas+(e.other||0),0); const outgoing = cashOut.reduce((a,e)=>a+e.amount,0); return <section className="content"><div className="recap-banner cashflow-banner"><div><span className="eyebrow light">Catatan keuangan</span><h2>Arus masuk & keluar</h2><p>Pantau pergerakan kas warga.</p></div></div><form className="expense-form" onSubmit={addExpense}><div><span className="eyebrow">Input pengeluaran</span><h2>Catat uang keluar</h2></div><label>Periode pengeluaran<select value={expensePeriod} onChange={e=>setExpensePeriod(e.target.value)} required>{getDaftarPeriode().map(periode=><option key={periode} value={periode}>{periode}</option>)}</select></label><input placeholder="Keterangan pengeluaran" value={expenseTitle} onChange={e=>setExpenseTitle(e.target.value)} required /><div className="input-money"><span>Rp</span><input type="number" min="1" value={expenseAmount} onChange={e=>setExpenseAmount(e.target.value)} required /></div><button className="primary-btn" type="submit"><Plus size={18}/>{editingExpense ? "Simpan perubahan" : "Simpan pengeluaran"}</button></form><div className="recap-grid"><div className="big-card cash-balance"><span className="eyebrow">Saldo berjalan</span><strong>{money(runningBalance)}</strong></div><div className="big-card cash-in"><span className="eyebrow">Total uang masuk</span><strong>{money(incoming)}</strong></div><div className="big-card cash-out"><span className="eyebrow">Total uang keluar</span><strong>{money(outgoing)}</strong></div></div><div className="table-card"><table><thead><tr><th>Tanggal</th><th>Keterangan</th><th>Jenis</th><th>Nominal</th></tr></thead><tbody>{entries.map(e=><tr key={`in-${e.id}`}><td>{e.date}</td><td>Iuran {e.resident}</td><td><span className="flow-in">Masuk</span></td><td>{money(e.clean+e.kas+(e.other||0))}</td></tr>)}{cashOut.map(e=><tr key={`out-${e.id}`}><td>{e.date}</td><td>{e.title}</td><td><span className="flow-out">Keluar</span></td><td>{money(e.amount)}</td></tr>)}</tbody></table></div></section>; }
function ResidentReport({ entries, cashOut, runningBalance, openingBalance = 0, manualAdjustment = 0 }) { const [reportPeriod, setReportPeriod] = useState(getPeriodeSekarang()); const reportEntries = entries.filter(e => cocokPeriode(e, reportPeriod)); const reportExpenses = cashOut.filter(e => cocokPeriode(e, reportPeriod)); const incoming = reportEntries.reduce((a,e)=>a+Number(e.clean||0)+Number(e.kas||0)+Number(e.other||0),0); const outgoing = reportExpenses.reduce((a,e)=>a+Number(e.amount||0),0); const reportBalance = reportPeriod === "Semua periode" ? runningBalance : Number(openingBalance||0) + incoming - outgoing + Number(manualAdjustment||0); return <section className="content report-page"><div className="report-toolbar"><div><span className="eyebrow">Dokumen warga</span><h2>Laporan Pertanggungjawaban</h2><p className="subcopy">Ringkasan kas RT 26 / RW 04.</p></div><button className="primary-btn" onClick={()=>window.print()}><FileText size={17}/> Cetak / Simpan PDF</button></div><label className="report-period-control">Pilih bulan laporan<select value={reportPeriod} onChange={e=>setReportPeriod(e.target.value)}><option>Semua periode</option>{getDaftarPeriode().map(periode=><option key={periode} value={periode}>{periode}</option>)}</select></label><div className="report-paper"><div className="report-heading"><div><div className="report-kicker">RT 26 / RW 04</div><h1>Laporan Pertanggungjawaban Kas</h1><p>{reportPeriod === "Semua periode" ? "Semua periode" : `Periode ${reportPeriod}`}</p></div></div><p className="report-intro">Ringkasan pemasukan dan pengeluaran kas warga untuk diketahui bersama.</p><div className="report-stats"><div><span>Total pemasukan</span><b>{money(incoming)}</b></div><div><span>Total pengeluaran</span><b>{money(outgoing)}</b></div><div><span>Saldo kas saat ini</span><b>{money(reportBalance)}</b></div></div><h3>Rincian pemasukan</h3><table><tbody>{reportEntries.length ? reportEntries.map(e=><tr key={e.id}><td>{e.resident}</td><td>{e.date}</td><td>{money(e.clean+e.kas+(e.other||0))}</td></tr>) : <tr><td colSpan="3">Tidak ada pemasukan pada periode ini.</td></tr>}</tbody></table><h3>Rincian pengeluaran</h3><table><tbody>{reportExpenses.length ? reportExpenses.map(e=><tr key={e.id}><td>{e.title}</td><td>{e.date}</td><td>{money(e.amount)}</td></tr>) : <tr><td colSpan="3">Tidak ada pengeluaran pada periode ini.</td></tr>}</tbody></table><div className="report-footer"><p>Demikian laporan ini dibuat untuk dapat diketahui bersama.</p><div><span>Bendahara RT</span><strong>Bendahara RT</strong></div></div></div></section>; }
function PublicSummary({ entries, cashOut, openingBalance, manualAdjustment }) {
  const [period, setPeriod] = useState(getPeriodeSekarang());
  const periodEntries = entries.filter((e) => cocokPeriode(e, period));
  const periodExpenses = cashOut.filter((e) => cocokPeriode(e, period));
  const cleanTotal = periodEntries.reduce((sum, e) => sum + Number(e.clean || 0), 0);
  const kasTotal = periodEntries.reduce((sum, e) => sum + Number(e.kas || 0), 0);
  const otherTotal = periodEntries.reduce((sum, e) => sum + Number(e.other || 0), 0);
  const income = cleanTotal + kasTotal + otherTotal;
  const expense = periodExpenses.reduce((sum, e) => sum + Number(e.amount || 0), 0);
  const balance = Number(openingBalance || 0) + income - expense + Number(manualAdjustment || 0);
  return <section className="content public-summary">
    <div className="public-heading"><div><span className="eyebrow">Informasi untuk warga</span><h2>Laporan Kas Warga</h2><p>RT 26 / RW 04 · Ringkasan terbuka pengelolaan kas</p></div><div className="public-controls"><label>Pilih periode<select value={period} onChange={e => setPeriod(e.target.value)}><option>Semua periode</option>{getDaftarPeriode().map(periode => <option key={periode} value={periode}>{periode}</option>)}</select></label><button className="primary-btn" onClick={() => window.print()}><FileText size={16}/> Cetak / Simpan PDF</button></div></div>
    <div className="privacy-note"><span>✓</span><div><b>Tampilan publik, data pribadi disembunyikan</b><p>Tidak menampilkan nama warga, nomor rumah, blok, atau rincian pembayaran per orang.</p></div></div>
    <div className="public-kpis"><article className="public-kpi income"><span>Total pemasukan</span><strong>{money(income)}</strong><small>Periode {period}</small></article><article className="public-kpi expense"><span>Total pengeluaran</span><strong>{money(expense)}</strong><small>Periode {period}</small></article><article className="public-kpi balance"><span>Saldo akhir</span><strong>{money(balance)}</strong><small>Saldo awal + masuk - keluar + koreksi</small></article></div>
    <div className="public-breakdown"><div className="section-head"><div><span className="eyebrow">Rangkuman kategori</span><h2>Rincian agregat</h2></div></div><div className="public-breakdown-row"><span>Iuran kebersihan</span><b>{money(cleanTotal)}</b></div><div className="public-breakdown-row"><span>Iuran kas RT</span><b>{money(kasTotal)}</b></div><div className="public-breakdown-row"><span>Iuran lainnya</span><b>{money(otherTotal)}</b></div><div className="public-breakdown-row"><span>Total pengeluaran</span><b>{money(expense)}</b></div></div>
    <footer className="public-footer"><span>Kas Warga RT 26 / RW 04</span><span>Ringkasan agregat · {period}</span></footer>
  </section>;
}
function Recap({ entries, total, residents = [], recapSearch, setRecapSearch }) { const [recapPeriod, setRecapPeriod] = useState(getPeriodeSekarang()); const monthEntries = entries.filter((e) => cocokPeriode(e, recapPeriod)); const filteredEntries = monthEntries.filter((e) => (e.resident || "").toLocaleLowerCase("id-ID").includes(recapSearch.trim().toLocaleLowerCase("id-ID"))); const periodTotal = monthEntries.reduce((sum, e) => sum + e.clean + e.kas + (e.other || 0), 0); const paidCount = new Set(monthEntries.map(e => e.residentId || e.resident || `${e.block}-${e.subBlock}`)).size; const paidPercent = residents.length ? Math.round((paidCount / residents.length) * 100) : 0; return <section className="content recap-page"><div className="recap-banner"><div><span className="eyebrow light">Laporan periode</span><h2>Rekap {recapPeriod}</h2><p>Ringkasan pemasukan kas warga.</p></div><div className="recap-actions"><label className="recap-period-label">Pilih bulan<select className="recap-period-select" value={recapPeriod} onChange={e => setRecapPeriod(e.target.value)}><option>Semua periode</option>{getDaftarPeriode().map(periode => <option key={periode} value={periode}>{periode}</option>)}</select></label><button className="outline-light" onClick={() => exportExcel(monthEntries, recapPeriod)}><FileText size={16} /> Ekspor Excel</button><button className="outline-light" onClick={() => window.print()}><FileText size={16} /> Ekspor PDF</button></div></div><div className="recap-grid"><div className="big-card"><span className="eyebrow">Total terkumpul</span><strong>{money(periodTotal)}</strong><div className="bar"><i style={{ width: `${Math.min(paidPercent, 100)}%` }} /></div><small>{residents.length ? `${paidPercent}% warga sudah membayar` : "Belum ada data warga"}</small></div><div className="big-card"><span className="eyebrow">Warga sudah bayar</span><strong>{paidCount} / {residents.length}</strong><small>{residents.length ? `${paidPercent}% tingkat pembayaran` : "Belum ada data warga"}</small></div></div><div className="section-head"><div><span className="eyebrow">Per warga</span><h2 className="recap-details-title">Rincian iuran · {recapPeriod}</h2></div><div className="search"><Search size={16} /><input placeholder="Cari nama warga..." value={recapSearch} onChange={e => setRecapSearch(e.target.value)} /></div></div><div className="table-card"><table><thead><tr><th>Nama warga</th><th>Periode</th><th>Blok</th><th>Kebersihan</th><th>Kas RT</th><th>Iuran lainnya</th><th>Total</th><th>Status</th></tr></thead><tbody>{filteredEntries.length ? filteredEntries.map(e => <tr key={e.id}><td><b>{e.resident || residents.find(r => r.block === e.block && r.subBlock === e.subBlock)?.name || "Nama belum diisi"}</b></td><td>{e.period || recapPeriod}</td><td><span className="block-tag">{e.subBlock || "-"}</span></td><td>{money(e.clean)}</td><td>{money(e.kas)}</td><td>{money(e.other || 0)}</td><td><b>{money(e.clean + e.kas + (e.other || 0))}</b></td><td><span className="status-paid">Lunas</span></td></tr>) : <tr><td colSpan="8" className="empty-table">{monthEntries.length ? "Nama warga tidak ditemukan." : "Belum ada pembayaran pada periode ini."}</td></tr>}</tbody></table></div></section>; }
function Residents({ residents, residentName, setResidentName, residentBlock, setResidentBlock, residentSubBlock, setResidentSubBlock, addResident, editingResident, setEditingResident, startEditResident, deleteResident, residentSearch, setResidentSearch, residentFilter, setResidentFilter, residentSubFilter, setResidentSubFilter }) { const blocks = [...new Set(residents.map((r) => r.block).filter(Boolean))].sort(); const subBlocks = [...new Set(residents.map((r) => r.subBlock).filter(Boolean))].sort(); const query = residentSearch.trim().toLocaleLowerCase("id-ID"); const filteredResidents = residents.filter((r) => (residentFilter === "Semua blok" || r.block.toLocaleLowerCase("id-ID") === residentFilter.toLocaleLowerCase("id-ID")) && (residentSubFilter === "Semua sub blok" || (r.subBlock || "").toLocaleLowerCase("id-ID") === residentSubFilter.toLocaleLowerCase("id-ID")) && (r.name || "").toLocaleLowerCase("id-ID").includes(query)); return <section className="content narrow"><div className="section-head"><div><span className="eyebrow">Buku warga</span><h2>Daftar nama warga</h2><p className="subcopy">Kelola data yang muncul di menu input pembayaran.</p></div></div><form className="add-resident" onSubmit={addResident} noValidate><input placeholder="Nama lengkap warga" value={residentName} onChange={e => setResidentName(e.target.value)} /><input className="compact-field" value={residentBlock} onChange={e => setResidentBlock(e.target.value)} placeholder="Blok, contoh A" /><input className="compact-field" value={residentSubBlock} onChange={e => setResidentSubBlock(e.target.value)} placeholder="Sub blok, contoh A1" /><button className="primary-btn" type="submit"><Plus size={18} /> {editingResident ? "Simpan perubahan" : "Tambah warga"}</button>{editingResident && <button className="cancel-btn" type="button" onClick={() => { setEditingResident(null); setResidentName(""); }}><X size={16} /> Batal</button>}</form><div className="resident-toolbar"><div className="search resident-search"><Search size={16} /><input placeholder="Cari nama warga..." value={residentSearch} onChange={e => setResidentSearch(e.target.value)} /></div><select className="resident-filter" value={residentFilter} onChange={e => { setResidentFilter(e.target.value); setResidentSubFilter("Semua sub blok"); }}><option>Semua blok</option>{blocks.map(b => <option key={b}>{b}</option>)}</select><select className="resident-filter" value={residentSubFilter} onChange={e => setResidentSubFilter(e.target.value)}><option>Semua sub blok</option>{subBlocks.filter(sb => residentFilter === "Semua blok" || sb.startsWith(residentFilter)).map(sb => <option key={sb}>{sb}</option>)}</select><button type="button" className="print-btn" onClick={() => window.print()}><FileText size={15} /> Cetak daftar</button><span className="result-count">{filteredResidents.length} warga</span></div><div className="resident-grid">{filteredResidents.map(r => <div className="resident-card" key={r.id}><span className="resident-avatar">{r.name.split(" ").map(n => n[0]).join("").slice(0, 2)}</span><div><b>{r.name}</b><small>Blok {r.block} · Sub blok {r.subBlock || `${r.block}1`} · Rumah {r.house}</small></div><div className="resident-actions"><button className="icon-btn edit" title="Edit warga" onClick={() => startEditResident(r)}><Pencil size={15} /></button><button className="icon-btn delete" title="Hapus warga" onClick={() => deleteResident(r.id)}><Trash2 size={15} /></button></div></div>)}{filteredResidents.length === 0 && <div className="empty-residents">Tidak ada warga yang cocok.</div>}</div></section>; }
