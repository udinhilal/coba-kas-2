// Pinned subject: simple RT finance app for a neighborhood treasurer. Plan: warm paper ledger meets modern utility, with terracotta as the one owned accent, a fast entry flow, and monthly recaps that make cash movement obvious.
import { useEffect, useMemo, useRef, useState } from "react";
import { ArrowDownLeft, ArrowUpRight, BookOpen, CalendarDays, Check, ChevronDown, FileText, Globe2, Home, LogOut, Menu, Pencil, Plus, ReceiptText, Search, Settings2, Trash2, Users, WalletCards, X } from "lucide-react";
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
const DEFAULT_COMMUNITY_IDENTITY = {
  rtNumber: "26",
  rwNumber: "04",
  kelurahan: "",
  kecamatan: "",
  kabupatenKota: "",
};
const getCommunityIdentityLine = (identity = DEFAULT_COMMUNITY_IDENTITY) =>
  `RT ${String(identity.rtNumber || "26").trim()} / RW ${String(identity.rwNumber || "04").trim()}`;
const getCommunityLocationLine = (identity = DEFAULT_COMMUNITY_IDENTITY) =>
  [
    identity.kelurahan ? `Kel. ${String(identity.kelurahan).trim()}` : "",
    identity.kecamatan ? `Kec. ${String(identity.kecamatan).trim()}` : "",
    identity.kabupatenKota ? String(identity.kabupatenKota).trim() : "",
  ].filter(Boolean).join(" · ");

const STORAGE_KEY = "kas-warga-rt-data-v1";
const readStored = () => { try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "null"); } catch { return null; } };
const money = (n) => new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", maximumFractionDigits: 0 }).format(n);
// Format angka input dengan pemisah ribuan.
// Nilai yang disimpan tetap berupa angka tanpa titik.
const formatNominalInput = (value) => {
  const digits = String(value ?? "").replace(/\D/g, "");

  return digits
    ? new Intl.NumberFormat("id-ID").format(Number(digits))
    : "";
};
// Mengurutkan blok dan subblok secara alami:
// E.01, E.02, E.03_1, E.03_2, E.03_10, E.04
const compareNatural = (a, b) =>
  String(a || "").localeCompare(
    String(b || ""),
    "id-ID",
    { numeric: true, sensitivity: "base" }
  );

// Mengurutkan daftar pembayaran berdasarkan blok, subblok,
// kemudian nama warga. Data asli tidak diubah.
const sortPaymentEntries = (items = []) =>
  [...items].sort((a, b) => {
    const blokA = a.block || a.subBlock || "";
    const blokB = b.block || b.subBlock || "";

    return (
      compareNatural(blokA, blokB) ||
      compareNatural(a.subBlock || "", b.subBlock || "") ||
      compareNatural(a.resident || "", b.resident || "")
    );
  });
const exportExcel = (
  entries,
  period = "Semua periode",
  fileLabel = "",
  communityIdentityLine = "RT 26 / RW 04",
  communityLocationLine = ""
) => {
  // Urutkan data sebelum diekspor
  const sortedEntries = sortPaymentEntries(entries);

  // Hitung total masing-masing kategori
  const cleanTotal = sortedEntries.reduce(
    (sum, e) => sum + Number(e.clean || 0), 0
  );

  const kasTotal = sortedEntries.reduce(
    (sum, e) => sum + Number(e.kas || 0), 0
  );

  const otherTotal = sortedEntries.reduce(
    (sum, e) => sum + Number(e.other || 0), 0
  );

  const grandTotal = cleanTotal + kasTotal + otherTotal;

  // Rincian pembayaran
  const rows = sortedEntries.map((e, index) => ({
    No: index + 1,
    "Nama Warga": e.resident || "",
    Periode: e.period || period,
    "Tanggal Bayar": e.date || "",
    Blok: e.subBlock || e.block || "",
    "Iuran Kebersihan": Number(e.clean || 0),
    "Iuran Kas RT": Number(e.kas || 0),
    "Iuran Lainnya": Number(e.other || 0),
    Total:
      Number(e.clean || 0) +
      Number(e.kas || 0) +
      Number(e.other || 0),
    Status: e.status || "Lunas",
  }));

  const worksheet = XLSX.utils.aoa_to_sheet([
    ["KAS WARGA"],
    [communityIdentityLine],
    [communityLocationLine],
    [],
  ]);
  XLSX.utils.sheet_add_json(worksheet, rows, { origin: "A5" });

  // Tambahkan ringkasan total di bawah tabel pembayaran
  XLSX.utils.sheet_add_aoa(
    worksheet,
    [
      [],
      ["RINGKASAN TOTAL IURAN"],
      ["Total Iuran Kebersihan", cleanTotal],
      ["Total Iuran Kas RT", kasTotal],
      ["Total Iuran Lainnya", otherTotal],
      ["TOTAL PEMASUKAN", grandTotal],
    ],
    { origin: -1 }
  );

  // Lebarkan baris identitas ke seluruh lebar tabel Excel.
  worksheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 9 } },
  ];
  worksheet["!rows"] = [
    { hpt: 24 },
    { hpt: 19 },
    { hpt: 18 },
    { hpt: 8 },
  ];

  worksheet["!cols"] = [
    { wch: 6 },
    { wch: 24 },
    { wch: 20 },
    { wch: 18 },
    { wch: 14 },
    { wch: 20 },
    { wch: 17 },
    { wch: 18 },
    { wch: 18 },
    { wch: 12 },
  ];

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    `Rekap ${period}`.slice(0, 31)
  );

  XLSX.writeFile(
  workbook,
  `rekap-kas-rt-${period
    .toLowerCase()
    .replace(/\s+/g, "-")}${
    fileLabel
      ? "-" + fileLabel.toLowerCase().replace(/\s+/g, "-")
      : ""
  }.xlsx`
);
};
export default function App({ user, onSignOut }) {

  // Mengambil nama pengguna dari akun yang sedang login
  const namaPengguna =
    user?.user_metadata?.username?.trim() ||
    user?.user_metadata?.full_name?.trim() ||
    user?.user_metadata?.name?.trim() ||
    user?.email?.split("@")[0] ||
    "Pengelola";
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
  const [communityIdentity, setCommunityIdentity] = useState(() => ({
    ...DEFAULT_COMMUNITY_IDENTITY,
    ...(stored?.communityIdentity || {}),
  }));
  const communityIdentityLine = getCommunityIdentityLine(communityIdentity);
  const communityLocationLine = getCommunityLocationLine(communityIdentity);
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
  // Modal konfirmasi khusus aplikasi
  const [confirmDialog, setConfirmDialog] = useState(null);
  const requestConfirmation = (message, action) => {setConfirmDialog({ message, action });};
  const handleConfirmAction = () => {
  const action = confirmDialog?.action;setConfirmDialog(null);if (action) action();};
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const fileInput = { current: null };

  useEffect(() => {
    const online = () => { setIsOnline(true); const now = new Date().toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" }); localStorage.setItem("kas-warga-rt-last-sync", now); setLastSync(now); setNotice("Sinkronisasi otomatis selesai"); setTimeout(() => setNotice(""), 2200); }; const offline = () => setIsOnline(false);
    window.addEventListener("online", online); window.addEventListener("offline", offline);
    return () => { window.removeEventListener("online", online); window.removeEventListener("offline", offline); };
  }, []);
  // Keluar otomatis hanya setelah 15 menit tanpa aktivitas.
  // Jangan menjalankan logout pada pagehide/beforeunload karena event tersebut
  // juga terpanggil saat refresh, pindah halaman, atau WebView memuat ulang app.
  // Sesi login Supabase harus tetap tersimpan agar pengguna tidak terjebak
  // kembali ke halaman login setiap aplikasi dibuka ulang.
  useEffect(() => {
    if (typeof onSignOut !== "function") return undefined;

    let idleTimer;
    let isEffectActive = true;
    const IDLE_TIMEOUT_MS = 15 * 60 * 1000;

    // Reset penanda ketika effect aktif kembali (misalnya StrictMode).
    signOutStarted.current = false;

    const jalankanLogout = async () => {
      if (!isEffectActive || signOutStarted.current) return;

      signOutStarted.current = true;
      window.clearTimeout(idleTimer);

      try {
        await onSignOut();
      } catch (error) {
        console.error("Gagal menjalankan logout otomatis:", error);
        // Izinkan percobaan berikutnya jika proses logout gagal.
        signOutStarted.current = false;
        aturUlangTimer();
      }
    };

    const aturUlangTimer = () => {
      window.clearTimeout(idleTimer);
      if (isEffectActive && !signOutStarted.current) {
        idleTimer = window.setTimeout(jalankanLogout, IDLE_TIMEOUT_MS);
      }
    };

    const aktivitas = ["pointerdown", "keydown", "mousemove", "touchstart", "scroll"];
    aktivitas.forEach((eventName) =>
      window.addEventListener(eventName, aturUlangTimer, { passive: true })
    );

    aturUlangTimer();

    return () => {
      isEffectActive = false;
      window.clearTimeout(idleTimer);
      aktivitas.forEach((eventName) =>
        window.removeEventListener(eventName, aturUlangTimer)
      );
    };
  }, [onSignOut]);

  useEffect(() => {
    setEntries((current) => current.map((item) => item.resident ? item : { ...item, resident: residents.find((r) => r.block === item.block && r.subBlock === item.subBlock)?.name || "Nama belum diisi" }));
  }, []);
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ residents, entries, cashOut, openingBalance, manualAdjustment, balanceHistory, communityIdentity }));
  }, [residents, entries, cashOut, openingBalance, manualAdjustment, balanceHistory, communityIdentity]);

  const total = useMemo(
    () => entries.reduce((sum, e) => sum + Number(e.clean || 0) + Number(e.kas || 0) + Number(e.other || 0), 0),
    [entries]
  );
  const runningBalance = useMemo(
    () => Number(openingBalance || 0) + total - cashOut.reduce((sum, e) => sum + Number(e.amount || 0), 0) + Number(manualAdjustment || 0),
    [openingBalance, total, cashOut, manualAdjustment]
  );
  const backupData = () => {
    const payload = { app: `Kas Warga ${communityIdentityLine}`, version: 1, exportedAt: new Date().toISOString(), residents, entries, cashOut, openingBalance, manualAdjustment, balanceHistory, communityIdentity };
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
    reader.onload = () => { try { const data = JSON.parse(reader.result); if (!Array.isArray(data.residents) || !Array.isArray(data.entries)) throw new Error(); setResidents(data.residents); setEntries(data.entries); if (Number.isFinite(Number(data.openingBalance))) setOpeningBalance(Number(data.openingBalance)); if (Number.isFinite(Number(data.manualAdjustment))) setManualAdjustment(Number(data.manualAdjustment)); if (Array.isArray(data.balanceHistory)) setBalanceHistory(data.balanceHistory); if (data.communityIdentity && typeof data.communityIdentity === "object") setCommunityIdentity({ ...DEFAULT_COMMUNITY_IDENTITY, ...data.communityIdentity }); setNotice("Data berhasil dipulihkan"); } catch { setNotice("File backup tidak valid"); } setTimeout(() => setNotice(""), 2600); event.target.value = ""; };
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
  e?.preventDefault?.();

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
  if (!target) return;
  const originMenu = active;

  requestConfirmation(
    `Apakah Anda yakin ingin menghapus pencatatan keluar "${target.title}"?`,
    () => {
      setCashOut(cashOut.filter((item) => item.id !== id));
      setActive(originMenu);
      setSettingsOpen(false);
      setNotice("Pengeluaran dihapus");
      setTimeout(() => setNotice(""), 2200);
    }
  );
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
  if (!target) return;
  const originMenu = active;

  requestConfirmation(
    `Apakah Anda yakin ingin menghapus pembayaran ${target.resident}?`,
    () => {
      setEntries(entries.filter((item) => item.id !== id));
      setActive(originMenu);
      setSettingsOpen(false);
      setNotice("Pembayaran dihapus");
      setTimeout(() => setNotice(""), 2200);
    }
  );
};
  const startEditResident = (r) => {
    setEditingResident(r.id);
    setResidentName(r.name);
    setResidentBlock(r.block);
    setResidentSubBlock(r.subBlock || `${r.block}1`);
  };
  const deleteResident = (id) => {
  const target = residents.find((r) => r.id === id);
  if (!target) return;
  const originMenu = active;

  requestConfirmation(
    `Apakah Anda yakin ingin menghapus data warga ${target.name}?`,
    () => {
      setResidents(residents.filter((r) => r.id !== id));
      setActive(originMenu);
      setSettingsOpen(false);
      setNotice("Data warga dihapus");
      setTimeout(() => setNotice(""), 2200);
    }
  );
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
/* =========================================================
   NAVY BLUE + WHITE — penyegaran visual seluruh aplikasi
   Hanya mengubah tampilan; tidak mengubah state, data, atau alur fitur.
   ========================================================= */
:root {
  --navy: #173B65;
  --navy-deep: #102A43;
  --navy-mid: #245789;
  --blue: #2563EB;
  --blue-soft: #DBEAFE;
  --page-bg: #F4F7FB;
  --surface: #FFFFFF;
  --border-soft: #DCE5EF;
  --text-main: #1F2937;
  --text-muted: #64748B;
  --positive: #15803D;
  --positive-bg: #DCFCE7;
  --negative: #B91C1C;
  --negative-bg: #FEE2E2;
  font-family: 'Inter', sans-serif;
  color-scheme: light;
}
body, .app-shell, .main { background: var(--page-bg) !important; color: var(--text-main) !important; }
.app-shell, .app-shell *, .login-screen, .login-screen * { font-family: 'Inter', sans-serif !important; }
.sidebar { background: var(--navy-deep) !important; color: #F8FAFC !important; border-color: #244563 !important; box-shadow: 8px 0 28px rgba(16,42,67,.06); }
.brand strong, .brand small, .side-label, .sidebar .nav-item, .sidebar .profile, .sidebar .profile small { color: #EAF2FA !important; }
.brand-mark { background: #FFFFFF !important; color: var(--navy) !important; border-radius: 11px !important; }
.sidebar .nav-item { border: 1px solid transparent !important; border-radius: 10px !important; transition: background .18s ease, color .18s ease !important; }
.sidebar .nav-item:hover, .sidebar .nav-item.active { background: rgba(219,234,254,.14) !important; color: #FFFFFF !important; }
.sidebar .nav-item.active { box-shadow: inset 3px 0 0 #93C5FD !important; }
.sidebar .nav-item.active svg, .brand-mark svg { color: #BFDBFE !important; }
.nav-dot, .connection-status span { background: #93C5FD !important; }
.main { min-width: 0; }
.topbar h1, .section-head h2, .form-intro h2, .report-toolbar h2, .public-heading h2 { color: var(--navy-deep) !important; font-family: 'Inter', sans-serif !important; letter-spacing: -.035em; }
.eyebrow, .side-label, .report-kicker { color: #64748B !important; }
.hero-card { background: linear-gradient(125deg, #102A43 0%, #173B65 58%, #2563A6 100%) !important; color: #FFFFFF !important; border: 1px solid rgba(255,255,255,.12) !important; box-shadow: 0 16px 34px rgba(23,59,101,.14); }
.hero-card .eyebrow.light, .hero-card .hero-note { color: #DBEAFE !important; }
.hero-card .sun { background: rgba(191,219,254,.24) !important; }
.hero-card .h1 { background: rgba(7,25,45,.30) !important; }
.hero-card .h2 { background: rgba(147,197,253,.30) !important; }
.stat-card, .big-card, .table-card, .form-card, .add-resident, .resident-card, .expense-form, .history-card, .balance-settings, .report-paper, .public-breakdown { background: #FFFFFF !important; border: 1px solid var(--border-soft) !important; border-radius: 13px !important; box-shadow: 0 5px 18px rgba(16,42,67,.045); }
.stat-card strong, .big-card strong, .total-row strong, .public-kpi strong { color: var(--navy-deep) !important; font-family: 'Inter', sans-serif !important; }
.stat-card small, .big-card small, .muted, .subcopy, .field-hint, .history-empty { color: var(--text-muted) !important; }
.stat-icon, .notebook-icon, .avatar, .resident-avatar { background: var(--blue-soft) !important; color: var(--navy) !important; }
.green, .flow-in { background: var(--positive-bg) !important; color: var(--positive) !important; }
.red, .flow-out { background: var(--negative-bg) !important; color: var(--negative) !important; }
.gold { background: #FEF3C7 !important; color: #92400E !important; }
.primary-btn, .login-btn { background: var(--navy) !important; color: #FFFFFF !important; border: 1px solid var(--navy) !important; border-radius: 9px !important; box-shadow: 0 4px 10px rgba(23,59,101,.12); }
.primary-btn:hover, .login-btn:hover { background: var(--navy-deep) !important; border-color: var(--navy-deep) !important; }
.text-btn { color: var(--blue) !important; }
input, select, textarea, .month-pill, .search, .input-money { border-color: var(--border-soft) !important; background: #FFFFFF !important; color: var(--text-main) !important; border-radius: 9px !important; }
input:focus, select:focus, textarea:focus, .input-money:focus-within { border-color: #60A5FA !important; box-shadow: 0 0 0 3px rgba(37,99,235,.11) !important; }
.input-money span { background: #F8FAFC !important; color: var(--text-muted) !important; border-color: var(--border-soft) !important; }
.tip, .privacy-note { background: #EFF6FF !important; border: 1px solid #BFDBFE !important; color: var(--navy) !important; border-radius: 11px !important; }
.tip p, .privacy-note p { color: #475569 !important; }
.recap-banner { background: linear-gradient(120deg, #102A43, #173B65) !important; color: #FFFFFF !important; border-radius: 13px !important; }
.recap-banner .eyebrow.light, .recap-banner p { color: #DBEAFE !important; }
.outline-light { color: #FFFFFF !important; border-color: #93C5FD !important; border-radius: 8px !important; }
.outline-light:hover { background: rgba(255,255,255,.12) !important; }
.bar { background: #DBEAFE !important; }
.bar i { background: var(--blue) !important; }
th { background: #F1F5F9 !important; color: #475569 !important; }
th, td { border-bottom-color: #E8EEF5 !important; }
tbody tr:hover { background: #F8FAFC !important; }
.block-tag { background: #EFF6FF !important; color: var(--navy) !important; }
.status-paid { background: var(--positive-bg) !important; color: var(--positive) !important; }
.icon-btn, .cancel-btn, .print-btn { border-color: var(--border-soft) !important; border-radius: 8px !important; }
.icon-btn.edit, .print-btn { color: var(--navy) !important; }
.icon-btn.edit { background: #EFF6FF !important; }
.icon-btn.delete { color: var(--negative) !important; background: var(--negative-bg) !important; }
.cancel-btn { background: #FFFFFF !important; color: #475569 !important; }
.toast { background: var(--navy-deep) !important; color: #FFFFFF !important; }
.login-screen { background: var(--page-bg) !important; }
.login-art { background: linear-gradient(145deg, #102A43 0%, #173B65 60%, #2563A6 100%) !important; }
.login-card { background: #FFFFFF !important; border: 1px solid var(--border-soft) !important; border-radius: 18px !important; box-shadow: 0 18px 48px rgba(16,42,67,.10) !important; }
.login-card h1 { color: var(--navy-deep) !important; }
.login-copy, .login-note { color: var(--text-muted) !important; }
.profile { border-color: #31516F !important; }
.mobile-menu-toggle { background: #FFFFFF !important; color: var(--navy) !important; border-color: var(--border-soft) !important; border-radius: 10px !important; }
.mobile-menu-backdrop { background: rgba(16,42,67,.48) !important; }
.connection-status { color: var(--text-muted) !important; }
.public-kpi { border-radius: 12px !important; }
.public-kpi.income { background: #EFF6FF !important; color: var(--navy) !important; }
.public-kpi.expense { background: #FEF2F2 !important; color: #991B1B !important; }
.public-kpi.balance { background: #E0F2FE !important; color: #075985 !important; }
@media screen and (max-width: 760px) {
  .main { padding: 20px 14px !important; }
  .topbar { margin-bottom: 22px !important; }
  .topbar h1 { font-size: clamp(21px, 6vw, 27px) !important; }
  .hero-card { min-height: 175px !important; padding: 22px !important; }
  .hero-number { font-size: clamp(27px, 8vw, 34px) !important; }
  .stat-grid { gap: 10px !important; margin-bottom: 25px !important; }
  .stat-card, .big-card { padding: 14px !important; }
  .form-card { padding: 17px !important; }
  .section-head { gap: 10px !important; flex-wrap: wrap !important; }
  .table-card { border-radius: 10px !important; }
  .login-card { margin: 14px !important; }
}


        /* =========================================================
           FIX TEMA NAVY BLUE — banner arus kas, laporan, dan cetak PDF
           ========================================================= */
        .recap-banner,
        .recap-banner.cashflow-banner,
        .cashflow-banner,
        .report-heading {
          background: linear-gradient(125deg, #102A43 0%, #173B65 58%, #2563A6 100%) !important;
          color: #FFFFFF !important;
          border-color: #245789 !important;
        }
        .recap-banner h2,
        .recap-banner p,
        .cashflow-banner .eyebrow,
        .cashflow-banner h2,
        .cashflow-banner p,
        .report-heading .report-kicker,
        .report-heading h1,
        .report-heading p {
          color: #FFFFFF !important;
        }
        /* Label LAPORAN PERIODE disamakan dengan teks ringkasan */
        .recap-banner .eyebrow.light,
        .recap-banner p,
        .cashflow-banner p { color: #DBEAFE !important; }
        .report-heading { border-bottom-color: #93C5FD !important; }

        @media print {
          .report-toolbar { border-bottom-color: #173B65 !important; }
          .report-toolbar h2,
          .public-heading h2,
          .recap-banner h2,
          .hero-card h2,
          .report-heading h1,
          .section-head h2,
          .report-paper h3,
          .public-breakdown h2,
          .big-card strong,
          .public-kpi strong,
          .report-stats b,
          .report-footer strong {
            color: #102A43 !important;
          }
          .recap-banner,
          .hero-card,
          .report-heading,
          .public-heading {
            color: #102A43 !important;
            background: #EFF6FF !important;
            border: 1px solid #BFDBFE !important;
          }
          .recap-banner p,
          .recap-banner .eyebrow.light,
          .hero-card p,
          .report-heading p,
          .report-toolbar .subcopy,
          .public-heading p { color: #475569 !important; }
          .eyebrow,
          .report-kicker { color: #245789 !important; }
          .report-heading { border-bottom: 2px solid #173B65 !important; }
          .report-heading .report-kicker,
          .report-heading h1,
          .report-heading p { color: #102A43 !important; }
          .report-stats div { background: #EFF6FF !important; }
          .section-head { border-bottom-color: #DCE5EF !important; }
          .report-paper h3 { border-bottom-color: #DCE5EF !important; }
          th, .table-card th {
            color: #173B65 !important;
            background: #EFF6FF !important;
          }
          tbody tr:nth-child(even) { background: #F8FAFC !important; }
          .status-paid, .block-tag {
            color: #166534 !important;
            background: #DCFCE7 !important;
            border-color: #BBF7D0 !important;
          }
          .report-footer { border-top-color: #DCE5EF !important; }
        }


        /* Navigasi mobile melayang */
        .mobile-floating-nav { display: none; }
        @media screen and (max-width: 800px) {
          .app-shell { padding-bottom: 104px !important; }
          .mobile-floating-nav {
            display: flex !important; position: fixed !important;
            left: 12px !important; right: 12px !important;
            bottom: max(12px, env(safe-area-inset-bottom)) !important;
            z-index: 1200 !important; min-height: 66px;
            align-items: center; justify-content: space-around; gap: 3px;
            padding: 7px 6px; border: 1px solid rgba(220,229,239,.95);
            border-radius: 20px; background: rgba(255,255,255,.96);
            box-shadow: 0 10px 32px rgba(16,42,67,.18);
            backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
          }
          .mobile-floating-nav button {
            appearance: none; border: 0; background: transparent; color: #64748B;
            display: flex; flex: 1 1 0; min-width: 0; min-height: 50px;
            flex-direction: column; align-items: center; justify-content: center;
            gap: 4px; padding: 4px 2px; border-radius: 13px;
            font: 600 10px/1.1 'Inter', sans-serif; cursor: pointer;
            -webkit-tap-highlight-color: transparent;
          }
          .mobile-floating-nav button svg { flex: 0 0 auto; }
          .mobile-floating-nav button.active { color: #173B65; background: #EAF2FC; }
          .mobile-floating-nav button.active svg { stroke-width: 2.5; }
          .main { padding-bottom: 24px !important; }
          .connection-status { bottom: 91px !important; right: 14px !important; max-width: calc(100% - 28px); }
        }
        @media print { .mobile-floating-nav { display: none !important; } .app-shell { padding-bottom: 0 !important; } }

        /* Pintasan halaman tambahan di beranda */
        .home-shortcuts { margin: 22px 0 26px; }
        .home-shortcuts-heading { display:flex; align-items:center; justify-content:space-between; margin-bottom:12px; }
        .home-shortcuts-heading h2 { margin:4px 0 0; font-size:18px; color:#102A43; }
        .home-shortcuts-grid { display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:10px; }
        .home-shortcut {
          display:flex; align-items:center; gap:11px; min-width:0; text-align:left;
          padding:13px 12px; border:1px solid #DCE5EF; border-radius:15px;
          background:#FFFFFF; color:#102A43; cursor:pointer; transition:transform .18s ease, box-shadow .18s ease;
        }
        .home-shortcut:hover { transform:translateY(-2px); box-shadow:0 8px 22px rgba(16,42,67,.08); }
        .home-shortcut-icon {
          width:40px; height:40px; flex:0 0 40px; display:grid; place-items:center;
          border-radius:12px; background:#EAF2FC; color:#173B65;
        }
        .home-shortcut b { display:block; font-size:12px; line-height:1.3; }
        .home-shortcut small { display:block; margin-top:3px; color:#64748B; font-size:10px; line-height:1.3; }
        .home-shortcut-logout .home-shortcut-icon { background:#FEF2F2; color:#B91C1C; }
        .home-shortcut-logout b { color:#991B1B; }

        @media screen and (max-width:800px) {
          /* Pada mobile, navigasi samping tidak digunakan lagi */
          .sidebar, .mobile-menu-backdrop, .mobile-menu-toggle { display:none !important; }
          .app-shell { padding-bottom:100px !important; }
          .main { margin-left:0 !important; width:100% !important; }
          .topbar { padding-left:18px !important; padding-right:18px !important; }
          .home-shortcuts { margin:17px 0 22px; }
          .home-shortcuts-heading { margin-bottom:9px; }
          .home-shortcuts-heading h2 { font-size:16px; }
          .home-shortcuts-grid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:9px; }
          .home-shortcut { min-height:66px; padding:10px; gap:9px; border-radius:13px; }
          .home-shortcut-icon { width:36px; height:36px; flex-basis:36px; border-radius:11px; }
          .home-shortcut b { font-size:11px; }
          .home-shortcut small { font-size:9px; }
          .hero-card { padding: 12px 15px !important; min-height: 0 !important; border-radius: 13px !important; }
          .hero-card .hero-number { font-size: clamp(23px, 6vw, 29px) !important; line-height: 1.2 !important; margin: 5px 0 !important; }
          .hero-card .hero-note { font-size: 11px !important; }
          /* Rincian iuran pada kartu Saldo Kas Berjalan */
          .hero-income-details { display: flex; flex-direction: column; gap: 7px; margin-top: 14px; padding-top: 11px; border-top: 1px solid rgba(255,255,255,.22); font-family: inherit; font-size: 12px; font-weight: 400; line-height: 1.5; color: #DBEAFE; }
          .hero-income-details > div { display: flex; align-items: center; justify-content: space-between; gap: 16px; }
          .hero-income-details span,
          .hero-income-details b { font-family: inherit; font-size: inherit; font-weight: inherit; color: inherit; }
          .hero-income-details b { text-align: right; white-space: nowrap; }
          .hero-art { display:none !important; }
          .stat-grid { gap:9px !important; }
          .stat-card { padding:12px !important; }
          .stat-card strong { font-size:17px !important; }
          .mobile-floating-nav button { flex:1 1 0; }
        }
          .recap-income-details { display: flex; flex-wrap: wrap; gap: 4px 12px; margin: 7px 0 10px; }
          .recap-income-details span { color: #64748B; font-size: 10px; line-height: 1.5; }
        @media screen and (max-width:800px) {
          .recap-income-details { gap: 3px 9px; }
          .recap-income-details span { font-size: 9px; }}
        @media screen and (max-width:380px) {
          .home-shortcut { gap:7px; padding:9px 8px; }
          .home-shortcut-icon { width:32px; height:32px; flex-basis:32px; }
          .home-shortcut b { font-size:10px; }
          .home-shortcut small { font-size:9px; }
        }
        @media print {
          .home-shortcuts, .mobile-floating-nav { display:none !important; }
        }
        @media screen and (max-width:800px) {
          .hero-income-details { gap: 6px; margin-top: 11px; padding-top: 9px; font-size: 11px; }
        }
      /* Rincian iuran ringkas pada kartu rekap */
.recap-income-details {
  display: flex;
  flex-direction: column;
  gap: 7px;
  margin-top: 15px;
  padding-top: 12px;
  border-top: 1px solid #e2e8f0;
  font-size: 12px;
  line-height: 1.5;
  color: #64748b;
}

.recap-income-details > div {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.recap-income-details span {
  font-size: 12px;
  font-weight: 400;
}

.recap-income-details b {
  color: #173b65;
  font-size: 12px;
  font-weight: 600;
  white-space: nowrap;
}

/* Ukuran kartu rekap di HP */
@media screen and (max-width: 760px) {
  .recap-grid {
    grid-template-columns: 1fr !important;
    gap: 10px !important;
  }

  .recap-grid .big-card {
    padding: 14px !important;
  }

  .recap-grid .big-card strong {
    font-size: 23px !important;
    margin: 7px 0 10px !important;
  }

  .recap-income-details {
    margin-top: 12px;
    padding-top: 10px;
    gap: 5px;
  }
}
/* =====================================
   PERAPIAN PILIH PERIODE RINGKASAN PUBLIK
   ===================================== */

.public-heading {
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  gap: 20px;
}

.public-controls {
  display: flex;
  align-items: flex-end;
  gap: 12px;
  flex-wrap: wrap;
}

.public-controls label {
  display: flex;
  flex-direction: column;
  gap: 8px;
  min-width: 195px;
  font-size: 13px;
  font-weight: 600;
  color: #475569;
}

.public-controls select {
  width: 100%;
  min-width: 195px;
  min-height: 48px;
  padding: 0 14px;
  font-size: 15px;
  font-weight: 500;
  border: 1px solid #DCE5EF;
  border-radius: 10px;
  background: #FFFFFF;
  color: #102A43;
  cursor: pointer;
}

.public-controls .primary-btn {
  min-height: 48px;
  padding: 0 18px;
  white-space: nowrap;
}

/* Tampilan HP */
@media screen and (max-width: 760px) {
  .public-heading {
    display: flex;
    flex-direction: column;
    align-items: stretch;
    gap: 18px;
  }

  .public-controls {
    display: flex;
    flex-direction: column;
    align-items: stretch;
    width: 100%;
    gap: 12px;
  }

  .public-controls label {
    width: 100%;
    min-width: 0;
    font-size: 13px;
  }

  .public-controls select {
    width: 100%;
    min-width: 0;
    min-height: 48px;
    font-size: 15px;
  }

  .public-controls .primary-btn {
    width: 100%;
    min-height: 46px;
    justify-content: center;
  }
}


        /* CSS catatan pembayaran dan popup detail */
/* Beranda mobile: catatan pembayaran ringkas tanpa scroll horizontal */
@media screen and (max-width: 760px) {
  .recent-payment-card { overflow: visible !important; }
  .recent-payment-table { width: 100% !important; min-width: 0 !important; table-layout: fixed !important; }
  .recent-payment-table th:nth-child(4),
  .recent-payment-table td:nth-child(4),
  .recent-payment-table th:nth-child(5),
  .recent-payment-table td:nth-child(5),
  .recent-payment-table th:nth-child(6),
  .recent-payment-table td:nth-child(6),
  .recent-payment-table th:nth-child(8),
  .recent-payment-table td:nth-child(8) { display: none !important; }
  .recent-payment-table th:nth-child(1),
  .recent-payment-table td:nth-child(1) { width: 22%; }
  .recent-payment-table th:nth-child(2),
  .recent-payment-table td:nth-child(2) { width: 34%; }
  .recent-payment-table th:nth-child(3),
  .recent-payment-table td:nth-child(3) { width: 18%; }
  .recent-payment-table th:nth-child(7),
  .recent-payment-table td:nth-child(7) { width: 26%; }
  .recent-payment-table td { padding: 11px 7px !important; font-size: 11px !important; white-space: normal !important; overflow-wrap: anywhere !important; }
  .recent-payment-table th { padding: 9px 7px !important; font-size: 10px !important; }
  .recent-payment-row { cursor: pointer; }
  .recent-payment-row:active { background: #EFF6FF !important; }
  .recent-payment-table td:nth-child(2) b { display: block; line-height: 1.35; }
}
.payment-detail-backdrop {
  position: fixed; inset: 0; z-index: 9999; display: flex;
  align-items: center; justify-content: center; padding: 18px;
  background: rgba(15, 32, 52, .55); backdrop-filter: blur(3px);
}
.payment-detail-modal {
  position: relative; width: min(100%, 430px); max-height: 90vh; overflow-y: auto;
  padding: 24px; border: 1px solid #DCE5EF; border-radius: 18px;
  background: #FFFFFF; color: #1F2937; box-shadow: 0 22px 65px rgba(16,42,67,.22);
}
.payment-detail-modal h2 { margin: 5px 0 18px; color: #102A43; font-size: 21px; }
.payment-detail-close {
  position: absolute; top: 14px; right: 14px; display: grid; place-items: center;
  width: 34px; height: 34px; border: 1px solid #DCE5EF; border-radius: 9px;
  background: #F8FAFC; color: #475569; cursor: pointer;
}
.payment-detail-person {
  display: flex; justify-content: space-between; align-items: center; gap: 10px;
  padding: 13px; margin-bottom: 13px; border-radius: 11px; background: #EFF6FF;
}
.payment-detail-person b { color: #102A43; }
.payment-detail-person span { color: #173B65; font-size: 12px; font-weight: 700; }
.payment-detail-lines { display: grid; gap: 0; }
.payment-detail-lines div {
  display: flex; justify-content: space-between; gap: 12px; padding: 10px 2px;
  border-bottom: 1px solid #E8EEF5; font-size: 13px;
}
.payment-detail-lines span { color: #64748B; }
.payment-detail-lines b { text-align: right; color: #1F2937; }
.payment-detail-total {
  display: flex; justify-content: space-between; align-items: center; gap: 12px;
  margin-top: 14px; padding: 14px; border-radius: 11px; background: #F1F5F9;
}
.payment-detail-total strong { color: #102A43; font-size: 18px; }
.payment-detail-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; }
.payment-detail-actions button { display: inline-flex; align-items: center; justify-content: center; gap: 6px; }
@media screen and (max-width: 480px) {
  .payment-detail-modal { padding: 20px 16px; border-radius: 15px; }
  .payment-detail-actions {
    display: grid;
    grid-template-columns: 1fr 1fr 1.2fr;
    gap: 8px;
  }
  .payment-detail-actions button {
    box-sizing: border-box;
    width: 100%;
    min-width: 0;
    min-height: 40px;
    padding: 10px 9px;
    font-size: 12px;
    line-height: 1.2;
    white-space: nowrap;
  }
  .payment-detail-actions .icon-btn.delete {
    width: 100%;
    height: auto;
    min-height: 40px;
    padding: 10px 9px;
    border-radius: 9px;
    gap: 6px;
    font-size: 12px;
  }
}


/* Rekap bulanan per warga: baris responsif dan popup hanya-baca */
.recap-payment-row { cursor: pointer; }
.recap-payment-row:active { background: #EFF6FF !important; }
.recap-payment-actions { display: flex; justify-content: flex-end; }
@media screen and (max-width: 760px) {
  .recap-payment-card { overflow: visible !important; }
  .recap-payment-table { width: 100% !important; min-width: 0 !important; table-layout: fixed !important; }
  .recap-payment-table th:nth-child(3), .recap-payment-table td:nth-child(3),
  .recap-payment-table th:nth-child(5), .recap-payment-table td:nth-child(5),
  .recap-payment-table th:nth-child(6), .recap-payment-table td:nth-child(6),
  .recap-payment-table th:nth-child(7), .recap-payment-table td:nth-child(7) { display: none !important; }
  .recap-payment-table th:nth-child(1), .recap-payment-table td:nth-child(1) { width: 18%; }
  .recap-payment-table th:nth-child(2), .recap-payment-table td:nth-child(2) { width: 28%; }
  .recap-payment-table th:nth-child(4), .recap-payment-table td:nth-child(4) { width: 16%; }
  .recap-payment-table th:nth-child(8), .recap-payment-table td:nth-child(8) { width: 23%; }
  .recap-payment-table th:nth-child(9), .recap-payment-table td:nth-child(9) { width: 15%; }
  .recap-payment-table td { padding: 11px 6px !important; font-size: 10px !important; white-space: normal !important; overflow-wrap: anywhere !important; }
  .recap-payment-table th { padding: 9px 6px !important; font-size: 9px !important; }
  .recap-payment-table td:nth-child(2) b { display: block; line-height: 1.35; }
  .recap-payment-table .status-paid { font-size: 9px; padding: 4px 5px; white-space: nowrap; }
}
/* Toolbar rincian iuran */
.recap-detail-head {
  align-items: flex-end;
  gap: 14px;
}

.recap-detail-toolbar {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.recap-resident-search {
  min-width: 210px;
}

.recap-status-filter {
  min-height: 40px;
  padding: 0 12px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: var(--white);
  color: var(--text);
  font: inherit;
  font-size: 13px;
}

.recap-detail-export {
  min-height: 40px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 0 12px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: var(--white);
  color: var(--dark);
  font: inherit;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
}

/* Status belum lunas */
.status-unpaid {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 4px 7px;
  border-radius: 20px;
  background: #FEE2E2;
  color: #B91C1C;
  font-size: 11px;
  font-weight: 700;
  white-space: nowrap;
}

.recap-unpaid-row {
  background: #FFFAFA;
}

/* Tampilan mobile */
@media screen and (max-width: 760px) {

  .recap-detail-head {
    display: block;
  }

  .recap-detail-toolbar {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(112px, .85fr);
    gap: 7px;
    margin-top: 12px;
    width: 100%;
  }

  .recap-resident-search {
    min-width: 0;
    width: 100%;
    grid-column: 1 / -1;
  }

  .recap-resident-search input {
    min-width: 0;
    width: 100%;
    font-size: 12px;
  }

  .recap-status-filter {
    width: 100%;
    min-width: 0;
    min-height: 38px;
    padding: 0 8px;
    font-size: 11px;
  }

  .recap-detail-export {
    width: 100%;
    min-height: 36px;
    padding: 0 7px;
    font-size: 11px;
  }
}
/* Samakan ukuran status Belum Lunas dengan Lunas di mobile */
@media screen and (max-width: 760px) {
  .status-unpaid {
    display: inline-flex !important;
    align-items: center;
    justify-content: center;
    width: auto !important;
    min-width: 0 !important;
    max-width: 100%;
    padding: 3px 7px !important;
    font-size: 11px !important;
    line-height: 1.3 !important;
    font-weight: 700;
    border-radius: 20px;
    white-space: nowrap;
    box-sizing: border-box;
  }
}
/* Samakan ukuran label Lunas dan Belum Lunas */
.recap-payment-table .status-paid,
.recap-payment-table .status-unpaid {
  display: inline-flex !important;
  align-items: center;
  justify-content: center;
  width: auto !important;
  min-width: 0 !important;
  max-width: none !important;
  padding: 3px 7px !important;
  font-size: 11px !important;
  line-height: 1.3 !important;
  font-weight: 600 !important;
  border-radius: 20px !important;
  white-space: nowrap !important;
  box-sizing: border-box;
}

/* Warna status tetap berbeda */
.recap-payment-table .status-unpaid {
  background: #FEE2E2 !important;
  color: #B91C1C !important;
}

/* Perbaikan khusus mobile */
@media screen and (max-width: 760px) {
  .recap-payment-table td[data-label="Status"] {
    text-align: center !important;
    padding-left: 3px !important;
    padding-right: 3px !important;
  }

  .recap-payment-table .status-paid,
  .recap-payment-table .status-unpaid {
    font-size: 10px !important;
    padding: 3px 5px !important;
  }

  /* Filter status lebih ringkas */
  .recap-detail-toolbar {
    display: grid !important;
    grid-template-columns: 1fr 140px !important;
    gap: 7px !important;
    align-items: center;
  }

  .recap-resident-search {
    grid-column: 1 / -1;
  }

  .recap-status-filter {
    width: 140px !important;
    max-width: 140px !important;
    min-width: 0 !important;
    height: 38px !important;
    min-height: 38px !important;
    padding: 0 8px !important;
    font-size: 11px !important;
    border-radius: 9px !important;
    box-shadow: none !important;
  }
}
/* Perbaikan lebar toolbar Rincian Iuran mobile */
@media screen and (max-width: 760px) {

  .recap-payment-card,
  .recap-detail-head,
  .recap-detail-toolbar {
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
  }

  .recap-detail-head {
    display: flex !important;
    flex-direction: column !important;
    align-items: stretch !important;
    overflow: hidden;
  }

  .recap-detail-toolbar {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) minmax(0, 125px) !important;
    gap: 7px !important;
    padding: 0 !important;
    margin: 10px 0 0 !important;
  }

  /* Kolom pencarian */
  .recap-resident-search {
    grid-column: 1 / -1 !important;
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
  }

  .recap-resident-search input {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
  }

  /* Filter status */
  .recap-detail-toolbar .recap-status-filter {
    grid-column: 1 / 2 !important;
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    height: 36px !important;
    min-height: 36px !important;
    padding: 0 7px !important;
    margin: 0 !important;
    font-size: 11px !important;
    box-sizing: border-box !important;
  }

  /* Mencegah elemen anak memperlebar kartu */
  .recap-detail-toolbar > * {
    min-width: 0 !important;
    max-width: 100% !important;
  }

  .recap-detail-toolbar select {
    overflow: hidden;
    text-overflow: ellipsis;
  }

  /* Kartu tabel tidak melewati layar */
  .recap-payment-card {
    overflow-x: auto !important;
  }
}

/* Final mobile sizing: status badge must fit entirely inside its table cell */
.recap-payment-table .status-paid,
.recap-payment-table .status-unpaid,
.recap-payment-modal .status-paid,
.recap-payment-modal .status-unpaid {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  width: auto !important;
  min-width: 0 !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  padding: 3px 6px !important;
  border-radius: 999px !important;
  font-size: 10px !important;
  line-height: 1.2 !important;
  font-weight: 600 !important;
  white-space: normal !important;
  text-align: center !important;
  overflow-wrap: anywhere !important;
}

.recap-payment-table .status-unpaid,
.recap-payment-modal .status-unpaid {
  background: #FEE2E2 !important;
  color: #B91C1C !important;
}

@media screen and (max-width: 760px) {
  .recap-payment-card { width: 100% !important; max-width: 100% !important; overflow: hidden !important; }
  .recap-payment-table { width: 100% !important; max-width: 100% !important; min-width: 0 !important; table-layout: fixed !important; }
  .recap-payment-table th:nth-child(9),
  .recap-payment-table td:nth-child(9) { width: 17% !important; padding-left: 2px !important; padding-right: 2px !important; text-align: center !important; }
  .recap-payment-table th:nth-child(1), .recap-payment-table td:nth-child(1) { width: 16% !important; }
  .recap-payment-table th:nth-child(2), .recap-payment-table td:nth-child(2) { width: 29% !important; }
  .recap-payment-table th:nth-child(4), .recap-payment-table td:nth-child(4) { width: 15% !important; }
  .recap-payment-table th:nth-child(8), .recap-payment-table td:nth-child(8) { width: 23% !important; }
  .recap-payment-table .status-paid,
  .recap-payment-table .status-unpaid { font-size: 8px !important; padding: 2px 3px !important; letter-spacing: -.2px !important; }
  .recap-detail-head, .recap-detail-toolbar { width: 100% !important; max-width: 100% !important; min-width: 0 !important; box-sizing: border-box !important; }
  .recap-detail-toolbar { grid-template-columns: minmax(0, 1fr) 112px !important; }
  .recap-detail-toolbar .recap-status-filter { width: 112px !important; max-width: 112px !important; min-width: 0 !important; height: 34px !important; min-height: 34px !important; padding: 0 5px !important; font-size: 10px !important; }
  .recap-payment-modal .payment-detail-lines > div { gap: 12px; }
  .recap-payment-modal .payment-detail-lines b { text-align: right; }
}
/* Tombol aksi popup rekap iuran */
.recap-payment-actions {
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 9px;
}

.recap-payment-actions button {
  min-height: 40px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 9px 14px;
  border-radius: 9px;
  font-size: 12px;
  font-weight: 600;
}

/* Khusus tampilan mobile */
@media screen and (max-width: 480px) {

  .recap-payment-modal {
    width: 100%;
    max-width: 390px;
    padding: 20px 15px;
    box-sizing: border-box;
  }

  .recap-payment-actions {
    display: grid !important;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 7px;
    width: 100%;
  }

  .recap-payment-actions button {
    width: 100% !important;
    min-width: 0 !important;
    min-height: 39px;
    padding: 8px 4px !important;
    font-size: 11px !important;
    white-space: nowrap;
    gap: 4px;
    box-sizing: border-box;
  }

  .recap-payment-actions button svg {
    width: 14px;
    height: 14px;
    flex-shrink: 0;
  }

  .recap-payment-actions .icon-btn.delete {
    background: #FEE2E2 !important;
    color: #B91C1C !important;
    border: 1px solid #FECACA !important;
  }

}

/* Dropdown popup kustom — tema Navy Blue + Putih */
.themed-dropdown {
  position: relative;
  display: block;
  width: 100%;
  min-width: 0;
  color: #102A43;
  z-index: 20;
}
.themed-dropdown-trigger {
  width: 100% !important;
  min-width: 0 !important;
  min-height: 44px !important;
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 10px !important;
  padding: 0 13px !important;
  border: 1px solid #C9D8E8 !important;
  border-radius: 10px !important;
  background: #FFFFFF !important;
  color: #102A43 !important;
  box-shadow: 0 1px 3px rgba(16,42,67,.05) !important;
  font: 600 13px/1.25 Inter, sans-serif !important;
  text-align: left !important;
  cursor: pointer;
}
.themed-dropdown-trigger:focus-visible {
  outline: none !important;
  border-color: #2563A6 !important;
  box-shadow: 0 0 0 3px rgba(37,99,166,.14) !important;
}
.themed-dropdown-trigger:disabled {
  background: #F1F5F9 !important;
  color: #94A3B8 !important;
  border-color: #E2E8F0 !important;
  cursor: not-allowed;
}
.themed-dropdown-value {
  display: block;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.themed-dropdown-trigger > svg {
  flex: 0 0 auto;
  color: #173B65;
  transition: transform .18s ease;
}
.themed-dropdown.is-open { z-index: 100001; }
.themed-dropdown.is-open .themed-dropdown-trigger > svg { transform: rotate(180deg); }
.themed-dropdown-menu {
  position: absolute;
  top: calc(100% + 7px);
  left: 0;
  right: 0;
  z-index: 100002;
  max-height: min(310px, 52vh);
  overflow-y: auto;
  overscroll-behavior: contain;
  padding: 6px;
  border: 1px solid #DCE5EF;
  border-radius: 12px;
  background: #FFFFFF;
  box-shadow: 0 16px 38px rgba(16,42,67,.18);
  animation: themedDropdownIn .14s ease-out;
}
@keyframes themedDropdownIn {
  from { opacity: 0; transform: translateY(-4px); }
  to { opacity: 1; transform: translateY(0); }
}
.themed-dropdown-option {
  width: 100%;
  min-height: 42px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 9px 11px;
  border: 0;
  border-radius: 8px;
  background: transparent;
  color: #334155;
  font: 500 13px/1.3 Inter, sans-serif;
  text-align: left;
  cursor: pointer;
}
.themed-dropdown-option:hover,
.themed-dropdown-option:focus-visible {
  outline: none;
  background: #F1F5F9;
  color: #102A43;
}
.themed-dropdown-option.selected {
  background: #EFF6FF;
  color: #173B65;
  font-weight: 700;
}
.themed-dropdown-option.selected svg { color: #2563A6; flex: 0 0 auto; }
.recap-detail-head,
.recap-detail-toolbar { overflow: visible !important; }
.recap-detail-toolbar .themed-dropdown { min-width: 0; }
.recap-detail-toolbar .recap-status-filter .themed-dropdown-trigger,
.recap-detail-toolbar .recap-block-filter .themed-dropdown-trigger {
  min-height: 43px !important;
  font-size: 13px !important;
}
@media screen and (max-width: 760px) {
  .themed-dropdown-trigger { min-height: 46px !important; font-size: 14px !important; padding: 0 11px !important; }
  .themed-dropdown-menu { max-height: min(270px, 46vh); padding: 6px; border-radius: 11px; }
  .themed-dropdown-option { min-height: 44px; padding: 10px 11px; font-size: 14px; }
  .recap-detail-toolbar {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
    align-items: center !important;
    gap: 8px !important;
  }
  .recap-detail-toolbar .recap-resident-search {
    grid-column: 1 / -1 !important;
    width: 100% !important;
  }
  .recap-detail-toolbar .recap-status-filter,
  .recap-detail-toolbar .recap-block-filter {
    grid-column: auto !important;
    width: 100% !important;
    max-width: none !important;
    min-width: 0 !important;
  }
  .recap-detail-toolbar .recap-status-filter .themed-dropdown-trigger,
  .recap-detail-toolbar .recap-block-filter .themed-dropdown-trigger {
    min-height: 43px !important;
    padding: 0 9px !important;
    font-size: 13px !important;
  }
}

/* =====================================================
   FIX LAYOUT FILTER: tidak saling menumpuk di desktop/HP
   ===================================================== */
.resident-toolbar {
  display: grid !important;
  grid-template-columns: minmax(180px, 1fr) minmax(125px, 150px) minmax(145px, 170px) auto auto !important;
  align-items: center !important;
  gap: 9px !important;
  width: 100% !important;
  min-width: 0 !important;
}
.resident-toolbar > * {
  min-width: 0 !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}
.resident-toolbar .resident-search {
  grid-column: auto !important;
  width: 100% !important;
  min-width: 0 !important;
}
.resident-toolbar .resident-filter {
  position: relative !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  flex: initial !important;
}
.resident-toolbar .resident-filter .themed-dropdown-trigger {
  width: 100% !important;
  min-width: 0 !important;
  min-height: 42px !important;
  padding-left: 10px !important;
  padding-right: 9px !important;
  font-size: 12px !important;
}
.resident-toolbar .print-btn {
  justify-self: start !important;
  white-space: nowrap !important;
}
.resident-toolbar .result-count {
  justify-self: end !important;
  white-space: nowrap !important;
}

/* Filter bulan tetap satu kolom dan tidak bertumpuk dengan tombol ekspor. */
.recap-banner,
.recap-actions,
.recap-period-label {
  min-width: 0 !important;
}
.recap-actions {
  display: flex !important;
  align-items: flex-end !important;
  flex-wrap: wrap !important;
  gap: 9px !important;
}
.recap-actions .recap-period-label {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) !important;
  flex: 1 1 160px !important;
  min-width: 0 !important;
}
.recap-actions .recap-period-select {
  display: block !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
}
.recap-actions > button {
  flex: 0 1 auto !important;
  min-width: 0 !important;
}

@media screen and (max-width: 760px) {
  .resident-toolbar {
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
    gap: 8px !important;
    margin-top: 14px !important;
  }
  .resident-toolbar .resident-search {
    grid-column: 1 / -1 !important;
  }
  .resident-toolbar .resident-block-filter,
  .resident-toolbar .resident-subblock-filter {
    grid-column: auto !important;
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
  }
  .resident-toolbar .resident-filter .themed-dropdown-trigger {
    min-height: 42px !important;
    padding: 0 8px !important;
    font-size: 12px !important;
  }
  .resident-toolbar .print-btn {
    grid-column: 1 / 2 !important;
    width: 100% !important;
    justify-content: center !important;
  }
  .resident-toolbar .result-count {
    grid-column: 2 / 3 !important;
    justify-self: end !important;
  }

  .recap-actions {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) !important;
    width: 100% !important;
    gap: 9px !important;
  }
  .recap-actions .recap-period-label {
    display: grid !important;
    grid-column: 1 / -1 !important;
    width: 100% !important;
    min-width: 0 !important;
  }
  .recap-actions .recap-period-select {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
  }
  .recap-actions > button {
    width: 100% !important;
    min-width: 0 !important;
    justify-content: center !important;
  }
}


/* FINAL FIX: ukuran filter Status dan Blok pada Rekap Bulanan */
.recap-detail-toolbar {
  display: grid !important;
  grid-template-columns: minmax(190px, 1.35fr) minmax(135px, 1fr) minmax(125px, 1fr) !important;
  align-items: stretch !important;
  gap: 8px !important;
  width: 100% !important;
  min-width: 0 !important;
  box-sizing: border-box !important;
}
.recap-detail-toolbar .recap-resident-search,
.recap-detail-toolbar > .recap-status-filter,
.recap-detail-toolbar > .recap-block-filter {
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  height: 43px !important;
  min-height: 43px !important;
  max-height: 43px !important;
  margin: 0 !important;
  box-sizing: border-box !important;
}
.recap-detail-toolbar > .recap-status-filter,
.recap-detail-toolbar > .recap-block-filter {
  padding: 0 !important;
  border: 0 !important;
  background: transparent !important;
}
.recap-detail-toolbar .recap-status-filter .themed-dropdown-trigger,
.recap-detail-toolbar .recap-block-filter .themed-dropdown-trigger {
  display: flex !important;
  width: 100% !important;
  min-width: 0 !important;
  height: 43px !important;
  min-height: 43px !important;
  max-height: 43px !important;
  padding: 0 10px !important;
  box-sizing: border-box !important;
  font-size: 13px !important;
  line-height: 1.2 !important;
}
.recap-detail-toolbar .recap-status-filter .themed-dropdown-value,
.recap-detail-toolbar .recap-block-filter .themed-dropdown-value {
  min-width: 0 !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
}
@media screen and (max-width: 760px) {
  .recap-detail-toolbar {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
    align-items: stretch !important;
    gap: 8px !important;
    width: 100% !important;
    margin-top: 10px !important;
  }
  .recap-detail-toolbar .recap-resident-search {
    grid-column: 1 / -1 !important;
    width: 100% !important;
    height: 42px !important;
    min-height: 42px !important;
    max-height: 42px !important;
  }
  .recap-detail-toolbar > .recap-status-filter,
  .recap-detail-toolbar > .recap-block-filter {
    grid-column: auto !important;
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    height: 42px !important;
    min-height: 42px !important;
    max-height: 42px !important;
  }
  .recap-detail-toolbar .recap-status-filter .themed-dropdown-trigger,
  .recap-detail-toolbar .recap-block-filter .themed-dropdown-trigger {
    width: 100% !important;
    min-width: 0 !important;
    height: 42px !important;
    min-height: 42px !important;
    max-height: 42px !important;
    padding: 0 8px !important;
    font-size: 12px !important;
  }
  .recap-detail-toolbar .recap-status-filter .themed-dropdown-trigger > svg,
  .recap-detail-toolbar .recap-block-filter .themed-dropdown-trigger > svg {
    width: 15px !important;
    height: 15px !important;
    margin-left: 5px !important;
  }
}



/* =========================================================
   FINAL MOBILE FILTER FIX — warga dan rekap bulanan
   Menetralkan style filter lama agar tidak muncul bingkai/
   panah ganda dan memastikan setiap kontrol berada di kolomnya.
   ========================================================= */
.resident-toolbar {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
  align-items: center !important;
  gap: 9px !important;
  width: 100% !important;
  min-width: 0 !important;
  box-sizing: border-box !important;
}
.resident-toolbar > * {
  min-width: 0 !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}
.resident-toolbar .resident-search {
  grid-column: 1 / -1 !important;
  width: 100% !important;
  min-width: 0 !important;
}
.resident-toolbar .resident-filter {
  display: block !important;
  position: relative !important;
  grid-column: auto !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  min-height: 0 !important;
  height: auto !important;
  padding: 0 !important;
  border: 0 !important;
  background: transparent !important;
  box-shadow: none !important;
}
.resident-toolbar .resident-filter::before,
.resident-toolbar .resident-filter::after {
  content: none !important;
  display: none !important;
}
.resident-toolbar .resident-filter .themed-dropdown-trigger {
  appearance: none !important;
  -webkit-appearance: none !important;
  display: flex !important;
  position: relative !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  height: 44px !important;
  min-height: 44px !important;
  max-height: 44px !important;
  margin: 0 !important;
  padding: 0 10px !important;
  box-sizing: border-box !important;
  border: 1px solid #C9D8E8 !important;
  border-radius: 10px !important;
  background: #FFFFFF !important;
  box-shadow: 0 1px 3px rgba(16,42,67,.05) !important;
  font-size: 13px !important;
}
.resident-toolbar .resident-filter .themed-dropdown-trigger::before,
.resident-toolbar .resident-filter .themed-dropdown-trigger::after {
  content: none !important;
  display: none !important;
}
.resident-toolbar .resident-filter .themed-dropdown-value {
  flex: 1 1 auto !important;
  min-width: 0 !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
}
.resident-toolbar .resident-filter .themed-dropdown-trigger > svg {
  display: block !important;
  flex: 0 0 16px !important;
  width: 16px !important;
  height: 16px !important;
  margin: 0 0 0 5px !important;
}
.resident-toolbar .print-btn {
  grid-column: 1 / 2 !important;
  justify-self: stretch !important;
  width: 100% !important;
  min-width: 0 !important;
}
.resident-toolbar .result-count {
  grid-column: 2 / 3 !important;
  justify-self: end !important;
  white-space: nowrap !important;
}

/* Filter bulan dan tombol ekspor pada banner Rekap Bulanan */
.recap-page .recap-banner {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) !important;
  align-items: start !important;
  gap: 14px !important;
  width: 100% !important;
  min-width: 0 !important;
  box-sizing: border-box !important;
}
.recap-page .recap-banner > * {
  min-width: 0 !important;
  max-width: 100% !important;
}
.recap-page .recap-actions {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) !important;
  align-items: stretch !important;
  gap: 9px !important;
  width: 100% !important;
  min-width: 0 !important;
  margin: 0 !important;
}
.recap-page .recap-actions .recap-period-label {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) !important;
  gap: 6px !important;
  width: 100% !important;
  min-width: 0 !important;
  margin: 0 !important;
  box-sizing: border-box !important;
}
.recap-page .recap-actions .recap-period-select,
.recap-page .recap-actions .recap-period-select .themed-dropdown-trigger {
  display: flex !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}
.recap-page .recap-actions .recap-period-select::before,
.recap-page .recap-actions .recap-period-select::after {
  content: none !important;
  display: none !important;
}
.recap-page .recap-actions > button {
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  justify-content: center !important;
  box-sizing: border-box !important;
}

@media screen and (max-width: 760px) {
  .resident-toolbar {
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
    gap: 8px !important;
  }
  .resident-toolbar .resident-search {
    grid-column: 1 / -1 !important;
  }
  .resident-toolbar .resident-block-filter,
  .resident-toolbar .resident-subblock-filter {
    grid-column: auto !important;
    width: 100% !important;
    min-width: 0 !important;
  }
  .resident-toolbar .resident-filter .themed-dropdown-trigger {
    height: 43px !important;
    min-height: 43px !important;
    max-height: 43px !important;
    padding: 0 8px !important;
    font-size: 12px !important;
  }
  .resident-toolbar .print-btn {
    grid-column: 1 / 2 !important;
  }
  .resident-toolbar .result-count {
    grid-column: 2 / 3 !important;
  }
  .recap-page .recap-banner {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) !important;
    gap: 12px !important;
    padding: 18px !important;
  }
  .recap-page .recap-actions {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) !important;
    gap: 8px !important;
    width: 100% !important;
  }
  .recap-page .recap-actions .recap-period-label,
  .recap-page .recap-actions .recap-period-select {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
  }
  .recap-page .recap-actions > button {
    width: 100% !important;
    min-height: 42px !important;
  }
}


  .cashflow-transaction-card.is-clickable { cursor:pointer; transition:border-color .16s ease, box-shadow .16s ease; }
  .cashflow-transaction-card.is-clickable:hover { border-color:#9BB8D8; box-shadow:0 4px 14px rgba(23,59,101,.08); }
  .cashflow-tap-hint { color:#8792A0; font-size:10px; font-weight:500; }
  .cashflow-detail-backdrop { position:fixed; inset:0; z-index:100010; display:flex; align-items:center; justify-content:center; padding:18px; background:rgba(15,23,42,.56); }
  .cashflow-detail-modal { width:min(100%,420px); max-height:85vh; overflow:auto; padding:20px; border:1px solid #DCE5EF; border-radius:18px; background:#fff; box-shadow:0 22px 60px rgba(15,23,42,.24); }
  .cashflow-detail-heading { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; margin-bottom:18px; }
  .cashflow-detail-heading h3 { margin:4px 0 0; color:#173B65; font-size:18px; overflow-wrap:anywhere; }
  .cashflow-detail-close { display:grid; place-items:center; width:36px; height:36px; flex:0 0 36px; border:1px solid #DCE5EF; border-radius:10px; background:#F4F7FB; color:#173B65; }
  .cashflow-detail-amount { display:flex; flex-direction:column; gap:5px; padding:14px; margin-bottom:12px; border-radius:12px; background:#FDECEC; }
  .cashflow-detail-amount span { color:#8E3333; font-size:12px; font-weight:600; }
  .cashflow-detail-amount strong { color:#B42318; font-size:23px; overflow-wrap:anywhere; }
  .cashflow-detail-row { display:flex; justify-content:space-between; align-items:flex-start; gap:14px; padding:11px 2px; border-bottom:1px solid #EDF1F6; font-size:13px; }
  .cashflow-detail-row span { color:#718096; }
  .cashflow-detail-row strong { color:#26384B; text-align:right; overflow-wrap:anywhere; }
  .cashflow-detail-actions { display:grid; grid-template-columns:1fr 1fr; gap:9px; margin-top:18px; }
  .cashflow-detail-actions button { min-height:43px; display:flex; align-items:center; justify-content:center; gap:7px; padding:9px; font-size:12px; }
  .cashflow-detail-actions .danger-btn { border:0; border-radius:9px; background:#B42318; color:#fff; font-weight:700; cursor:pointer; }
  @media(max-width:420px) { .cashflow-detail-modal { padding:16px; } .cashflow-detail-actions { grid-template-columns:1fr; } }

  /* Desain ringkas Arus Kas */
  .cashflow-redesign { display:flex; flex-direction:column; gap:16px; min-width:0; }
  .cashflow-filter-row { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
  .cashflow-filter-row label { display:flex; flex-direction:column; gap:6px; min-width:0; color:var(--text); font-size:13px; font-weight:700; }
  .cashflow-filter-row .input-select { width:100%; min-width:0; }
  .cashflow-summary-grid { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; min-width:0; }
  .cashflow-summary-card,.cashflow-category-card,.cashflow-activity { min-width:0; background:#fff; border:1px solid #DCE5EF; border-radius:16px; padding:16px; box-shadow:0 5px 18px rgba(16,42,67,.045); }
  .cashflow-summary-card { display:flex; flex-direction:column; gap:8px; }
  .cashflow-summary-card>span { display:flex; align-items:center; gap:7px; color:#52657A; font-size:13px; font-weight:700; }
  .cashflow-summary-card strong { color:#173B65; font-size:clamp(18px,2vw,25px); line-height:1.25; overflow-wrap:anywhere; }
  .cashflow-summary-card small { color:#7A8796; font-size:11px; }
  .cashflow-summary-card.incoming>span,.cashflow-transaction-amount.masuk { color:#16734A; }
  .cashflow-summary-card.outgoing>span,.cashflow-transaction-amount.keluar { color:#B42318; }
  .cashflow-summary-card.balance>span { color:#173B65; }
  .cashflow-section-heading { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:12px; min-width:0; }
  .cashflow-section-heading h3 { margin:3px 0 0; color:#173B65; font-size:18px; }
  .cashflow-section-heading>small { color:#718096; white-space:nowrap; }
  .cashflow-category-row { display:flex; justify-content:space-between; align-items:center; gap:12px; padding:11px 0; border-top:1px solid #EDF1F6; min-width:0; }
  .cashflow-category-row>span { display:flex; align-items:center; gap:9px; color:#40536A; font-size:14px; min-width:0; }
  .cashflow-category-row strong { color:#173B65; text-align:right; overflow-wrap:anywhere; }
  .category-dot { width:9px; height:9px; border-radius:50%; flex:0 0 9px; background:#2563A6; }
  .category-dot.clean { background:#2563A6; } .category-dot.kas { background:#0F766E; } .category-dot.other { background:#B7791F; }
  .cashflow-expense-details { border:1px solid #DCE5EF; border-radius:14px; background:#fff; overflow:hidden; }
  .cashflow-expense-details>summary { display:flex; align-items:center; gap:8px; padding:14px 16px; color:#173B65; font-weight:700; cursor:pointer; list-style:none; }
  .cashflow-expense-details>summary::-webkit-details-marker { display:none; }
  .cashflow-expense-details[open]>summary { border-bottom:1px solid #E5EAF1; }
  .cashflow-expense-details .expense-form { margin:0; border:0; box-shadow:none; border-radius:0; }
  .cashflow-form-actions { display:flex; gap:8px; flex-wrap:wrap; }
  .cashflow-activity { padding:16px; }
  .cashflow-transaction-list { display:flex; flex-direction:column; gap:9px; }
  .cashflow-transaction-card { display:flex; align-items:center; gap:11px; min-width:0; padding:12px; border:1px solid #E7EDF4; border-radius:13px; background:#fff; }
  .cashflow-transaction-icon { width:38px; height:38px; flex:0 0 38px; display:grid; place-items:center; border-radius:11px; }
  .cashflow-transaction-icon.masuk { background:#E7F6EE; color:#16734A; }
  .cashflow-transaction-icon.keluar { background:#FDECEC; color:#B42318; }
  .cashflow-transaction-main { display:flex; flex:1; flex-direction:column; gap:4px; min-width:0; }
  .cashflow-transaction-main strong { color:#26384B; font-size:14px; overflow-wrap:anywhere; }
  .cashflow-transaction-main span { color:#8792A0; font-size:12px; }
  .cashflow-transaction-amount { display:flex; flex-direction:column; align-items:flex-end; gap:5px; color:#16734A; font-size:13px; font-weight:800; text-align:right; overflow-wrap:anywhere; max-width:46%; }
  .cashflow-transaction-actions { display:flex; gap:5px; }
  .cashflow-transaction-actions button { display:grid; place-items:center; width:28px; height:28px; border:1px solid #DCE5EF; border-radius:8px; background:#F4F7FB; color:#173B65; }
  .cashflow-transaction-actions button:last-child { background:#FEECEC; color:#B42318; }
  .cashflow-empty { padding:24px 12px; text-align:center; color:#8792A0; border:1px dashed #DCE5EF; border-radius:12px; }
  @media(max-width:760px) {
    .cashflow-redesign { gap:11px !important; padding-left:0 !important; padding-right:0 !important; }
    .cashflow-filter-row { grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }
    .cashflow-filter-row label { font-size:12px; }
    .cashflow-summary-grid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }
    .cashflow-summary-card { padding:12px 10px; border-radius:13px; gap:6px; }
    .cashflow-summary-card:last-child { grid-column:1 / -1; }
    .cashflow-summary-card>span { font-size:11px; }
    .cashflow-summary-card strong { font-size:18px; }
    .cashflow-summary-card small { font-size:10px; }
    .cashflow-category-card,.cashflow-activity { padding:12px; border-radius:13px; }
    .cashflow-section-heading h3 { font-size:16px; }
    .cashflow-category-row { padding:10px 0; }
    .cashflow-category-row>span { font-size:12px; }
    .cashflow-category-row strong { font-size:12px; }
    .cashflow-transaction-card { gap:8px; padding:10px 8px; }
    .cashflow-transaction-icon { width:32px; height:32px; flex-basis:32px; }
    .cashflow-transaction-main strong { font-size:12px; }
    .cashflow-transaction-main span { font-size:10px; }
    .cashflow-transaction-amount { font-size:11px; max-width:45%; }
  }

/* FIX MOBILE: filter bulan Rekap Bulanan tidak boleh bertumpuk/overflow */
.recap-page .recap-banner,
.recap-page .recap-actions,
.recap-page .recap-actions .recap-period-label,
.recap-page .recap-actions .recap-period-select {
  min-width: 0 !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}
.recap-page .recap-actions .recap-period-label {
  display: flex !important;
  flex-direction: column !important;
  align-items: stretch !important;
  gap: 6px !important;
  width: 100% !important;
  flex: 1 1 100% !important;
  margin: 0 !important;
}
.recap-page .recap-actions .recap-period-select {
  position: relative !important;
  display: block !important;
  width: 100% !important;
  flex: 1 1 auto !important;
  overflow: visible !important;
}
.recap-page .recap-actions .recap-period-select .themed-dropdown-trigger {
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  min-height: 44px !important;
  box-sizing: border-box !important;
  margin: 0 !important;
}
.recap-page .recap-actions .recap-period-select .themed-dropdown-value {
  min-width: 0 !important;
  max-width: calc(100% - 28px) !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
}
@media (max-width: 760px) {
  .recap-page .recap-banner {
    grid-template-columns: minmax(0, 1fr) !important;
    width: 100% !important;
    padding: 15px !important;
    gap: 12px !important;
  }
  .recap-page .recap-actions {
    display: flex !important;
    flex-direction: column !important;
    align-items: stretch !important;
    flex-wrap: nowrap !important;
    gap: 9px !important;
    width: 100% !important;
  }
  .recap-page .recap-actions .recap-period-label,
  .recap-page .recap-actions .recap-period-select,
  .recap-page .recap-actions .recap-period-select .themed-dropdown-trigger,
  .recap-page .recap-actions > button {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    flex: 0 0 auto !important;
  }
  .recap-page .recap-actions > button {
    min-height: 43px !important;
  }
}


/* FIX FINAL FILTER BULAN: satu kontrol dropdown kustom saja */
.recap-page .recap-period-label,
.recap-page .recap-period-label::before,
.recap-page .recap-period-label::after {
  display: flex !important;
  flex-direction: column !important;
  align-items: stretch !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  padding: 0 !important;
  border: 0 !important;
  border-radius: 0 !important;
  background: transparent !important;
  box-shadow: none !important;
  content: none !important;
  appearance: none !important;
  -webkit-appearance: none !important;
}
.recap-page .recap-period-select {
  position: relative !important;
  display: block !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  margin: 0 !important;
  padding: 0 !important;
  border: 0 !important;
  border-radius: 0 !important;
  background: transparent !important;
  background-image: none !important;
  box-shadow: none !important;
  appearance: none !important;
  -webkit-appearance: none !important;
}
.recap-page .recap-period-select::before,
.recap-page .recap-period-select::after,
.recap-page .recap-period-select .themed-dropdown-trigger::before,
.recap-page .recap-period-select .themed-dropdown-trigger::after {
  content: none !important;
  display: none !important;
}
.recap-page .recap-period-select select,
.recap-page .recap-period-select > select {
  display: none !important;
}
.recap-page .recap-period-select .themed-dropdown-trigger {
  position: relative !important;
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  min-height: 46px !important;
  height: 46px !important;
  margin: 0 !important;
  padding: 0 13px !important;
  border: 1px solid #DCE5EF !important;
  border-radius: 10px !important;
  background: #FFFFFF !important;
  background-image: none !important;
  box-shadow: 0 1px 3px rgba(16,42,67,.05) !important;
  appearance: none !important;
  -webkit-appearance: none !important;
}
.recap-page .recap-period-select .themed-dropdown-trigger > svg {
  display: block !important;
  flex: 0 0 18px !important;
  width: 18px !important;
  height: 18px !important;
  margin: 0 !important;
}
.recap-page .recap-period-select .themed-dropdown-value {
  flex: 1 1 auto !important;
  min-width: 0 !important;
  max-width: calc(100% - 28px) !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
}

/* Popup ekspor JPG HD: dialog internal aplikasi, bukan alert browser. */
.recap-export-notice-backdrop {
  position: fixed;
  inset: 0;
  z-index: 100120;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px;
  background: rgba(16,42,67,.56);
}
.recap-export-notice {
  width: min(100%, 390px);
  box-sizing: border-box;
  padding: 24px;
  border: 1px solid #DCE5EF;
  border-radius: 18px;
  background: #FFFFFF;
  box-shadow: 0 22px 60px rgba(16,42,67,.22);
  text-align: left;
}
.recap-export-notice-icon {
  display: grid;
  place-items: center;
  width: 48px;
  height: 48px;
  margin-bottom: 15px;
  border-radius: 14px;
}
.recap-export-notice-icon.success { color: #15803D; background: #DCFCE7; }
.recap-export-notice-icon.error { color: #B91C1C; background: #FEE2E2; }
.recap-export-notice h3 { margin: 5px 0 8px; color: #102A43; font-size: 19px; line-height: 1.3; }
.recap-export-notice p { margin: 0 0 21px; color: #64748B; font-size: 14px; line-height: 1.6; }
.recap-export-notice .primary-btn { width: 100%; min-height: 44px; justify-content: center; }
@media (max-width: 480px) {
  .recap-export-notice { padding: 21px 19px; border-radius: 17px; }
}


/* =========================================================
   FINAL FIX REKAP MOBILE:
   - Pembungkus filter bukan label kontrol bawaan.
   - Hanya satu bingkai dan satu ikon panah pada dropdown custom.
   - Dialog ekspor memakai tampilan aplikasi Navy + Putih.
   ========================================================= */
.recap-page .recap-banner .recap-actions {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) !important;
  align-items: stretch !important;
  gap: 10px !important;
  width: 100% !important;
  min-width: 0 !important;
}
.recap-page .recap-banner .recap-period-label {
  display: flex !important;
  flex-direction: column !important;
  align-items: stretch !important;
  justify-content: flex-start !important;
  gap: 7px !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  min-height: 0 !important;
  height: auto !important;
  margin: 0 !important;
  padding: 0 !important;
  border: 0 !important;
  border-radius: 0 !important;
  background: transparent !important;
  background-image: none !important;
  box-shadow: none !important;
  color: #DBEAFE !important;
  font: 600 13px/1.35 Inter, sans-serif !important;
  appearance: none !important;
  -webkit-appearance: none !important;
}
.recap-page .recap-period-label::before,
.recap-page .recap-period-label::after,
.recap-page .recap-period-title::before,
.recap-page .recap-period-title::after,
.recap-page .recap-period-select::before,
.recap-page .recap-period-select::after,
.recap-page .recap-period-select .themed-dropdown-trigger::before,
.recap-page .recap-period-select .themed-dropdown-trigger::after {
  content: none !important;
  display: none !important;
}
.recap-page .recap-period-title {
  display: block !important;
  width: auto !important;
  margin: 0 !important;
  padding: 0 !important;
  border: 0 !important;
  background: transparent !important;
  box-shadow: none !important;
  color: #DBEAFE !important;
  font: 600 13px/1.35 Inter, sans-serif !important;
}
.recap-page .recap-period-select {
  position: relative !important;
  display: block !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  height: auto !important;
  min-height: 0 !important;
  margin: 0 !important;
  padding: 0 !important;
  border: 0 !important;
  border-radius: 0 !important;
  background: transparent !important;
  background-image: none !important;
  box-shadow: none !important;
  overflow: visible !important;
}
.recap-page .recap-period-select .themed-dropdown-trigger {
  position: relative !important;
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 10px !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  height: 46px !important;
  min-height: 46px !important;
  max-height: 46px !important;
  margin: 0 !important;
  padding: 0 13px !important;
  border: 1px solid #C9D8E8 !important;
  border-radius: 10px !important;
  background: #FFFFFF !important;
  background-image: none !important;
  box-shadow: 0 1px 3px rgba(16,42,67,.05) !important;
  color: #102A43 !important;
  font: 600 14px/1.25 Inter, sans-serif !important;
  appearance: none !important;
  -webkit-appearance: none !important;
}
.recap-page .recap-period-select .themed-dropdown-value {
  display: block !important;
  flex: 1 1 auto !important;
  min-width: 0 !important;
  max-width: calc(100% - 30px) !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
}
.recap-page .recap-period-select .themed-dropdown-trigger > svg {
  display: block !important;
  flex: 0 0 18px !important;
  width: 18px !important;
  height: 18px !important;
  margin: 0 !important;
  color: #173B65 !important;
  transform: none;
}
.recap-page .recap-period-select.is-open .themed-dropdown-trigger > svg {
  transform: rotate(180deg) !important;
}
.recap-page .recap-actions > button {
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  min-height: 44px !important;
  justify-content: center !important;
  box-sizing: border-box !important;
}

/* Dialog ekspor JPG HD: gaya dialog aplikasi, bukan dialog alamat situs. */
.recap-export-notice-backdrop {
  z-index: 100120 !important;
  padding: 18px !important;
  background: rgba(10, 27, 46, .62) !important;
  backdrop-filter: blur(3px);
  -webkit-backdrop-filter: blur(3px);
}
.recap-export-notice {
  width: min(100%, 410px) !important;
  padding: 0 22px 22px !important;
  overflow: hidden !important;
  border: 1px solid #DCE5EF !important;
  border-radius: 20px !important;
  background: #FFFFFF !important;
  box-shadow: 0 24px 65px rgba(7, 25, 45, .28) !important;
}
.recap-export-notice::before {
  content: "KAS WARGA";
  display: block;
  margin: 0 -22px 20px;
  padding: 17px 22px;
  background: linear-gradient(120deg, #102A43, #173B65 68%, #2563A6);
  color: #FFFFFF;
  font: 700 12px/1.3 Inter, sans-serif;
  letter-spacing: .07em;
}
.recap-export-notice-icon {
  width: 54px !important;
  height: 54px !important;
  margin: 0 0 12px !important;
  border-radius: 16px !important;
}
.recap-export-notice h3 {
  margin: 5px 0 8px !important;
  color: #102A43 !important;
  font-size: 19px !important;
  font-weight: 800 !important;
}
.recap-export-notice p {
  margin: 0 0 20px !important;
  color: #64748B !important;
  font-size: 14px !important;
  line-height: 1.6 !important;
}
.recap-export-notice .primary-btn {
  min-height: 46px !important;
  border-radius: 11px !important;
  background: #173B65 !important;
  border-color: #173B65 !important;
  font-size: 14px !important;
  font-weight: 700 !important;
}
@media screen and (max-width: 760px) {
  .recap-page .recap-banner {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) !important;
    gap: 13px !important;
    width: 100% !important;
    min-width: 0 !important;
    padding: 17px !important;
    box-sizing: border-box !important;
  }
  .recap-page .recap-banner > div,
  .recap-page .recap-actions,
  .recap-page .recap-period-label,
  .recap-page .recap-period-select {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
  }
  .recap-page .recap-period-select .themed-dropdown-trigger {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    height: 46px !important;
    min-height: 46px !important;
    max-height: 46px !important;
  }
  .recap-export-notice {
    padding: 0 19px 19px !important;
    border-radius: 19px !important;
  }
  .recap-export-notice::before {
    margin: 0 -19px 18px;
    padding: 16px 19px;
  }
}


/* FINAL MOBILE ACTION LAYOUT:
   PDF dan JPG HD berdampingan, Excel tetap satu baris penuh. */
.recap-page .recap-banner .recap-actions {
  display: grid !important;
  grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
  align-items: stretch !important;
  gap: 8px !important;
  width: 100% !important;
  min-width: 0 !important;
}
.recap-page .recap-banner .recap-actions .recap-period-label {
  grid-column: 1 / -1 !important;
  width: 100% !important;
  min-width: 0 !important;
}
.recap-page .recap-banner .recap-actions > button:nth-of-type(1) {
  grid-column: 1 / -1 !important;
}
.recap-page .recap-banner .recap-actions > button:nth-of-type(2) {
  grid-column: 1 !important;
}
.recap-page .recap-banner .recap-actions > button:nth-of-type(3) {
  grid-column: 2 !important;
}
.recap-page .recap-banner .recap-actions > button {
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 6px !important;
  width: 100% !important;
  min-width: 0 !important;
  max-width: 100% !important;
  min-height: 43px !important;
  height: auto !important;
  box-sizing: border-box !important;
  padding: 9px 7px !important;
  white-space: normal !important;
  text-align: center !important;
  font-size: 12px !important;
  line-height: 1.2 !important;
}
.recap-page .recap-banner .recap-actions > button svg {
  flex: 0 0 16px !important;
  width: 16px !important;
  height: 16px !important;
}
@media screen and (max-width: 760px) {
  .recap-page .recap-banner .recap-actions {
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
    gap: 8px !important;
  }
  .recap-page .recap-banner .recap-actions .recap-period-label {
    grid-column: 1 / -1 !important;
  }
  .recap-page .recap-banner .recap-actions > button:nth-of-type(1) {
    grid-column: 1 / -1 !important;
  }
  .recap-page .recap-banner .recap-actions > button:nth-of-type(2) {
    grid-column: 1 !important;
  }
  .recap-page .recap-banner .recap-actions > button:nth-of-type(3) {
    grid-column: 2 !important;
  }
}


/* Tampilan filter status & blok yang dicetak (nilai aktif, bukan kontrol interaktif). */
.recap-print-filters { display:none !important; }

@media print {
  /* Kontrol periode dan tombol ekspor tidak ikut tercetak. */
  .app-shell .recap-page .recap-banner .recap-actions,
  .app-shell .recap-page .recap-banner .recap-actions * { display:none !important; }
  .app-shell .recap-page .recap-detail-head .recap-detail-toolbar { display:none !important; }

  .app-shell .recap-page .recap-print-filters {
    display:flex !important; align-items:center !important; justify-content:flex-end !important;
    flex-wrap:wrap !important; gap:6px !important; width:100% !important;
    margin:-1px 0 6px !important; padding:0 !important;
    color:#475569 !important; font-size:7.5pt !important; line-height:1.2 !important;
  }
  .app-shell .recap-page .recap-print-filters > span {
    display:inline-block !important; padding:3px 6px !important;
    border:1px solid #DCE5EF !important; border-radius:4px !important;
    background:#F8FAFC !important; color:#334155 !important; white-space:nowrap !important;
  }
  .app-shell .recap-page .recap-print-filters b { color:#173B65 !important; font-weight:700 !important; }

  /* PDF mengikuti ukuran dan susunan ringkas pada JPG HD. */
  .app-shell .recap-page .recap-banner {
    display:block !important; min-height:0 !important; height:auto !important;
    margin:0 0 8px !important; padding:10px 12px !important;
    border:1px solid #BFDBFE !important; border-radius:7px !important;
    background:#EFF6FF !important; color:#102A43 !important; box-shadow:none !important;
    break-inside:avoid !important; page-break-inside:avoid !important;
  }
  .app-shell .recap-page .recap-banner h2 {
    margin:2px 0 4px !important; color:#102A43 !important;
    font-size:17pt !important; line-height:1.2 !important;
  }
  .app-shell .recap-page .recap-banner p {
    margin:2px 0 0 !important; color:#475569 !important; font-size:8.5pt !important;
  }
  .app-shell .recap-page .recap-banner .eyebrow {
    color:#245789 !important; font-size:7pt !important; line-height:1.15 !important;
  }
  .app-shell .recap-page .recap-grid {
    display:grid !important; grid-template-columns:repeat(2,minmax(0,1fr)) !important;
    align-items:start !important; gap:7px !important; margin:7px 0 8px !important;
  }
  .app-shell .recap-page .recap-grid > .big-card {
    display:block !important; align-self:start !important;
    width:auto !important; min-width:0 !important; min-height:0 !important; height:auto !important;
    margin:0 !important; padding:7px 8px !important;
    border:1px solid #DCE5EF !important; border-radius:6px !important;
    background:#FFFFFF !important; color:#20352E !important; box-shadow:none !important;
    break-inside:avoid !important; page-break-inside:avoid !important;
  }
  .app-shell .recap-page .recap-grid > .big-card .eyebrow {
    display:block !important; margin:0 0 2px !important;
    color:#245789 !important; font-size:7pt !important; line-height:1.15 !important;
  }
  .app-shell .recap-page .recap-grid > .big-card strong {
    display:block !important; margin:1px 0 !important; color:#102A43 !important;
    font-size:13pt !important; line-height:1.12 !important;
  }
  .app-shell .recap-page .recap-grid > .big-card small {
    display:block !important; margin:0 !important; color:#52665D !important;
    font-size:7.5pt !important; line-height:1.15 !important;
  }
  .app-shell .recap-page .recap-grid .bar {
    height:4px !important; margin:3px 0 !important; background:#DBEAFE !important;
  }
  .app-shell .recap-page .recap-grid .bar i { background:#2563EB !important; }
  .app-shell .recap-page .recap-income-details {
    display:flex !important; flex-direction:column !important; gap:0 !important;
    margin:4px 0 0 !important; padding-top:4px !important;
    border-top:1px solid #E2E8F0 !important; line-height:1.15 !important;
  }
  .app-shell .recap-page .recap-income-details > div {
    display:flex !important; align-items:center !important; justify-content:space-between !important;
    gap:8px !important; min-height:0 !important; padding:2px 0 !important;
    line-height:1.15 !important; font-size:7.2pt !important;
  }
  .app-shell .recap-page .recap-income-details span,
  .app-shell .recap-page .recap-income-details b {
    color:#52665D !important; font-size:7.2pt !important; line-height:1.15 !important;
  }
  .app-shell .recap-page .recap-income-details b {
    color:#173B65 !important; font-weight:700 !important; white-space:nowrap !important;
  }
  .app-shell .recap-page .recap-detail-head {
    display:flex !important; justify-content:space-between !important; align-items:flex-end !important;
    gap:8px !important; margin:7px 0 4px !important; padding-bottom:4px !important;
    border-bottom:1px solid #DCE5EF !important;
  }
  .app-shell .recap-page .recap-detail-head h2 {
    margin:1px 0 2px !important; color:#102A43 !important;
    font-size:10.5pt !important; line-height:1.2 !important;
  }
  .app-shell .recap-page .recap-detail-head .eyebrow {
    color:#245789 !important; font-size:7pt !important; line-height:1.1 !important;
  }
  .app-shell .recap-page .recap-payment-card {
    display:block !important; width:100% !important; min-width:0 !important; max-width:100% !important;
    margin:0 !important; padding:0 !important; overflow:visible !important;
    border:1px solid #DCE5EF !important; border-radius:6px !important;
    background:#FFFFFF !important; box-shadow:none !important;
  }
  .app-shell .recap-page .recap-payment-table {
    display:table !important; width:100% !important; min-width:0 !important; max-width:100% !important;
    margin:0 !important; border-collapse:collapse !important; table-layout:fixed !important;
    font-size:7pt !important;
  }
  .app-shell .recap-page .recap-payment-table th,
  .app-shell .recap-page .recap-payment-table td {
    padding:4px !important; line-height:1.2 !important; white-space:normal !important;
    overflow-wrap:anywhere !important; vertical-align:top !important; border-bottom:1px solid #E2EAE5 !important;
  }
  .app-shell .recap-page .recap-payment-table th {
    color:#475569 !important; background:#F1F5F9 !important;
    font-size:6.5pt !important; font-weight:700 !important;
  }
  .app-shell .recap-page .recap-payment-table .status-paid,
  .app-shell .recap-page .recap-payment-table .status-unpaid,
  .app-shell .recap-page .recap-payment-table .block-tag {
    display:inline-block !important; padding:2px 4px !important; border-radius:4px !important;
    font-size:6.5pt !important; white-space:nowrap !important;
  }
}


        .community-report-identity { display:block; margin-top:4px; color:#DBEAFE; font-size:11px; line-height:1.35; }
        .report-community-location { margin:3px 0 0; color:#52677F; font-size:9pt; line-height:1.35; }
        .community-location-label { display:block; max-width:220px; margin-top:2px; color:#93A9C1 !important; font-size:10px !important; line-height:1.3 !important; }
        .identity-settings {
          display:flex; flex-direction:column; gap:16px; margin:0 0 18px; padding:20px;
          background:#FFFFFF; border:1px solid #DCE5EF; border-radius:14px;
          box-shadow:0 5px 18px rgba(16,42,67,.045);
        }
        .identity-settings-heading h3 { margin:4px 0 5px; color:#102A43; font-size:18px; }
        .identity-settings-heading p { margin:0; color:#64748B; font-size:13px; line-height:1.5; }
        .identity-settings-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
        .identity-settings-grid label { display:flex; flex-direction:column; gap:6px; color:#334155; font-size:13px; font-weight:600; min-width:0; }
        .identity-settings-grid input { width:100%; min-width:0; min-height:43px; box-sizing:border-box; padding:10px 12px; border:1px solid #DCE5EF; border-radius:9px; background:#fff; color:#1F2937; font:inherit; }
        .identity-settings-grid .identity-wide { grid-column:1 / -1; }
        .identity-preview { display:flex; flex-direction:column; gap:4px; padding:14px 16px; border:1px solid #BFDBFE; border-radius:11px; background:#EFF6FF; }
        .identity-preview .eyebrow { color:#64748B !important; font-size:10px; letter-spacing:.08em; }
        .identity-preview strong { color:#173B65; font-size:18px; letter-spacing:.03em; }
        .identity-preview b { color:#173B65; font-size:15px; }
        .identity-preview small { color:#52677F; font-size:12px; line-height:1.4; }
        .identity-save-btn { align-self:flex-start; min-height:43px; }
        .resident-print-title { display:none; }
        @media print {
          .resident-print-title {
            display:block !important;
            margin:0 0 12px !important;
            padding:0 0 7px !important;
            border-bottom:2px solid #173B65 !important;
            color:#102A43 !important;
            font:700 17pt/1.2 Arial, sans-serif !important;
          }
          .content.narrow .section-head:after { content:none !important; display:none !important; }
        }
        @media(max-width:600px) {
          .identity-settings { padding:15px; gap:13px; }
          .identity-settings-grid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }
          .identity-settings-grid .identity-wide { grid-column:1 / -1; }
          .identity-save-btn { width:100%; justify-content:center; }
        }


        @media print {
          .recap-page .community-report-identity {
            display:block !important;
            margin:4px 0 0 !important;
            color:#244563 !important;
            font-size:8pt !important;
            font-weight:700 !important;
            line-height:1.3 !important;
          }
          .report-page .report-kicker,
          .public-summary .public-heading p,
          .public-summary .public-footer span:first-child {
            color:#244563 !important;
            font-weight:700 !important;
          }
          .report-page .report-community-location {
            color:#52677F !important;
            font-size:8.5pt !important;
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
        <div className="brand"><Logo /><div><strong>kas warga</strong><small>{communityIdentityLine}</small>{communityLocationLine && <small className="community-location-label">{communityLocationLine}</small>}</div></div>
        <div className="side-label">Menu utama</div>
        <nav>
          {[["overview", Home, "Ringkasan"], ["input", ReceiptText, "Input pembayaran"], ["cashflow", WalletCards, "Arus keuangan"], ["recap", CalendarDays, "Rekap bulanan"], ["report", FileText, "Laporan warga"], ["public", Users, "Ringkasan publik"], ["residents", Users, "Data warga"]].map(([id, Icon, label]) => (
            <button className={`nav-item ${active === id ? "active" : ""}`} key={id} onClick={() => { setActive(id); setSettingsOpen(false); setMobileMenuOpen(false); if (id === "input") { setPaymentPeriod(getPeriodeSekarang()); setBlock(""); setPaymentSubBlock(""); setResident(""); setClean(""); setKas(""); setOther(""); setEditingEntry(null); } }}><Icon size={18} />{label}{id === "input" && <span className="nav-dot" />}</button>
          ))}
        </nav>
        <div className="side-bottom"><button className={`nav-item ${settingsOpen ? "active" : ""}`} onClick={() => {setSettingsOpen(!settingsOpen);if (!settingsOpen) setActive("settings");setMobileMenuOpen(false);}}><Settings2 size={18} />Pengaturan</button><div className="profile"><span className="avatar">{(user?.user_metadata?.full_name || user?.email || "W").slice(0,2).toUpperCase()}</span><div><b>{namaPengguna}</b><small>{user?.email || "Akun terverifikasi"}</small></div><button
  type="button"
  className="signout-btn"
  onClick={() => {
    setMobileMenuOpen(false);
    requestConfirmation(
      "Apakah Anda yakin ingin keluar dari aplikasi?",
      () => onSignOut?.()
    );
  }}
>
  Keluar
</button></div></div>
      </aside>
      <nav className="mobile-floating-nav" aria-label="Navigasi utama mobile">
        <button type="button" className={active === "overview" ? "active" : ""} onClick={() => { setActive("overview"); setSettingsOpen(false); setMobileMenuOpen(false); }}>
          <Home size={20} /><span>Beranda</span>
        </button>
        <button type="button" className={active === "input" ? "active" : ""} onClick={() => { setActive("input"); setSettingsOpen(false); setMobileMenuOpen(false); setPaymentPeriod(getPeriodeSekarang()); setBlock(""); setPaymentSubBlock(""); setResident(""); setClean(""); setKas(""); setOther(""); setEditingEntry(null); }}>
          <ReceiptText size={20} /><span>Bayar</span>
        </button>
        <button type="button" className={active === "cashflow" ? "active" : ""} onClick={() => { setActive("cashflow"); setSettingsOpen(false); setMobileMenuOpen(false); }}>
          <WalletCards size={20} /><span>Arus kas</span>
        </button>
        <button type="button" className={active === "report" ? "active" : ""} onClick={() => { setActive("report"); setSettingsOpen(false); setMobileMenuOpen(false); }}>
          <FileText size={20} /><span>Laporan</span>
        </button>
        <button type="button" className={active === "residents" ? "active" : ""} onClick={() => { setActive("residents"); setSettingsOpen(false); setMobileMenuOpen(false); }}>
          <Users size={20} /><span>Warga</span>
        </button>
      </nav>
      <main className="main">
        <div className="connection-status" style={{ display: "none" }}></div>
        <header className="topbar">

  <button
    type="button"
    className="mobile-menu-toggle"
    onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
    aria-label="Buka atau tutup menu"
  >
    {mobileMenuOpen ? <X size={22} /> : <Menu size={22} />}
  </button>

  <div>
    <span className="eyebrow">
      {formatTanggalIndonesia()}
    </span>

    <h1>
      {active === "overview"
        ? `Selamat datang, ${namaPengguna}`
        : active === "input"
        ? "Input pembayaran"
        : active === "cashflow"
        ? "Arus keuangan"
        : active === "recap"
        ? "Rekap bulanan"
        : active === "report"
        ? "Laporan warga"
        : active === "public"
        ? "Ringkasan publik"
        : active === "settings"
        ? "Pengaturan"
        : "Data warga"}
    </h1>
  </div>

</header>
        {notice && <div className="toast"><span>✓</span>{notice}</div>}
        {active === "overview" && !settingsOpen && <Overview
  communityIdentityLine={communityIdentityLine}
  entries={entries}
  cashOut={cashOut}
  residents={residents}
  runningBalance={runningBalance}
  setActive={setActive}
  setSettingsOpen={setSettingsOpen}
  onSignOut={onSignOut}
  requestConfirmation={requestConfirmation}
/>}
        {active === "input" && !settingsOpen && <InputView {...{ block, setBlock, autoSubBlock, paymentSubBlock, setPaymentSubBlock, resident, setResident, clean, setClean, kas, setKas, other, setOther, paymentPeriod, setPaymentPeriod, addEntry, residents, editingEntry, setEditingEntry }} />}
        {active === "cashflow" && !settingsOpen && <Cashflow entries={entries} cashOut={cashOut} runningBalance={runningBalance} addExpense={addExpense} expenseTitle={expenseTitle} setExpenseTitle={setExpenseTitle} expenseAmount={expenseAmount} setExpenseAmount={setExpenseAmount} expensePeriod={expensePeriod} setExpensePeriod={setExpensePeriod} editingExpense={editingExpense} setEditingExpense={setEditingExpense} startEditExpense={startEditExpense} deleteExpense={deleteExpense} />}
        {active === "recap" && !settingsOpen && <Recap
  communityIdentityLine={communityIdentityLine}
  communityLocationLine={communityLocationLine}
  entries={entries}
  total={total}
  residents={residents}
  recapSearch={recapSearch}
  setRecapSearch={setRecapSearch}
  startEditEntry={startEditEntry}
  deleteEntry={deleteEntry}
/>}
        {active === "report" && !settingsOpen && <ResidentReport entries={entries} cashOut={cashOut} runningBalance={runningBalance} openingBalance={openingBalance} manualAdjustment={manualAdjustment} residents={residents} communityIdentityLine={communityIdentityLine} communityLocationLine={communityLocationLine} />}
        {active === "public" && !settingsOpen && <PublicSummary entries={entries} cashOut={cashOut} openingBalance={openingBalance} manualAdjustment={manualAdjustment} communityIdentityLine={communityIdentityLine} communityLocationLine={communityLocationLine} />}
        {settingsOpen && <SettingsPanel backupData={backupData} restoreData={restoreData} resetData={resetData} openingBalance={openingBalance} setOpeningBalance={setOpeningBalance} manualAdjustment={manualAdjustment} setManualAdjustment={setManualAdjustment} balanceHistory={balanceHistory} setBalanceHistory={setBalanceHistory} setNotice={setNotice} communityIdentity={communityIdentity} setCommunityIdentity={setCommunityIdentity} communityIdentityLine={communityIdentityLine} communityLocationLine={communityLocationLine} />}
        {active === "residents" && !settingsOpen && <Residents {...{ residents, residentName, setResidentName, residentBlock, setResidentBlock, residentSubBlock, setResidentSubBlock, addResident, editingResident, setEditingResident, startEditResident, deleteResident, residentSearch, setResidentSearch, residentFilter, setResidentFilter, residentSubFilter, setResidentSubFilter, communityIdentityLine, communityLocationLine }} />}
      </main>
      {confirmDialog && (
  <div
    style={{
      position: "fixed",
      inset: 0,
      zIndex: 99999,
      background: "rgba(15, 23, 42, 0.55)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: "20px"
    }}
    onClick={() => setConfirmDialog(null)}
  >
    <div
      role="alertdialog"
      aria-modal="true"
      aria-labelledby="confirm-title"
      aria-describedby="confirm-message"
      onClick={(e) => e.stopPropagation()}
      style={{
        width: "100%",
        maxWidth: "390px",
        background: "#FFFFFF",
        borderRadius: "16px",
        padding: "24px",
        boxShadow: "0 20px 60px rgba(0,0,0,.22)",
        border: "1px solid #DCE5EF"
      }}
    >
      <div style={{
        width: "44px",
        height: "44px",
        borderRadius: "12px",
        background: "#DBEAFE",
        color: "#173B65",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: "23px",
        marginBottom: "16px"
      }}>
        !
      </div>

      <h3
        id="confirm-title"
        style={{
          margin: "0 0 10px",
          color: "#102A43",
          fontSize: "18px",
          fontWeight: 700
        }}
      >
        Konfirmasi tindakan
      </h3>

      <p
        id="confirm-message"
        style={{
          margin: "0 0 24px",
          color: "#64748B",
          fontSize: "14px",
          lineHeight: 1.6
        }}
      >
        {confirmDialog.message}
      </p>

      <div style={{
        display: "flex",
        justifyContent: "flex-end",
        gap: "10px"
      }}>
        <button
          type="button"
          onClick={() => setConfirmDialog(null)}
          style={{
            padding: "11px 18px",
            borderRadius: "9px",
            border: "1px solid #DCE5EF",
            background: "#FFFFFF",
            color: "#475569",
            fontWeight: 600,
            cursor: "pointer"
          }}
        >
          Batal
        </button>

        <button
          type="button"
          onClick={handleConfirmAction}
          style={{
            padding: "11px 18px",
            borderRadius: "9px",
            border: "none",
            background: "#173B65",
            color: "#FFFFFF",
            fontWeight: 600,
            cursor: "pointer"
          }}
        >
          Ya, lanjutkan
        </button>
      </div>
    </div>
  </div>
)}
    </div>
  );
}

function Logo({ large = false }) { return <span className={`brand-mark ${large ? "large" : ""}`} aria-label="Logo Kas Warga"><svg viewBox="0 0 48 48" aria-hidden="true"><path d="M12 10v28M12 24l20-14M12 24l21 14" fill="none" stroke="currentColor" strokeWidth="4.2" strokeLinecap="round" strokeLinejoin="round"/><path d="M29 10h7v7M30 38h6v-7" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round"/><path d="M8 42h32" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" opacity=".65"/></svg></span>; }
function LoginScreen({ onLogin, communityIdentityLine = "RT 26 / RW 04", communityLocationLine = "" }) { return <main className="login-screen"><div className="login-art"><div className="login-orbit orbit-one" /><div className="login-orbit orbit-two" /><div className="login-sun" /><div className="login-logo-lockup"><Logo large /></div></div><section className="login-card"><div className="brand login-brand"><Logo /><div><strong>kas warga</strong><small>{communityIdentityLine}</small>{communityLocationLine && <small>{communityLocationLine}</small>}</div></div><span className="eyebrow">Ruang bendahara</span><h1>Kelola kas warga,<br/><em>tanpa ribet.</em></h1><p className="login-copy">Masuk untuk mencatat iuran, memantau arus kas, dan membagikan laporan yang transparan.</p><button className="login-btn" type="button" onClick={onLogin}>Masuk ke aplikasi <span>→</span></button><small className="login-note">Akses lokal untuk pengelola kas RT</small></section></main>; }

function SettingsPanel({
  backupData, restoreData, resetData,
  openingBalance, setOpeningBalance,
  manualAdjustment, setManualAdjustment,
  balanceHistory, setBalanceHistory, setNotice,
  communityIdentity, setCommunityIdentity,
  communityIdentityLine, communityLocationLine
}) {
  const [identityDraft, setIdentityDraft] = useState({
    ...DEFAULT_COMMUNITY_IDENTITY,
    ...(communityIdentity || {}),
  });

  useEffect(() => {
    setIdentityDraft({
      ...DEFAULT_COMMUNITY_IDENTITY,
      ...(communityIdentity || {}),
    });
  }, [communityIdentity]);

  const updateIdentityField = (field, value) =>
    setIdentityDraft((current) => ({ ...current, [field]: value }));

  const saveIdentity = (event) => {
    event.preventDefault();
    const nextIdentity = {
      rtNumber: String(identityDraft.rtNumber || "").trim(),
      rwNumber: String(identityDraft.rwNumber || "").trim(),
      kelurahan: String(identityDraft.kelurahan || "").trim(),
      kecamatan: String(identityDraft.kecamatan || "").trim(),
      kabupatenKota: String(identityDraft.kabupatenKota || "").trim(),
    };

    if (!nextIdentity.rtNumber || !nextIdentity.rwNumber) {
      setNotice("Nomor RT dan RW wajib diisi.");
      setTimeout(() => setNotice(""), 2600);
      return;
    }
    if (!/^[0-9A-Za-z.-]{1,5}$/.test(nextIdentity.rtNumber) ||
        !/^[0-9A-Za-z.-]{1,5}$/.test(nextIdentity.rwNumber)) {
      setNotice("Nomor RT dan RW maksimal 5 karakter.");
      setTimeout(() => setNotice(""), 2600);
      return;
    }

    setCommunityIdentity(nextIdentity);
    setNotice("Identitas berhasil diperbarui. Perubahan akan disinkronkan ke Supabase.");
    setTimeout(() => setNotice(""), 3200);
  };

  const saveBalanceChange = (event) => {
    event.preventDefault();
    setBalanceHistory([{ id: Date.now(), date: new Date().toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" }), opening: openingBalance, correction: manualAdjustment }, ...balanceHistory]);
    setNotice("Perubahan saldo dicatat");
    setTimeout(() => setNotice(""), 2200);
  };

  return <section className="content settings-panel">
    <div className="form-intro">
      <div className="notebook-icon"><Settings2 size={25} /></div>
      <div><h2>Pengaturan data</h2><p>Kelola identitas lingkungan dan data kas warga.</p></div>
    </div>

    <form className="identity-settings" onSubmit={saveIdentity}>
      <div className="identity-settings-heading">
        <span className="eyebrow">Identitas aplikasi</span>
        <h3>Atur nama RT / RW</h3>
        <p>Identitas ini digunakan pada halaman login, header aplikasi, laporan, dan hasil ekspor.</p>
      </div>
      <div className="identity-settings-grid">
        <label>Nomor RT
          <input type="text" inputMode="numeric" maxLength={5} value={identityDraft.rtNumber} onChange={(e) => updateIdentityField("rtNumber", e.target.value)} placeholder="Contoh: 26" required />
        </label>
        <label>Nomor RW
          <input type="text" inputMode="numeric" maxLength={5} value={identityDraft.rwNumber} onChange={(e) => updateIdentityField("rwNumber", e.target.value)} placeholder="Contoh: 04" required />
        </label>
        <label>Desa / Kelurahan
          <input type="text" maxLength={100} value={identityDraft.kelurahan} onChange={(e) => updateIdentityField("kelurahan", e.target.value)} placeholder="Nama desa atau kelurahan" />
        </label>
        <label>Kecamatan
          <input type="text" maxLength={100} value={identityDraft.kecamatan} onChange={(e) => updateIdentityField("kecamatan", e.target.value)} placeholder="Nama kecamatan" />
        </label>
        <label className="identity-wide">Kabupaten / Kota
          <input type="text" maxLength={100} value={identityDraft.kabupatenKota} onChange={(e) => updateIdentityField("kabupatenKota", e.target.value)} placeholder="Contoh: Kabupaten Malang" />
        </label>
      </div>
      <div className="identity-preview">
        <span className="eyebrow">Pratinjau identitas</span>
        <strong>KAS WARGA</strong>
        <b>RT {identityDraft.rtNumber || "—"} / RW {identityDraft.rwNumber || "—"}</b>
        {(identityDraft.kelurahan || identityDraft.kecamatan || identityDraft.kabupatenKota) && (
          <small>{[
            identityDraft.kelurahan ? `Kel. ${identityDraft.kelurahan}` : "",
            identityDraft.kecamatan ? `Kec. ${identityDraft.kecamatan}` : "",
            identityDraft.kabupatenKota || "",
          ].filter(Boolean).join(" · ")}</small>
        )}
      </div>
      <button className="primary-btn identity-save-btn" type="submit"><Check size={18} /> Simpan identitas</button>
    </form>

    <form className="balance-settings" onSubmit={saveBalanceChange}>
      <div><span className="eyebrow">Saldo awal kas</span><h3>Atur saldo awal</h3><p>Nilai ini menjadi titik awal perhitungan saldo berjalan.</p></div>
      <div className="settings-inputs">
        <label>Saldo awal<div className="input-money"><span>Rp</span><input type="number" min="0" value={openingBalance} onChange={e => setOpeningBalance(Number(e.target.value || 0))} /></div></label>
        <label>Koreksi manual<div className="input-money"><span>Rp</span><input type="number" value={manualAdjustment} onChange={e => setManualAdjustment(Number(e.target.value || 0))} /></div></label>
        <button className="primary-btn" type="submit"><Plus size={18} /> Simpan koreksi</button>
      </div>
    </form>

    <div className="history-card">
      <div className="section-head"><div><span className="eyebrow">Audit saldo</span><h2>Riwayat perubahan</h2></div></div>
      {balanceHistory.length === 0 ? <p className="history-empty">Belum ada perubahan saldo atau koreksi.</p> : <div className="history-list">{balanceHistory.map(item => <div className="history-row" key={item.id}><span className="history-dot">↻</span><div><b>Saldo diperbarui</b><small>{item.date}</small></div><div className="history-values"><span>Awal: {money(item.opening)}</span><span>Koreksi: {item.correction >= 0 ? "+" : ""}{money(item.correction)}</span></div></div>)}</div>}
    </div>

    <div className="backup-card">
      <div><span className="eyebrow">Backup dan pemulihan</span><h3>Simpan salinan data kas</h3><p>Backup menyimpan identitas lingkungan, daftar warga, dan seluruh catatan pembayaran dalam file JSON.</p></div>
      <div className="backup-actions">
        <button className="primary-btn" onClick={backupData}><FileText size={17} /> Backup data</button>
        <label className="restore-btn"><UploadIcon /> Pulihkan data<input type="file" accept="application/json,.json" onChange={restoreData} /></label>
      </div>
    </div>
    <div className="tip"><span>✦</span><div><b>Saran aman</b><p>Simpan file backup di tempat aman dan buat salinan baru setelah ada perubahan penting.</p></div></div>
  </section>;
}
function UploadIcon() { return <span className="upload-icon">↥</span>; }
function Overview({
  communityIdentityLine = "RT 26 / RW 04",
  entries = [],
  cashOut = [],
  runningBalance,
  setActive,
  setSettingsOpen,
  residents = [],
  onSignOut,
  requestConfirmation
}) {
  // Semua nilai pada Beranda memakai akumulasi seluruh periode,
  // agar konsisten dengan saldo kas berjalan.
  const cleanTotal = entries.reduce(
    (sum, e) => sum + Number(e.clean || 0),
    0
  );
  const kasTotal = entries.reduce(
    (sum, e) => sum + Number(e.kas || 0),
    0
  );
  const otherTotal = entries.reduce(
    (sum, e) => sum + Number(e.other || 0),
    0
  );
  const totalIncome = cleanTotal + kasTotal + otherTotal;
  const totalExpense = cashOut.reduce(
    (sum, e) => sum + Number(e.amount || 0),
    0
  );
  // Kartu warga di Beranda menampilkan total seluruh warga dan jumlah
  // warga yang sudah membayar pada periode kalender yang sedang berjalan.
  const currentPeriod = getPeriodeSekarang();
  const currentPeriodEntries = entries.filter((entry) =>
    cocokPeriode(entry, currentPeriod)
  );
  const paidResidentCount = residents.filter((resident) =>
    currentPeriodEntries.some((payment) => {
      const sameId =
        payment.residentId != null && resident.id != null &&
        String(payment.residentId) === String(resident.id);

      const paymentName = String(payment.resident || "").trim().toLocaleLowerCase("id-ID");
      const residentName = String(resident.name || "").trim().toLocaleLowerCase("id-ID");
      const sameName = Boolean(paymentName && residentName && paymentName === residentName);

      const paymentBlock = String(payment.block || "").trim().toLocaleUpperCase("id-ID");
      const residentBlock = String(resident.block || "").trim().toLocaleUpperCase("id-ID");
      const paymentSubBlock = String(payment.subBlock || "").trim().toLocaleUpperCase("id-ID");
      const residentSubBlock = String(resident.subBlock || "").trim().toLocaleUpperCase("id-ID");
      const sameAddress = Boolean(
        paymentBlock && residentBlock && paymentBlock === residentBlock &&
        paymentSubBlock && residentSubBlock && paymentSubBlock === residentSubBlock
      );

      return sameId || sameName || sameAddress;
    })
  ).length;
  const paidPercent = residents.length
    ? Math.round((paidResidentCount / residents.length) * 100)
    : 0;

  return (
    <section className="content">
    <div className="hero-card">
  <div>
    <span className="eyebrow light">
  Saldo kas berjalan · Keseluruhan
</span>

<div className="hero-number">
  {money(runningBalance)}
</div>

<span className="hero-note">
  <WalletCards size={15} />
  Saldo awal + total pemasukan − total pengeluaran + koreksi
</span>
    <div className="hero-income-details">
      <div>
        <span>Total iuran kebersihan</span>
        <b>{money(cleanTotal)}</b>
      </div>

      <div>
        <span>Total iuran kas RT</span>
        <b>{money(kasTotal)}</b>
      </div>

      <div>
        <span>Total iuran lainnya</span>
        <b>{money(otherTotal)}</b>
      </div>
    </div>
  </div>

  <div className="hero-art">
    <div className="sun" />
    <div className="hill h1" />
    <div className="hill h2" />
    <WalletCards size={42} />
  </div>
</div>
    <div className="section-head home-summary-head">
  <div>
    <span className="eyebrow">Ringkasan keuangan</span>
    <h2>Ringkasan kas</h2>
  </div>

  <button
    className="text-btn"
    onClick={() => setActive("input")}
  >
    + Input pembayaran
  </button>
</div>
    <div className="stat-grid">
      <Stat
        title="Total pemasukan"
        value={money(totalIncome)}
        icon={ArrowDownLeft}
        tone="green"
        meta={`${entries.length} transaksi pembayaran`}
      />
      <Stat
        title="Total pengeluaran"
        value={money(totalExpense)}
        icon={ArrowUpRight}
        tone="red"
        meta={`${cashOut.length} transaksi pengeluaran`}
      />
      <Stat
        title="Warga sudah bayar"
        value={`${paidResidentCount} / ${residents.length}`}
        icon={Users}
        tone="gold"
        meta={residents.length ? `${paidPercent}% warga membayar pada ${currentPeriod}` : "Belum ada data warga"}
      />
    </div>
    <section className="home-shortcuts" aria-label="Menu lainnya">
      <div className="home-shortcuts-heading"><div><span className="eyebrow">Akses cepat</span><h2>Akses cepat</h2></div></div>
      <div className="home-shortcuts-grid">
        <button type="button" className="home-shortcut" onClick={() => setActive("recap")}><span className="home-shortcut-icon"><CalendarDays size={21}/></span><span><b>Rekap Bulanan</b><small>Ringkasan iuran</small></span></button>
        <button type="button" className="home-shortcut" onClick={() => setActive("public")}><span className="home-shortcut-icon"><Globe2 size={21}/></span><span><b>Ringkasan Publik</b><small>Informasi kas warga</small></span></button>
        <button
  type="button"
  className="home-shortcut"onClick={() => { setActive("settings"); setSettingsOpen(true); }}
><span className="home-shortcut-icon"><Settings2 size={21} /></span><span><b>Pengaturan</b><small>Data dan cadangan</small></span></button>
        <button
  type="button"
  className="home-shortcut home-shortcut-logout"
  onClick={() => {
    requestConfirmation(
      "Apakah Anda yakin ingin keluar dari aplikasi?",
      () => onSignOut?.()
    );
  }}
>
  <span className="home-shortcut-icon">
    <LogOut size={21} />
  </span>
  <span>
    <b>Keluar</b>
    <small>Akhiri sesi akun</small>
  </span>
</button>
      </div>
    </section>
  </section>
  );
}
function ThemedDropdown({
  value,
  onChange,
  options = [],
  className = "",
  disabled = false,
  placeholder = "Pilih pilihan",
  ariaLabel = "Pilih pilihan",
}) {
  const [open, setOpen] = useState(false);
  const dropdownRef = useRef(null);
  const normalizedOptions = options.map((option) =>
    typeof option === "string"
      ? { value: option, label: option }
      : option
  );
  const selectedOption = normalizedOptions.find(
    (option) => String(option.value) === String(value)
  );

  useEffect(() => {
    const handleOutsideClick = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setOpen(false);
      }
    };
    const handleEscape = (event) => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("pointerdown", handleOutsideClick);
    document.addEventListener("keydown", handleEscape);
    return () => {
      document.removeEventListener("pointerdown", handleOutsideClick);
      document.removeEventListener("keydown", handleEscape);
    };
  }, []);

  return (
    <div
      ref={dropdownRef}
      className={`themed-dropdown ${className} ${open ? "is-open" : ""}`.trim()}
    >
      <button
        type="button"
        className="themed-dropdown-trigger"
        aria-label={ariaLabel}
        aria-haspopup="listbox"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((current) => !current)}
      >
        <span className="themed-dropdown-value">
          {selectedOption?.label ?? placeholder}
        </span>
        <ChevronDown size={17} aria-hidden="true" />
      </button>

      {open && !disabled && (
        <div className="themed-dropdown-menu" role="listbox" aria-label={ariaLabel}>
          {normalizedOptions.map((option) => {
            const selected = String(option.value) === String(value);
            return (
              <button
                key={`${String(option.value)}-${option.label}`}
                type="button"
                role="option"
                aria-selected={selected}
                className={`themed-dropdown-option ${selected ? "selected" : ""}`}
                onClick={() => {
                  onChange(option.value);
                  setOpen(false);
                }}
              >
                <span>{option.label}</span>
                {selected && <Check size={17} strokeWidth={2.7} aria-hidden="true" />}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function Stat({ title, value, icon: Icon, tone, meta }) { return <div className="stat-card"><div className={`stat-icon ${tone}`}><Icon size={18} /></div><span className="eyebrow">{title}</span><strong>{value}</strong><small>{meta}</small></div>; }
function InputView({ block, setBlock, autoSubBlock, paymentSubBlock, setPaymentSubBlock, resident, setResident, clean, setClean, kas, setKas, other, setOther, paymentPeriod, setPaymentPeriod, addEntry, residents, editingEntry, setEditingEntry }) {
  const [previewOpen, setPreviewOpen] = useState(false);
  const totalPreview =
  Number(clean || 0) +
  Number(kas || 0) +
  Number(other || 0);
  const normalizedBlock = block.trim().toUpperCase();
  const subBlocksForBlock = [...new Set(residents.filter((r) => r.block.trim().toUpperCase() === normalizedBlock).map((r) => r.subBlock).filter(Boolean))].sort();
  const matchingResident = residents.find((r) => r.block.trim().toUpperCase() === normalizedBlock && r.subBlock === paymentSubBlock);
  const syncResident = (nextBlock, nextSubBlock) => {
    const match = residents.find((r) => r.block.trim().toUpperCase() === nextBlock.trim().toUpperCase() && r.subBlock === nextSubBlock);
    setResident(match?.name || "");
  };
  return <section className="content narrow"><div className="form-intro"><div className="notebook-icon"><BookOpen size={25} /></div><div><h2>Catat iuran warga</h2><p>Isi sekali, total dihitung otomatis. Praktis buat bendahara.</p></div></div><form
  className="form-card"
  onSubmit={(e) => {
    e.preventDefault();
    setPreviewOpen(true);
  }}
><div className="form-grid"><label>
  Periode pembayaran
  <ThemedDropdown
    value={paymentPeriod}
    onChange={setPaymentPeriod}
    options={getDaftarPeriode().map((periode) => ({ value: periode, label: periode }))}
    className="input-select"
    ariaLabel="Periode pembayaran"
  />
</label><label>Blok<input value={block} onChange={(e) => { const nextBlock = e.target.value.toUpperCase(); const firstSub = residents.find((r) => r.block.trim().toUpperCase() === nextBlock)?.subBlock || ""; setBlock(nextBlock); setPaymentSubBlock(firstSub); syncResident(nextBlock, firstSub); }} placeholder="Contoh: A" required /></label><label>Sub blok<ThemedDropdown
  value={paymentSubBlock}
  onChange={(nextValue) => { setPaymentSubBlock(nextValue); syncResident(block, nextValue); }}
  disabled={!normalizedBlock}
  placeholder={normalizedBlock ? "Pilih sub blok" : "Isi blok dahulu"}
  options={[
    { value: "", label: normalizedBlock ? "Pilih sub blok" : "Isi blok dahulu" },
    ...subBlocksForBlock.map((sub) => ({
      value: sub,
      label: `${sub} - ${residents.find((r) => r.block.trim().toUpperCase() === normalizedBlock && (r.subBlock || "").trim() === sub.trim())?.name || "Nama belum diisi"}`
    }))
  ]}
  className="input-select"
  ariaLabel="Sub blok"
/><small className="field-hint">Pilihan diambil dari Data Warga</small></label><label>Nama warga<input value={matchingResident?.name || resident} readOnly placeholder="Otomatis dari blok dan sub blok" /><small className="field-hint">Terisi otomatis dari Data Warga</small></label>{

/* INPUT IURAN KEBERSIHAN */}
<label>
  Iuran kebersihan
  <div className="input-money">
    <span>Rp</span>
    <input
      type="text"
      inputMode="numeric"
      pattern="[0-9.]*"
      value={formatNominalInput(clean)}
      onChange={(e) =>
        setClean(e.target.value.replace(/\D/g, ""))
      }
    />
  </div>
</label>

{/* INPUT IURAN KAS RT */}
<label>
  Iuran kas RT
  <div className="input-money">
    <span>Rp</span>
    <input
      type="text"
      inputMode="numeric"
      pattern="[0-9.]*"
      value={formatNominalInput(kas)}
      onChange={(e) =>
        setKas(e.target.value.replace(/\D/g, ""))
      }
    />
  </div>
</label>

{/* INPUT IURAN LAINNYA */}
<label>
  Iuran lainnya <em>opsional</em>
  <div className="input-money">
    <span>Rp</span>
    <input
      type="text"
      inputMode="numeric"
      pattern="[0-9.]*"
      value={formatNominalInput(other)}
      onChange={(e) =>
        setOther(e.target.value.replace(/\D/g, ""))
      }
    />
  </div>
</label></div><div className="total-row"><div><span>Total pembayaran</span><strong>{money(Number(clean || 0) + Number(kas || 0) + Number(other || 0))}</strong></div><button className="primary-btn" type="submit"><Plus size={18} /> {editingEntry ? "Simpan perubahan" : "Simpan pembayaran"}</button>{editingEntry && <button className="cancel-btn" type="button" onClick={() => setEditingEntry(null)}><X size={16} /> Batal</button>}
</div>
</form>

{/* POPUP KONFIRMASI PEMBAYARAN */}
{previewOpen && (
  <div className="payment-preview-backdrop">
    <div
      className="payment-preview-modal"
      role="dialog"
      aria-modal="true"
    >
      <div className="payment-preview-header">
        <div>
          <span className="eyebrow">KONFIRMASI TRANSAKSI</span>
          <h2>Periksa Pembayaran</h2>
        </div>
        <button
          type="button"
          onClick={() => setPreviewOpen(false)}
          className="payment-preview-close"
        >
          ×
        </button>
      </div>

      <div className="payment-preview-info">
        <div>
          <span>Nama warga</span>
          <strong>{matchingResident?.name || resident || "-"}</strong>
        </div>
        <div>
          <span>Blok / Sub blok</span>
          <strong>{block} / {paymentSubBlock || "-"}</strong>
        </div>
        <div>
          <span>Periode pembayaran</span>
          <strong>{paymentPeriod}</strong>
        </div>
        <div>
          <span>Tanggal pembayaran</span>
          <strong>
            {editingEntry?.date || getTanggalTransaksi()}
          </strong>
        </div>
      </div>

      <div className="payment-preview-details">
        <h3>Rincian Pembayaran</h3>

        <div>
          <span>Iuran kebersihan</span>
          <strong>{money(Number(clean || 0))}</strong>
        </div>

        <div>
          <span>Iuran kas RT</span>
          <strong>{money(Number(kas || 0))}</strong>
        </div>

        <div>
          <span>Iuran lainnya</span>
          <strong>{money(Number(other || 0))}</strong>
        </div>
      </div>

      <div className="payment-preview-total">
        <span>Total Pembayaran</span>
        <strong>{money(totalPreview)}</strong>
      </div>

      <p className="payment-preview-note">
        Pastikan nama warga dan nominal pembayaran
        sudah benar sebelum menyimpan.
      </p>

      <div className="payment-preview-actions">
        <button
          type="button"
          className="cancel-btn"
          onClick={() => setPreviewOpen(false)}
        >
          Batal
        </button>

        <button
          type="button"
          className="primary-btn"
          onClick={() => {
            addEntry();
            setPreviewOpen(false);
          }}
        >
          ✓ {editingEntry ? "Simpan Perubahan" : "Simpan Pembayaran"}
        </button>
      </div>
    </div>
  </div>
)}

<div className="tip">
<span>✦</span><div><b>Nama warga otomatis</b><p>Pilih blok dan sub blok, nama akan diambil langsung dari Data Warga.</p></div></div></section>;
}
function Cashflow({
  entries,
  cashOut,
  runningBalance,
  addExpense,
  expenseTitle,
  setExpenseTitle,
  expenseAmount,
  setExpenseAmount,
  expensePeriod,
  setExpensePeriod,
  editingExpense,
  setEditingExpense,
  startEditExpense,
  deleteExpense
}) {
  const currentYear = String(new Date().getFullYear());
  const getYear = (value) => {
    const match = String(value || "").match(/\b(?:19|20)\d{2}\b/);
    return match ? match[0] : currentYear;
  };
  const years = [...new Set([
    currentYear,
    ...entries.map((item) => getYear(item.date || item.period)),
    ...cashOut.map((item) => getYear(item.date || item.period))
  ])].sort((a, b) => Number(b) - Number(a));
  const [selectedYear, setSelectedYear] = useState(currentYear);
  const [cashType, setCashType] = useState("semua");
  const [selectedExpense, setSelectedExpense] = useState(null);

  useEffect(() => {
    if (selectedExpense && !cashOut.some((item) => item.id === selectedExpense.id)) {
      setSelectedExpense(null);
    }
  }, [cashOut, selectedExpense]);

  const yearEntries = entries.filter((item) =>
    getYear(item.date || item.period) === selectedYear
  );
  const yearExpenses = cashOut.filter((item) =>
    getYear(item.date || item.period) === selectedYear
  );

  const cleanTotal = yearEntries.reduce((sum, item) => sum + Number(item.clean || 0), 0);
  const kasTotal = yearEntries.reduce((sum, item) => sum + Number(item.kas || 0), 0);
  const otherTotal = yearEntries.reduce((sum, item) => sum + Number(item.other || 0), 0);
  const incoming = cleanTotal + kasTotal + otherTotal;
  const outgoing = yearExpenses.reduce((sum, item) => sum + Number(item.amount || 0), 0);

  const transactions = [
    ...yearEntries.map((item) => ({
      id: `in-${item.id}`,
      date: item.date || item.period || "-",
      title: `Iuran ${item.resident || "warga"}`,
      type: "masuk",
      amount: Number(item.clean || 0) + Number(item.kas || 0) + Number(item.other || 0),
      raw: item
    })),
    ...yearExpenses.map((item) => ({
      id: `out-${item.id}`,
      date: item.date || item.period || "-",
      title: item.title || "Pengeluaran",
      type: "keluar",
      amount: Number(item.amount || 0),
      raw: item
    }))
  ].filter((item) => cashType === "semua" || item.type === cashType);

  const dateValue = (value) => {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) return parsed.getTime();
    const monthIndex = namaBulan.findIndex((month) =>
      String(value).toLowerCase().includes(month.toLowerCase())
    );
    const year = Number(getYear(value));
    const day = Number(String(value).match(/\b\d{1,2}\b/)?.[0] || 1);
    return monthIndex >= 0 ? new Date(year, monthIndex, day).getTime() : 0;
  };
  transactions.sort((a, b) => dateValue(b.date) - dateValue(a.date));

  return (
    <section className="content cashflow-page cashflow-redesign">
      <div className="recap-banner cashflow-banner">
        <div>
          <span className="eyebrow light">Catatan keuangan</span>
          <h2>Arus Kas</h2>
          <p>Pantau pemasukan, pengeluaran, dan saldo kas warga.</p>
        </div>
      </div>

      <div className="cashflow-filter-row">
        <label>
          <span>Tahun</span>
          <ThemedDropdown
            value={selectedYear}
            onChange={setSelectedYear}
            options={years.map((year) => ({ value: year, label: year }))}
            className="input-select"
            ariaLabel="Filter tahun arus kas"
          />
        </label>
        <label>
          <span>Jenis kas</span>
          <ThemedDropdown
            value={cashType}
            onChange={setCashType}
            options={[
              { value: "semua", label: "Semua kas" },
              { value: "masuk", label: "Kas masuk" },
              { value: "keluar", label: "Kas keluar" }
            ]}
            className="input-select"
            ariaLabel="Filter jenis kas"
          />
        </label>
      </div>

      <div className="cashflow-summary-grid">
        <article className="cashflow-summary-card incoming">
          <span><ArrowDownLeft size={17} /> Total kas masuk</span>
          <strong>{money(incoming)}</strong>
          <small>Tahun {selectedYear}</small>
        </article>
        <article className="cashflow-summary-card outgoing">
          <span><ArrowUpRight size={17} /> Total kas keluar</span>
          <strong>{money(outgoing)}</strong>
          <small>Tahun {selectedYear}</small>
        </article>
        <article className="cashflow-summary-card balance">
          <span><WalletCards size={17} /> Saldo kas berjalan</span>
          <strong>{money(runningBalance)}</strong>
          <small>Akumulasi seluruh periode</small>
        </article>
      </div>

      <section className="cashflow-category-card">
        <div className="cashflow-section-heading">
          <div>
            <span className="eyebrow">Rincian pemasukan</span>
            <h3>Total iuran tahun {selectedYear}</h3>
          </div>
        </div>
        <div className="cashflow-category-row">
          <span><i className="category-dot clean" /> Iuran kebersihan</span>
          <strong>{money(cleanTotal)}</strong>
        </div>
        <div className="cashflow-category-row">
          <span><i className="category-dot kas" /> Iuran kas RT</span>
          <strong>{money(kasTotal)}</strong>
        </div>
        <div className="cashflow-category-row">
          <span><i className="category-dot other" /> Iuran lainnya</span>
          <strong>{money(otherTotal)}</strong>
        </div>
      </section>

      <details className="cashflow-expense-details" open={Boolean(editingExpense)}>
        <summary><Plus size={18} /> {editingExpense ? "Ubah pengeluaran" : "Input pengeluaran"}</summary>
        <form className="expense-form" onSubmit={addExpense}>
          <div>
            <span className="eyebrow">Catatan kas keluar</span>
            <h2>{editingExpense ? "Simpan perubahan" : "Catat uang keluar"}</h2>
          </div>
          <label>
            Periode pengeluaran
            <ThemedDropdown
              value={expensePeriod}
              onChange={setExpensePeriod}
              options={getDaftarPeriode().map((period) => ({ value: period, label: period }))}
              className="input-select"
              ariaLabel="Periode pengeluaran"
            />
          </label>
          <input
            placeholder="Keterangan pengeluaran"
            value={expenseTitle}
            onChange={(event) => setExpenseTitle(event.target.value)}
            required
          />
          <div className="input-money">
            <span>Rp</span>
            <input
              type="number"
              min="1"
              value={expenseAmount}
              onChange={(event) => setExpenseAmount(event.target.value)}
              required
            />
          </div>
          <div className="cashflow-form-actions">
            <button className="primary-btn" type="submit">
              <Plus size={18} /> {editingExpense ? "Simpan perubahan" : "Simpan pengeluaran"}
            </button>
            {editingExpense && (
              <button type="button" className="secondary-btn" onClick={() => setEditingExpense(null)}>
                Batal
              </button>
            )}
          </div>
        </form>
      </details>

      <section className="cashflow-activity">
        <div className="cashflow-section-heading">
          <div>
            <span className="eyebrow">Riwayat transaksi</span>
            <h3>Aktivitas kas</h3>
          </div>
          <small>{transactions.length} transaksi</small>
        </div>
        {transactions.length ? (
          <div className="cashflow-transaction-list">
            {transactions.map((item) => (
              <article
                className={`cashflow-transaction-card ${item.type === "keluar" ? "is-clickable" : ""}`}
                key={item.id}
                role={item.type === "keluar" ? "button" : undefined}
                tabIndex={item.type === "keluar" ? 0 : undefined}
                onClick={item.type === "keluar" ? () => setSelectedExpense(item.raw) : undefined}
                onKeyDown={item.type === "keluar" ? (event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    setSelectedExpense(item.raw);
                  }
                } : undefined}
              >
                <div className={`cashflow-transaction-icon ${item.type}`}>
                  {item.type === "masuk" ? <ArrowDownLeft size={19} /> : <ArrowUpRight size={19} />}
                </div>
                <div className="cashflow-transaction-main">
                  <strong>{item.title}</strong>
                  <span>{item.date} · {item.type === "masuk" ? "Kas masuk" : "Kas keluar"}</span>
                </div>
                <div className={`cashflow-transaction-amount ${item.type}`}>
                  {item.type === "masuk" ? "+" : "−"}{money(item.amount)}
                  {item.type === "keluar" && <small className="cashflow-tap-hint">Lihat rincian</small>}
                </div>
              </article>
            ))}
          </div>
        ) : (
          <div className="cashflow-empty">Belum ada transaksi untuk filter yang dipilih.</div>
        )}
      </section>

      {selectedExpense && (
        <div className="cashflow-detail-backdrop" onClick={() => setSelectedExpense(null)}>
          <section
            className="cashflow-detail-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="cashflow-detail-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="cashflow-detail-heading">
              <div>
                <span className="eyebrow">Rincian pengeluaran</span>
                <h3 id="cashflow-detail-title">{selectedExpense.title || "Pengeluaran"}</h3>
              </div>
              <button type="button" className="cashflow-detail-close" aria-label="Tutup rincian" onClick={() => setSelectedExpense(null)}>
                <X size={19} />
              </button>
            </div>
            <div className="cashflow-detail-amount">
              <span>Total pengeluaran</span>
              <strong>{money(selectedExpense.amount)}</strong>
            </div>
            <div className="cashflow-detail-row">
              <span>Tanggal transaksi</span>
              <strong>{selectedExpense.date || "-"}</strong>
            </div>
            <div className="cashflow-detail-row">
              <span>Periode</span>
              <strong>{selectedExpense.period || "-"}</strong>
            </div>
            <div className="cashflow-detail-actions">
              <button
                type="button"
                className="primary-btn"
                onClick={() => {
                  const expense = selectedExpense;
                  setSelectedExpense(null);
                  startEditExpense(expense);
                }}
              >
                <Pencil size={16} /> Edit pengeluaran
              </button>
              <button
                type="button"
                className="danger-btn"
                onClick={() => {
                  const id = selectedExpense.id;
                  setSelectedExpense(null);
                  deleteExpense(id);
                }}
              >
                <Trash2 size={16} /> Hapus pengeluaran
              </button>
            </div>
          </section>
        </div>
      )}
    </section>
  );
}

function ResidentReport({
  entries,
  cashOut,
  runningBalance,
  openingBalance = 0,
  manualAdjustment = 0,
  communityIdentityLine = "RT 26 / RW 04",
  communityLocationLine = ""
}) {
  const [reportPeriod, setReportPeriod] = useState(getPeriodeSekarang());

  // Data pemasukan sesuai periode yang dipilih
  const reportEntries = sortPaymentEntries(
    entries.filter(e => cocokPeriode(e, reportPeriod))
  );

  const reportExpenses = cashOut.filter(
    e => cocokPeriode(e, reportPeriod)
  );

  // Total pemasukan keseluruhan
  const incoming = reportEntries.reduce(
    (sum, e) =>
      sum +
      Number(e.clean || 0) +
      Number(e.kas || 0) +
      Number(e.other || 0),
    0
  );

  // Total pengeluaran
  const outgoing = reportExpenses.reduce(
    (sum, e) => sum + Number(e.amount || 0),
    0
  );

  // Saldo laporan
  const reportBalance =
    reportPeriod === "Semua periode"
      ? runningBalance
      : Number(openingBalance || 0) +
        incoming -
        outgoing +
        Number(manualAdjustment || 0);

  // =====================================
  // REKAP PEMASUKAN BERDASARKAN BLOK
  // DAN PERIODE IURAN
  // =====================================

  const groupedIncome = Object.values(
    reportEntries.reduce((groups, e) => {

      const block = String(e.block || "Tanpa blok")
        .trim()
        .toUpperCase();

      // Gunakan periode iuran, bukan tanggal pembayaran
      const period =
        e.period || periodeDariTanggal(e.date) || "Periode tidak diketahui";

      const key = `${block}|${period}`;

      if (!groups[key]) {
        groups[key] = {
          block,
          period,
          total: 0
        };
      }

      // Jumlahkan seluruh kategori iuran warga
      groups[key].total +=
        Number(e.clean || 0) +
        Number(e.kas || 0) +
        Number(e.other || 0);

      return groups;

    }, {})
  );

  // Urutkan berdasarkan periode, kemudian blok
  groupedIncome.sort((a, b) => {
    const [monthA, yearA] = a.period.split(" ");
    const [monthB, yearB] = b.period.split(" ");

    const indexA = namaBulan.indexOf(monthA);
    const indexB = namaBulan.indexOf(monthB);

    const orderA =
      Number(yearA) * 12 + (indexA >= 0 ? indexA : 0);

    const orderB =
      Number(yearB) * 12 + (indexB >= 0 ? indexB : 0);

    return (
      orderA - orderB ||
      compareNatural(a.block, b.block)
    );
  });

  return (
    <section className="content report-page">

      <div className="report-toolbar">
        <div>
          <span className="eyebrow">Dokumen warga</span>
          <h2>Laporan Pertanggungjawaban</h2>
          <p className="subcopy">
            Ringkasan kas {communityIdentityLine}.
          </p>
        </div>

        <button
          className="primary-btn"
          onClick={() => window.print()}
        >
          <FileText size={17} />
          Cetak / Simpan PDF
        </button>
      </div>

      <label className="report-period-control">
        Pilih bulan laporan

        <ThemedDropdown
          value={reportPeriod}
          onChange={setReportPeriod}
          options={[{ value: "Semua periode", label: "Semua periode" }, ...getDaftarPeriode().map(periode => ({ value: periode, label: periode }))]}
          className="input-select"
          ariaLabel="Periode laporan"
        />
      </label>

      <div className="report-paper">

        <div className="report-heading">
          <div>
            <div className="report-kicker">
              {communityIdentityLine}
            </div>
            {communityLocationLine && <p className="report-community-location">{communityLocationLine}</p>}

            <h1>Laporan Pertanggungjawaban Kas</h1>

            <p>
              {reportPeriod === "Semua periode"
                ? "Semua periode"
                : `Periode ${reportPeriod}`}
            </p>
          </div>
        </div>

        <p className="report-intro">
          Ringkasan pemasukan dan pengeluaran kas warga
          untuk diketahui bersama.
        </p>

        {/* RINGKASAN KEUANGAN */}
        <div className="report-stats">
          <div>
            <span>Total pemasukan</span>
            <b>{money(incoming)}</b>
          </div>

          <div>
            <span>Total pengeluaran</span>
            <b>{money(outgoing)}</b>
          </div>

          <div>
            <span>Saldo kas saat ini</span>
            <b>{money(reportBalance)}</b>
          </div>
        </div>

        {/* RINCIAN PEMASUKAN PER BLOK */}
        <h3>Rincian pemasukan</h3>

        <table>
          <tbody>
            {groupedIncome.length > 0 ? (
              groupedIncome.map((item, index) => (
                <tr key={`${item.block}-${item.period}-${index}`}>
                  <td>
                    Pemasukan iuran dari blok {item.block}
                    {" "}periode {item.period}
                  </td>

                  <td>
                    <strong>{money(item.total)}</strong>
                  </td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="2">
                  Tidak ada pemasukan pada periode ini.
                </td>
              </tr>
            )}
          </tbody>
        </table>

        {/* RINCIAN PENGELUARAN TETAP */}
        <h3>Rincian pengeluaran</h3>

        <table>
          <tbody>
            {reportExpenses.length > 0 ? (
              reportExpenses.map(e => (
                <tr key={e.id}>
                  <td>{e.title}</td>
                  <td>{e.date}</td>
                  <td>{money(e.amount)}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="3">
                  Tidak ada pengeluaran pada periode ini.
                </td>
              </tr>
            )}
          </tbody>
        </table>

        <div className="report-footer">
          <p>
            Demikian laporan ini dibuat untuk dapat
            diketahui bersama.
          </p>

          <div>
            <span>Bendahara RT</span>
            <strong>Bendahara RT</strong>
          </div>
        </div>

      </div>
    </section>
  );
}
function PublicSummary({ entries, cashOut, openingBalance, manualAdjustment, communityIdentityLine = "RT 26 / RW 04", communityLocationLine = "" }) {
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
    <div className="public-heading"><div><span className="eyebrow">Informasi untuk warga</span><h2>Laporan Kas Warga</h2><p>{communityIdentityLine}{communityLocationLine ? ` · ${communityLocationLine}` : ""} · Ringkasan terbuka pengelolaan kas</p></div><div className="public-controls"><label>Pilih periode<ThemedDropdown value={period} onChange={setPeriod} options={[{value:"Semua periode",label:"Semua periode"}, ...getDaftarPeriode().map(periode=>({value:periode,label:periode}))]} className="public-period-select" ariaLabel="Pilih periode ringkasan publik" /></label><button className="primary-btn" onClick={() => window.print()}><FileText size={16}/> Cetak / Simpan PDF</button></div></div>
    <div className="privacy-note"><span>✓</span><div><b>Tampilan publik, data pribadi disembunyikan</b><p>Tidak menampilkan nama warga, nomor rumah, blok, atau rincian pembayaran per orang.</p></div></div>
    <div className="public-kpis"><article className="public-kpi income"><span>Total pemasukan</span><strong>{money(income)}</strong><small>Periode {period}</small></article><article className="public-kpi expense"><span>Total pengeluaran</span><strong>{money(expense)}</strong><small>Periode {period}</small></article><article className="public-kpi balance"><span>Saldo akhir</span><strong>{money(balance)}</strong><small>Saldo awal + masuk - keluar + koreksi</small></article></div>
    <div className="public-breakdown"><div className="section-head"><div><span className="eyebrow">Rangkuman kategori</span><h2>Rangkuman kategori iuran</h2></div></div><div className="public-breakdown-row"><span>Iuran kebersihan</span><b>{money(cleanTotal)}</b></div><div className="public-breakdown-row"><span>Iuran kas RT</span><b>{money(kasTotal)}</b></div><div className="public-breakdown-row"><span>Iuran lainnya</span><b>{money(otherTotal)}</b></div><div className="public-breakdown-row"><span>Total pengeluaran</span><b>{money(expense)}</b></div></div>
    <footer className="public-footer"><span>Kas Warga {communityIdentityLine}{communityLocationLine ? ` · ${communityLocationLine}` : ""}</span><span>Ringkasan agregat · {period}</span></footer>
  </section>;
}

// Memuat html2canvas hanya saat pengguna meminta ekspor JPG HD.
const loadHtml2Canvas = () =>
  new Promise((resolve, reject) => {
    if (typeof window !== "undefined" && typeof window.html2canvas === "function") {
      resolve(window.html2canvas);
      return;
    }

    const existingScript = document.querySelector("script[data-html2canvas-jpg]");
    if (existingScript) {
      existingScript.addEventListener("load", () => resolve(window.html2canvas), { once: true });
      existingScript.addEventListener("error", () => reject(new Error("Library ekspor JPG gagal dimuat.")), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
    script.async = true;
    script.dataset.html2canvasJpg = "true";
    script.onload = () => {
      if (typeof window.html2canvas === "function") resolve(window.html2canvas);
      else reject(new Error("Library ekspor JPG tidak tersedia."));
    };
    script.onerror = () => reject(new Error("Tidak dapat memuat library ekspor JPG. Periksa koneksi internet."));
    document.head.appendChild(script);
  });

// JPG dibuat dari konten Rekap Bulanan yang sama dengan PDF:
// periode, ringkasan, kategori iuran, dan tabel mengikuti filter aktif.
const exportRecapJpgHD = async (period) => {
  const html2canvas = await loadHtml2Canvas();
  const source = document.querySelector(".recap-page");
  if (!source) throw new Error("Area Rekap Bulanan tidak ditemukan.");

  const host = document.createElement("div");
  host.className = "recap-jpg-export-host";
  host.style.cssText = [
    "position:fixed",
    "left:-12000px",
    "top:0",
    "width:794px",
    "min-width:794px",
    "max-width:794px",
    "z-index:-1",
    "opacity:1",
    "pointer-events:none",
    "background:#ffffff",
    "overflow:visible"
  ].join(";");

  const clone = source.cloneNode(true);
  clone.classList.add("recap-jpg-capture");
  // Lebar area konten dibuat setara area cetak A4 (margin kiri/kanan 10 mm).
  // Margin atas/bawah diterapkan lagi pada setiap kanvas halaman di bawah.
  clone.style.cssText = "width:794px!important;min-width:794px!important;max-width:794px!important;margin:0!important;padding:0 24px!important;overflow:visible!important;background:#fff!important;box-sizing:border-box!important;";

  // PDF tidak menampilkan kontrol filter, tombol, maupun popup detail.
  clone.querySelectorAll(
    ".recap-actions, .recap-detail-toolbar, .payment-detail-backdrop, button, input, select, textarea"
  ).forEach((node) => node.remove());

  const printLikeStyle = document.createElement("style");
  printLikeStyle.textContent = `
    .recap-jpg-capture, .recap-jpg-capture * {
      box-sizing: border-box !important;
      font-family: Arial, Helvetica, sans-serif !important;
      animation: none !important;
      transition: none !important;
      text-shadow: none !important;
    }
    .recap-jpg-capture {
      display:block!important; width:794px!important; min-width:794px!important;
      max-width:794px!important; margin:0!important; padding:0 24px!important;
      color:#20352E!important; background:#FFFFFF!important; overflow:visible!important;
      font-size:9pt!important;
    }
    .recap-jpg-capture .recap-banner {
      display:block!important; min-height:0!important; height:auto!important;
      margin:0 0 8px!important; padding:10px 12px!important;
      color:#102A43!important; background:#EFF6FF!important;
      border:1px solid #BFDBFE!important; border-radius:7px!important;
      box-shadow:none!important;
    }
    .recap-jpg-capture .recap-banner h2 {
      margin:2px 0 4px!important; color:#102A43!important;
      font-size:17pt!important; line-height:1.2!important;
    }
    .recap-jpg-capture .recap-banner p {
      margin:2px 0 0!important; color:#475569!important; font-size:8.5pt!important;
    }
    .recap-jpg-capture .recap-banner .eyebrow {
      color:#245789!important; font-size:7pt!important; font-weight:700!important;
    }
    .recap-jpg-capture .community-report-identity {
      display:block!important; margin:4px 0 0!important; padding:0!important;
      color:#244563!important; font-size:8pt!important; font-weight:700!important;
      line-height:1.3!important; overflow-wrap:anywhere!important;
    }
    .recap-jpg-capture .recap-grid {
      display:grid!important; grid-template-columns:repeat(2,minmax(0,1fr))!important;
      align-items:start!important; gap:7px!important; margin:7px 0 8px!important;
    }
    .recap-jpg-capture .big-card {
      min-width:0!important; min-height:0!important; height:auto!important;
      margin:0!important; padding:8px 9px!important; color:#20352E!important;
      background:#FFFFFF!important; border:1px solid #DCE5EF!important;
      border-radius:6px!important; box-shadow:none!important;
      break-inside:avoid!important; page-break-inside:avoid!important;
    }
    .recap-jpg-capture .big-card strong {
      display:block!important; margin:2px 0!important; color:#102A43!important;
      font-size:13pt!important; line-height:1.2!important;
    }
    .recap-jpg-capture .big-card small,
    .recap-jpg-capture .big-card span { color:#52665D!important; font-size:7.5pt!important; }
    .recap-jpg-capture .bar { height:4px!important; margin:4px 0!important; background:#DBEAFE!important; }
    .recap-jpg-capture .bar i { background:#2563EB!important; }
    .recap-jpg-capture .recap-income-details { margin-top:6px!important; }
    .recap-jpg-capture .recap-income-details > div {
      display:flex!important; justify-content:space-between!important; gap:8px!important;
      padding:3px 0!important; font-size:7.5pt!important;
    }
    .recap-jpg-capture .section-head {
      display:flex!important; justify-content:space-between!important; align-items:end!important;
      gap:6px!important; margin:9px 0 5px!important; padding:0 0 4px!important;
      border-bottom:1px solid #DCE5EF!important;
    }
    .recap-jpg-capture .section-head h2 {
      margin:1px 0 3px!important; color:#102A43!important;
      font-size:10.5pt!important; line-height:1.25!important;
    }
    .recap-jpg-capture .eyebrow {
      color:#245789!important; font-size:7pt!important; font-weight:700!important;
    }
    .recap-jpg-capture .table-card {
      display:block!important; width:100%!important; min-width:0!important;
      max-width:100%!important; margin:0!important; padding:0!important;
      overflow:visible!important; background:#fff!important;
      border:1px solid #DCE5EF!important; border-radius:6px!important;
      box-shadow:none!important;
    }
    .recap-jpg-capture table {
      display:table!important; width:100%!important; min-width:0!important;
      max-width:100%!important; margin:0!important; border-collapse:collapse!important;
      table-layout:fixed!important; font-size:7pt!important; color:#263B33!important;
    }
    .recap-jpg-capture thead { display:table-header-group!important; }
    .recap-jpg-capture tbody { display:table-row-group!important; }
    .recap-jpg-capture tr { break-inside:avoid!important; page-break-inside:avoid!important; }
    .recap-jpg-capture th, .recap-jpg-capture td {
      padding:4px!important; line-height:1.25!important; white-space:normal!important;
      overflow-wrap:anywhere!important; word-break:normal!important; vertical-align:top!important;
      border-bottom:1px solid #E2EAE5!important;
    }
    .recap-jpg-capture th {
      color:#475569!important; background:#F1F5F9!important;
      font-size:6.5pt!important; font-weight:700!important;
    }
    .recap-jpg-capture tbody tr:nth-child(even) { background:#F8FBF9!important; }
    .recap-jpg-capture .status-paid, .recap-jpg-capture .status-unpaid,
    .recap-jpg-capture .block-tag {
      display:inline-block!important; padding:2px 4px!important; font-size:6.5pt!important;
      border-radius:4px!important;
    }
    .recap-jpg-capture .status-paid {
      color:#065F46!important; background:#E6F4ED!important; border:1px solid #C8E7D7!important;
    }
    .recap-jpg-capture .status-unpaid {
      color:#991B1B!important; background:#FEE2E2!important; border:1px solid #FECACA!important;
    }
    .recap-jpg-capture .block-tag {
      color:#173B65!important; background:#EFF6FF!important; border:1px solid #BFDBFE!important;
    }
    .recap-jpg-capture .empty-table { color:#64748B!important; }

    /* Kartu ringkasan JPG HD dibuat lebih ringkas. */
    .recap-jpg-capture .big-card {
      align-self:start!important; display:block!important; padding:7px 8px!important;
    }
    .recap-jpg-capture .big-card .eyebrow {
      display:block!important; margin:0 0 2px!important; line-height:1.15!important;
    }
    .recap-jpg-capture .big-card strong {
      margin:1px 0!important; line-height:1.12!important;
    }
    .recap-jpg-capture .big-card small {
      display:block!important; margin:0!important; line-height:1.15!important;
    }
    .recap-jpg-capture .bar { margin:3px 0!important; }
    .recap-jpg-capture .recap-income-details {
      margin:4px 0 0!important; padding-top:4px!important; gap:0!important; line-height:1.15!important;
    }
    .recap-jpg-capture .recap-income-details > div {
      min-height:0!important; padding:2px 0!important; line-height:1.15!important; font-size:7.2pt!important;
    }
    .recap-jpg-capture .recap-income-details span,
    .recap-jpg-capture .recap-income-details b {
      font-size:7.2pt!important; line-height:1.15!important;
    }
  `;

  host.appendChild(printLikeStyle);
  host.appendChild(clone);
  document.body.appendChild(host);

  try {
    if (document.fonts?.ready) await document.fonts.ready;
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

    const pageWidth = 2480;
    const pageHeight = 3508;
    const marginX = 118; // sekitar 10 mm pada 300 DPI
    const marginY = 118; // sekitar 10 mm pada 300 DPI
    const contentWidth = pageWidth - (marginX * 2);
    const pageContentHeight = pageHeight - (marginY * 2);
    const cssContentWidth = 794 - (24 * 2);
    const scale = contentWidth / cssContentWidth;
    const fullCanvas = await html2canvas(clone, {
      backgroundColor: "#FFFFFF",
      scale,
      width: 794,
      windowWidth: 794,
      scrollX: 0,
      scrollY: 0,
      useCORS: true,
      allowTaint: false,
      logging: false,
      imageTimeout: 15000
    });

    if (!fullCanvas.width || !fullCanvas.height) {
      throw new Error("Konten Rekap Bulanan tidak berhasil dirender.");
    }

    const pageCount = Math.ceil(fullCanvas.height / pageContentHeight);
    const slug = String(period || "semua-periode")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "");

    for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = pageWidth;
      pageCanvas.height = pageHeight;
      const context = pageCanvas.getContext("2d");
      context.fillStyle = "#FFFFFF";
      context.fillRect(0, 0, pageWidth, pageHeight);

      const sourceY = pageIndex * pageContentHeight;
      const sliceHeight = Math.min(pageContentHeight, fullCanvas.height - sourceY);
      context.drawImage(
        fullCanvas,
        0, sourceY, fullCanvas.width, sliceHeight,
        0, marginY, pageWidth, sliceHeight
      );

      const blob = await new Promise((resolve, reject) => {
        pageCanvas.toBlob(
          (result) => result ? resolve(result) : reject(new Error("Gagal membuat file JPG.")),
          "image/jpeg",
          0.95
        );
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `rekap-bulanan-${slug}${pageCount > 1 ? `-halaman-${pageIndex + 1}` : ""}-hd.jpg`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 15000);

      // Beri ruang agar browser memproses setiap unduhan untuk laporan multi-halaman.
      if (pageIndex < pageCount - 1) {
        await new Promise((resolve) => window.setTimeout(resolve, 350));
      }
    }

    return pageCount;
  } finally {
    host.remove();
  }
};

function Recap({
  communityIdentityLine = "RT 26 / RW 04",
  communityLocationLine = "",
  entries,
  total,
  residents = [],
  recapSearch,
  setRecapSearch,
  startEditEntry,
  deleteEntry
}) { const [selectedRecapPayment, setSelectedRecapPayment] = useState(null);
const [isExportingJpg, setIsExportingJpg] = useState(false);
const [exportNotice, setExportNotice] = useState(null);
const [recapPeriod, setRecapPeriod] = useState(getPeriodeSekarang());

// Filter status iuran
const [recapStatusFilter, setRecapStatusFilter] = useState("lunas");
const [recapBlockFilter, setRecapBlockFilter] = useState("semua");

const monthEntries = sortPaymentEntries(
  entries.filter((e) => cocokPeriode(e, recapPeriod))
);

// Memeriksa apakah warga sudah membayar pada periode terpilih
const isResidentPaid = (resident) =>
  monthEntries.some((payment) => {
    const sameId =
      payment.residentId != null && resident.id != null &&
      String(payment.residentId) === String(resident.id);

    const paymentName = (payment.resident || "").trim().toLocaleLowerCase("id-ID");
    const residentName = (resident.name || "").trim().toLocaleLowerCase("id-ID");
    const sameName = Boolean(paymentName && residentName && paymentName === residentName);

    const paymentBlock = String(payment.block || "").trim().toLocaleLowerCase("id-ID");
    const residentBlock = String(resident.block || "").trim().toLocaleLowerCase("id-ID");
    const paymentSubBlock = String(payment.subBlock || "").trim().toLocaleLowerCase("id-ID");
    const residentSubBlock = String(resident.subBlock || "").trim().toLocaleLowerCase("id-ID");
    const sameAddress = Boolean(paymentBlock && residentBlock && paymentBlock === residentBlock &&
      paymentSubBlock && residentSubBlock && paymentSubBlock === residentSubBlock);

    return sameId || sameName || sameAddress;
  });

// Membuat daftar warga yang belum membayar
const unpaidRows = residents
  .filter((resident) => !isResidentPaid(resident))
  .map((resident) => ({
    id: `unpaid-${resident.id}`,
    residentId: resident.id,
    resident: resident.name,
    block: resident.block,
    subBlock: resident.subBlock,
    date: "-",
    period: recapPeriod,
    clean: null,
    kas: null,
    other: null,
    status: "Belum Lunas",
    isUnpaidPlaceholder: true,
  }));

// Menggabungkan data lunas dan belum lunas
const allRecapRows = [
  ...monthEntries.map((entry) => ({
    ...entry,
    status: "Lunas",
  })),
  ...unpaidRows,
];

// Filter nama warga
const searchedRows = allRecapRows.filter((entry) =>
  (entry.resident || "")
    .toLocaleLowerCase("id-ID")
    .includes(recapSearch.trim().toLocaleLowerCase("id-ID"))
);

// Filter status
const recapBlockOptions = [...new Set([
  ...residents.map((resident) => String(resident.block || "").trim().toUpperCase()),
  ...monthEntries.map((entry) => String(entry.block || "").trim().toUpperCase())
].filter(Boolean))].sort((a, b) => a.localeCompare(b, "id-ID", { numeric: true }));

const filteredEntries = searchedRows.filter((entry) => {
  const matchesStatus = recapStatusFilter === "semua" ||
    (recapStatusFilter === "lunas" && entry.status === "Lunas") ||
    (recapStatusFilter === "belum" && entry.status === "Belum Lunas");
  const matchesBlock = recapBlockFilter === "semua" ||
    String(entry.block || "").trim().toUpperCase() === recapBlockFilter;
  return matchesStatus && matchesBlock;
}); const periodTotal = monthEntries.reduce((sum, e) => sum + e.clean + e.kas + (e.other || 0), 0); const cleanTotal = monthEntries.reduce((sum, e) => sum + Number(e.clean || 0), 0); const kasTotal = monthEntries.reduce((sum, e) => sum + Number(e.kas || 0), 0); const otherTotal = monthEntries.reduce((sum, e) => sum + Number(e.other || 0), 0); const paidCount = new Set(monthEntries.map(e => e.residentId || e.resident || `${e.block}-${e.subBlock}`)).size; const paidPercent = residents.length ? Math.round((paidCount / residents.length) * 100) : 0;
const handleExportJpgHD = async () => {
  if (isExportingJpg) return;
  setIsExportingJpg(true);
  try {
    const pages = await exportRecapJpgHD(recapPeriod);
    setExportNotice({
      type: "success",
      title: "Export JPG HD berhasil",
      message: pages > 1
        ? `${pages} gambar A4 berhasil dibuat dan siap dibagikan.`
        : "Gambar JPG HD berhasil dibuat dan siap dibagikan."
    });
  } catch (error) {
    console.error("Gagal mengekspor Rekap Bulanan ke JPG HD:", error);
    setExportNotice({
      type: "error",
      title: "Export JPG HD belum berhasil",
      message: error?.message || "Silakan periksa koneksi internet dan coba kembali."
    });
  } finally {
    setIsExportingJpg(false);
  }
};
return <section className="content recap-page"><div className="recap-banner"><div><span className="eyebrow light">Laporan periode</span><h2>Rekap {recapPeriod}</h2><p>Ringkasan pemasukan kas warga.</p><small className="community-report-identity">{communityIdentityLine}{communityLocationLine ? ` · ${communityLocationLine}` : ""}</small></div><div className="recap-actions"><div className="recap-period-label"><span className="recap-period-title">Pilih bulan</span><ThemedDropdown className="recap-period-select" value={recapPeriod} onChange={setRecapPeriod} options={[{value:"Semua periode",label:"Semua periode"}, ...getDaftarPeriode().map(periode=>({value:periode,label:periode}))]} ariaLabel="Pilih bulan rekap" /></div><button className="outline-light" onClick={() => exportExcel(filteredEntries, recapPeriod, recapStatusFilter === "belum" ? "belum-lunas" : recapStatusFilter === "lunas" ? "lunas" : "semua-status", communityIdentityLine, communityLocationLine)}><FileText size={16} /> Ekspor Excel</button><button className="outline-light" onClick={() => window.print()}><FileText size={16} /> Ekspor PDF</button><button className="outline-light" type="button" onClick={handleExportJpgHD} disabled={isExportingJpg}><FileText size={16} /> {isExportingJpg ? "Menyiapkan JPG HD..." : "Ekspor JPG HD"}</button></div></div><div className="recap-grid">
  {/* Kartu total terkumpul */}
  <div className="big-card recap-total-card">
    <span className="eyebrow">Total terkumpul</span>
    <strong>{money(periodTotal)}</strong>

    <div className="bar">
      <i style={{ width: `${Math.min(paidPercent, 100)}%` }} />
    </div>

    <small>
      {residents.length
        ? `${paidPercent}% warga sudah membayar`
        : "Belum ada data warga"}
    </small>

    {/* Rincian iuran ukuran kecil */}
    <div className="recap-income-details">
      <div>
        <span>Iuran kebersihan</span>
        <b>{money(cleanTotal)}</b>
      </div>
      <div>
        <span>Iuran kas RT</span>
        <b>{money(kasTotal)}</b>
      </div>
      <div>
        <span>Iuran lainnya</span>
        <b>{money(otherTotal)}</b>
      </div>
    </div>
  </div>

  {/* Kartu warga sudah bayar */}
  <div className="big-card">
    <span className="eyebrow">Warga sudah bayar</span>
    <strong>{paidCount} / {residents.length}</strong>
    <small>
      {residents.length
        ? `${paidPercent}% tingkat pembayaran`
        : "Belum ada data warga"}
    </small>
  </div>
</div>
<div className="section-head recap-detail-head">
  <div>
    <span className="eyebrow">Per warga</span>
    <h2 className="recap-details-title">
      Rincian iuran · {recapPeriod}
    </h2>
  </div>

  <div className="recap-detail-toolbar">

    {/* Pencarian nama warga */}
    <div className="search recap-resident-search">
      <Search size={16} />
      <input
        placeholder="Cari nama warga..."
        value={recapSearch}
        onChange={e => setRecapSearch(e.target.value)}
      />
    </div>

    {/* Filter status pembayaran */}
    <ThemedDropdown
      className="recap-status-filter"
      value={recapStatusFilter}
      onChange={setRecapStatusFilter}
      ariaLabel="Filter status pembayaran"
      options={[
        { value: "semua", label: "Semua status" },
        { value: "lunas", label: "Lunas" },
        { value: "belum", label: "Belum lunas" }
      ]}
    />

    {/* Filter berdasarkan blok */}
    <ThemedDropdown
      className="recap-block-filter"
      value={recapBlockFilter}
      onChange={setRecapBlockFilter}
      ariaLabel="Filter blok warga"
      options={[
        { value: "semua", label: "Semua blok" },
        ...recapBlockOptions.map((blok) => ({ value: blok, label: `Blok ${blok}` }))
      ]}
    />

  </div>
</div>
<div className="recap-print-filters" aria-label="Filter rekap yang digunakan saat dicetak">
  <span><b>Status pembayaran:</b> {recapStatusFilter === "semua" ? "Semua status" : recapStatusFilter === "lunas" ? "Lunas" : "Belum lunas"}</span>
  <span><b>Blok:</b> {recapBlockFilter === "semua" ? "Semua blok" : `Blok ${recapBlockFilter}`}</span>
</div>
<div className="table-card recap-payment-card"><table className="recap-payment-table"><thead><tr><th>Tanggal</th><th>Nama warga</th><th>Periode</th><th>Blok</th><th>Kebersihan</th><th>Kas RT</th><th>Iuran lainnya</th><th>Total</th><th>Status</th></tr></thead><tbody>{filteredEntries.length ? (
  filteredEntries.map((e) => (
    <tr
      key={e.id}
      className={`recap-payment-row ${
        e.status === "Belum Lunas" ? "recap-unpaid-row" : ""
      }`}
      onClick={() => setSelectedRecapPayment(e)}
    >
      <td data-label="Tanggal">{e.date || "-"}</td>

      <td data-label="Nama warga">
        <b>{e.resident || "Nama belum diisi"}</b>
      </td>

      <td data-label="Periode">{e.period || recapPeriod}</td>

      <td data-label="Blok">
        <span className="block-tag">
          {e.subBlock || e.block || "-"}
        </span>
      </td>

      <td data-label="Kebersihan">{money(e.clean || 0)}</td>

      <td data-label="Kas RT">{money(e.kas || 0)}</td>

      <td data-label="Iuran lainnya">{money(e.other || 0)}</td>

      <td data-label="Total">
        <b>
          {money(
            Number(e.clean || 0) +
            Number(e.kas || 0) +
            Number(e.other || 0)
          )}
        </b>
      </td>

      <td data-label="Status">
  <span className={
    e.status === "Belum Lunas"
      ? "status-unpaid"
      : "status-paid"
  }>
    {e.status || "Lunas"}
  </span>
</td>
    </tr>
  ))
) : (
  <tr>
    <td colSpan="9" className="empty-table">
      Tidak ada data warga sesuai filter.
    </td>
  </tr>
)}</tbody></table></div>
{exportNotice && (
  <div className="recap-export-notice-backdrop" onMouseDown={(event) => {
    if (event.target === event.currentTarget) setExportNotice(null);
  }}>
    <section className="recap-export-notice" role="dialog" aria-modal="true" aria-labelledby="recap-export-notice-title">
      <div className={`recap-export-notice-icon ${exportNotice.type}`}>
        {exportNotice.type === "success" ? <Check size={25} strokeWidth={2.8} /> : <X size={24} strokeWidth={2.8} />}
      </div>
      <span className="eyebrow">Kas Warga · Rekap Bulanan</span>
      <h3 id="recap-export-notice-title">{exportNotice.title}</h3>
      <p>{exportNotice.message}</p>
      <button type="button" className="primary-btn" onClick={() => setExportNotice(null)}>Mengerti</button>
    </section>
  </div>
)}
{selectedRecapPayment && (() => {
  const payment = selectedRecapPayment;
  const isUnpaid = payment.status === "Belum Lunas" || payment.isUnpaidPlaceholder === true;
  const paymentTotal = Number(payment.clean || 0) + Number(payment.kas || 0) + Number(payment.other || 0);
  const resident = payment.resident || residents.find(r =>
    (payment.residentId != null && String(r.id) === String(payment.residentId)) ||
    (r.block === payment.block && r.subBlock === payment.subBlock)
  )?.name || "Nama belum diisi";
  return <div className="payment-detail-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) setSelectedRecapPayment(null); }}>
    <section className="payment-detail-modal recap-payment-modal" role="dialog" aria-modal="true" aria-labelledby="recap-payment-title">
      <button type="button" className="payment-detail-close" aria-label="Tutup" onClick={() => setSelectedRecapPayment(null)}><X size={19}/></button>
      <span className="eyebrow">Rekap bulanan · Per warga</span>
      <h2 id="recap-payment-title">Rincian iuran</h2>
      <div className="payment-detail-person">
        <b>{resident}</b>
        <span>Blok {payment.subBlock || payment.block || "-"}</span>
      </div>
      <div className="payment-detail-lines">
        <div><span>Periode iuran</span><b>{payment.period || recapPeriod}</b></div>
        {isUnpaid ? (
          <div><span>Tanggal pembayaran</span><b>Belum ada pembayaran</b></div>
        ) : (
          <>
            <div><span>Tanggal pembayaran</span><b>{payment.date || "-"}</b></div>
            <div><span>Iuran kebersihan</span><b>{money(payment.clean || 0)}</b></div>
            <div><span>Iuran kas RT</span><b>{money(payment.kas || 0)}</b></div>
            <div><span>Iuran lainnya</span><b>{money(payment.other || 0)}</b></div>
          </>
        )}
        <div>
          <span>Status pembayaran</span>
          <b><span className={isUnpaid ? "status-unpaid" : "status-paid"}>{isUnpaid ? "Belum Lunas" : "Lunas"}</span></b>
        </div>
      </div>
      <div className="payment-detail-total">
        <span>{isUnpaid ? "Status tagihan" : "Total pembayaran"}</span>
        <strong>{isUnpaid ? "Belum dibayar" : money(paymentTotal)}</strong>
      </div>
      <div className="payment-detail-actions recap-payment-actions">

  <button
    type="button"
    className="cancel-btn"
    onClick={() => setSelectedRecapPayment(null)}
  >
    Tutup
  </button>

  {!isUnpaid && (
    <>
      <button
        type="button"
        className="primary-btn"
        onClick={() => {
          setSelectedRecapPayment(null);
          startEditEntry(payment);
        }}
      >
        <Pencil size={15} />
        Edit
      </button>

      <button
        type="button"
        className="icon-btn delete"
        onClick={() => {
          setSelectedRecapPayment(null);
          deleteEntry(payment.id);
        }}
      >
        <Trash2 size={15} />
        Hapus
      </button>
    </>
  )}

</div>
    </section>
  </div>;
})()}
</section>; }
function Residents({ residents, residentName, setResidentName, residentBlock, setResidentBlock, residentSubBlock, setResidentSubBlock, addResident, editingResident, setEditingResident, startEditResident, deleteResident, residentSearch, setResidentSearch, residentFilter, setResidentFilter, residentSubFilter, setResidentSubFilter, communityIdentityLine = "RT 26 / RW 04", communityLocationLine = "" }) { const blocks = [...new Set(residents.map((r) => r.block).filter(Boolean))].sort(); const subBlocks = [...new Set(residents.map((r) => r.subBlock).filter(Boolean))].sort(); const query = residentSearch.trim().toLocaleLowerCase("id-ID"); const filteredResidents = residents.filter((r) => (residentFilter === "Semua blok" || r.block.toLocaleLowerCase("id-ID") === residentFilter.toLocaleLowerCase("id-ID")) && (residentSubFilter === "Semua sub blok" || (r.subBlock || "").toLocaleLowerCase("id-ID") === residentSubFilter.toLocaleLowerCase("id-ID")) && (r.name || "").toLocaleLowerCase("id-ID").includes(query)); return <section className="content narrow"><div className="resident-print-title">Daftar warga · {communityIdentityLine}{communityLocationLine ? ` · ${communityLocationLine}` : ""}</div><div className="section-head"><div><span className="eyebrow">Buku warga</span><h2>Daftar nama warga</h2><p className="subcopy">Kelola data yang muncul di menu input pembayaran.</p></div></div><form className="add-resident" onSubmit={addResident} noValidate><input placeholder="Nama lengkap warga" value={residentName} onChange={e => setResidentName(e.target.value)} /><input className="compact-field" value={residentBlock} onChange={e => setResidentBlock(e.target.value)} placeholder="Blok, contoh A" /><input className="compact-field" value={residentSubBlock} onChange={e => setResidentSubBlock(e.target.value)} placeholder="Sub blok, contoh A1" /><button className="primary-btn" type="submit"><Plus size={18} /> {editingResident ? "Simpan perubahan" : "Tambah warga"}</button>{editingResident && <button className="cancel-btn" type="button" onClick={() => { setEditingResident(null); setResidentName(""); }}><X size={16} /> Batal</button>}</form><div className="resident-toolbar"><div className="search resident-search"><Search size={16} /><input placeholder="Cari nama warga..." value={residentSearch} onChange={e => setResidentSearch(e.target.value)} /></div><ThemedDropdown className="resident-filter resident-block-filter" value={residentFilter} onChange={(nextValue) => { setResidentFilter(nextValue); setResidentSubFilter("Semua sub blok"); }} ariaLabel="Filter berdasarkan blok" options={[{value:"Semua blok",label:"Semua blok"}, ...blocks.map(b=>({value:b,label:`Blok ${b}`}))]} /><ThemedDropdown className="resident-filter resident-subblock-filter" value={residentSubFilter} onChange={setResidentSubFilter} ariaLabel="Filter berdasarkan sub blok" options={[{value:"Semua sub blok",label:"Semua sub blok"}, ...subBlocks.filter(sb => residentFilter === "Semua blok" || sb.startsWith(residentFilter)).map(sb=>({value:sb,label:sb}))]} /><button type="button" className="print-btn" onClick={() => window.print()}><FileText size={15} /> Cetak daftar</button><span className="result-count">{filteredResidents.length} warga</span></div><div className="resident-grid">{filteredResidents.map(r => <div className="resident-card" key={r.id}><span className="resident-avatar">{r.name.split(" ").map(n => n[0]).join("").slice(0, 2)}</span><div><b>{r.name}</b><small>Blok {r.block} · Sub blok {r.subBlock || `${r.block}1`} · Rumah {r.house}</small></div><div className="resident-actions"><button className="icon-btn edit" title="Edit warga" onClick={() => startEditResident(r)}><Pencil size={15} /></button><button className="icon-btn delete" title="Hapus warga" onClick={() => deleteResident(r.id)}><Trash2 size={15} /></button></div></div>)}{filteredResidents.length === 0 && <div className="empty-residents">Tidak ada warga yang cocok.</div>}</div></section>; }

