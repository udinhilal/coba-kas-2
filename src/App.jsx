// Pinned subject: simple RT finance app for a neighborhood treasurer. Plan: warm paper ledger meets modern utility, with terracotta as the one owned accent, a fast entry flow, and monthly recaps that make cash movement obvious.
import { useEffect, useMemo, useState } from "react";
import { ArrowDownLeft, ArrowUpRight, BookOpen, CalendarDays, ChevronDown, FileText, Home, Pencil, Plus, ReceiptText, Search, Settings2, Trash2, Users, WalletCards, X } from "lucide-react";
import * as XLSX from "xlsx";
import "./styles.css";

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
  const [editingExpense, setEditingExpense] = useState(null);
  const [block, setBlock] = useState("");
  const [paymentSubBlock, setPaymentSubBlock] = useState("");
  const [resident, setResident] = useState("");
  const selectedResident = residents.find((r) => r.name === resident);
  const autoSubBlock = selectedResident?.subBlock || "";
  const [clean, setClean] = useState("");
  const [kas, setKas] = useState("");
  const [other, setOther] = useState("");
  const [paymentPeriod, setPaymentPeriod] = useState("");
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
  const [summaryPeriod, setSummaryPeriod] = useState("Oktober 2026");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const fileInput = { current: null };

  useEffect(() => {
    const online = () => { setIsOnline(true); const now = new Date().toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" }); localStorage.setItem("kas-warga-rt-last-sync", now); setLastSync(now); setNotice("Sinkronisasi otomatis selesai"); setTimeout(() => setNotice(""), 2200); }; const offline = () => setIsOnline(false);
    window.addEventListener("online", online); window.addEventListener("offline", offline);
    return () => { window.removeEventListener("online", online); window.removeEventListener("offline", offline); };
  }, []);
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
    const record = { id: editingExpense || Date.now(), date: editingExpense ? editingExpense.date : "23 Okt 2026", title: expenseTitle.trim(), amount: Number(expenseAmount) };
    setCashOut(editingExpense ? cashOut.map((item) => item.id === editingExpense.id ? record : item) : [record, ...cashOut]);
    setEditingExpense(null);
    setExpenseTitle("");
    setExpenseAmount("");
    setNotice(editingExpense ? "Pengeluaran diperbarui" : "Pengeluaran berhasil dicatat");
    setTimeout(() => setNotice(""), 2200);
  };
  const addEntry = (e) => {
    e.preventDefault();
    const record = { id: editingEntry || Date.now(), date: editingEntry ? editingEntry.date : `23 ${paymentPeriod.replace("2026", "")}`.trim(), period: paymentPeriod, resident, block: block.trim().toUpperCase(), subBlock: paymentSubBlock || autoSubBlock, clean: Number(clean), kas: Number(kas), other: Number(other) };
    setEntries(editingEntry ? entries.map((item) => item.id === editingEntry.id ? record : item) : [record, ...entries]);
    setEditingEntry(null);
    setPaymentPeriod(""); setBlock(""); setResident(""); setClean(""); setKas(""); setOther(""); setPaymentSubBlock("");
    setNotice(editingEntry ? "Pembayaran diperbarui" : "Pembayaran tersimpan");
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
      <aside className="sidebar">
        <div className="brand"><Logo /><div><strong>kas warga</strong><small>RT 26 / RW 04</small></div></div>
        <div className="side-label">Menu utama</div>
        <nav>
          {[["overview", Home, "Ringkasan"], ["input", ReceiptText, "Input pembayaran"], ["cashflow", WalletCards, "Arus keuangan"], ["recap", CalendarDays, "Rekap bulanan"], ["report", FileText, "Laporan warga"], ["public", Users, "Ringkasan publik"], ["residents", Users, "Data warga"]].map(([id, Icon, label]) => (
            <button className={`nav-item ${active === id ? "active" : ""}`} key={id} onClick={() => { setActive(id); setSettingsOpen(false); if (id === "input") { setPaymentPeriod(""); setBlock(""); setPaymentSubBlock(""); setResident(""); setClean(""); setKas(""); setOther(""); setEditingEntry(null); } }}><Icon size={18} />{label}{id === "input" && <span className="nav-dot" />}</button>
          ))}
        </nav>
        <div className="side-bottom"><button className={`nav-item ${settingsOpen ? "active" : ""}`} onClick={() => { setSettingsOpen(!settingsOpen); if (!settingsOpen) setActive("settings"); }}><Settings2 size={18} />Pengaturan</button><div className="profile"><span className="avatar">{(user?.user_metadata?.full_name || user?.email || "W").slice(0,2).toUpperCase()}</span><div><b>{user?.user_metadata?.full_name || "Pengelola"}</b><small>{user?.email || "Akun terverifikasi"}</small></div><button type="button" className="signout-btn" onClick={onSignOut}>Keluar</button></div></div>
      </aside>
      <main className="main">
        <div className="connection-status"><span /> {isOnline ? "Online" : "Mode offline"} · {isOnline ? `Sinkron terakhir: ${lastSync}` : "Data tersimpan di perangkat"}</div>
        <header className="topbar"><div><span className="eyebrow">Jumat, 23 Oktober 2026</span><h1>{active === "overview" ? `Selamat datang, ${user?.user_metadata?.full_name || "Pengelola"}` : active === "input" ? "Input pembayaran" : active === "cashflow" ? "Arus keuangan" : active === "recap" ? "Rekap bulanan" : active === "report" ? "Laporan warga" : active === "public" ? "Ringkasan publik" : active === "settings" ? "Pengaturan" : "Data warga"}</h1></div><button className="month-pill"><CalendarDays size={16} /> Oktober 2026 <ChevronDown size={15} /></button></header>
        {notice && <div className="toast"><span>✓</span>{notice}</div>}
        {active === "overview" && !settingsOpen && <Overview entries={entries} residents={residents} total={total} outgoingTotal={cashOut.reduce((a, e) => a + e.amount, 0)} runningBalance={runningBalance} summaryPeriod={summaryPeriod} setSummaryPeriod={setSummaryPeriod} setActive={setActive} startEditEntry={startEditEntry} deleteEntry={deleteEntry} />}
        {active === "input" && !settingsOpen && <InputView {...{ block, setBlock, autoSubBlock, paymentSubBlock, setPaymentSubBlock, resident, setResident, clean, setClean, kas, setKas, other, setOther, paymentPeriod, setPaymentPeriod, addEntry, residents, editingEntry, setEditingEntry }} />}
        {active === "cashflow" && !settingsOpen && <Cashflow entries={entries} cashOut={cashOut} runningBalance={runningBalance} addExpense={addExpense} expenseTitle={expenseTitle} setExpenseTitle={setExpenseTitle} expenseAmount={expenseAmount} setExpenseAmount={setExpenseAmount} editingExpense={editingExpense} setEditingExpense={setEditingExpense} startEditExpense={startEditExpense} deleteExpense={deleteExpense} />}
        {active === "recap" && !settingsOpen && <Recap entries={entries} total={total} residents={residents} recapSearch={recapSearch} setRecapSearch={setRecapSearch} />}
        {active === "report" && !settingsOpen && <ResidentReport entries={entries} cashOut={cashOut} runningBalance={runningBalance} residents={residents} />}
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
function Overview({ entries, total, outgoingTotal, runningBalance, summaryPeriod, setSummaryPeriod, setActive, startEditEntry, deleteEntry, residents = [] }) { const periodEntries = summaryPeriod === "Semua periode" ? entries : entries.filter((e) => e.period === summaryPeriod || e.date.includes(summaryPeriod.split(" ")[0] === "Oktober" ? "Okt" : summaryPeriod.split(" ")[0].slice(0, 3))); const periodTotal = periodEntries.reduce((a, e) => a + e.clean + e.kas + (e.other || 0), 0);
  return <section className="content">
    <div className="hero-card"><div><span className="eyebrow light">Saldo kas berjalan</span><div className="hero-number">{money(runningBalance)}</div><span className="hero-note"><ArrowUpRight size={15} /> Saldo diperbarui otomatis</span></div><div className="hero-art"><div className="sun" /><div className="hill h1" /><div className="hill h2" /><WalletCards size={42} /></div></div>
    <div className="stat-grid"><Stat title={`Pemasukan ${summaryPeriod === "Semua periode" ? "" : "periode"}`} value={money(periodTotal)} icon={ArrowDownLeft} tone="green" meta={`Dari ${periodEntries.length} pembayaran`} /><Stat title="Pengeluaran bulan ini" value={money(outgoingTotal)} icon={ArrowUpRight} tone="red" meta={`${outgoingTotal ? "Transaksi tercatat" : "Belum ada transaksi"}`} /><Stat title="Warga sudah bayar" value={`${periodEntries.length} / ${residents.length}`} icon={Users} tone="gold" meta={residents.length ? `${Math.round((periodEntries.length / residents.length) * 100)}% tingkat pembayaran` : "Belum ada data warga"} /></div>
    <div className="section-head"><div><span className="eyebrow">Aktivitas terbaru</span><h2>Catatan pembayaran</h2></div><select className="period-filter" value={summaryPeriod} onChange={e => setSummaryPeriod(e.target.value)}><option>Semua periode</option>{["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"].map((month) => <option key={month}>{month} 2026</option>)}</select><button className="text-btn" onClick={() => setActive("input")}>+ Input pembayaran</button></div>
    <div className="table-card"><table><thead><tr><th>Tanggal</th><th>Nama warga</th><th>Blok</th><th>Iuran kebersihan</th><th>Iuran kas RT</th><th>Iuran lainnya</th><th>Total</th><th>Aksi</th></tr></thead><tbody>{periodEntries.length === 0 ? <tr><td colSpan="8" className="empty-table">Belum ada data pembayaran untuk periode ini.</td></tr> : periodEntries.map((e) => <tr key={e.id}><td className="muted">{e.date}</td><td><b>{e.resident || residents.find((r) => r.block === e.block && r.subBlock === e.subBlock)?.name || "Nama belum diisi"}</b></td><td><span className="block-tag">{e.block}{e.subBlock ? ` · ${e.subBlock}` : ""}</span></td><td>{money(e.clean)}</td><td>{money(e.kas)}</td><td>{money(e.other || 0)}</td><td><b>{money(e.clean + e.kas + (e.other || 0))}</b></td><td><div className="entry-actions"><button className="icon-btn edit" title="Edit pembayaran" onClick={() => startEditEntry(e)}><Pencil size={15} /></button><button className="icon-btn delete" title="Hapus pembayaran" onClick={() => deleteEntry(e.id)}><Trash2 size={15} /></button></div></td></tr>)}</tbody></table></div>
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
  return <section className="content narrow"><div className="form-intro"><div className="notebook-icon"><BookOpen size={25} /></div><div><h2>Catat iuran warga</h2><p>Isi sekali, total dihitung otomatis. Praktis buat bendahara.</p></div></div><form className="form-card" onSubmit={addEntry}><div className="form-grid"><label>Periode<select value={paymentPeriod} onChange={e => setPaymentPeriod(e.target.value)}><option value="">Pilih periode</option>{["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"].map((month) => <option key={month}>{month} 2026</option>)}</select></label><label>Blok<input value={block} onChange={(e) => { const nextBlock = e.target.value.toUpperCase(); const firstSub = residents.find((r) => r.block.trim().toUpperCase() === nextBlock)?.subBlock || ""; setBlock(nextBlock); setPaymentSubBlock(firstSub); syncResident(nextBlock, firstSub); }} placeholder="Contoh: A" required /></label><label>Sub blok<select value={paymentSubBlock} onChange={(e) => { setPaymentSubBlock(e.target.value); syncResident(block, e.target.value); }} disabled={!normalizedBlock}><option value="">{normalizedBlock ? "Pilih sub blok" : "Isi blok dahulu"}</option>{subBlocksForBlock.map((sub) => <option key={sub} value={sub}>{sub}</option>)}</select><small className="field-hint">Pilihan diambil dari Data Warga</small></label><label>Nama warga<input value={matchingResident?.name || resident} readOnly placeholder="Otomatis dari blok dan sub blok" /><small className="field-hint">Terisi otomatis dari Data Warga</small></label><label>Iuran kebersihan<div className="input-money"><span>Rp</span><input type="number" min="0" value={clean} onChange={(e) => setClean(e.target.value)} /></div></label><label>Iuran kas RT<div className="input-money"><span>Rp</span><input type="number" min="0" value={kas} onChange={(e) => setKas(e.target.value)} /></div></label><label>Iuran lainnya <em>opsional</em><div className="input-money"><span>Rp</span><input type="number" min="0" value={other} onChange={(e) => setOther(e.target.value)} /></div></label></div><div className="total-row"><div><span>Total pembayaran</span><strong>{money(Number(clean || 0) + Number(kas || 0) + Number(other || 0))}</strong></div><button className="primary-btn" type="submit"><Plus size={18} /> {editingEntry ? "Simpan perubahan" : "Simpan pembayaran"}</button>{editingEntry && <button className="cancel-btn" type="button" onClick={() => setEditingEntry(null)}><X size={16} /> Batal</button>}</div></form><div className="tip"><span>✦</span><div><b>Nama warga otomatis</b><p>Pilih blok dan sub blok, nama akan diambil langsung dari Data Warga.</p></div></div></section>;
}
function Cashflow({ entries, cashOut, runningBalance, addExpense, expenseTitle, setExpenseTitle, expenseAmount, setExpenseAmount, editingExpense, setEditingExpense, startEditExpense, deleteExpense }) { const incoming = entries.reduce((a,e)=>a+e.clean+e.kas+(e.other||0),0); const outgoing = cashOut.reduce((a,e)=>a+e.amount,0); return <section className="content"><div className="recap-banner cashflow-banner"><div><span className="eyebrow light">Catatan keuangan</span><h2>Arus masuk & keluar</h2><p>Pantau pergerakan kas warga.</p></div></div><form className="expense-form" onSubmit={addExpense}><div><span className="eyebrow">Input pengeluaran</span><h2>Catat uang keluar</h2></div><input placeholder="Keterangan pengeluaran" value={expenseTitle} onChange={e=>setExpenseTitle(e.target.value)} required /><div className="input-money"><span>Rp</span><input type="number" min="1" value={expenseAmount} onChange={e=>setExpenseAmount(e.target.value)} required /></div><button className="primary-btn" type="submit"><Plus size={18}/>{editingExpense ? "Simpan perubahan" : "Simpan pengeluaran"}</button></form><div className="recap-grid"><div className="big-card cash-balance"><span className="eyebrow">Saldo berjalan</span><strong>{money(runningBalance)}</strong></div><div className="big-card cash-in"><span className="eyebrow">Total uang masuk</span><strong>{money(incoming)}</strong></div><div className="big-card cash-out"><span className="eyebrow">Total uang keluar</span><strong>{money(outgoing)}</strong></div></div><div className="table-card"><table><thead><tr><th>Tanggal</th><th>Keterangan</th><th>Jenis</th><th>Nominal</th></tr></thead><tbody>{entries.map(e=><tr key={`in-${e.id}`}><td>{e.date}</td><td>Iuran {e.resident}</td><td><span className="flow-in">Masuk</span></td><td>{money(e.clean+e.kas+(e.other||0))}</td></tr>)}{cashOut.map(e=><tr key={`out-${e.id}`}><td>{e.date}</td><td>{e.title}</td><td><span className="flow-out">Keluar</span></td><td>{money(e.amount)}</td></tr>)}</tbody></table></div></section>; }
function ResidentReport({ entries, cashOut, runningBalance }) { const incoming=entries.reduce((a,e)=>a+e.clean+e.kas+(e.other||0),0); const outgoing=cashOut.reduce((a,e)=>a+e.amount,0); return <section className="content report-page"><div className="report-toolbar"><div><span className="eyebrow">Dokumen warga</span><h2>Laporan Pertanggungjawaban</h2><p className="subcopy">Ringkasan kas RT 26 / RW 04.</p></div><button className="primary-btn" onClick={()=>window.print()}><FileText size={17}/> Cetak / Simpan PDF</button></div><div className="report-paper"><div className="report-heading"><div><div className="report-kicker">RT 26 / RW 04</div><h1>Laporan Pertanggungjawaban Kas</h1><p>Periode berjalan</p></div></div><p className="report-intro">Ringkasan pemasukan dan pengeluaran kas warga untuk diketahui bersama.</p><div className="report-stats"><div><span>Total pemasukan</span><b>{money(incoming)}</b></div><div><span>Total pengeluaran</span><b>{money(outgoing)}</b></div><div><span>Saldo akhir</span><b>{money(runningBalance)}</b></div></div><h3>Rincian pemasukan</h3><table><tbody>{entries.map(e=><tr key={e.id}><td>{e.resident}</td><td>{e.date}</td><td>{money(e.clean+e.kas+(e.other||0))}</td></tr>)}</tbody></table><h3>Rincian pengeluaran</h3><table><tbody>{cashOut.map(e=><tr key={e.id}><td>{e.title}</td><td>{e.date}</td><td>{money(e.amount)}</td></tr>)}</tbody></table><div className="report-footer"><p>Demikian laporan ini dibuat untuk dapat diketahui bersama.</p><div><span>Bendahara RT</span><strong>Bendahara RT</strong></div></div></div></section>; }
function PublicSummary({ entries, cashOut, openingBalance, manualAdjustment }) {
  const [period, setPeriod] = useState("Oktober 2026");
  const months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
  const monthPrefix = period.split(" ")[0].slice(0, 3).toLocaleLowerCase("id-ID");
  const periodEntries = entries.filter((e) => period === "Semua periode" || e.period === period || (e.date || "").toLocaleLowerCase("id-ID").includes(monthPrefix));
  const periodExpenses = cashOut.filter((e) => period === "Semua periode" || e.period === period || (e.date || "").toLocaleLowerCase("id-ID").includes(monthPrefix));
  const cleanTotal = periodEntries.reduce((sum, e) => sum + Number(e.clean || 0), 0);
  const kasTotal = periodEntries.reduce((sum, e) => sum + Number(e.kas || 0), 0);
  const otherTotal = periodEntries.reduce((sum, e) => sum + Number(e.other || 0), 0);
  const income = cleanTotal + kasTotal + otherTotal;
  const expense = periodExpenses.reduce((sum, e) => sum + Number(e.amount || 0), 0);
  const balance = Number(openingBalance || 0) + income - expense + Number(manualAdjustment || 0);
  return <section className="content public-summary">
    <div className="public-heading"><div><span className="eyebrow">Informasi untuk warga</span><h2>Laporan Kas Warga</h2><p>RT 26 / RW 04 · Ringkasan terbuka pengelolaan kas</p></div><div className="public-controls"><label>Pilih periode<select value={period} onChange={e => setPeriod(e.target.value)}><option>Semua periode</option>{months.map(month => <option key={month}>{month} 2026</option>)}</select></label><button className="primary-btn" onClick={() => window.print()}><FileText size={16}/> Cetak / Simpan PDF</button></div></div>
    <div className="privacy-note"><span>✓</span><div><b>Tampilan publik, data pribadi disembunyikan</b><p>Tidak menampilkan nama warga, nomor rumah, blok, atau rincian pembayaran per orang.</p></div></div>
    <div className="public-kpis"><article className="public-kpi income"><span>Total pemasukan</span><strong>{money(income)}</strong><small>Periode {period}</small></article><article className="public-kpi expense"><span>Total pengeluaran</span><strong>{money(expense)}</strong><small>Periode {period}</small></article><article className="public-kpi balance"><span>Saldo akhir</span><strong>{money(balance)}</strong><small>Saldo awal + masuk - keluar + koreksi</small></article></div>
    <div className="public-breakdown"><div className="section-head"><div><span className="eyebrow">Rangkuman kategori</span><h2>Rincian agregat</h2></div></div><div className="public-breakdown-row"><span>Iuran kebersihan</span><b>{money(cleanTotal)}</b></div><div className="public-breakdown-row"><span>Iuran kas RT</span><b>{money(kasTotal)}</b></div><div className="public-breakdown-row"><span>Iuran lainnya</span><b>{money(otherTotal)}</b></div><div className="public-breakdown-row"><span>Total pengeluaran</span><b>{money(expense)}</b></div></div>
    <footer className="public-footer"><span>Kas Warga RT 26 / RW 04</span><span>Ringkasan agregat · {period}</span></footer>
  </section>;
}
function Recap({ entries, total, residents = [], recapSearch, setRecapSearch }) { const [recapPeriod, setRecapPeriod] = useState("Oktober 2026"); const months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"]; const matchesPeriod = (e) => recapPeriod === "Semua periode" || e.period === recapPeriod || (e.date || "").toLocaleLowerCase("id-ID").includes(recapPeriod.split(" ")[0].toLocaleLowerCase("id-ID").slice(0,3)); const monthEntries = entries.filter(matchesPeriod); const filteredEntries = monthEntries.filter((e) => (e.resident || "").toLocaleLowerCase("id-ID").includes(recapSearch.trim().toLocaleLowerCase("id-ID"))); const periodTotal = monthEntries.reduce((sum, e) => sum + e.clean + e.kas + (e.other || 0), 0); const paidCount = new Set(monthEntries.map(e => e.residentId || e.resident || `${e.block}-${e.subBlock}`)).size; const paidPercent = residents.length ? Math.round((paidCount / residents.length) * 100) : 0; return <section className="content recap-page"><div className="recap-banner"><div><span className="eyebrow light">Laporan periode</span><h2>Rekap {recapPeriod}</h2><p>Ringkasan pemasukan kas warga.</p></div><div className="recap-actions"><label className="recap-period-label">Pilih bulan<select className="recap-period-select" value={recapPeriod} onChange={e => setRecapPeriod(e.target.value)}><option>Semua periode</option>{months.map(month => <option key={month}>{month} 2026</option>)}</select></label><button className="outline-light" onClick={() => exportExcel(monthEntries, recapPeriod)}><FileText size={16} /> Ekspor Excel</button><button className="outline-light" onClick={() => window.print()}><FileText size={16} /> Ekspor PDF</button></div></div><div className="recap-grid"><div className="big-card"><span className="eyebrow">Total terkumpul</span><strong>{money(periodTotal)}</strong><div className="bar"><i style={{ width: `${Math.min(paidPercent, 100)}%` }} /></div><small>{residents.length ? `${paidPercent}% warga sudah membayar` : "Belum ada data warga"}</small></div><div className="big-card"><span className="eyebrow">Warga sudah bayar</span><strong>{paidCount} / {residents.length}</strong><small>{residents.length ? `${paidPercent}% tingkat pembayaran` : "Belum ada data warga"}</small></div></div><div className="section-head"><div><span className="eyebrow">Per warga</span><h2 className="recap-details-title">Rincian iuran · {recapPeriod}</h2></div><div className="search"><Search size={16} /><input placeholder="Cari nama warga..." value={recapSearch} onChange={e => setRecapSearch(e.target.value)} /></div></div><div className="table-card"><table><thead><tr><th>Nama warga</th><th>Periode</th><th>Blok</th><th>Kebersihan</th><th>Kas RT</th><th>Iuran lainnya</th><th>Total</th><th>Status</th></tr></thead><tbody>{filteredEntries.length ? filteredEntries.map(e => <tr key={e.id}><td><b>{e.resident || residents.find(r => r.block === e.block && r.subBlock === e.subBlock)?.name || "Nama belum diisi"}</b></td><td>{e.period || recapPeriod}</td><td><span className="block-tag">{e.block}{e.subBlock ? ` · ${e.subBlock}` : ""}</span></td><td>{money(e.clean)}</td><td>{money(e.kas)}</td><td>{money(e.other || 0)}</td><td><b>{money(e.clean + e.kas + (e.other || 0))}</b></td><td><span className="status-paid">Lunas</span></td></tr>) : <tr><td colSpan="8" className="empty-table">{monthEntries.length ? "Nama warga tidak ditemukan." : "Belum ada pembayaran pada periode ini."}</td></tr>}</tbody></table></div></section>; }
function Residents({ residents, residentName, setResidentName, residentBlock, setResidentBlock, residentSubBlock, setResidentSubBlock, addResident, editingResident, setEditingResident, startEditResident, deleteResident, residentSearch, setResidentSearch, residentFilter, setResidentFilter, residentSubFilter, setResidentSubFilter }) { const blocks = [...new Set(residents.map((r) => r.block).filter(Boolean))].sort(); const subBlocks = [...new Set(residents.map((r) => r.subBlock).filter(Boolean))].sort(); const query = residentSearch.trim().toLocaleLowerCase("id-ID"); const filteredResidents = residents.filter((r) => (residentFilter === "Semua blok" || r.block.toLocaleLowerCase("id-ID") === residentFilter.toLocaleLowerCase("id-ID")) && (residentSubFilter === "Semua sub blok" || (r.subBlock || "").toLocaleLowerCase("id-ID") === residentSubFilter.toLocaleLowerCase("id-ID")) && (r.name || "").toLocaleLowerCase("id-ID").includes(query)); return <section className="content narrow"><div className="section-head"><div><span className="eyebrow">Buku warga</span><h2>Daftar nama warga</h2><p className="subcopy">Kelola data yang muncul di menu input pembayaran.</p></div></div><form className="add-resident" onSubmit={addResident} noValidate><input placeholder="Nama lengkap warga" value={residentName} onChange={e => setResidentName(e.target.value)} /><input className="compact-field" value={residentBlock} onChange={e => setResidentBlock(e.target.value)} placeholder="Blok, contoh A" /><input className="compact-field" value={residentSubBlock} onChange={e => setResidentSubBlock(e.target.value)} placeholder="Sub blok, contoh A1" /><button className="primary-btn" type="submit"><Plus size={18} /> {editingResident ? "Simpan perubahan" : "Tambah warga"}</button>{editingResident && <button className="cancel-btn" type="button" onClick={() => { setEditingResident(null); setResidentName(""); }}><X size={16} /> Batal</button>}</form><div className="resident-toolbar"><div className="search resident-search"><Search size={16} /><input placeholder="Cari nama warga..." value={residentSearch} onChange={e => setResidentSearch(e.target.value)} /></div><select className="resident-filter" value={residentFilter} onChange={e => { setResidentFilter(e.target.value); setResidentSubFilter("Semua sub blok"); }}><option>Semua blok</option>{blocks.map(b => <option key={b}>{b}</option>)}</select><select className="resident-filter" value={residentSubFilter} onChange={e => setResidentSubFilter(e.target.value)}><option>Semua sub blok</option>{subBlocks.filter(sb => residentFilter === "Semua blok" || sb.startsWith(residentFilter)).map(sb => <option key={sb}>{sb}</option>)}</select><button type="button" className="print-btn" onClick={() => window.print()}><FileText size={15} /> Cetak daftar</button><span className="result-count">{filteredResidents.length} warga</span></div><div className="resident-grid">{filteredResidents.map(r => <div className="resident-card" key={r.id}><span className="resident-avatar">{r.name.split(" ").map(n => n[0]).join("").slice(0, 2)}</span><div><b>{r.name}</b><small>Blok {r.block} · Sub blok {r.subBlock || `${r.block}1`} · Rumah {r.house}</small></div><div className="resident-actions"><button className="icon-btn edit" title="Edit warga" onClick={() => startEditResident(r)}><Pencil size={15} /></button><button className="icon-btn delete" title="Hapus warga" onClick={() => deleteResident(r.id)}><Trash2 size={15} /></button></div></div>)}{filteredResidents.length === 0 && <div className="empty-residents">Tidak ada warga yang cocok.</div>}</div></section>; }
