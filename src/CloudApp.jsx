// Sinkronisasi data Kas Warga ke Supabase.
// Memuat data akun dari Supabase saat login, lalu menyimpan setiap perubahan aplikasi ke tabel Supabase.
// Data lama di browser dikosongkan: sumber data utama sekarang Supabase.
import { createElement as h, useEffect, useRef, useState } from "react";
import App from "./App.jsx";
import { supabase, supabaseConfigured } from "./lib/supabaseClient";

const STORAGE_KEY = "kas-warga-rt-data-v1";
const SYNC_KEY = "kas-warga-rt-last-sync";
const MONTHS = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const originalSetItem = Storage.prototype.setItem;

const pad = (n) => String(n).padStart(2, "0");
const toIsoDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const todayIso = () => toIsoDate(new Date());
const nowLabel = () => new Date().toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" });
const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);
const nonNegative = (v) => Math.max(0, num(v));

const monthIndex = (word) => {
  const w = String(word || "").toLowerCase().slice(0, 3);
  return w.length === 3 ? MONTHS.findIndex((m) => m.toLowerCase().slice(0, 3) === w) : -1;
};
const parseDisplayDate = (text, fallbackYear) => {
  const match = String(text || "").match(/(\d{1,2})\s+([A-Za-z]+)(?:\s+(\d{4}))?/);
  if (!match) return null;
  const m = monthIndex(match[2]);
  const year = Number(match[3] || fallbackYear || new Date().getFullYear());
  const day = Number(match[1]);
  if (m < 0 || day < 1 || day > 31) return null;
  return `${year}-${pad(m + 1)}-${pad(day)}`;
};
const periodToMonth = (period) => {
  const match = String(period || "").match(/([A-Za-z]+)\s+(\d{4})/);
  if (!match) return null;
  const m = monthIndex(match[1]);
  return m < 0 ? null : `${match[2]}-${pad(m + 1)}-01`;
};
const formatShortDate = (iso) => {
  const [y, m, d] = String(iso || todayIso()).slice(0, 10).split("-");
  return `${Number(d)} ${MONTHS[Number(m) - 1].slice(0, 3)} ${y}`;
};
const formatPeriod = (iso) => {
  const [y, m] = String(iso || todayIso()).slice(0, 10).split("-");
  return `${MONTHS[Number(m) - 1]} ${y}`;
};

// ID aplikasi berupa angka (Date.now()). Supabase memakai UUID.
// Angka diubah ke UUID yang stabil agar edit dan hapus tetap mengenai baris yang sama.
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const NUMERIC_UUID_RE = /^00000000-0000-4000-8000-([0-9a-f]{12})$/i;
const fallbackIds = new Map();
const rawId = (id) => (id && typeof id === "object" ? id.id : id);
const idToUuid = (id) => {
  const value = rawId(id);
  if (typeof value === "string" && UUID_RE.test(value)) return value.toLowerCase();
  const n = Number(value);
  if (Number.isSafeInteger(n) && n >= 0 && n < 0xffffffffffff) return `00000000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
  const key = String(value);
  if (!fallbackIds.has(key)) fallbackIds.set(key, crypto.randomUUID());
  return fallbackIds.get(key);
};
const uuidToId = (uuid) => {
  const match = String(uuid).match(NUMERIC_UUID_RE);
  return match ? parseInt(match[1], 16) : uuid;
};
const timeFromId = (id) => {
  const n = Number(rawId(id));
  return Number.isSafeInteger(n) && n > 1500000000000 && n < 4102444800000 ? new Date(n).toISOString() : null;
};

const residentRow = (r, owner) => ({
  id: idToUuid(r.id),
  owner_id: owner,
  name: String(r.name || "").trim() || "Tanpa nama",
  block: String(r.block || "").trim() || "-",
  sub_block: String(r.subBlock || "").trim(),
  house: r.house ? String(r.house) : null,
});
const paymentRow = (e, owner) => {
  const periodMonth = periodToMonth(e.period);
  const paymentDate = parseDisplayDate(e.date, periodMonth ? periodMonth.slice(0, 4) : undefined) || periodMonth || todayIso();
  return {
    id: idToUuid(e.id),
    owner_id: owner,
    resident_id: null,
    resident_name: String(e.resident || "").trim() || "Nama belum diisi",
    block: String(e.block || "").trim() || "-",
    sub_block: e.subBlock ? String(e.subBlock) : null,
    payment_date: paymentDate,
    period_month: periodMonth || `${paymentDate.slice(0, 7)}-01`,
    cleaning_amount: nonNegative(e.clean),
    rt_fee_amount: nonNegative(e.kas),
    other_amount: nonNegative(e.other),
  };
};
const expenseRow = (x, owner) => ({
  id: idToUuid(x.id),
  owner_id: owner,
  expense_date: parseDisplayDate(x.date) || todayIso(),
  title: String(x.title || "").trim() || "Pengeluaran",
  amount: num(x.amount),
});
const historyRow = (item, owner) => ({
  id: idToUuid(item.id),
  owner_id: owner,
  opening_balance: num(item.opening),
  correction_amount: num(item.correction),
  changed_at: item.changedAt || timeFromId(item.id) || new Date().toISOString(),
});

const TABLES = [
  ["residents", "residents", residentRow],
  ["payments", "entries", paymentRow],
  ["expenses", "cashOut", expenseRow],
  ["balance_history", "balanceHistory", historyRow],
];

const buildSnapshot = (data, owner) => {
  const snapshot = {};
  for (const [table, key, build] of TABLES) {
    const rows = new Map();
    for (const item of Array.isArray(data?.[key]) ? data[key] : []) {
      const row = build(item, owner);
      rows.set(row.id, row);
    }
    snapshot[table] = rows;
  }
  snapshot.cash_settings = { owner_id: owner, opening_balance: num(data?.openingBalance), manual_adjustment: num(data?.manualAdjustment) };
  return snapshot;
};

const raise = (error, label) => {
  if (error) throw new Error(`${label}: ${error.message}`);
};

const pushDiff = async (prev, next, owner) => {
  const stamp = new Date().toISOString();
  for (const [table] of TABLES) {
    const before = prev[table];
    const after = next[table];
    const upserts = [...after.values()]
      .filter((row) => !before.has(row.id) || JSON.stringify(before.get(row.id)) !== JSON.stringify(row))
      .map((row) => (table === "balance_history" ? row : { ...row, updated_at: stamp }));
    if (upserts.length) {
      const { error } = await supabase.from(table).upsert(upserts, { onConflict: "id" });
      raise(error, `Simpan ${table}`);
    }
  }
  for (const [table] of [...TABLES].reverse()) {
    const removed = [...prev[table].keys()].filter((id) => !next[table].has(id));
    if (removed.length) {
      const { error } = await supabase.from(table).delete().eq("owner_id", owner).in("id", removed);
      raise(error, `Hapus ${table}`);
    }
  }
  if (JSON.stringify(prev.cash_settings) !== JSON.stringify(next.cash_settings)) {
    const { error } = await supabase.from("cash_settings").upsert({ ...next.cash_settings, updated_at: stamp }, { onConflict: "owner_id" });
    raise(error, "Simpan saldo");
  }
};

const loadFromSupabase = async (owner) => {
  const [residents, payments, expenses, history, settings] = await Promise.all([
    supabase.from("residents").select("*").eq("owner_id", owner).order("created_at", { ascending: true }),
    supabase.from("payments").select("*").eq("owner_id", owner).order("payment_date", { ascending: false }).order("created_at", { ascending: false }),
    supabase.from("expenses").select("*").eq("owner_id", owner).order("expense_date", { ascending: false }).order("created_at", { ascending: false }),
    supabase.from("balance_history").select("*").eq("owner_id", owner).order("changed_at", { ascending: false }),
    supabase.from("cash_settings").select("*").eq("owner_id", owner).maybeSingle(),
  ]);
  raise(residents.error, "Muat data warga");
  raise(payments.error, "Muat pembayaran");
  raise(expenses.error, "Muat pengeluaran");
  raise(history.error, "Muat riwayat saldo");
  raise(settings.error, "Muat saldo");
  return {
    residents: (residents.data || []).map((r) => ({ id: uuidToId(r.id), block: r.block, subBlock: r.sub_block || "", name: r.name, house: r.house || "" })),
    entries: (payments.data || []).map((p) => ({
      id: uuidToId(p.id),
      date: formatShortDate(p.payment_date),
      period: formatPeriod(p.period_month),
      resident: p.resident_name,
      block: p.block,
      subBlock: p.sub_block || "",
      clean: num(p.cleaning_amount),
      kas: num(p.rt_fee_amount),
      other: num(p.other_amount),
    })),
    cashOut: (expenses.data || []).map((x) => ({ id: uuidToId(x.id), date: formatShortDate(x.expense_date), title: x.title, amount: num(x.amount) })),
    openingBalance: num(settings.data?.opening_balance),
    manualAdjustment: num(settings.data?.manual_adjustment),
    balanceHistory: (history.data || []).map((item) => ({
      id: uuidToId(item.id),
      date: new Date(item.changed_at).toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" }),
      opening: num(item.opening_balance),
      correction: num(item.correction_amount),
      changedAt: item.changed_at,
    })),
  };
};

let statusTimer;
const showStatus = (text, isError = false) => {
  let el = document.getElementById("cloud-sync-status");
  if (!el) {
    el = document.createElement("div");
    el.id = "cloud-sync-status";
    el.setAttribute("role", "status");
    el.style.cssText = "position:fixed;left:50%;bottom:18px;transform:translateX(-50%);z-index:50;padding:10px 14px;border-radius:8px;font:600 13px DM Sans,sans-serif;box-shadow:0 8px 25px #0002;max-width:90vw;text-align:center";
    document.body.appendChild(el);
  }
  el.textContent = text;
  el.style.background = isError ? "#8f2f22" : "#292622";
  el.style.color = "#fff";
  el.style.display = "block";
  clearTimeout(statusTimer);
  if (!isError) statusTimer = setTimeout(() => { el.style.display = "none"; }, 2200);
};

const createSyncer = (owner, initialSnapshot) => {
  let synced = initialSnapshot;
  let latest = null;
  let running = null;
  let timer = null;
  const run = async () => {
    while (latest !== null) {
      const raw = latest;
      latest = null;
      let data;
      try { data = JSON.parse(raw); } catch { continue; }
      const next = buildSnapshot(data, owner);
      try {
        await pushDiff(synced, next, owner);
        synced = next;
        originalSetItem.call(window.localStorage, SYNC_KEY, nowLabel());
        showStatus("Tersimpan ke Supabase");
      } catch (error) {
        console.error(error);
        latest = latest ?? raw;
        showStatus("Gagal menyimpan ke Supabase. Periksa koneksi; akan dicoba lagi.", true);
        await new Promise((resolve) => setTimeout(resolve, 8000));
      }
    }
  };
  const flush = () => {
    clearTimeout(timer);
    timer = null;
    if (!running && latest !== null) running = run().finally(() => { running = null; if (latest !== null) flush(); });
    return running || Promise.resolve();
  };
  return {
    queue(raw) {
      latest = raw;
      clearTimeout(timer);
      timer = setTimeout(flush, 500);
    },
    flush,
  };
};

export default function CloudApp({ user, onSignOut }) {
  const [state, setState] = useState({ status: "loading", error: "" });
  const [attempt, setAttempt] = useState(0);
  const syncerRef = useRef(null);
  const owner = user?.id;

  useEffect(() => {
    if (!supabaseConfigured || !owner) {
      setState({ status: "error", error: "Supabase belum dikonfigurasi atau sesi login tidak ditemukan." });
      return undefined;
    }
    let alive = true;
    let restore = null;
    setState({ status: "loading", error: "" });
    window.localStorage.removeItem(STORAGE_KEY);
    loadFromSupabase(owner)
      .then((data) => {
        if (!alive) return;
        const raw = JSON.stringify(data);
        originalSetItem.call(window.localStorage, STORAGE_KEY, raw);
        originalSetItem.call(window.localStorage, SYNC_KEY, nowLabel());
        const syncer = createSyncer(owner, buildSnapshot(data, owner));
        syncerRef.current = syncer;
        Storage.prototype.setItem = function patchedSetItem(key, value) {
          originalSetItem.call(this, key, value);
          if (this === window.localStorage && key === STORAGE_KEY) syncer.queue(String(value));
        };
        restore = () => { Storage.prototype.setItem = originalSetItem; };
        const retry = () => syncer.flush();
        window.addEventListener("online", retry);
        const previousRestore = restore;
        restore = () => { window.removeEventListener("online", retry); previousRestore(); };
        setState({ status: "ready", error: "" });
      })
      .catch((error) => {
        if (alive) setState({ status: "error", error: error.message || "Data Supabase gagal dimuat." });
      });
    return () => {
      alive = false;
      if (restore) restore();
      syncerRef.current = null;
    };
  }, [owner, attempt]);

  const signOut = async () => {
    try { await syncerRef.current?.flush(); } catch (error) { console.error(error); }
    window.localStorage.removeItem(STORAGE_KEY);
    window.localStorage.removeItem(SYNC_KEY);
    await onSignOut?.();
  };

  if (state.status === "loading") return h("main", { className: "auth-loading" }, "Memuat data kas dari Supabase…");
  if (state.status === "error") {
    return h("main", { className: "auth-shell" },
      h("section", { className: "auth-card" },
        h("h1", null, "Data belum bisa dimuat"),
        h("p", { className: "auth-message", role: "alert" }, state.error),
        h("button", { className: "primary-btn auth-submit", type: "button", onClick: () => setAttempt((n) => n + 1) }, "Coba lagi"),
        h("div", { className: "auth-links" }, h("button", { type: "button", onClick: signOut }, "Keluar"))));
  }
  return h(App, { user, onSignOut: signOut });
}
