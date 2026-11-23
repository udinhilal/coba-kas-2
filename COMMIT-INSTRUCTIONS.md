# Pembaruan fitur Pemasukan Lain

Bundle ini berisi isi root project Vite. Saat memperbarui repository GitHub yang sudah ada, unggah dan timpa file dengan nama yang sama sesuai struktur di bawah ini. Jangan mengunggah folder pembungkus ZIP sebagai satu tingkat tambahan.

```text
repository-root/
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
├── README.md
├── SUPABASE-PEMASUKAN-LAIN.sql
├── PETUNJUK-PEMASUKAN-LAIN.md
└── src/
    ├── App.jsx
    ├── CloudApp.jsx
    ├── AuthGate.jsx
    ├── LoadingScreen.jsx
    ├── auth.css
    ├── styles.css
    ├── main.jsx
    └── lib/supabaseClient.js
```

## Urutan pembaruan

1. Sebelum deployment kode, buka Supabase Dashboard → SQL Editor → New query.
2. Salin isi `SUPABASE-PEMASUKAN-LAIN.sql`, lalu jalankan satu kali. SQL hanya membuat tabel `other_income` beserta RLS; tabel transaksi lama tidak dihapus.
3. Di GitHub, pilih **Add file → Upload files**. Unggah file/folder proyek hasil ekstrak ke root repository dan timpa file dengan nama yang sama.
4. Commit perubahan, misalnya dengan pesan `Add other income records to cashflow`.
5. Pastikan Vercel **Root Directory** tetap menunjuk ke lokasi `package.json`, lalu pantau deployment.

Jangan mengunggah `.env.local` atau `service_role` key. Variabel `VITE_SUPABASE_URL` dan `VITE_SUPABASE_PUBLISHABLE_KEY` tetap dikelola melalui Vercel Environment Variables.
