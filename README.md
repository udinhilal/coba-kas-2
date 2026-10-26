# Kas Warga RT 26 / RW 04

## Struktur repository wajib
Upload **isi folder ini** ke root repository GitHub, bukan folder pembungkusnya. `index.html`, `package.json`, `vite.config.js`, `vercel.json`, dan `src/` harus sejajar di root.

```text
repository-root/
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
├── .env.example
└── src/
    ├── main.jsx
    ├── App.jsx
    ├── AuthGate.jsx
    ├── auth.css
    ├── styles.css
    └── lib/supabaseClient.js
```

`index.html` merujuk tepat ke `/src/main.jsx`. Jangan unggah file HTML hasil export artifact sebagai pengganti source project.

## Jalankan lokal
1. Pasang Node.js LTS.
2. Salin `.env.example` menjadi `.env.local`, lalu isi URL dan publishable key dari Supabase.
3. Jalankan `npm install`, lalu `npm run dev`.

## Deploy Vercel
Import repository di Vercel. Atur Framework Preset `Vite`, Build Command `npm run build`, Output Directory `dist`, dan Root Directory ke folder tempat `package.json` berada (umumnya `./`). Tambahkan `VITE_SUPABASE_URL` dan `VITE_SUPABASE_PUBLISHABLE_KEY` pada Environment Variables, lalu deploy ulang.

## Supabase Auth
Atur Supabase Authentication → URL Configuration: Site URL = domain Vercel, Redirect URLs mencakup domain Vercel dan `http://localhost:5173/**`. Aktifkan Email provider.

Catatan: autentikasi disambungkan ke Supabase; data warga/transaksi saat ini tetap tersimpan lokal di browser, belum memakai tabel cloud. Jangan pernah memasukkan `service_role` key ke frontend.
