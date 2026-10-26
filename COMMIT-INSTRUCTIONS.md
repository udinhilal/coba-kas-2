# Patch struktur repository siap commit

Bundle ini berisi isi root project Vite, bukan folder pembungkus. Ekstrak ke root repository GitHub sehingga susunannya langsung seperti berikut:

```
.gitignore
.env.example
index.html
package.json
vite.config.js
vercel.json
src/main.jsx
src/App.jsx
src/AuthGate.jsx
src/auth.css
src/styles.css
src/lib/supabaseClient.js
```

## Cara commit ke repository yang sudah ada
1. Unduh dan ekstrak bundle.
2. Di repository GitHub, pilih **Add file → Upload files**.
3. Unggah file dan folder di atas dari hasil ekstrak, jangan unggah folder pembungkus.
4. Pastikan `index.html` dan `package.json` berada di root, dan `src/main.jsx` berada di `src/`.
5. Commit dengan pesan `Fix Vite root structure and add Supabase Auth`.
6. Pastikan Vercel **Root Directory** menunjuk ke lokasi `package.json`, lalu redeploy.

Jangan upload `.env.local` atau `service_role` key. Tambahkan `VITE_SUPABASE_URL` dan `VITE_SUPABASE_PUBLISHABLE_KEY` melalui Vercel Environment Variables.
