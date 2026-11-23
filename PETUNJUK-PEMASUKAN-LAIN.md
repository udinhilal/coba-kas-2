# Petunjuk aktivasi fitur Pemasukan Lain

Fitur Pemasukan Lain memakai tabel Supabase baru bernama `other_income`. Tabel ini terpisah dari `payments` dan `expenses`, serta hanya dapat diakses oleh akun pemilik data melalui kebijakan RLS.

## Urutan pemasangan

1. Buka **Supabase Dashboard → SQL Editor → New query**.
2. Buka file `SUPABASE-PEMASUKAN-LAIN.sql` pada folder utama proyek ini.
3. Salin seluruh isi SQL tersebut ke SQL Editor, lalu pilih **Run** satu kali.
4. Setelah tabel berhasil dibuat, unggah perubahan kode proyek ini ke GitHub dan biarkan Vercel melakukan deployment seperti biasanya.
5. Buka aplikasi dan periksa menu **Arus Kas**.

Jalankan SQL sebelum deployment kode baru. Versi kode baru akan membaca tabel `other_income` saat akun masuk ke aplikasi. Tanpa tabel tersebut, proses pemuatan data Supabase dapat gagal.

## Dampak ke data lama

- Tabel `residents`, `payments`, `expenses`, `balance_history`, `cash_settings`, dan `community_identity` tidak dihapus atau diubah oleh SQL ini.
- Baris pembayaran dan pengeluaran yang sudah ada tetap berada di tabelnya masing-masing.
- Tabel `other_income` dibuat kosong dan akan menerima catatan Pemasukan Lain yang baru.
- Backup JSON versi sebelumnya tetap dapat dipulihkan; jika belum memiliki properti `otherIncome`, daftar Pemasukan Lain akan dimulai kosong.
