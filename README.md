# Stok Dhiascent: app telefon

Dengan app ni, staf boleh rekod stok dari telefon, dan app boleh di-install macam app biasa. Data kekal dalam Google Sheet **Dhiascent Stok Operasi**.

- **Paparan app:** GitHub Pages (`index.html`, `manifest.webmanifest`, `sw.js`, `icons/`)
- **Jambatan ke sheet:** Apps Script, dalam `apps-script/Code.gs`
- **PIN:** tab **Tetapan**, sel **B2** dalam sheet

## Pasang (sekali sahaja)

### 1. Apps Script (jambatan ke sheet)
1. Buka https://script.google.com, pilih **New project**, dan namakan projek **Dhiascent Stok API**.
2. Padam kod contoh, kemudian tampal semua isi `apps-script/Code.gs`. Tekan **Save**.
3. Di bar atas, pilih fungsi `test` dan tekan **Run**. Bila diminta, benarkan akses (Allow).
4. Pergi ke **Deploy → New deployment → Web app** dan tetapkan:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Salin URL yang berakhir dengan `/exec`.

### 2. Sambung app ke Apps Script
Buka `config.js` dan tampal URL `/exec` tadi pada `API_URL`.

### 3. GitHub Pages
Pergi ke **Settings → Pages**, kemudian tetapkan Source: **Deploy from a branch**, Branch: **main** dan folder **/ (root)**.
Link app akan jadi `https://<username>.github.io/dhiascent-stok/`.

### 4. Install di telefon
- **iPhone:** buka link dalam Safari, tekan **Share**, kemudian **Add to Home Screen**.
- **Android:** buka link dalam Chrome, tekan **⋮**, kemudian **Install app**.

## Cara rekod
| Langkah | Bila guna |
|---|---|
| Stock out | Produk diambil untuk packing. Untuk **Set Sinaran**, isi jumlah set sahaja. App akan tolak GA, KK dan Kotak Sinaran sendiri. Perfume 11ml yang customer pilih, isi pada bau tu. |
| Stock in | Refill siap masuk stor. Bahan ditolak automatik ikut resepi. |
| Masuk balik | **Return pelanggan sahaja**, bukan baki stor. |
| Kira stok | Kiraan fizikal. Baki app akan reset kepada nombor yang diisi. Buat selepas semua rekod hari tu dimasukkan. |
| Guna bahan | Bahan rosak, bahan packing dan fridge magnet gift. |
| Terima bahan | Barang baru sampai dari supplier. |

## Kalau ubah kod Apps Script
Selepas ubah kod, pergi ke **Deploy → Manage deployments**, tekan ikon pensel pada deployment yang ada, pilih **Version: New version**, kemudian **Deploy**. Cara ni kekalkan URL `/exec` yang sama.
