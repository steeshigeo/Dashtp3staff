# 🍏 iOS Liquid Glass Sales Dashboard

Dashboard pemantauan pencapaian sales staff dengan UI bertema **iOS Liquid Glass** modern. Terhubung otomatis dengan file Excel di OneDrive melalui GitHub Actions.

## 🚀 Fitur Utama
- **Auto Sync OneDrive:** GitHub Action mendownload & memperbarui data Excel setiap jam secara otomatis.
- **Auto Parsing Staff & Tanggal:** Membaca format khusus `22013473 / DEWI ANDANSARI` dan tanggal `DD-MM-YYYY`.
- **Auto Categorization:** 
  - **VAS:** Terdeteksi otomatis via kata kunci `KLA`, `TSL`, `IDT`, dan `XXL`.
  - **Device & Accessories:** Terkelompokkan secara cerdas.
- **LOB Focus Target Tracker:** Sisa kuota item target (`iPhone 15`, `iPad A16`, `MBN`, `AW SE`, `AirPods`) berkurang otomatis saat terjadi penjualan.
- **LFL WoW Comparison:** Perbandingan pencapaian week berjalan dengan week sebelumnya per staff.
- **Manual Target Input:** Input fleksibel untuk target Monthly, Weekly, Daily, dan Qty LOB.

---

## 🛠️ Cara Deploy ke GitHub

1. Buat **Repository Baru** di GitHub (misal: `sales-dashboard`).
2. Buat file-file di atas sesuai dengan struktur folder.
3. Tambahkan **Secret GitHub** (Opsional jika ingin mengganti link OneDrive):
   - Masuk ke **Settings** > **Secrets and variables** > **Actions**.
   - Tambahkan secret baru: `ONEDRIVE_DIRECT_URL`.
   - Isi nilainya dengan link sharing OneDrive Anda.
4. Aktifkan **GitHub Pages**:
   - Masuk ke **Settings** > **Pages**.
   - Pilih `Source: Deploy from a branch` -> Branch `main` / `root`.
   - Klik **Save**. Dashboard Anda akan langsung tayang secara publik!
