# PDI Inspection System

Next.js application for the OQC rev.3 PDI form.

## Fitur
- Inspector, Inspection Date, Inspection Line LH/RH, Inspection Lot, Shift.
- Model dropdown mengikuti Line yang dipilih.
- Prefix Part Number otomatis; suffix 6 karakter divalidasi.
- Maksimal 20 Inspection Item; total qty otomatis.
- Mengisi workbook template asli `OQC rev.3` dan mengunduh XLSX.
- Draft dan History disimpan di Local Storage browser.
- Tombol buka Drive dan WhatsApp untuk record yang sudah tersimpan ke Drive.
- Apps Script backend untuk membuat folder tahun/bulan/tanggal, menyimpan Excel, dan mencatat Google Sheets.

## 1. Install dan jalankan
```bash
npm install
npm run dev
```

## 2. Deploy Apps Script
1. Buka https://script.google.com/ dan buat project.
2. Salin isi `google-apps-script/Code.gs`.
3. Isi `CONFIG.ROOT_FOLDER_ID` dan `CONFIG.SPREADSHEET_ID` sesuai ID milik Anda.
4. Jalankan fungsi `setup` satu kali dari editor dan berikan izin Drive + Sheets.
5. Deploy > New deployment > Web app.
6. Execute as: Me. Who has access: Anyone (atau Anyone within domain jika dibatasi organisasi).
7. Salin URL `/exec`.

## 3. Konfigurasi Vercel
Buat Environment Variable:
`NEXT_PUBLIC_GAS_WEB_APP_URL` = URL Web App Apps Script.
Lalu deploy ulang.

## 4. Catatan penting
- Auto-save draft dan History saat ini menggunakan Local Storage. Data hanya tersedia pada browser/perangkat yang sama; bukan sinkronisasi antarperangkat.
- Menghapus record dari History tidak menghapus file Excel di Google Drive.
- B dan C dipertahankan dari template. Pengisian B/C dilakukan setelah file Excel diunduh, menggunakan Excel.
- Hasil file menggunakan sheet `OQC rev.3` dari workbook sumber. Template disimpan di `public/templates/OQC-rev3-template.xlsx`.
- Folder Drive dibuat sebagai `Tahun/Bulan/Tanggal` di dalam folder induk yang ditentukan.
- Untuk keamanan, Apps Script dieksekusi sebagai pemilik deployment. URL web app bersifat endpoint publik jika akses deployment disetel Anyone; pertimbangkan pembatasan akses/domain atau token sebelum digunakan untuk data internal perusahaan.
# Report-PDI-Inspection
