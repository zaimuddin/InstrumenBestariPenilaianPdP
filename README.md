# Instrumen Bestari Penilaian PdP

Aplikasi percuma ini dibangunkan oleh [SolusiBestariGuru](https://t.me/solusibestariguru) untuk membantu pengurusan Penilaian PdP bagi sesebuah sekolah di Malaysia. Aplikasi ini merupakan sebuah **Google Apps Script web app** yang beroperasi sepenuhnya di dalam satu Google Sheet.

Buat masa sekarang, aplikasi ini menyediakan kemudahan pengurusan bersepadu bagi Instrumen Standard Penilaian Pembudayaan KBAT, di mana sekolah dapat merekodkan penilaian pembudayaan KBAT secara dalam talian dan menghantar rekod penilaian kepada KPM melalui borang Google Forms rasmi bagi setiap negeri.

---

## 🚀 Panduan Pemasangan

Untuk menjadikan aplikasi ini milik sekolah anda secara percuma, sila ikuti langkah-langkah di bawah:

### 1. Salin Google Sheet

1. Buka pautan [Salinan v2026.09](https://docs.google.com/spreadsheets/d/1epAXlmcAco672fQo1izhsnUhcYWhe809izgoweulMog/copy) (disarankan untuk menggunakan komputer/komputer riba).
2. Pastikan anda menggunakan akaun DELIMa peribadi (`g-12345678@moe-dl.edu.my`) atau akaun DELIMa sekolah (`sekolah-0000-xxx@moe-dl.edu.my`) untuk membuka pautan di atas.
3. Tekan butang `Make a copy`.

### 2. Konfigurasi Asas

1. Tukar nama fail Google Sheet yang disalin kepada nama yang deskriptif (contoh: buang "Copy of " dan tambah nama sekolah anda di hujung nama Google Sheet).
2. Buka sheet **"GURU"** dan isikan emel DELIMa guru, nama guru sekolah, serta jawatan, jantina dan opsyen (boleh dibiarkan kosong dahulu untuk diisikan sendiri oleh guru berkenaan). Bagi lajur 'Pencerap', tulis 'ya' jika guru tersebut adalah pencerap yang akan menilai, sebaliknya tulis 'tidak' atau biarkan kosong. Bagi 'Pencerap', senarai guru ini dapat diuruskan kemudian dalam antara muka aplikasi.
3. Buka sheet **"TETAPAN"** dan isikan maklumat sekolah.

### 3. Tetapan Apps Script

1. Pada menu atas Google Sheet, klik **Extensions** → **Apps Script**. (Tetingkap Editor Apps Script akan dibuka).
2. Klik butang biru `Deploy` dan pilih **New deployment**.
3. Isikan _New description_ (jika perlu).
4. Tetapkan _Execute as_: **Me** (emel anda/sekolah), contoh 'Me (g-12345678@moe-dl.edu.my)'.
5. Tetapkan _Who has access_: **Anyone within Ministry Of Education Malaysia**.
6. Klik butang `Deploy`.
7. Klik `Authorize access`, log masuk akaun Google anda, dan **WAJIB tandakan (tick)** `Select all` sebelum menekan `Continue`.
8. Selepas dikembalikan ke Apps Script, klik butang `Copy` untuk menyalin URL baharu ini, kemudian klik `Done`. (Anda kini boleh menutup tab Apps Script Editor).
9. Kembali ke Google Sheet sheet **"TETAPAN"**, dan tampalkan URL ini (_Paste_ / _Ctrl+V_) pada **sel B4** (bersebelahan teks **URL Web App**).

### 4. Pemasangan Sistem

1. Pada bar menu atas Google Sheet, klik **⚙️ Instrumen Bestari PPdP** → **🚀 Pasang dan Sahkan Sistem**.
2. Tunggu sebentar sehingga tetingkap "Pemasangan Selesai" dipaparkan. Anda boleh menutupnya menggunakan butang **"OK"**.
3. **Selesai!** Anda boleh melalukan kursor pada URL di sel B4 dan klik pautan _pop-over_ tersebut untuk melancarkan aplikasi.

> 💡 **MAKLUMAN:** Kongsikan **URL di sel B4** kepada rakan guru untuk kegunaan sekolah anda.

---

## 🎓 Cara Penggunaan bagi Guru Mata Pelajaran - Mod Guru/Kendiri

1. Pastikan anda berada dalam `Mod Guru` sebagaimana terpapar di bahagian atas aplikasi.
2. Di bahagian kad **BUTIRAN GURU**, sila pilih nama anda dalam senarai `1. Nama Guru Dicerap` dan kemaskinikan maklumat diri anda.
   - Jika nama anda tiada dalam senarai, sila maklumkan kepada mana-mana guru yang berstatus 'Pencerap' untuk memasukkan nama anda ke dalam senarai guru.
3. Di bahagian kad **BUTIRAN PENCERAPAN**, sila lengkapkan semua maklumat yang diperlukan.
4. Di setiap bahagian penilaian, sila lengkapkan semuanya dengan skor berdasarkan rubrik yang disediakan. Hanya klik butang-butang skor yang sesuai bagi setiap item.
5. Di bahagian kad **REFLEKSI KENDIRI GURU**, sila tuliskan sedikit refleksi berkaitan dengan PdP yang dijalankan.
6. Setelah selesai, klik butang `Simpan Pengisian` (hijau) di bahagian bawah kad **REFLEKSI KENDIRI GURU** atau di bahagian atas halaman.
   - Fail _.json_ yang mengandungi maklumat penilaian kendiri anda akan dimuat turun ke dalam peranti anda, maklumat penilaian juga disimpan secara dalam talian, dan emel pengesahan juga dihantar kepada anda.
   - Nama anda akan dipadamkan daripada senarai `1. Nama Guru Dicerap` supaya tidak lagi boleh diisi ulang.

---

## 📝 Cara Penggunaan bagi Pencerap - Mod Pencerap

1. Pastikan anda berada dalam `Mod Pencerap` sebagaimana terpapar di bahagian atas aplikasi.
2. Di bahagian kad **BUTIRAN GURU**, sila pilih nama seseorang guru dalam senarai `1. Nama Guru Dicerap` dan klik pada butang `📤Muat Rekod` (biru) di bahagian atas kanan kad **BUTIRAN GURU**.
3. Di bahagian kad **BUTIRAN PENCERAPAN**, sila semak semula maklumat yang sedia ada.
4. Di setiap bahagian penilaian, sila lengkapkan semuanya dengan skor berdasarkan rubrik yang disediakan. Hanya klik butang-butang skor yang sesuai bagi setiap item.
5. Di bahagian kad **RUMUSAN & PENGESAHAN PENCERAP**, sila tuliskan sedikit rumusan keseluruhan berkaitan dengan PdP yang dilaksanakan.
6. Setelah selesai, klik butang `Simpan Pengisian` (hijau) di bahagian bawah kad **RUMUSAN & PENGESAHAN PENCERAP** atau di bahagian atas halaman.
   - Maklumat penilaian anda akan disimpan secara dalam talian.
   - Seterusnya, anda dapat membuat penilaian pencerapan bagi guru yang lain pula dengan memilih nama guru dalam senarai `1. Nama Guru Dicerap` dan klik pada butang `📤Muat Rekod` (biru) di bahagian atas kanan kad **BUTIRAN GURU**.
   - Jika anda ingin memadamkan rekod penilaian pencerapan bagi seseorang guru, sila klik pada butang `🗑️Hapus Pencerapan` (merah) di bahagian atas kanan kad **BUTIRAN PENCERAPAN**.

---

## 📋 Cara Menghantar Rekod Pencerapan ke Google Forms - Mod Pencerap

1. Pastikan anda berada dalam `Mod Pencerap` sebagaimana terpapar di bahagian atas aplikasi.
2. Klik pada butang `📋Hantar ke Google Forms` (hijau) di bahagian atas halaman Aplikasi. Tetingkap baharu akan dibuka.
3. Di dalam tetingkap baharu tersebut, tampalkan pautan/URL Google Forms Pengisian Skor ISPPK JPN pada bahagian input yang disediakan.
4. Klik butang `⚡Proses URL` (hijau) atau `🔎Dapatkan Elemen Soalan` (biru) (dalam tab `📝1. Senarai Elemen Soalan`) untuk memuat turun elemen soalan dari Google Forms.
5. Buka tab `🔄2. Pemetaan Elemen Soalan`, kemudian klik butang `⚡Auto-Padan Pintar` untuk membuat pemetaan elemen soalan secara automatik dengan data pencerapan sekolah anda.
6. Lakukan **Pemetaan Manual** bagi elemen soalan `NEGERI:`, `PPD:` dan `Nama Sekolah` anda. Pastikan semua pemetaan manual bertanda label `✅Tetap` (hijau).
7. Klik butang `💾Simpan Pemetaan` (hijau) di sebelah butang `⚡Auto-Padan Pintar` untuk menyimpan pemetaan elemen soalan anda.
8. Buka tab `🚀3. Hantar ke Google Forms`.
9. Klik pada setiap butang `📄Buka Pra-Isi Pintar` di hujung baris setiap data pencerapan guru untuk membuka Google Form dengan data yang telah dipra-isi.
10. Klik `Submit` atau `Hantar` pada Google Form untuk menghantar data pencerapan bagi guru berkenaan, kemudian anda boleh menutup Google Form tersebut.
11. Anda akan disoal untuk pengesahan penghantaran tersebut. Sila jawab mengikut tindakan anda dalam Google Form tersebut.

---

## 📞 Sokongan & Pertanyaan

### Sekiranya terdapat sebarang ralat, cadangan penambahbaikan, atau pertanyaan lanjut, sila sertai ruang komuniti **SolusiBestariGuru** di Telegram:

**Bincang & Sokongan:**

- 📱 **Kumpulan Sokongan:** [@bincangsolusibestariguru](https://t.me/bincangsolusibestariguru)
- 📱 **Saluran Maklumat:** [@solusibestariguru](https://t.me/solusibestariguru)
- 📱 **Hubungi Ustaz:** [@zaimuddinhassan](https://t.me/zaimuddinhassan)

**Peraturan Kumpulan:**

- ✅ Bertanya untuk bantuan adalah digalakkan
- ⚠️ Elakkan spam atau iklan tidak berkaitan
- 🤝 Bantu guru lain jika anda tahu jawapannya

### 🐛 Laporan Ralat & Permintaan Fitur

**Via GitHub Issues:**

1. Lawati: [InstrumenBestariPenilaianPdP/issues](https://github.com/zaimuddin/InstrumenBestariPenilaianPdP/issues)
2. Klik **"New Issue"**
3. Pilih kategori:
   - 🐛 **Bug Report** - untuk melaporkan masalah atau ralat
   - ✨ **Feature Request** - untuk idea fitur baharu
4. Tulis perihalan secara terperinci
5. Kepilkan tangkap layar jika dapat
6. Submit issue

---

## 👥 Penyumbang & Lesen

### Maklumat Penyumbang

**Penulis Utama:**

- **Ustaz Zaimuddin Hassan**
  - 👨‍🏫 Guru Al-Quran dan Bahasa Arab
  - 🏠 SMK Padang Pak Amat, Pasir Puteh, Kelantan
  - 💬 Telegram: [@zaimuddinhassan](https://t.me/zaimuddinhassan)

### Lesen - Attribution-NonCommercial-ShareAlike 4.0 International

Projek ini dilesenkan di bawah **Attribution-NonCommercial-ShareAlike 4.0 International**.

- Anda bebas untuk:

1. **Kongsi** — menyalin dan menyebarkan semula bahan tersebut dalam sebarang medium atau format.
2. **Adaptasi** — menggabungkan, mengubah, dan membina berasaskan bahan tersebut.
3. Pemberi lesen tidak boleh menarik balik kebebasan ini selagi anda mematuhi terma lesen.

- Di bawah terma-terma berikut:

1. **Pengiktirafan** — Anda mesti memberikan _kredit yang sesuai_, menyediakan pautan ke lesen, dan _menyatakan jika perubahan telah dibuat_. Anda boleh melakukannya dalam apa jua cara yang munasabah, tetapi bukan dengan cara yang mencadangkan pemberi lesen menyokong anda atau penggunaan anda.
2. **Bukan Komersial** — Anda tidak boleh menggunakan bahan ini untuk _tujuan komersial_.
3. **Perkongsian Serupa** — Jika anda menggabungkan, mengubah, atau membina berasaskan bahan ini, anda mesti mengedarkan sumbangan anda di bawah _lesen yang sama_ dengan yang asal.
4. **Tiada sekatan tambahan** — Anda tidak boleh mengenakan terma undang-undang atau _langkah-langkah teknologi_ yang menyekat orang lain secara sah daripada melakukan apa-apa yang dibenarkan oleh lesen.

**Teks Penuh Lesen:** Lihat fail `LICENSE` dalam repositori ini.

---

**Selamat menggunakan Instrumen Bestari Penilaian PdP! 🎉**
