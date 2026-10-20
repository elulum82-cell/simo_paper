---
name: paper-generator
description: Menyusun paper ilmiah (jurnal/prosiding) dari dokumen, data, laporan, atau bahan mentah lain mengikuti alur penulisan paper yang baik (IMRaD). Gunakan skill ini setiap kali pengguna meminta membuat, menyusun, menulis, mengubah laporan/kajian/data/skripsi/proposal menjadi paper, artikel jurnal, manuskrip, atau naskah publikasi, termasuk saat pengguna menyebut "paper", "jurnal", "manuskrip", "abstrak", "literature review", "Scopus", "Sinta", atau mengunggah data/dokumen lalu ingin hasilnya jadi tulisan ilmiah. Use also when the user wants to turn documents, datasets, or reports into a research paper, even if they do not say "paper" explicitly.
---

# Paper Generator

Mengubah bahan mentah (dokumen, laporan, data, catatan, hasil model/simulasi) menjadi manuskrip ilmiah yang runtut, jujur terhadap data, dan siap disubmit. Paper yang baik dibangun dari **satu argumen inti** yang didukung bukti, bukan dari menyalin isi dokumen sumber. Karena itu alurnya bertahap: pahami bahan → tentukan klaim → susun kerangka → tulis → uji kualitas.

## Prinsip yang tidak boleh dilanggar

1. **Jangan mengarang.** Angka, hasil, sitasi, dan nama penulis hanya boleh berasal dari bahan pengguna atau dari sumber yang benar-benar ditemukan (web search / dokumen). Alasannya: satu referensi atau angka palsu bisa membuat paper ditolak atau ditarik. Bila data tidak ada, tulis penanda `[DATA DIBUTUHKAN: ...]` dan jangan menebak.
2. **Setiap klaim punya sumber.** Klaim dari bahan pengguna dirujuk ke bagiannya; klaim dari literatur diberi sitasi. Referensi yang tidak bisa diverifikasi ditandai `[VERIFIKASI]`.
3. **Bahasa mengikuti target.** Tanyakan bahasa naskah (Indonesia/Inggris) bila belum jelas. Jangan campur dalam satu naskah.
4. **Transparan soal keterbatasan.** Keterbatasan data dan metode ditulis jujur di Pembahasan; ini justru menaikkan kredibilitas di mata reviewer.

## Alur kerja

### Tahap 0 — Intake (tanya seperlunya, maksimal sekali tanya)
Ambil dulu dari percakapan dan file yang diunggah. Tanyakan hanya yang belum jelas dan menentukan arah:
- Target publikasi: nama jurnal/prosiding, atau tingkat (Scopus Q?, Sinta?, konferensi). Bila ada, minta *author guidelines* (format sitasi, batas kata, struktur).
- Bahasa naskah dan panjang target.
- Jenis paper: penelitian empiris, kajian/studi kasus, pemodelan/simulasi, tinjauan literatur, atau teknis/kebijakan.
- Apakah boleh mencari literatur tambahan di web.

Jika pengguna sudah memberi cukup konteks, langsung lanjut tanpa bertanya.

### Tahap 1 — Inventaris bahan
Baca semua dokumen/data (gunakan skill `file-reading`, `pdf-reading`, `xlsx` sesuai tipe file). Buat tabel ringkas:

| Bahan | Isi utama | Dipakai untuk bagian |
|---|---|---|

Catat: pertanyaan yang dijawab bahan ini, metode yang dipakai, hasil kunci (angka, tabel, grafik), dan **celah** (apa yang belum ada).

### Tahap 2 — Tentukan klaim inti (jantung paper)
Rumuskan dalam 4 kalimat dan tunjukkan ke pengguna sebelum menulis panjang:
- **Masalah/gap**: apa yang belum diketahui atau belum terselesaikan.
- **Tujuan/pertanyaan penelitian**: satu kalimat tegas.
- **Kontribusi/kebaruan**: apa yang baru dibanding literatur.
- **Temuan utama**: 2–3 hasil paling penting dan kuat buktinya.

Bila bahan tidak cukup untuk kebaruan yang jelas, katakan terus terang dan tawarkan reframing (misalnya fokus ke konteks lokal, metode, atau studi kasus).

### Tahap 3 — Kerangka (outline) dan cek kelayakan
Pilih struktur sesuai jenis paper (lihat `references/struktur-jenis-paper.md`). Default IMRaD:
Judul → Abstrak → Kata kunci → Pendahuluan → Tinjauan Pustaka (bila diminta) → Metode → Hasil → Pembahasan → Kesimpulan → Ucapan terima kasih → Daftar pustaka.

Untuk tiap bagian tulis 1–2 baris isi dan bahan sumbernya. Ini juga tempat menyoroti celah data. Untuk permintaan sederhana, tampilkan kerangka singkat lalu lanjut menulis; untuk paper besar, minta persetujuan kerangka dulu.

### Tahap 4 — Literatur (bila diizinkan)
Gunakan web search untuk menemukan 15–40 referensi relevan (prioritaskan jurnal bereputasi, 5–10 tahun terakhir, plus karya klasik yang wajib). Untuk tiap referensi catat penulis, tahun, judul, venue, DOI, dan satu kalimat relevansinya. Jangan menulis sitasi dari ingatan saja; tandai `[VERIFIKASI]` bila tidak bisa dikonfirmasi lewat pencarian. Jangan menyalin teks sumber; parafrasekan dan sitasi.

### Tahap 5 — Menulis per bagian
Tulis berurutan. **Abstrak ditulis terakhir** agar sesuai isi.

- **Judul**: spesifik, memuat objek, metode/pendekatan, dan konteks; hindari kata mubazir.
- **Pendahuluan** (pola corong): konteks luas → masalah → apa yang sudah dilakukan peneliti lain → **gap** → tujuan dan kontribusi → peta isi paper. Gap harus didukung sitasi, bukan sekadar pernyataan.
- **Metode**: cukup rinci agar bisa direplikasi: lokasi/objek, sumber dan periode data, variabel, rumus/model, asumsi, alat/perangkat lunak, validasi/uji. Jelaskan alasan memilih metode.
- **Hasil**: sajikan temuan secara objektif, tanpa tafsir panjang. Satu tabel/gambar untuk satu pesan; beri nomor dan keterangan yang mandiri; tabel/gambar harus dirujuk di teks. Sertakan satuan dan ketidakpastian bila ada.
- **Pembahasan**: tafsirkan (mengapa hasil begini), bandingkan dengan literatur (sesuai/bertentangan dan alasannya), implikasi praktis/kebijakan, **keterbatasan**, dan arah riset lanjutan.
- **Kesimpulan**: menjawab pertanyaan penelitian satu per satu, tanpa informasi atau angka baru, singkat.
- **Abstrak** (150–250 kata, sesuaikan jurnal): latar singkat, tujuan, metode, hasil kunci dengan angka, kesimpulan/implikasi. Satu paragraf, tanpa sitasi.
- **Kata kunci**: 4–6 istilah yang dipakai di judul/abstrak.

Gaya: kalimat aktif dan ringkas, satu ide per paragraf, istilah konsisten, singkatan dijelaskan saat pertama muncul, hindari klaim berlebihan ("membuktikan", "pertama kali") kecuali ada dasar.

### Tahap 6 — Uji kualitas sebelum diserahkan
Jalankan daftar periksa di `references/checklist-kualitas.md`. Ringkasnya:
- Pertanyaan penelitian ↔ metode ↔ hasil ↔ kesimpulan saling menjawab (tidak ada yang menggantung).
- Semua angka di abstrak/kesimpulan cocok dengan bagian Hasil dan bahan sumber.
- Semua tabel/gambar dirujuk; semua sitasi ada di daftar pustaka dan sebaliknya.
- Tidak ada teks yang menyalin sumber; tidak ada penanda `[DATA DIBUTUHKAN]` / `[VERIFIKASI]` yang terlewat tanpa dilaporkan.
- Format sesuai pedoman jurnal target.

Lalu tulis **laporan singkat** untuk pengguna: kekuatan paper, 3–5 risiko penolakan reviewer yang paling mungkin, daftar data/verifikasi yang masih dibutuhkan.

## Format keluaran

- Default: file **.docx** (gunakan skill `docx`) dengan judul, heading bertingkat, tabel, dan daftar pustaka. Jika pengguna meminta file lain (PDF, Markdown), ikuti. Bila diminta cepat/draf awal, boleh tampil di chat dulu lalu tawarkan file.
- Gambar/grafik: buat dari data pengguna dengan Python (matplotlib) dan sisipkan dengan keterangan; jangan membuat grafik dari angka rekaan.
- Pisahkan keluaran: (1) naskah paper, (2) lampiran catatan kerja (inventaris bahan, daftar referensi bertanda verifikasi, daftar celah data).

## Menangani situasi umum

- **Bahan tipis/tidak lengkap**: tulis sebisanya dengan penanda `[DATA DIBUTUHKAN]`, dan di akhir daftar apa yang perlu dilengkapi. Jangan menambal dengan data karangan.
- **Bahan berupa laporan proyek/kajian teknis**: ubah gaya dari "laporan kegiatan" menjadi "argumen ilmiah". Buang rincian administratif, angkat pertanyaan penelitian, metode, dan temuan.
- **Banyak dokumen yang saling bertentangan**: laporkan pertentangannya ke pengguna dan minta keputusan sebelum menulis Hasil.
- **Revisi**: bila pengguna mengunggah naskah dan komentar reviewer, buat tabel respons (komentar → perubahan → lokasi) dan edit naskah yang ada, bukan menulis ulang dari nol.
- **Etika**: ingatkan bila ada kebutuhan izin data, persetujuan etik, atau pengakuan penggunaan alat AI sesuai kebijakan jurnal target.
