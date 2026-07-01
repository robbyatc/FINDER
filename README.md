# FINDER — DAT vs STREAM Reconciliation System

FINDER mendeteksi data penerbangan pada **DAT DEP** dan **DAT ARR** yang belum tercatat di **STREAM** untuk mendukung validasi billing.

## Aplikasi online

Buka FINDER dari komputer, tablet, atau ponsel melalui:

**https://finder-dat-stream.streamlit.app/**

## Fitur utama

- Tiga area upload: DAT DEP, DAT ARR, dan STREAM.
- Mendukung CSV, TSV, XLS, XLSX, serta laporan tabel HTML yang berekstensi `.xls`.
- Pemetaan otomatis untuk Flight Number/Callsign, ADEP, ADES, EOBD, ATD, ATA, gate, runway, register, dan movement type.
- Normalisasi tanggal, jam, kapitalisasi, spasi, tanda hubung, serta tanda petik bawaan DAT.
- Deduplikasi DAT berbasis movement time, completeness, timestamp, dan message number.
- Validasi STREAM berdasarkan status flight dan selisih waktu 15/30/60 menit.
- Hasil: Missing in Stream, Matched, Need Review, Extra in Stream, dan Duplicate DAT.
- Hard exclude Non-Billable/Internal Movement sebelum recovery, deduplikasi, dan reconciliation.
- Adjacent Date / Midnight Recovery dari RAW DAT yang sudah dinormalisasi.
- Matching utama STREAM melalui tanggal actual movement (ATD untuk departure, ATA untuk arrival), lalu original/recovered date sebagai fallback.
- Special remark DIVERT/RTB/RETURN/REROUTE/ALTERNATE dapat dicocokkan tanpa route dalam window tanggal ±1 hari dan diarahkan ke Perlu Review STREAM.
- Recovered movement date validation dan AC-register mismatch routing ke Need Review.
- Deduplikasi per flight instance agar pergerakan berbeda tidak saling menimpa.
- Audit reason untuk setiap hasil dan candidate STREAM.
- Kolom Actual Movement Date dari ATA untuk arrival atau ATD untuk departure, tanpa mengubah EOBD sebagai base key.
- Filter Missing in Stream berdasarkan tanggal, flight number, aerodrome, TO FROM, dan movement.
- Kolom VALIDASI: ADA DI STREAM, ADA DI DAT TIDAK ADA DI STREAM, atau PERLU REVIEW STREAM.
- Dashboard metrik Validasi dan accuracy percentage.
- Laporan Excel sembilan sheet, termasuk Validasi, Audit Detail, dan Excluded Non-Billable.

## Menjalankan aplikasi

### macOS

Klik dua kali `run_app.command`, atau jalankan:

```bash
cd /lokasi/flight-data-comparator
./run_app.command
```

### Windows

Klik dua kali `run_app.bat`.

Pada penggunaan pertama, launcher membuat virtual environment dan memasang dependensi. Browser membuka `http://localhost:8501`.

## Alur operator

1. Buka menu **Upload Data**.
2. Unggah DAT DEP, DAT ARR, dan STREAM.
3. Pastikan validation checklist berwarna hijau.
4. Klik **PROCESS RECONCILIATION**.
5. Buka **Reconciliation Result** dan periksa tab **Validasi**.
6. Gunakan filter untuk mempersempit pemeriksaan.
7. Buka **Export Report** dan klik **Download Excel Report**.

## Logika pencocokan

Record diprioritaskan berdasarkan tanggal/waktu actual movement dan kombinasi:

- ATD untuk departure atau ATA untuk arrival
- Flight Number / Callsign
- ADEP / Aerodrome
- ADES / TO FROM
- Movement Type (D/A/L/O)

Date of Flight/EOBD tetap ditampilkan dan digunakan sebagai fallback bersama recovered date. Sheet Validasi hanya memuat DAT yang tidak memiliki kandidat STREAM pada seluruh jalur pencarian; kandidat dengan masalah status, tanggal, waktu, register, atau special remark diberi VALIDASI `PERLU REVIEW STREAM`.

Untuk STREAM ber-remark `DIVERT`, `DIVERSION`, `DVT`, `RTB`, `RETURN TO BASE`, `RETURN`, `REROUTE`, `ALTN`, atau `ALTERNATE`, AERODROME dan TO FROM boleh diabaikan. Flight Number, movement type, register (jika tersedia), serta ATD/ATA dalam tolerance tetap wajib sesuai. Audit menyimpan remark, keyword, flag special remark, dan flag route ignored.

Untuk setiap base key, FINDER memilih satu DAT terbaik: movement time terisi, completeness score tertinggi, timestamp terbaru, lalu message number terbesar. Hanya record yang tidak terpilih masuk Duplicate DAT; record terbaik selalu dipakai dalam reconciliation.

Candidate STREAM harus memiliki status valid, tanggal movement yang sesuai, dan selisih waktu dalam tolerance. Selisih di atas tolerance sampai 120 menit masuk Need Review; selisih lebih dari 120 menit, tanggal berbeda, waktu invalid, atau status invalid masuk Missing in Stream. Accuracy dihitung sebagai jumlah Matched dibagi total DAT Unique.

Pada versi lokal, file diproses di komputer operator. Pada versi cloud, file diproses oleh sesi aplikasi Streamlit dan tidak dikomit ke repository GitHub oleh FINDER.
