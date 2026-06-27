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
- Deduplikasi DAT sebelum pencocokan.
- Hasil: Missing in Stream, Matched, Extra in Stream, dan Duplicate DAT.
- Filter Missing in Stream berdasarkan tanggal, flight number, aerodrome, TO FROM, dan movement.
- Dashboard metrik dan accuracy percentage.
- Laporan Excel dengan sheet Summary, Missing in Stream, Matched, Extra in Stream, dan Duplicate DAT.

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
5. Buka **Reconciliation Result** dan periksa tab **Missing in Stream**.
6. Gunakan filter untuk mempersempit pemeriksaan.
7. Buka **Export Report** dan klik **Download Excel Report**.

## Logika pencocokan

Record dicocokkan menggunakan kombinasi:

- Date of Flight / EOBD
- Flight Number / Callsign
- ADEP / Aerodrome
- ADES / TO FROM
- Movement Type (D/A/L/O)

Duplicate DAT dikeluarkan sebelum pencocokan. Accuracy dihitung sebagai jumlah Matched dibagi total DAT unik.

Pada versi lokal, file diproses di komputer operator. Pada versi cloud, file diproses oleh sesi aplikasi Streamlit dan tidak dikomit ke repository GitHub oleh FINDER.
