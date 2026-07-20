# Generate PowerPoint Cutover Presentation

Repository ini berisi script Python untuk membuat file presentasi PowerPoint (`.pptx`) terkait rencana cutover implementasi aplikasi.

## File Utama
- `generate_ppt.py` → script pembuat file PowerPoint
- `materi_presentasi_cutover.txt` → outline materi presentasi per slide

## Output
Saat script dijalankan, file berikut akan dibuat:

- `cutover_implementation_presentation.pptx`

## Prasyarat
Pastikan Python sudah terinstall di komputer Anda.

### Cek versi Python
```bash
python --version
```
atau
```bash
python3 --version
```

Disarankan menggunakan Python 3.9 atau lebih baru.

## Install Dependency
Script ini membutuhkan library `python-pptx`.

Install dengan perintah:
```bash
pip install python-pptx
```

Jika menggunakan Python 3 secara terpisah:
```bash
pip3 install python-pptx
```

## Cara Menjalankan
Masuk ke folder repository, lalu jalankan:

```bash
python generate_ppt.py
```

atau jika environment Anda menggunakan `python3`:

```bash
python3 generate_ppt.py
```

## Hasil
Jika berhasil, file berikut akan terbentuk di folder yang sama:

```text
cutover_implementation_presentation.pptx
```

Script juga akan menampilkan pesan seperti berikut:

```text
Presentation created: cutover_implementation_presentation.pptx
```

## Struktur Slide
PowerPoint yang dihasilkan berisi slide berikut:
1. Judul
2. Latar Belakang Perubahan
3. Current State vs Target State
4. Tujuan Implementasi
5. Metode Cutover
6. Tahapan Implementasi
7. Alur Hari H Cutover
8. Risiko Utama dan Mitigasi
9. Backup Plan / Rollback Plan
10. Kesimpulan dan Permohonan Persetujuan

## Jika Terjadi Error
### 1. Module not found
Jika muncul error seperti:
```text
ModuleNotFoundError: No module named 'pptx'
```
solusinya jalankan:
```bash
pip install python-pptx
```

### 2. Command python tidak dikenali
Coba gunakan:
```bash
python3 generate_ppt.py
```

### 3. Permission error
Pastikan Anda memiliki hak akses tulis pada folder repository.

## Kustomisasi
Jika ingin mengubah isi presentasi, Anda dapat mengedit file:
- `generate_ppt.py` untuk layout, warna, dan struktur visual
- `materi_presentasi_cutover.txt` untuk referensi konten slide

## Saran Penggunaan
Setelah file `.pptx` berhasil dibuat:
- buka di Microsoft PowerPoint
- sesuaikan logo perusahaan
- sesuaikan warna tema perusahaan
- tambahkan nama presenter, tanggal, dan unit kerja
- lakukan review final sebelum presentasi ke manajemen

## Ringkasan Cepat
```bash
pip install python-pptx
python generate_ppt.py
```

Output:
```text
cutover_implementation_presentation.pptx
```
