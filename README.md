# ReferMatch - Rekonsiliasi Referensi Transaksi Otomatis

ReferMatch adalah tool Python yang saya rancang untuk menjawab kebutuhan nyata: mencocokkan referensi pembayaran antar dua sumber data '.xslx.
Proyek ini lahir dari pengalaman harian di kantor - ketika data datang berceceran, tapi yang dibutuhkan hanya satu hal: *apakah referensi ini benar-benar cocok?* 

-----

## Fitur Utama

- Input dua file '.xlsx' (data transaksi dari Bank & referensi pembayaran)
- Deteksi Referensi & Nilai secara otomatis
- Tampil hasil cocok/tidak cocok dalam format tabel
- Proses Cepat: 700+ baris dalam waktu <1 detik!
- GUI sederhana untuk akses non-teknis
- 100% Gratis, Lisensi MIT, dan bisa digunakan untuk semua orang

## Struktur Folder
- 'assets' 
- 'dummy_data\' - berisi file dummy excel untuk demo
- 'src\' - file python utama
- LICENSE - Licensi open-source (MIT)
- README.md - Dokumen ini

---

## File Dummy

Folder 'dummy_data/' berisi contoh transaksi dan referensi untuk demo YouTube. Bukan file transaksi dan Referensi yang sebenarnya.
Dirancang menyerupai format umum laporan keuangan kantor, dengan elemen seperti:

- Kolom Referensi / Kode Transaksi
- Nominal / Jumlah Pembayaran
- Skenario cocok, tidak cocok, dan kasus kosong

## Batasan saat ini

- Program hanya bekerja pada extensi '.xlsx' belum '.xls' atau '.pdf' yang didalamnya terdapat tabel transaksi.
- Program belum bisa membaca Multiple Sheet pada file excel, baru bisa mendukung 1 sheet dalam file excel.


## Catatan Pribadi

ReferMatch bukan proyek besar, namun ia lahir dari tantangan kecil yang berulang setiap minggu.
Jika kamu juga punya data yang cocok, semoga alat ini bisa mempermudah proses kerjamu.

---

## Lisensi
Proyek ini tersedia secara bebas di bawah [MIT License](./License), dan dapat digunakan, dimodifikasi, serta dibagikan, selama mencantumkan pembuat asli: Fathir Afatar.

---

## Download GUI versi .exe
Versi executable (.exe) tersedia di [GitHub Releases]().  
File ini tidak disertakan dalam repo utama untuk menjaga ukuran tetap ringan.


## Kontribusi & Kontak
Ingin memberi saran, melaporkan bug, atau berkontribusi ke proyek ini?.
Silakan kunjungi halaman [GitHub repositori](https://github.com/Fathir-Afatar/) dan beri bintang bila project ini bermanfaat.
