#  Aplikasi Pencatatan Keuangan Harian (Python GUI)

Aplikasi desktop berbasis Python dengan antarmuka grafis (GUI) yang dirancang untuk membantu mengelola dan mencatat keuangan harian secara praktis. Dilengkapi dengan fitur ekspor data ke format Excel (`.xlsx`) agar laporan keuangan lebih mudah dianalisis.

---

##  Fitur Utama

- **Pencatatan Transaksi Harian:** Catat pemasukan dan pengeluaran dengan mudah dan cepat.
- **Antarmuka Grafis (GUI):** Dibangun dengan antarmuka yang intuitif dan ramah pengguna untuk pengalaman navigasi yang mulus.
- **Ekspor ke Excel (.xlsx):** Simpan dan unduh laporan catatan keuangan secara otomatis ke dalam format berkas Excel yang rapi dan terformat.
- **Ikon Kustom:** Dilengkapi dengan logo ikon khusus untuk memberikan identitas visual pada aplikasi.

---

## Teknologi yang Digunakan

- **Bahasa Pemrograman:** Python
- **GUI Framework:** *Tkinter*
- **Pengolahan Data & Ekspor Excel:** *openpyxl*

## Penerapan Konsep OOP (Object-Oriented Programming)

Aplikasi pencatatan keuangan ini dibangun dengan paradigma **Object-Oriented Programming (OOP)** di Python untuk menjaga struktur kode tetap modular, terorganisir, dan mudah dikembangkan.

### 1. Class & Object
Program memisahkan tanggung jawab antara antarmuka pengguna (GUI), logika bisnis/transaksi, dan pengelolaan file Excel ke dalam class terpisah:
* **`Transaksi` (Model Data):** Merepresentasikan setiap entitas catatan keuangan (tanggal, jenis transaksi, kategori, nominal, dan keterangan) sebagai sebuah *object*.
* **`PengelolaKeuangan` / `ExcelManager` (Data & Export Handler):** Bertanggung jawab menyimpan daftar objek transaksi, menghitung total pemasukan/pengeluaran/saldo, serta mengekspor data ke format `.xlsx`.
* **`AplikasiKeuanganGUI` (View & Controller):** Mengelola komponen antarmuka grafis (window, form input, tabel riwayat, dan tombol aksi) serta menghubungkan interaksi pengguna dengan logika data.

### 2. Encapsulation (Enkapsulasi)
Data transaksi dan logika pemrosesan dibungkus di dalam class masing-masing agar tidak dimodifikasi sembarangan dari luar:
* Atribut internal (seperti daftar transaksi, path file `.xlsx`, atau komponen widget GUI) dikelola melalui atribut instance (`self._daftar_transaksi`, `self.__saldo`) dan hanya diakses atau diubah melalui *method* resmi seperti `tambah_transaksi()`, `hapus_transaksi()`, dan `hitung_saldo()`.
* Validasi input (misalnya memastikan nominal berupa angka positif sebelum dicatat) dilakukan di dalam *method* class terkait.

### 3. Inheritance (Pewarisan)
* Class antarmuka utama (`AplikasiKeuanganGUI`) mewarisi class dasar dari library GUI Python (seperti `tk.Tk` atau `ctk.CTk`), sehingga langsung memiliki properti jendela utama, *event loop*, dan pengaturan tata letak bawaan tanpa harus membangunnya dari nol.

### 4. Abstraction (Abstraksi)
* Pengguna atau bagian kode GUI tidak perlu mengetahui detail teknis bagaimana sel-sel Excel diformat, diwarnai, atau disimpan oleh library (`openpyxl`/`pandas`). Proses kompleks tersebut disembunyikan di balik satu pemanggilan *method* sederhana seperti `export_ke_excel()`.

### 5. Polymorphism (Polimorfisme)
* Penerapan *method* dengan antarmuka seragam untuk menangani perilaku yang berbeda—misalnya perhitungan dampak terhadap saldo akhir berdasarkan jenis transaksi (`Pemasukan` menambah saldo, sedangkan `Pengeluaran` mengurangi saldo) atau representasi string objek melalui *magic method* `__str__()`.

