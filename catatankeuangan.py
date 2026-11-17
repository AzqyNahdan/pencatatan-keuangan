from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FILE_EXCEL = "Laporan_Keuangan_Harian.xlsx"


class AplikasiKeuanganExcel:

    def __init__(self, root):
        self.root = root
        self.root.title("Pencatat Keuangan Harian")
        self.root.geometry("680x550")
        self.root.resizable(False, False)

        # Inisialisasi file Excel jika belum ada
        self.inisialisasi_excel()

        
        title_label = tk.Label(
            root, text="PENCATAT KEUANGAN HARIAN", font=("Arial", 16, "bold")
        )
        title_label.pack(pady=10)

        self.label_saldo = tk.Label(
            root, text="Total Saldo: Rp 0", font=("Arial", 12, "bold"), fg="green"
        )
        self.label_saldo.pack(pady=5)
        frame_input = tk.LabelFrame(
            root,
            text=" Form Transaksi Baru ",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=15,
        )
        frame_input.pack(fill="x", padx=20, pady=10)

        # Jenis Transaksi
        tk.Label(frame_input, text="Jenis:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        self.jenis_var = tk.StringVar(value="Pemasukan")
        tk.Radiobutton(
            frame_input,
            text="Pemasukan",
            variable=self.jenis_var,
            value="Pemasukan",
        ).grid(row=0, column=1, sticky="w")
        tk.Radiobutton(
            frame_input,
            text="Pengeluaran",
            variable=self.jenis_var,
            value="Pengeluaran",
        ).grid(row=0, column=2, sticky="w")

        # Jumlah
        tk.Label(frame_input, text="Jumlah (Rp):").grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.entry_jumlah = tk.Entry(frame_input, width=35)
        self.entry_jumlah.grid(
            row=1, column=1, columnspan=2, sticky="w", pady=5
        )

        # Kategori
        tk.Label(frame_input, text="Kategori:").grid(
            row=2, column=0, sticky="w", pady=5
        )
        self.combo_kategori = ttk.Combobox(
            frame_input,
            values=[
                "Makanan & Minuman",
                "Transportasi",
                "Gaji / Uang Saku",
                "Belanja",
                "Tagihan",
                "Lainnya",
            ],
            width=32,
        )
        self.combo_kategori.grid(
            row=2, column=1, columnspan=2, sticky="w", pady=5
        )
        self.combo_kategori.set("Makanan & Minuman")

        # Keterangan
        tk.Label(frame_input, text="Keterangan:").grid(
            row=3, column=0, sticky="w", pady=5
        )
        self.entry_keterangan = tk.Entry(frame_input, width=35)
        self.entry_keterangan.grid(
            row=3, column=1, columnspan=2, sticky="w", pady=5
        )

        # Tombol Simpan & Buka Excel
        frame_tombol = tk.Frame(frame_input)
        frame_tombol.grid(row=4, column=1, columnspan=2, pady=10, sticky="w")

        btn_simpan = tk.Button(
            frame_tombol,
            text="Simpan ke Excel",
            bg="#2E7D32",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            command=self.tambah_transaksi,
        )
        btn_simpan.pack(side="left", padx=(0, 10))

        btn_buka = tk.Button(
            frame_tombol,
            text="Buka File Excel",
            bg="#1565C0",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            command=self.buka_file_excel,
        )
        btn_buka.pack(side="left")

        # TABEL RIWAYAT
        frame_tabel = tk.LabelFrame(
            root,
            text=" Preview Data dari Excel ",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10,
        )
        frame_tabel.pack(fill="both", expand=True, padx=20, pady=10)

        columns = ("Tanggal", "Jenis", "Jumlah", "Kategori", "Keterangan")
        self.tree = ttk.Treeview(
            frame_tabel, columns=columns, show="headings", height=6
        )

        lebar_kolom = [110, 95, 120, 130, 150]
        for col, lebar in zip(columns, lebar_kolom):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=lebar)

        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            frame_tabel, orient="vertical", command=self.tree.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Muat data yang sudah ada di Excel ke tampilan GUI
        self.muat_data_dari_excel()

    def inisialisasi_excel(self):
        """Membuat file Excel baru dengan header yang rapi jika file belum ada."""
        if not os.path.exists(FILE_EXCEL):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Catatan Keuangan"

            # Judul Laporan di Baris 1
            ws.merge_cells("A1:E1")
            ws["A1"] = "LAPORAN PENCATATAN KEUANGAN HARIAN"
            ws["A1"].font = Font(name="Calibri", size=14, bold=True)
            ws["A1"].alignment = Alignment(horizontal="center")


            headers = [
                "Tanggal & Waktu",
                "Jenis Transaksi",
                "Jumlah (Rp)",
                "Kategori",
                "Keterangan",
            ]
            ws.append([])  
            ws.append(headers)

            # Styling Header
            header_fill = PatternFill(
                start_color="1F4E78", end_color="1F4E78", fill_type="solid"
            )
            header_font = Font(
                name="Calibri", size=11, bold=True, color="FFFFFF"
            )
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            for col_num in range(1, 6):
                cell = ws.cell(row=3, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                cell.border = thin_border

            lebar_kolom = {
                "A": 20,
                "B": 18,
                "C": 18,
                "D": 22,
                "E": 35,
            }
            for col_letter, width in lebar_kolom.items():
                ws.column_dimensions[col_letter].width = width

            wb.save(FILE_EXCEL)

    def tambah_transaksi(self):
        jenis = self.jenis_var.get()
        jumlah_str = self.entry_jumlah.get()
        kategori = self.combo_kategori.get()
        keterangan = self.entry_keterangan.get()

        try:
            jumlah = float(jumlah_str)
            if jumlah <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Error", "Masukkan nominal angka yang valid (lebih dari 0)!"
            )
            return

        tanggal_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M")
        ket_final = keterangan if keterangan.strip() else "-"

        try:
            wb = openpyxl.load_workbook(FILE_EXCEL)
            ws = wb["Catatan Keuangan"]

            # Tambahkan baris baru ke Excel
            baris_baru = [
                tanggal_sekarang,
                jenis,
                jumlah,
                kategori,
                ket_final,
            ]
            ws.append(baris_baru)

            row_idx = ws.max_row
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            # Terapkan border dan format sel untuk baris yang baru ditambahkan
            for col_num in range(1, 6):
                cell = ws.cell(row=row_idx, column=col_num)
                cell.border = thin_border
                if col_num in [1, 2]:
                    cell.alignment = Alignment(horizontal="center")
                elif col_num == 3:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                    # Warna teks hijau untuk pemasukan, merah untuk pengeluaran
                    cell.font = Font(
                        color="008000" if jenis == "Pemasukan" else "FF0000",
                        bold=True,
                    )

            wb.save(FILE_EXCEL)

            # Refresh tabel GUI dan reset form
            self.muat_data_dari_excel()
            self.entry_jumlah.delete(0, tk.END)
            self.entry_keterangan.delete(0, tk.END)
            messagebox.showinfo(
                "Sukses", f"Data berhasil disimpan ke {FILE_EXCEL}!"
            )

        except PermissionError:
            messagebox.showerror(
                "File Terbuka",
                f"Tutup file '{FILE_EXCEL}' di Microsoft Excel terlebih dahulu sebelum menyimpan data baru!",
            )

    def muat_data_dari_excel(self):
        """Membaca ulang seluruh data dari Excel untuk ditampilkan di GUI dan menghitung saldo."""
        for row in self.tree.get_children():
            self.tree.delete(row)

        if not os.path.exists(FILE_EXCEL):
            return

        wb = openpyxl.load_workbook(FILE_EXCEL)
        ws = wb["Catatan Keuangan"]

        total_saldo = 0
        # Data dimulai dari baris ke-4 (karena baris 1 judul, baris 3 header)
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[0] is not None:
                tgl, jenis, jumlah, kategori, ket = row
                if jenis == "Pemasukan":
                    total_saldo += jumlah
                else:
                    total_saldo -= jumlah

                self.tree.insert(
                    "",
                    "end",
                    values=(tgl, jenis, f"Rp {jumlah:,.0f}", kategori, ket),
                )

        self.label_saldo.config(
            text=f"Total Saldo: Rp {total_saldo:,.0f}",
            fg="green" if total_saldo >= 0 else "red",
        )

    def buka_file_excel(self):
        """Membuka file Excel langsung dari aplikasi."""
        if os.path.exists(FILE_EXCEL):
            os.startfile(FILE_EXCEL)  # Khusus Windows
        else:
            messagebox.showwarning("Info", "File Excel belum dibuat.")


if __name__ == "__main__":
    root = tk.Tk()
    app = AplikasiKeuanganExcel(root)
    root.mainloop()