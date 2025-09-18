import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from refermatch_logic import proses_referensi
import time
import os
import sys


# Splash Screen 
def show_splash_screen():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.configure(bg="black")

    # Splash size

    width, height = 400, 200
    screen_width = splash.winfo_screenmmwidth()
    screen_height = splash.winfo_screenheight()
    x = int((screen_width - width) / 2)
    y = int((screen_height - height) / 2)
    splash.geometry(f"{width}x{height}+{x}+{y}")

    # Splash Content

    tk.Label(
        splash,
        text="ReferMatch by Fathir Afatar",
        font=("Segoe UI", 16, "bold"),
        bg="black",
        fg="white"
    ).pack(pady=(40,10))

    tk.Label(
        splash,
        text="(c) 2025 * MIT License",
        font=("Segoe UI", 10),
        bg="black",
        fg="lightgray"
    ).pack()

    splash.update_idletasks()
    splash.after(2000, splash.destroy)
    splash.mainloop()

show_splash_screen()

# GUI Initialization

root = tk.Tk()
root.title("ReferMatch by Fathir Afatar")
root.geometry("700x400")



# Global Variable

file_transaksi = ""
file_referensi = ""
selected_kolom_transaksi = tk.StringVar()
selected_kolom_nomor = tk.StringVar()
selected_kolom_nama = tk.StringVar()
selected_kolom_tanggal = tk.StringVar()

# File Select Function

def pilih_file_transaksi():
    global file_transaksi
    file_transaksi = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    lbl_file1.config(text=file_referensi.split("/")[-1])
    update_dropdown(file_transaksi, cb_transaksi)
    

def pilih_file_referensi():
    global file_referensi
    file_referensi = filedialog.askopenfilename(filetypes=[("Excel files", "*xlsx")])
    lbl_file2.config(text=file_referensi.split("/")[-1])
    update_dropdown(file_referensi, cb_nomor)
    update_dropdown(file_referensi, cb_nama)
    update_dropdown(file_referensi, cb_tanggal)

# Update Dropdown Function Based on File

def update_dropdown(filepath, combobox):
    try:
        df = pd.read_excel(filepath, nrows=1)
        columns = df.columns.tolist()
        combobox['values'] = columns
        if columns:
            combobox.set(columns[0])
    except:
        messagebox.showerror("Error", f"Tidak bisa membaca file:\n{filepath}")

# Function Proses

def jalankan_proses():
    try:
        start = time.time()

        hasil = proses_referensi(
            file_transaksi,
            file_referensi,
            selected_kolom_transaksi.get().strip(),
            selected_kolom_nomor.get().strip(),
            selected_kolom_nama.get().strip(),
            selected_kolom_tanggal.get().strip()
        )
        proses_selesai = time.time()
        durasi =round(proses_selesai - start, 2)

        simpan_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if simpan_path:
            hasil.to_excel(simpan_path, index=False)

            info = f"""Proses selesai dan berhasil disimpan!
    Lokasi : {simpan_path}
    
    Statistik:
    -Jumlah transaksi: {len(hasil)} baris
    -Waktu Pemrosesan: {durasi} detik"""

            messagebox.showinfo("Selesai", info)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Layout GUI

tk.Label(root, text="ReferMatch by Fathir Afatar", font=("Segoe UI", 14, "bold")).pack(pady=10)

# Choose File Transaction

tk.Button(root, text="Pilih File Transaksi", command=pilih_file_transaksi).pack()
lbl_file1 = tk.Label(root, text="Belum dipilih")

tk.Label(root, text="Kolom untuk mencocokkan (mis. REMARK)").pack()
cb_transaksi = ttk.Combobox(root, textvariable=selected_kolom_transaksi)
cb_transaksi.pack()

# Label  Free 100%
label_gratis = tk.Label(
    root,
    text="Aplikasi ini 100% Gratis & Terbuka Untuk Semua Orang",
    font=("Segoe UI", 9),
    fg="gray",
    bg="white"
)
label_gratis.pack(side="bottom", padx=(0, 2))



# Label License
label_license = tk.Label(
    root,
    text="MIT License (c) 2025 Fathir Afatar",
    font=("Segoe UI", 8),
    fg="blue",
    cursor="hand2",
    bg="white"
)
label_license.pack(side="bottom", fill="x", pady=(0, 5))
label_license.bind("<Button-1>", lambda e: tampilkan_license_popup())


# POP Up License
def tampilkan_license_popup():
    license_path = get_license_path()
    try:
        with open(license_path, "r", encoding="utf-8") as file:
            license_text = file.read()
        messagebox.showinfo("MIT License", license_text)
    except FileNotFoundError:
        messagebox.showerror("Lisensi Tidak Ditemukan", f"File LICENSE tidak ditemukan di:\n{license_path}")


def get_license_path(filename="LICENSE"):
    if getattr(sys, 'frozen', False):

        base_path = os.path.dirname(sys.executable)
    else:

        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

# Choose File Reference

tk.Button(root, text="Pilih File Referensi", command=pilih_file_referensi).pack(pady=5)
lbl_file2 = tk.Label(root, text="Belum dipilih")
lbl_file2.pack()

tk.Label(root, text="Kolom Nomor Referensi"). pack()
cb_nomor = ttk.Combobox(root, textvariable=selected_kolom_nomor)
cb_nomor.pack()

tk.Label(root, text="Kolom Tanggal").pack()
cb_tanggal = ttk.Combobox(root, textvariable=selected_kolom_tanggal)
cb_tanggal.pack()

tk.Label(root, text="Kolom Nama / Penerima").pack()
cb_nama = ttk.Combobox(root, textvariable=selected_kolom_nama)
cb_nama.pack()

tk.Button(root, text="Proses & Simpan", command=jalankan_proses).pack(pady=15)

# Start GUI Loop
root.mainloop()
