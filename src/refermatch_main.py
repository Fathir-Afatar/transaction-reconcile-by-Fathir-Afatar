# Standard Library
import time
import os
import sys

# Third-party
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

# Local Application
from refermatch_logic import proses_referensi

TEXTS = {
    "id": {
        "title": "ReferMatch by Fathir Afatar",
        "btn_select_trans": "Pilih File Transaksi",
        "lbl_match_col": "Kolom untuk mencocokkan (mis. REMARK)",
        "btn_select_ref": "Pilih File Referensi",
        "lbl_ref_num": "Kolom Nomor Referensi",
        "lbl_date_col": "Kolom Tanggal",
        "lbl_name_col": "Kolom Nama / Penerima",
        "btn_process": "Proses & Simpan",
        "not_selected": "Belum dipilih",
        "info_free": "Aplikasi ini 100% Gratis & Terbuka Untuk Semua Orang",
        "err_read_file": "Tidak bisa membaca file:\n{}",
        "err_license_title": "Lisensi Tidak Ditemukan",
        "err_license_msg": "File LICENSE tidak ditemukan di:\n{}",
        "save_dialog_title": "Simpan Hasil Excel",
        "done_title": "Selesai",
        "done_info": (
            "Proses selesai dan berhasil disimpan!\nLokasi : {}\n\nStatistik:\n- "
            "Jumlah transaksi: {} baris\n- Jumlah Referensi: "
            "{} baris\n - Waktu Proses: {} detik"
        ),
    },
    "en": {
        "title": "ReferMatch by Fathir Afatar",
        "btn_select_trans": "Select Transaction File",
        "lbl_match_col": "Matching Column (e.g., REMARK)",
        "btn_select_ref": "Select Reference File",
        "lbl_ref_num": "Reference Number Column",
        "lbl_date_col": "Date Column",
        "lbl_name_col": "Name / Recipient Column",
        "btn_process": "Process & Save",
        "not_selected": "Not selected",
        "info_free": "This application is 100% Free & Open Source",
        "err_read_file": "Unable to read file:\n{}",
        "err_license_title": "License Not Found",
        "err_license_msg": "LICENSE file was not found at:\n{}",
        "save_dialog_title": "Save Excel Result",
        "done_title": "Completed",
        "done_info": (
            "Process completed and saved successfully!\nLocation : "
            "{}\n\nStatistics:\n- Total transactions: {} rows\n- Total reference: "
            "{} rows\n- Processing Time: {} seconds"
        ),
    },
}

CURRENT_LANG = "id"

# Splash Screen 
def show_splash_screen():
    """
    Display a centered splash screen for the ReferMatch application.

    Creates a dark-themed, borderless windows displaying the app titile and
    license information, forces an immediate UI refresh, and automatically
    destroys itself after 2000 milliseconds.
    """
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.configure(bg="black")

    # Splash size
    width, height = 400, 200
    screen_width = splash.winfo_screenwidth() # FIX Typo
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
        text="v2.0 • Multi-Language Edition",
        font=("Segoe UI", 10),
        bg="#1e1e1e",
        fg="#007acc"
    ).pack()
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
root.geometry("700x500")
root.resizable(False, False)

# Global Variable
file_transaksi = ""
file_referensi = ""

selected_kolom_transaksi = tk.StringVar()
selected_kolom_nomor = tk.StringVar()
selected_kolom_nama = tk.StringVar()
selected_kolom_tanggal = tk.StringVar()

# Handler GUI
def toggle_language():
    global CURRENT_LANG
    CURRENT_LANG = "en" if CURRENT_LANG == "id" else "id"
    update_ui_text()

def update_ui_text():
    lang = TEXTS[CURRENT_LANG]
    root.title(lang["title"])
    lbl_title.config(text=lang["title"])

    btn_trans.config(text=lang["btn_select_trans"])
    lbl_match_col.config(text=lang["lbl_match_col"])

    btn_ref.config(text=lang["btn_select_ref"])
    lbl_ref_num.config(text=lang["lbl_ref_num"])
    lbl_date_col.config(text=lang["lbl_date_col"])
    lbl_name_col.config(text=lang["lbl_name_col"])

    btn_process.config(text=lang["btn_process"])
    label_gratis.config(text=lang["info_free"])

    if file_transaksi:
        lbl_file1.config(text=os.path.basename(file_transaksi))
    else:
        lbl_file1.config(text=lang["not_selected"])

    if file_referensi:
        lbl_file2.config(text=os.path.basename(file_referensi))
    else:
        lbl_file2.config(text=lang["not_selected"])

# File Select Function
def pilih_file_transaksi():
    """
    Open a file dialog to select the transaction Excel file.

    Updates the global `file_transaksi` variable, extracts the file name
    to update the UI label, and refreshes the corresponding transaction
    column dropdown menu.
    """
    global file_transaksi
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        file_transaksi = path
        lbl_file1.config(text=os.path.basename(file_transaksi))
        update_dropdown(file_transaksi, cb_transaksi)

def pilih_file_referensi():
    """
    Open a file dialog to select the reference Excel file.

    Updates the global `file_referensi` variable, extracts the file name
    to update the UI label, and populates the reference-related dropdown
    menus (ID number, name, and date columns)
    """
    global file_referensi
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        file_referensi = path
        lbl_file2.config(text=os.path.basename(file_referensi))
        update_dropdown(file_referensi, cb_nomor)
        update_dropdown(file_referensi, cb_nama)
        update_dropdown(file_referensi, cb_tanggal)

# Update Dropdown Function Based on File
def update_dropdown(filepath, combobox):
    """
    Updates the values of a Tkinter Combobox with the column names of an Excel file.

    This function reads only the firs row of the specified Excel file to extract
    the column headers. It the populates the combobox and sets the default
    selection to the first column found. If an error occurs during reading,
    an error dialog is displayed.

    Args:
        filepath(str): The absolute or relative path to the Excel file.
        combobox (tkinter.ttk.Combobox): The Tkinter combobox widget to update.
    
    Returns:
        None
    """
    try:
        df = pd.read_excel(filepath, nrows=1)
        columns = df.columns.tolist()
        combobox['values'] = columns
        if columns:
            combobox.set(columns[0])
    except:
        messagebox.showerror(
            "Error", TEXTS[CURRENT_LANG]["err_read_file"].format(filepath)
        )

# Function Proses
def jalankan_proses():
    """
    Executes the reference matching process and saves the output to an Excel file.

    This function measures the execution time of `proses_referensi`, triggers
    a file dialog for the user to select a save location, and exports the
    resulting DataFrame. Upon success, it displays a popup containing
    processing statistics (row count and duration). Any exception caught
    during execution are displayed via an error messagebox.

    Returns:
        None

    Raises:
        Exception: Captures and displays any error that occurs during data
            processing or file saving.
    """
    lang = TEXTS[CURRENT_LANG]
    try:
        start = time.time()
        hasil, total_ref = proses_referensi(
            file_transaksi,
            file_referensi,
            selected_kolom_transaksi.get().strip(),
            selected_kolom_nomor.get().strip(),
            selected_kolom_nama.get().strip(),
            selected_kolom_tanggal.get().strip()
        )
        durasi = round(time.time() - start, 2)

        simpan_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if simpan_path:
            hasil.to_excel(simpan_path, index=False)
            info = lang["done_info"].format(simpan_path, len(hasil), total_ref, durasi)
            messagebox.showinfo(lang["done_title"], info)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# POP Up License
def tampilkan_license_popup():
    """
    Displays the content of the license file in a GUI info message box.

    This function attemps to find and read the license text file. If successful,
    it opens a Tkinter info popun showing the full license terms. If the file is
    missing, it catches the exception and displays an error messagebox instead.

    Returns:
        None
    """
    lang = TEXTS[CURRENT_LANG]
    license_path = get_license_path()
    try:
        with open(license_path, "r", encoding="utf-8") as file:
            license_text = file.read()
        messagebox.showinfo("MIT License", license_text)
    except FileNotFoundError:
        messagebox.showerror(
            lang["err_license_title"],
            lang["err_license_msg"].format(license_path),
        )

def get_license_path(filename="LICENSE"):
    """
    Determines the absolute path to the license file based on execution environment.

    This function checks wheter the script is running as a compiled standalone
    executable (e.g., packaged with PyInstaller) or as a standard Python script.
    It dynamically resolves the directory path accordingly to locate the file.

    Args:
        filename (str, optional): the name of the license file. Default to "LICENSE".
    
    Returns:
        str: The full absolute path to the specified license file. 
    """
    if getattr(sys, 'frozen', False):

        base_path = os.path.dirname(sys.executable)
    else:

        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

lang_frame = tk.Frame(root)
lang_frame.pack(anchor="ne", padx=10, pady=5)
btn_lang = tk.Button(lang_frame, text="🌐 EN / ID", command=toggle_language,
                     font=("Segoe UI", 8, "bold"),)
btn_lang.pack()

# Header (Judul + Versi)
lbl_title = tk.Label(root, font=("Segoe UI", 15, "bold"))
lbl_title.pack(pady=(0, 2))

lbl_version = tk.Label(root, text="v2.0", font=("Segoe UI", 9, "bold"), fg="#666666")
lbl_version.pack(pady=(0, 10))

# Bagian 1: File Transaksi
btn_trans = tk.Button(root, command=pilih_file_transaksi, width=25)
btn_trans.pack()
lbl_file1 = tk.Label(root, fg="gray")
lbl_file1.pack(pady=(2, 5))

lbl_match_col = tk.Label(root)
lbl_match_col.pack()
cb_transaksi = ttk.Combobox(root, textvariable=selected_kolom_transaksi, 
                            state="readonly", width=30)
cb_transaksi.pack(pady=(2, 15))

# Bagian 2: File Referensi
btn_ref = tk.Button(root, text=TEXTS[CURRENT_LANG]["btn_select_ref"], command=pilih_file_referensi, width=25)
btn_ref.pack(pady=(5, 0))

lbl_file2 = tk.Label(root, fg="gray")
lbl_file2.pack(pady=(2, 5))

lbl_ref_num = tk.Label(root)
lbl_ref_num.pack()
cb_nomor = ttk.Combobox(root, textvariable=selected_kolom_nomor, state="readonly", width=30)
cb_nomor.pack(pady=(2, 5))

lbl_date_col = tk.Label(root)
lbl_date_col.pack()
cb_tanggal = ttk.Combobox(root, textvariable=selected_kolom_tanggal, state="readonly", width=30)
cb_tanggal.pack(pady=(2, 5))

lbl_name_col = tk.Label(root)
lbl_name_col.pack()
cb_nama = ttk.Combobox(root, textvariable=selected_kolom_nama, state="readonly", width=30)
cb_nama.pack(pady=(2, 15))

# Tombol Eksekusi
btn_process = tk.Button(root, command=jalankan_proses, bg="#28a745", fg="white",
                        font=("Segoe UI", 10, "bold"), width=25,)
btn_process.pack(pady=5)

# Footer (Keterangan & Lisensi)
label_gratis = tk.Label(root, font=("Segoe UI", 8), fg="gray")
label_gratis.pack(side="bottom", pady=(0, 2))

label_license = tk.Label(root, text="MIT License (c) 2025 Fathir Afatar", 
                         font=("Segoe UI", 8, "underline"), fg="blue", cursor="hand2",)
label_license.pack(side="bottom")
label_license.bind("<Button-1>", lambda e: tampilkan_license_popup())

# Render Teks Awal & Jalankan Loop
update_ui_text()
root.mainloop()
