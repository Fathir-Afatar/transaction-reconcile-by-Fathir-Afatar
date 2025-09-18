import pandas as pd
import re

def extract_ids(text):
    if not isinstance(text, str):
        return[]
    return re.findall(r"\d{10,20}", text)

def proses_referensi(
    file_transaksi,
    file_referensi,
    kolom_transaksi,
    kolom_nomor,
    kolom_nama,
    kolom_tanggal
):
    
    # Load file
    df_transaksi = pd.read_excel(file_transaksi)
    df_referensi = pd.read_excel(file_referensi)

    # Normalize Column name

    #df_transaksi.columns[kolom_nomor] = df_referensi[kolom_nomor].astype(str).str.strip()
    #df_referensi.columns[kolom_tanggal] = (df_referensi.columns)

    # Normalize Column Reference

    df_referensi[kolom_nomor] = df_referensi[kolom_nomor].astype(str).str.strip()
    df_referensi[kolom_tanggal] = pd.to_datetime(df_referensi[kolom_tanggal], errors ="coerce")

    # Make a new empty column

    df_transaksi["Hasil ReferMatch"] = ""

    # line by line process

    for i, row in df_transaksi.iterrows():
        remark = row.get(kolom_transaksi, "")
        kandidat = extract_ids(remark)
        cocok = None

        for kode in kandidat:
            ref_match = df_referensi[df_referensi[kolom_nomor] == kode]
            if not ref_match.empty:
                ref_row = ref_match.iloc[0]
                tanggal = ref_row[kolom_tanggal].day if pd.notna(ref_row[kolom_tanggal]) else "?"
                nama = ref_row[kolom_nama]
                cocok = f"{tanggal}. {nama}"
                break
        df_transaksi.at[i, "Hasil ReferMatch"] = cocok if cocok else ""

    return df_transaksi

