"""
matcher.py
Logika inti: membaca file, normalisasi data, pencocokan, dan deteksi anomali
berdasarkan kolom NIP & Nama. Modul ini murni Python (tanpa Streamlit),
sehingga bisa di-unit test atau dipakai ulang di tempat lain (termasuk
fitur Lengkapi Data Otomatis).
"""

import re
import pandas as pd

REQUIRED_COLUMNS = ["NIP", "Nama"]


def read_file_raw(uploaded_file, header_row=0) -> pd.DataFrame:
    """
    Baca file mentah tanpa mengasumsikan nama kolom apapun.
    header_row=None -> baca tanpa header (kolom jadi angka 0,1,2,... untuk preview).
    header_row=N     -> baris ke-N (index 0-based) dipakai sebagai judul kolom.
    """
    uploaded_file.seek(0)
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, dtype=str, header=header_row)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, dtype=str, header=header_row)
    else:
        raise ValueError(f"Format file tidak didukung: {uploaded_file.name}")

    if header_row is not None:
        df.columns = [str(c).strip() for c in df.columns]
    return df


def finalize_dataframe(df: pd.DataFrame, nip_col: str, nama_col: str) -> pd.DataFrame:
    """
    Terapkan pemetaan kolom pilihan pengguna (kolom apa saja -> NIP, Nama standar),
    lalu buang baris kosong/spacer.
    """
    df = df.rename(columns={nip_col: "NIP", nama_col: "Nama"})
    df = drop_empty_rows(df)
    return df


def drop_empty_rows(df: pd.DataFrame, cols: list = None) -> pd.DataFrame:
    """
    Buang baris kosong/spacer. Secara default memeriksa kolom NIP & Nama;
    fitur lain (mis. Lengkapi Data) bisa memberi kolom acuan sendiri lewat
    parameter `cols` (contoh: cols=["NIP"] saja).
    File Excel sering punya baris kosong sisipan (mis. akibat merge cell
    di kolom lain) yang membuat jumlah baris fisik jauh lebih banyak
    dari jumlah data sebenarnya.
    """
    if cols is None:
        cols = [c for c in REQUIRED_COLUMNS if c in df.columns]
    else:
        cols = [c for c in cols if c in df.columns]
    if not cols:
        return df

    def _is_blank(col: pd.Series) -> pd.Series:
        text = col.astype(str).str.strip().str.lower()
        return col.isna() | text.isin(["", "nan", "none"])

    mask_blank = pd.concat([_is_blank(df[c]) for c in cols], axis=1)
    all_blank = mask_blank.all(axis=1)  # baris dianggap spacer jika SEMUA kolom acuan kosong

    return df[~all_blank].reset_index(drop=True)


def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Bersihkan & seragamkan kolom NIP dan Nama."""
    df = df.copy()
    df["NIP"] = df["NIP"].astype(str).str.strip()
    df["NIP"] = df["NIP"].str.lstrip("'")                      # buang tanda kutip depan (format teks Excel)
    df["NIP"] = df["NIP"].str.replace(r"\.0$", "", regex=True)  # buang .0 dari NIP numerik
    df["Nama"] = df["Nama"].astype(str).str.strip()
    df["Nama"] = df["Nama"].str.lstrip("'")
    df["Nama_key"] = df["Nama"].str.upper().str.replace(r"\s+", " ", regex=True)
    return df


def is_valid_nip(nip: str) -> bool:
    """Validasi format dasar: NIP harus berupa digit dengan panjang wajar (6-20 digit)."""
    if not nip or nip.lower() in ("nan", "none", ""):
        return False
    return bool(re.fullmatch(r"\d{6,20}", nip))


def find_duplicate_nip(df: pd.DataFrame) -> set:
    counts = df["NIP"].value_counts()
    return set(counts[counts > 1].index)


def analyze_single_file(df: pd.DataFrame) -> pd.DataFrame:
    """Mode satu file: deteksi anomali internal (tanpa pembanding)."""
    df = normalize(df)
    dup_nip = find_duplicate_nip(df)

    hasil = []
    for _, row in df.iterrows():
        nip, nama = row["NIP"], row["Nama"]
        masalah = []

        if not is_valid_nip(nip):
            masalah.append("Format NIP tidak valid")
        if nip in dup_nip:
            masalah.append("NIP duplikat")
        if not nama or nama.lower() in ("nan", "none", ""):
            masalah.append("Nama kosong")

        if not masalah:
            status = "Cocok"
        elif any(m in ("NIP duplikat", "Format NIP tidak valid") for m in masalah):
            status = "Anomali Berat"
        else:
            status = "Anomali Ringan"

        hasil.append({
            "NIP": nip,
            "Nama": nama,
            "Status": status,
            "Keterangan": "; ".join(masalah) if masalah else "Data valid",
        })

    return pd.DataFrame(hasil)


def analyze_two_files(df_a: pd.DataFrame, df_b: pd.DataFrame) -> pd.DataFrame:
    """Mode dua file: cocokkan File A vs File B berdasarkan NIP, bandingkan Nama."""
    df_a = normalize(df_a)
    df_b = normalize(df_b)

    dup_a = find_duplicate_nip(df_a)
    dup_b = find_duplicate_nip(df_b)

    map_a = df_a.drop_duplicates(subset="NIP", keep="first").set_index("NIP").to_dict("index")
    map_b = df_b.drop_duplicates(subset="NIP", keep="first").set_index("NIP").to_dict("index")

    all_nip = list(dict.fromkeys(list(df_a["NIP"]) + list(df_b["NIP"])))

    hasil = []
    for nip in all_nip:
        in_a = nip in map_a
        in_b = nip in map_b
        nama_a = map_a[nip]["Nama"] if in_a else ""
        nama_b = map_b[nip]["Nama"] if in_b else ""
        key_a = map_a[nip]["Nama_key"] if in_a else ""
        key_b = map_b[nip]["Nama_key"] if in_b else ""

        masalah = []
        if not is_valid_nip(nip):
            masalah.append("Format NIP tidak valid")
        if nip in dup_a:
            masalah.append("NIP duplikat di File A")
        if nip in dup_b:
            masalah.append("NIP duplikat di File B")

        if in_a and not in_b:
            masalah.append("Hilang di File B")
        elif in_b and not in_a:
            masalah.append("Hilang di File A")
        elif in_a and in_b and key_a != key_b:
            masalah.append(f'Nama beda: "{nama_a}" vs "{nama_b}"')

        if not masalah:
            status = "Cocok"
        elif any(("Hilang" in m or "duplikat" in m or "tidak valid" in m) for m in masalah):
            status = "Anomali Berat"
        else:
            status = "Anomali Ringan"

        hasil.append({
            "NIP": nip,
            "Nama (File A)": nama_a if in_a else "-",
            "Nama (File B)": nama_b if in_b else "-",
            "Di File A": "✓" if in_a else "—",
            "Di File B": "✓" if in_b else "—",
            "Status": status,
            "Keterangan": "; ".join(masalah) if masalah else "Data identik di kedua file",
        })

    return pd.DataFrame(hasil)
