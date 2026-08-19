"""
enrichment.py
Logika pengisian otomatis kolom yang kosong di File A (mis. Unit Kerja,
Jabatan, Pangkat), menggunakan data referensi dari File B, dicocokkan lewat NIP.
"""

import pandas as pd


def _clean_nip_series(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lstrip("'")
    s = s.str.replace(r"\.0$", "", regex=True)
    return s


def build_lookup(df_source: pd.DataFrame, nip_col: str, value_cols: list) -> dict:
    """Bangun kamus referensi: NIP -> {kolom: nilai}, diambil dari File B."""
    df = df_source.copy()
    df[nip_col] = _clean_nip_series(df[nip_col])
    df = df.drop_duplicates(subset=nip_col, keep="first")
    return df.set_index(nip_col)[value_cols].to_dict("index")


def fill_missing_columns(
    df_target: pd.DataFrame,
    nip_col: str,
    value_cols: list,
    lookup: dict,
    overwrite: bool = False,
) -> pd.DataFrame:
    """
    Isi kolom `value_cols` di df_target menggunakan `lookup` (hasil build_lookup),
    dicocokkan berdasarkan NIP.
    overwrite=False -> hanya sel kosong yang diisi.
    overwrite=True  -> semua sel ditimpa dengan data dari File B.
    """
    df = df_target.copy()
    df[nip_col] = _clean_nip_series(df[nip_col])

    for col in value_cols:
        if col not in df.columns:
            df[col] = ""

    status_list = []
    for idx, row in df.iterrows():
        nip = row[nip_col]
        ref = lookup.get(nip)
        if ref is None:
            status_list.append("NIP tidak ditemukan di File B")
            continue

        terisi = []
        for col in value_cols:
            current_val = str(row.get(col, "")).strip()
            kosong = current_val == "" or current_val.lower() in ("nan", "none")
            if kosong or overwrite:
                new_val = ref.get(col, "")
                if str(df.at[idx, col]) != str(new_val):
                    df.at[idx, col] = new_val
                    terisi.append(col)

        status_list.append(f"Terisi: {', '.join(terisi)}" if terisi else "Sudah lengkap / tidak diubah")

    df["Status Pengisian"] = status_list
    return df


def summarize_fill(df_result: pd.DataFrame) -> dict:
    """Ringkasan hasil pengisian: total, berhasil diisi, sudah lengkap, tidak ditemukan."""
    total = len(df_result)
    tidak_ketemu = int((df_result["Status Pengisian"] == "NIP tidak ditemukan di File B").sum())
    sudah_lengkap = int((df_result["Status Pengisian"] == "Sudah lengkap / tidak diubah").sum())
    terisi = total - tidak_ketemu - sudah_lengkap
    return {
        "total": total,
        "terisi": terisi,
        "sudah_lengkap": sudah_lengkap,
        "tidak_ketemu": tidak_ketemu,
    }
