"""
app.py
Aplikasi Streamlit: unggah 1 atau 2 file (xls/xlsx/csv), pilih baris header
dan pemetaan kolom NIP/Nama secara manual, cocokkan, tampilkan hasil, export Excel.

Jalankan dengan:
    streamlit run app.py
"""

import streamlit as st

import matcher
from excel_exporter import export_to_excel

st.set_page_config(page_title="Deteksi Anomali Data", page_icon="🔍", layout="wide")

st.title("🔍 Deteksi Anomali Data (NIP & Nama)")
st.caption("Unggah file .xls / .xlsx / .csv, atur kolom yang dicocokkan, lalu jalankan pencocokan.")


def configure_file(uploaded_file, label: str, key_prefix: str):
    """Tampilkan preview mentah, biarkan user pilih baris header & kolom NIP/Nama."""
    st.markdown(f"**{label} — `{uploaded_file.name}`**")

    raw_preview = matcher.read_file_raw(uploaded_file, header_row=None)
    st.caption("Pratinjau baris mentah (nomor baris mulai dari 1) — cek baris ke berapa judul kolom berada:")
    preview = raw_preview.head(8).copy()
    preview.index = [f"Baris {i+1}" for i in range(len(preview))]
    st.dataframe(preview, use_container_width=True)

    header_row = st.number_input(
        f"Baris ke berapa yang berisi judul kolom?",
        min_value=1, max_value=30, value=1, step=1,
        key=f"{key_prefix}_header_row",
    )

    df = matcher.read_file_raw(uploaded_file, header_row=header_row - 1)
    cols = list(df.columns)

    col1, col2 = st.columns(2)
    with col1:
        nip_col = st.selectbox(f"Kolom yang berisi NIP", cols, key=f"{key_prefix}_nip_col")
    with col2:
        default_nama_idx = 1 if len(cols) > 1 else 0
        nama_col = st.selectbox(f"Kolom yang berisi Nama", cols, index=default_nama_idx, key=f"{key_prefix}_nama_col")

    if nip_col == nama_col:
        st.error("Kolom NIP dan Nama tidak boleh sama.")
        return None

    final_df = matcher.finalize_dataframe(df, nip_col, nama_col)
    st.caption(f"✅ {len(final_df)} baris data terdeteksi setelah baris kosong dibuang.")
    return final_df


# ---------- UPLOAD ----------
compare_mode = st.toggle("Bandingkan 2 File", value=False)

file_a = None
file_b = None

if compare_mode:
    col1, col2 = st.columns(2)
    with col1:
        file_a = st.file_uploader("File A", type=["xls", "xlsx", "csv"], key="file_a")
    with col2:
        file_b = st.file_uploader("File B (pembanding)", type=["xls", "xlsx", "csv"], key="file_b")
else:
    file_a = st.file_uploader("File", type=["xls", "xlsx", "csv"], key="file_a_single")

df_a_final = None
df_b_final = None

if file_a:
    with st.expander("⚙️ Atur Kolom — File A", expanded=True):
        df_a_final = configure_file(file_a, "File A", "a")

if compare_mode and file_b:
    with st.expander("⚙️ Atur Kolom — File B", expanded=True):
        df_b_final = configure_file(file_b, "File B", "b")

run = st.button("Jalankan Pencocokan", type="primary", disabled=df_a_final is None)

# ---------- PROSES ----------
if run and df_a_final is not None:
    if compare_mode:
        if df_b_final is None:
            st.error("Mohon lengkapi pengaturan kolom untuk File B.")
            st.stop()
        result_df = matcher.analyze_two_files(df_a_final, df_b_final)
    else:
        result_df = matcher.analyze_single_file(df_a_final)

    st.session_state["result_df"] = result_df

# ---------- HASIL ----------
if "result_df" in st.session_state:
    result_df = st.session_state["result_df"]

    total = len(result_df)
    cocok = int((result_df["Status"] == "Cocok").sum())
    ringan = int((result_df["Status"] == "Anomali Ringan").sum())
    berat = int((result_df["Status"] == "Anomali Berat").sum())

    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Data", total)
    c2.metric("Data Cocok", cocok, f"{cocok/total*100:.1f}%" if total else "0%")
    c3.metric("Anomali Ringan", ringan)
    c4.metric("Anomali Berat", berat)

    st.divider()

    filter_choice = st.radio(
        "Tampilkan",
        ["Semua", "Cocok", "Anomali Ringan", "Anomali Berat"],
        horizontal=True,
    )
    search = st.text_input("Cari NIP / Nama", placeholder="Ketik NIP atau nama...")

    view_df = result_df.copy()
    if filter_choice != "Semua":
        view_df = view_df[view_df["Status"] == filter_choice]
    if search:
        mask = view_df.astype(str).apply(
            lambda col: col.str.contains(search, case=False, na=False)
        ).any(axis=1)
        view_df = view_df[mask]

    st.dataframe(view_df, use_container_width=True, hide_index=True)

    excel_bytes = export_to_excel(result_df)
    st.download_button(
        "⬇ Export ke Excel",
        data=excel_bytes,
        file_name="hasil_deteksi_anomali.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("Unggah file, atur kolom yang dicocokkan, lalu klik 'Jalankan Pencocokan'.")