"""
app.py
Aplikasi Streamlit dengan dua fitur, dipilih lewat sidebar:
1. Cek Anomali Data   - cocokkan & deteksi anomali antar file berdasarkan NIP & Nama
2. Lengkapi Data Otomatis - isi kolom kosong di File A (Unit Kerja, Jabatan, dll)
   menggunakan data referensi dari File B, dicocokkan lewat NIP.

Jalankan dengan:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import altair as alt

import matcher
import enrichment
from excel_exporter import export_to_excel, export_simple_excel

st.set_page_config(page_title="Deteksi Anomali Data", page_icon="🔍", layout="wide")

# ---------- CSS CUSTOM ----------
st.markdown("""
<style>
:root {
  --brand-red: #FF4B4B;
  --brand-gold: #E8B34E;
  --green: #3DDC97;
  --amber: #FBBF24;
}

.header-rule {
  height: 3px;
  border-radius: 3px;
  margin: 0.4rem 0 1.6rem 0;
  background: linear-gradient(90deg, var(--brand-red) 0%, var(--brand-gold) 45%, transparent 75%);
}

.metric-card {
  background: rgba(255,255,255,0.03);
  border: 1px solid rgba(255,255,255,0.09);
  border-radius: 10px;
  padding: 16px 18px;
  height: 100%;
}
.metric-card .top-row { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }
.metric-card .icon {
  width: 30px; height: 30px; border-radius: 8px;
  display:flex; align-items:center; justify-content:center; font-size:14px;
}
.metric-card .lbl { font-size: 11px; text-transform:uppercase; letter-spacing:1px; color: rgba(230,233,239,0.6); }
.metric-card .val { font-size: 26px; font-weight: 800; margin-top: 2px; }
.metric-card .delta { font-size: 11px; margin-top: 6px; font-weight: 600; }

.metric-total .icon  { background: rgba(232,179,78,0.14); color: var(--brand-gold); }
.metric-cocok .icon  { background: rgba(61,220,151,0.12); color: var(--green); }
.metric-ringan .icon { background: rgba(251,191,36,0.12); color: var(--amber); }
.metric-berat .icon  { background: rgba(255,75,75,0.12);  color: var(--brand-red); }
.metric-cocok .delta  { color: var(--green); }
.metric-ringan .delta { color: var(--amber); }
.metric-berat .delta  { color: var(--brand-red); }
.metric-total .delta  { color: rgba(230,233,239,0.6); }

div[role="radiogroup"] { gap: 6px; }
div[role="radiogroup"] label {
  border: 1px solid rgba(255,255,255,0.12) !important;
  border-radius: 8px !important;
  padding: 6px 14px !important;
  background: rgba(255,255,255,0.02);
}

button[kind="primary"] {
  background-color: var(--brand-red) !important;
  border-color: var(--brand-red) !important;
}
</style>
""", unsafe_allow_html=True)


# ---------- HELPER BERSAMA ----------
def preview_and_pick_header(uploaded_file, label: str, key_prefix: str) -> pd.DataFrame:
    """Preview baris mentah file, minta user pilih baris header, kembalikan DataFrame utuh."""
    st.markdown(f"**{label} — `{uploaded_file.name}`**")

    raw_preview = matcher.read_file_raw(uploaded_file, header_row=None)
    st.caption("Pratinjau baris mentah (nomor baris mulai dari 1) — cek baris ke berapa judul kolom berada:")
    preview = raw_preview.head(8).copy()
    preview.index = [f"Baris {i+1}" for i in range(len(preview))]
    st.dataframe(preview, use_container_width=True)

    header_row = st.number_input(
        "Baris ke berapa yang berisi judul kolom?",
        min_value=1, max_value=30, value=1, step=1,
        key=f"{key_prefix}_header_row",
    )

    df = matcher.read_file_raw(uploaded_file, header_row=header_row - 1)
    st.caption(f"{len(df)} baris terbaca (sebelum pembersihan baris kosong).")
    return df


def metric_card(container, css_class: str, label: str, icon: str, value, delta: str):
    with container:
        st.markdown(f"""
        <div class="metric-card {css_class}">
          <div class="top-row"><div class="lbl">{label}</div><div class="icon">{icon}</div></div>
          <div class="val">{value}</div>
          <div class="delta">{delta}</div>
        </div>""", unsafe_allow_html=True)


# =========================================================
# HALAMAN 1: CEK ANOMALI DATA
# =========================================================
def configure_file_anomali(uploaded_file, label: str, key_prefix: str):
    df = preview_and_pick_header(uploaded_file, label, key_prefix)
    cols = list(df.columns)

    col1, col2 = st.columns(2)
    with col1:
        nip_col = st.selectbox("Kolom yang berisi NIP", cols, key=f"{key_prefix}_nip_col")
    with col2:
        default_nama_idx = 1 if len(cols) > 1 else 0
        nama_col = st.selectbox("Kolom yang berisi Nama", cols, index=default_nama_idx, key=f"{key_prefix}_nama_col")

    if nip_col == nama_col:
        st.error("Kolom NIP dan Nama tidak boleh sama.")
        return None

    final_df = matcher.finalize_dataframe(df, nip_col, nama_col)
    st.caption(f"✅ {len(final_df)} baris data terdeteksi setelah baris kosong dibuang.")
    return final_df


def render_cek_anomali():
    st.markdown("<h1>🔍 Deteksi Anomali Data</h1>", unsafe_allow_html=True)
    st.caption("Unggah file .xls / .xlsx / .csv, atur kolom yang dicocokkan, lalu jalankan pencocokan.")
    st.markdown('<div class="header-rule"></div>', unsafe_allow_html=True)

    compare_mode = st.toggle("Bandingkan 2 File", value=False)

    file_a = None
    file_b = None

    if compare_mode:
        col1, col2 = st.columns(2)
        with col1:
            file_a = st.file_uploader("File A", type=["xls", "xlsx", "csv"], key="anomali_file_a")
        with col2:
            file_b = st.file_uploader("File B (pembanding)", type=["xls", "xlsx", "csv"], key="anomali_file_b")
    else:
        file_a = st.file_uploader("File", type=["xls", "xlsx", "csv"], key="anomali_file_a_single")

    df_a_final = None
    df_b_final = None

    if file_a:
        with st.expander("⚙️ Atur Kolom — File A", expanded=True):
            df_a_final = configure_file_anomali(file_a, "File A", "anomali_a")

    if compare_mode and file_b:
        with st.expander("⚙️ Atur Kolom — File B", expanded=True):
            df_b_final = configure_file_anomali(file_b, "File B", "anomali_b")

    run = st.button("Jalankan Pencocokan", type="primary", disabled=df_a_final is None)

    if run and df_a_final is not None:
        if compare_mode:
            if df_b_final is None:
                st.error("Mohon lengkapi pengaturan kolom untuk File B.")
                st.stop()
            result_df = matcher.analyze_two_files(df_a_final, df_b_final)
        else:
            result_df = matcher.analyze_single_file(df_a_final)

        st.session_state["result_df"] = result_df

    if "result_df" in st.session_state:
        result_df = st.session_state["result_df"]

        total = len(result_df)
        cocok = int((result_df["Status"] == "Cocok").sum())
        ringan = int((result_df["Status"] == "Anomali Ringan").sum())
        berat = int((result_df["Status"] == "Anomali Berat").sum())
        pct_cocok = f"{cocok/total*100:.1f}%" if total else "0%"

        st.divider()

        c1, c2, c3, c4 = st.columns(4)
        metric_card(c1, "metric-total", "Total Data", "📊", total, "Gabungan seluruh data")
        metric_card(c2, "metric-cocok", "Data Cocok", "✅", cocok, f"▲ {pct_cocok}")
        metric_card(c3, "metric-ringan", "Anomali Ringan", "⚠️", ringan, "Selisih ejaan nama")
        metric_card(c4, "metric-berat", "Anomali Berat", "🛑", berat, "Hilang / duplikat / format salah")

        st.write("")

        chart_df = pd.DataFrame({
            "Status": ["Cocok", "Anomali Ringan", "Anomali Berat"],
            "Jumlah": [cocok, ringan, berat],
        })
        color_scale = alt.Scale(
            domain=["Cocok", "Anomali Ringan", "Anomali Berat"],
            range=["#3DDC97", "#FBBF24", "#FF4B4B"],
        )
        chart = alt.Chart(chart_df).mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6).encode(
            x=alt.X("Jumlah:Q", title="Jumlah Data"),
            y=alt.Y("Status:N", sort="-x", title=None),
            color=alt.Color("Status:N", scale=color_scale, legend=None),
            tooltip=["Status", "Jumlah"],
        ).properties(height=150)
        st.altair_chart(chart, use_container_width=True)

        st.divider()

        filter_display = st.radio(
            "Tampilkan",
            [f"Semua · {total}", f"Cocok · {cocok}", f"Anomali Ringan · {ringan}", f"Anomali Berat · {berat}"],
            horizontal=True,
        )
        filter_choice = filter_display.split(" · ")[0]
        search = st.text_input("Cari NIP / Nama", placeholder="Ketik NIP atau nama...")

        view_df = result_df.copy()
        if filter_choice != "Semua":
            view_df = view_df[view_df["Status"] == filter_choice]
        if search:
            mask = view_df.astype(str).apply(
                lambda col: col.str.contains(search, case=False, na=False)
            ).any(axis=1)
            view_df = view_df[mask]

        STATUS_ORDER = {"Anomali Berat": 0, "Anomali Ringan": 1, "Cocok": 2}
        view_df = view_df.copy()
        view_df["_urut"] = view_df["Status"].map(STATUS_ORDER).fillna(3)
        view_df = view_df.sort_values("_urut").drop(columns="_urut").reset_index(drop=True)

        def warnai_baris(row):
            warna = {
                "Cocok": "background-color: #163d2b; color: #b7f0c9",
                "Anomali Ringan": "background-color: #4a3a12; color: #ffe0a3",
                "Anomali Berat": "background-color: #4a1f21; color: #ffb3b8",
            }
            return [warna.get(row["Status"], "")] * len(row)

        styled_df = view_df.style.apply(warnai_baris, axis=1)
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            excel_bytes_filtered = export_to_excel(view_df.drop(columns=["_urut"], errors="ignore"))
            st.download_button(
                f"⬇ Export Tampilan Saat Ini ({len(view_df)} baris)",
                data=excel_bytes_filtered,
                file_name="hasil_filter_saat_ini.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        with col_dl2:
            excel_bytes_all = export_to_excel(result_df)
            st.download_button(
                f"⬇ Export Semua Data ({len(result_df)} baris)",
                data=excel_bytes_all,
                file_name="hasil_deteksi_anomali.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("Unggah file, atur kolom yang dicocokkan, lalu klik 'Jalankan Pencocokan'.")


# =========================================================
# HALAMAN 2: LENGKAPI DATA OTOMATIS
# =========================================================
def render_lengkapi_data():
    st.markdown("<h1>🧩 Lengkapi Data Otomatis</h1>", unsafe_allow_html=True)
    st.caption(
        "Isi kolom yang kosong di File A (mis. Unit Kerja, Jabatan, Pangkat) "
        "menggunakan data dari File B, dicocokkan lewat NIP."
    )
    st.markdown('<div class="header-rule"></div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        file_a = st.file_uploader("File A (data yang mau dilengkapi)", type=["xls", "xlsx", "csv"], key="enrich_file_a")
    with col2:
        file_b = st.file_uploader("File B (data referensi/lengkap)", type=["xls", "xlsx", "csv"], key="enrich_file_b")

    df_a = None
    df_b = None

    if file_a:
        with st.expander("⚙️ Atur Kolom — File A", expanded=True):
            df_a = preview_and_pick_header(file_a, "File A", "enrich_a")

    if file_b:
        with st.expander("⚙️ Atur Kolom — File B", expanded=True):
            df_b = preview_and_pick_header(file_b, "File B", "enrich_b")

    if df_a is not None and df_b is not None:
        st.divider()
        st.markdown("**Pengaturan Pencocokan**")

        col1, col2 = st.columns(2)
        with col1:
            nip_col_a = st.selectbox("Kolom NIP di File A", list(df_a.columns), key="enrich_nip_a")
        with col2:
            nip_col_b = st.selectbox("Kolom NIP di File B", list(df_b.columns), key="enrich_nip_b")

        kolom_tersedia_b = [c for c in df_b.columns if c != nip_col_b]
        value_cols = st.multiselect(
            "Kolom yang ingin diambil dari File B (mis. Unit Kerja, Jabatan, Pangkat)",
            kolom_tersedia_b,
            key="enrich_value_cols",
        )

        overwrite = st.checkbox(
            "Timpa data yang sudah ada di File A (kalau tidak dicentang, hanya sel kosong yang diisi)",
            value=False,
        )

        run_enrich = st.button("Isi Data Otomatis", type="primary", disabled=not value_cols)

        if run_enrich:
            df_a_clean = matcher.drop_empty_rows(df_a, cols=[nip_col_a])
            df_b_clean = matcher.drop_empty_rows(df_b, cols=[nip_col_b])

            lookup = enrichment.build_lookup(df_b_clean, nip_col_b, value_cols)
            hasil_df = enrichment.fill_missing_columns(
                df_a_clean, nip_col_a, value_cols, lookup, overwrite=overwrite
            )
            st.session_state["enrich_result"] = hasil_df

    if "enrich_result" in st.session_state:
        hasil_df = st.session_state["enrich_result"]
        ringkasan = enrichment.summarize_fill(hasil_df)

        st.divider()
        c1, c2, c3 = st.columns(3)
        metric_card(c1, "metric-total", "Total Baris", "📄", ringkasan["total"], "Data File A")
        metric_card(c2, "metric-cocok", "Berhasil Diisi", "✅", ringkasan["terisi"], "Terisi otomatis dari File B")
        metric_card(c3, "metric-berat", "NIP Tidak Ditemukan", "🛑", ringkasan["tidak_ketemu"], "Perlu dicek manual")

        st.divider()
        st.dataframe(hasil_df, use_container_width=True, hide_index=True)

        excel_bytes = export_simple_excel(hasil_df, sheet_name="Data Lengkap")
        st.download_button(
            "⬇ Export Hasil ke Excel",
            data=excel_bytes,
            file_name="file_a_terlengkapi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.info("Unggah File A dan File B, atur kolom NIP, pilih kolom yang ingin diambil, lalu klik 'Isi Data Otomatis'.")


# =========================================================
# SIDEBAR & NAVIGASI
# =========================================================
st.sidebar.title("📁 Menu")
mode = st.sidebar.radio(
    "Pilih fitur",
    ["🔍 Cek Anomali Data", "🧩 Lengkapi Data Otomatis"],
    label_visibility="collapsed",
)
st.sidebar.divider()
st.sidebar.caption("**Cek Anomali Data** — cocokkan & deteksi selisih data antar 2 file berdasarkan NIP & Nama.")
st.sidebar.caption("**Lengkapi Data Otomatis** — isi kolom kosong di File A (Unit Kerja, Jabatan, dll) pakai referensi dari File B.")

if mode == "🔍 Cek Anomali Data":
    render_cek_anomali()
else:
    render_lengkapi_data()
