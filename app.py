from datetime import datetime
import openpyxl
import pandas as pd
import streamlit as st

# Konfigurasi Halaman Web
st.set_page_config(
    page_title="Dashboard Data Sorting September 2026",
    page_icon="🦀",
    layout="wide",
)

EXCEL_FILE = "DATA SORTING SEPTEMBER 2026.xlsx"


@st.cache_data
ools
def load_data():
    xls = pd.ExcelFile(EXCEL_FILE)
    sheet_names = xls.sheet_names

    # Pisahkan jenis sheet
    sortir_sheets = [s for s in sheet_names if "SORTIR" in s.upper()]
    other_sheets = [s for s in sheet_names if s not in sortir_sheets]

    data_others = {}
    for s in other_sheets:
        data_others[s] = pd.read_excel(EXCEL_FILE, sheet_name=s)

    return sheet_names, sortir_sheets, other_sheets, data_others


try:
    sheet_names, sortir_sheets, other_sheets, data_others = load_data()
except Exception as e:
    st.error(
        f"Gagal memuat file Excel. Pastikan file '{EXCEL_FILE}' ada di direktori yang sama. Error: {e}"
    )
    st.stop()

# Sidebar Navigasi
st.sidebar.title("📊 Navigasi Menu")
menu = st.sidebar.radio(
    "Pilih Menu:", ["Laporan Sortir Harian", "Data Suplier & Lainnya"]
)

st.title("🦀 Dashboard Data Sorting & Receiving September 2026")

if menu == "Laporan Sortir Harian":
    st.subheader("📅 Laporan Data Proses Harian Sortir")

    if not sortir_sheets:
        st.warning("Tidak ditemukan sheet laporan sortir harian.")
    else:
        # Pilih tanggal / sheet sortir
        selected_sheet = st.selectbox("Pilih Tanggal Laporan:", sortir_sheets)

        # Baca sheet sortir yang dipilih secara spesifik menggunakan pandas
        df_sortir = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet)

        st.markdown(f"### Menampilkan Data dari: **{selected_sheet}**")

        # Tampilkan DataFrame mentah atau bagian yang relevan
        with st.expander("Lihat Data Mentah Sheet Ini"):
            st.dataframe(df_sortir, use_container_width=True)

        # Coba ekstrak tabel ringkasan QTY (Kg) jika polanya konsisten
        try:
            # Baris tabel ringkasan biasanya mulai sekitar baris indeks tertentu
            # Mari tampilkan preview tabel yang bersih
            st.info(
                "Berikut adalah data laporan harian yang terbaca dari sheet terpilih:"
            )
            st.dataframe(df_sortir.dropna(how="all"), use_container_width=True)
        except Exception as ex:
            st.error(f"Gagal memparsing tabel terstruktur: {ex}")

elif menu == "Data Suplier & Lainnya":
    st.subheader("📁 Data Master & Pendukung (Suplier, RCV, dll)")

    selected_other = st.selectbox("Pilih Kategori Data:", other_sheets)

    df_selected = data_others[selected_other]

    st.markdown(f"### Data: **{selected_other}**")

    # Fitur Pencarian sederhana
    search_query = st.text_input(
        "🔍 Cari Data dalam tabel:",
        "",
        placeholder="Ketik kata kunci...",
    )

    if search_query:
        # Filter berdasarkan string di semua kolom
        mask = df_selected.astype(str).apply(
            lambda x: x.str.contains(search_query, case=False, na=False)
        ).any(axis=1)
        df_filtered = df_selected[mask]
    else:
        df_filtered = df_selected

    st.dataframe(df_filtered, use_container_width=True)

    # Tombol Download CSV
    csv = df_filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Data ini sebagai CSV",
        data=csv,
        file_name=f"{selected_other}_September_2026.csv",
        mime="text/csv",
    )

st.sidebar.markdown("---")
st.sidebar.info("Aplikasi Web Dashboard Operasional dikembangkan dengan Streamlit.")