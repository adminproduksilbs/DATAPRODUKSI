from datetime import datetime
import pandas as pd
import streamlit as st

# Konfigurasi Halaman Web agar tampil Full Width (Wide)
st.set_page_config(
    page_title="Sistem Manajemen Produksi & Sorting",
    page_icon="🦀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS untuk mempercantik tampilan (Warna, Kartu, Font, Tabel)
st.markdown(
    """
    <style>
        .main {
            background-color: #f8f9fa;
        }
        .stMetric {
            background-color: #ffffff;
            padding: 15px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-left: 5px solid #ff4b4b;
        }
        h1, h2, h3 {
            color: #1f2937;
        }
        .sidebar .sidebar-content {
            background-color: #f3f4f6;
        }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "DATA SORTING SEPTEMBER 2026.xlsx"


@st.cache_data
def load_data():
  xls = pd.ExcelFile(EXCEL_FILE)
  sheet_names = xls.sheet_names

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
      f"Gagal memuat file Excel '{EXCEL_FILE}'. Pastikan file ada di direktori"
      f" yang sama. Error: {e}"
  )
  st.stop()

# --- SIDEBAR NAVIGASI ---
st.sidebar.image(
    "https://img.icons8.com/color/96/crab.png", width=80
)  # Logo Maskot
st.sidebar.title("🏢 Portal Departemen")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Pilih Menu Utama:",
    [
        "📊 Dashboard Utama",
        "📥 Penerimaan (Receiving RM & Meat)",
        "⚙️ Laporan Sortir Harian",
        "🗂️ Data Suplier & Master",
    ],
)

st.sidebar.markdown("---")
st.sidebar.info(
    "💡 **Info:** Sistem ini terhubung langsung untuk akses seluruh"
    " departemen (Produksi, QC, Finance, & Manajemen)."
)

# --- HALAMAN 1: DASHBOARD UTAMA ---
if menu == "📊 Dashboard Utama":
  st.title("🦀 Dashboard Monitoring Produksi September 2026")
  st.markdown(
      "Selamat datang di pusat data operasional. Berikut adalah ringkasan"
      " cepat aktivitas produksi bulan ini."
  )

  # Kolom Metrik Ringkasan
  col1, col2, col3, col4 = st.columns(4)
  with col1:
    st.metric(label="Total Sheet Laporan", value=len(sortir_sheets))
  with col2:
    st.metric(label="Total Suplier Aktif", value="8+ Suplier")
  with col3:
    st.metric(label="Status Sistem", value="Online 🟢", delta="Normal")
  with col4:
    st.metric(label="Periode Data", value="September 2026")

  st.markdown("---")
  st.subheader("📌 Petunjuk Akses Departemen")
  col_a, col_b = st.columns(2)
  with col_a:
    st.markdown("""
        ### 🏭 Departemen Produksi & QC
        * Gunakan menu **Laporan Sortir Harian** untuk memantau hasil sortir kaleng, berat bersih, dan susut harian.
        * Periksa catatan rijek, shell & air, serta hasil mixing.
        """)
  with col_b:
    st.markdown("""
        ### 💰 Departemen Finance & Manajemen
        * Gunakan menu **Data Suplier & Master** untuk mengecek data rekening bank suplier.
        * Data dapat di-download langsung ke format CSV/Excel untuk laporan keuangan.
        """)

# --- HALAMAN 2: PENERIMAAN (RECEIVING) ---
elif menu == "📥 Penerimaan (Receiving RM & Meat)":
  st.title("📥 Penerimaan Raw Material & Meat")
  st.markdown(
      "Modul pencatatan bahan masuk dari suplier dan pencetakan nota/surat jalan."
  )

  tab1, tab2 = st.tabs(["📋 Data Receiving RC & ME", "➕ Input / Cetak Nota"])

  with tab1:
    if "RCV RC" in data_others and "RCV ME" in data_others:
      rcv_choice = st.selectbox("Pilih Jenis Receiving:", ["RCV RC", "RCV ME"])
      st.dataframe(
          data_others[rcv_choice].dropna(how="all"), use_container_width=True
      )
    else:
      st.info("Data receiving sedang dimuat.")

  with tab2:
    st.info(
        "Formulir Input Penerimaan & Generator Nota Otomatis (Segera"
        " diintegrasikan dengan database input live)."
    )
    with st.form("form_nota"):
      col1, col2 = st.columns(2)
      with col1:
        supplier_name = st.selectbox(
            "Nama Suplier",
            [
                "Novita (Palembang)",
                "Nur Halim (Seputih)",
                "Eka SW (Teladas)",
                "Juwandi (Tegal)",
            ],
        )
        material_type = st.selectbox(
            "Jenis Bahan", ["RC Crab", "Meat Crab", "Slipper Lobster"]
        )
      with col2:
        berat_kotor = st.number_input("Berat Kotor (Kg)", min_value=0.0)
        keranjang = st.number_input("Berat Keranjang / Potongan", min_value=0.0)

      submitted = st.form_submit_button("🖨️ Cetak Nota & Simpan")
      if submitted:
        st.success(
            f"Nota untuk suplier **{supplier_name}** berhasil dibuat! Berat"
            f" Bersih: {berat_kotor - keranjang} Kg"
        )

# --- HALAMAN 3: LAPORAN SORTIR HARIAN ---
elif menu == "⚙️ Laporan Sortir Harian":
  st.title("⚙️ Laporan Proses & Sortir Harian")
  st.markdown(
      "Pilih tanggal laporan harian untuk melihat rincian sortir secara"
      " transparan."
  )

  if not sortir_sheets:
    st.warning("Tidak ditemukan data laporan sortir.")
  else:
    selected_sheet = st.selectbox(
        "📅 Pilih Tanggal Laporan Harian:", sortir_sheets
    )

    df_sortir = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet)

    # Tampilkan Header Info
    st.markdown(f"### 📄 Hasil Laporan: **{selected_sheet}**")

    # Kolom pencarian data di dalam tabel harian
    search_sortir = st.text_input(
        "🔍 Cari jenis daging / data pada tanggal ini:",
        "",
        placeholder="Contoh: Jumbo, Jus, Backfin...",
    )
    if search_sortir:
      mask_s = df_sortir.astype(str).apply(
          lambda x: x.str.contains(search_sortir, case=False, na=False)
      ).any(axis=1)
      df_s_filtered = df_sortir[mask_s]
    else:
      df_s_filtered = df_sortir

    st.dataframe(df_s_filtered.dropna(how="all"), use_container_width=True)

    # Tombol Download
    csv_sortir = df_s_filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Laporan Tanggal Ini (.CSV)",
        data=csv_sortir,
        file_name=f"{selected_sheet}.csv",
        mime="text/csv",
    )

# --- HALAMAN 4: DATA SUPLIER & MASTER ---
elif menu == "🗂️ Data Suplier & Master":
  st.title("🗂️ Data Master Suplier & Pendukung")
  st.markdown("Informasi lengkap rekening bank dan data area suplier.")

  selected_other = st.selectbox(
      "Pilih Tabel Data Master:", list(data_others.keys())
  )

  df_selected = data_others[selected_other]

  # Fitur pencarian global
  search_query = st.text_input(
      "🔍 Cari Data Suplier / Keyword:",
      "",
      placeholder="Ketik nama bank, area, atau suplier...",
  )

  if search_query:
    mask = df_selected.astype(str).apply(
        lambda x: x.str.contains(search_query, case=False, na=False)
    ).any(axis=1)
    df_filtered = df_selected[mask]
  else:
    df_filtered = df_selected

  st.dataframe(df_filtered, use_container_width=True)

  # Tombol Download
  csv = df_filtered.to_csv(index=False).encode("utf-8")
  st.download_button(
      label="📥 Download Data Ini sebagai CSV",
      data=csv,
      file_name=f"{selected_other}_September_2026.csv",
      mime="text/csv",
  )

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>© 2026 Sistem Informasi"
    " Produksi & Sorting | Developed for All Departments</p>",
    unsafe_allow_html=True,
)
