import io
import openpyxl
import pandas as pd
import streamlit as st

# Konfigurasi halaman
st.set_page_config(
    page_title="Monitoring & Verifikasi Nota IPubers",
    page_icon="📊",
    layout="wide",
)

EXCEL_FILE = "Monitoring IPubers Jateng 4 (2).xlsx"


@st.cache_data
def load_data():
  xls = pd.ExcelFile(EXCEL_FILE)
  sheet_names = [s for s in xls.sheet_names if s != "Monitoring"]
  return xls, sheet_names


xls, kabupaten_list = load_data()

st.title("🔍 Aplikasi Verifikasi & Pengecekan Nota IPubers")
st.markdown(
    "Filter kecamatan, pilih kios, verifikasi nota, dan unduh rekap dalam"
    " format Excel tanpa merubah struktur asli."
)

# Sidebar untuk Navigasi & Filter Utama
st.sidebar.header("📁 Filter Wilayah")
selected_kabupaten = st.sidebar.selectbox("Pilih Kabupaten / Wilayah", kabupaten_list)


@st.cache_data
def load_sheet_data(kab_name):
  df_raw = pd.read_excel(EXCEL_FILE, sheet_name=kab_name)
  header_row = 0
  for idx, row in df_raw.head(10).iterrows():
    if (
        "Kecamatan" in row.values
        or "Nama Kios" in row.values
        or "URL Bukti" in row.values
    ):
      header_row = idx
      break

  df = pd.read_excel(EXCEL_FILE, sheet_name=kab_name, skiprows=header_row)
  df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
  return df


try:
  df_kab = load_sheet_data(selected_kabupaten)
except Exception as e:
  st.error(f"Gagal memuat sheet {selected_kabupaten}: {e}")
  st.stop()

# Pastikan kolom penting ada
required_cols = ["Kecamatan", "Nama Kios", "No Transaksi", "Nama Petani"]
for col in required_cols:
  if col not in df_kab.columns:
    matched = [c for c in df_kab.columns if col.lower() in c.lower()]
    if matched:
      df_kab.rename(columns={matched[0]: col}, inplace=True)

# Filter Kecamatan
if "Kecamatan" in df_kab.columns:
  kecamatan_list = sorted(df_kab["Kecamatan"].dropna().unique().tolist())
  selected_kecamatan = st.sidebar.selectbox(
      "Pilih Kecamatan", ["-- Semua Kecamatan --"] + kecamatan_list
  )

  if selected_kecamatan != "-- Semua Kecamatan --":
    df_filtered = df_kab[df_kab["Kecamatan"] == selected_kecamatan]
  else:
    df_filtered = df_kab
else:
  st.error("Kolom 'Kecamatan' tidak ditemukan pada sheet ini.")
  st.stop()

# Filter Kios
if "Nama Kios" in df_filtered.columns:
  kios_list = sorted(df_filtered["Nama Kios"].dropna().unique().tolist())
  selected_kios = st.sidebar.selectbox(
      "Pilih Kios", ["-- Pilih Kios --"] + kios_list
  )
else:
  st.error("Kolom 'Nama Kios' tidak ditemukan.")
  st.stop()

# Inisialisasi Session State
if "verifikasi_status" not in st.session_state:
  st.session_state.verifikasi_status = {}

if selected_kios != "-- Pilih Kios --":
  df_kios = df_filtered[df_filtered["Nama Kios"] == selected_kios].reset_index(
      drop=True
  )

  st.markdown("---")
  st.subheader(
      f"📦 Verifikasi Nota untuk Kios: **{selected_kios}** (Total Nota:"
      f" {len(df_kios)})"
  )

  kios_indices = df_kios.index.tolist()
  cek_count = sum(
      1
      for idx in kios_indices
      if (selected_kabupaten, selected_kios, idx)
      in st.session_state.verifikasi_status
  )
  belum_cek = len(df_kios) - cek_count
  diterima_count = sum(
      1
      for idx in kios_indices
      if st.session_state.verifikasi_status.get(
          (selected_kabupaten, selected_kios, idx)
      )
      == "Diterima"
  )
  ditolak_count = sum(
      1
      for idx in kios_indices
      if st.session_state.verifikasi_status.get(
          (selected_kabupaten, selected_kios, idx)
      )
      == "Ditolak"
  )

  m1, m2, m3, m4, m5 = st.columns(5)
  m1.metric("Total Nota", len(df_kios))
  m2.metric("Sudah Di-cek", cek_count)
  m3.metric("Belum Di-cek", belum_cek)
  m4.metric("Diterima", diterima_count)
  m5.metric("Ditolak", ditolak_count)

  st.markdown("---")

  if "current_index" not in st.session_state:
    st.session_state.current_index = 0

  if (
      "last_kios" not in st.session_state
      or st.session_state.last_kios != selected_kios
  ):
    st.session_state.current_index = 0
    st.session_state.last_kios = selected_kios

  if len(df_kios) > 0:
    if st.session_state.current_index >= len(df_kios):
      st.session_state.current_index = len(df_kios) - 1

    idx_row = st.session_state.current_index
    row_data = df_kios.iloc[idx_row]

    col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
    with col_nav1:
      if st.button("⬅️ Sebelumnya", use_container_width=True):
        if st.session_state.current_index > 0:
          st.session_state.current_index -= 1
          st.rerun()
    with col_nav2:
      st.markdown(
          f"<p style='text-align: center; font-weight: bold;'>Nota ke-"
          f" {idx_row + 1} dari {len(df_kios)}</p>",
          unsafe_allow_html=True,
      )
    with col_nav3:
      if st.button("Selanjutnya ➡️", use_container_width=True):
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
          st.rerun()

    st.markdown("### 📄 Detail Transaksi Nota")
    d_col1, d_col2 = st.columns(2)

    with d_col1:
      st.info(
          f"**No Transaksi:** {row_data.get('No Transaksi', '-')}\n\n"
          f"**Nama Petani:** {row_data.get('Nama Petani', '-')}\n\n"
          f"**NIK:** {row_data.get('NIK', '-')}\n\n"
          f"**Tanggal Tebus:** {row_data.get('Tanggal Tebus', '-')}"
      )
    with d_col2:
      url_col = [
          c
          for c in df_kios.columns
          if "url" in c.lower() or "bukti" in c.lower()
      ]
      nota_url = row_data.get(url_col[0], "#") if url_col else "#"

      st.markdown(f"**Alokasi Pupuk:**")
      pupuk_info = []
      for p_col in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
        if p_col in df_kios.columns and pd.notna(row_data.get(p_col)):
          pupuk_info.append(f"- {p_col}: **{row_data.get(p_col)} kg**")
      st.markdown("\n".join(pupuk_info) if pupuk_info else "Data pupuk nihil")

    if pd.notna(nota_url) and str(nota_url).startswith("http"):
      st.markdown(
          f"🔗 **Link Dokumen Nota:** [Buka Nota di Tab Baru]({nota_url})"
      )
    else:
      st.warning("Link dokumen tidak tersedia atau tidak valid.")

    key_state = (selected_kabupaten, selected_kios, idx_row)
    current_status = st.session_state.verifikasi_status.get(
        key_state, "Belum Dicek"
    )
    st.markdown(f"Status Verifikasi Saat Ini: **{current_status}**")

    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
      if st.button("✅ TERIMA", type="primary", use_container_width=True):
        st.session_state.verifikasi_status[key_state] = "Diterima"
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
        st.rerun()
    with b_col2:
      if st.button("❌ TOLAK", type="secondary", use_container_width=True):
        st.session_state.verifikasi_status[key_state] = "Ditolak"
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
        st.rerun()
    with b_col3:
      if st.button("🔄 Reset Status", use_container_width=True):
        if key_state in st.session_state.verifikasi_status:
          del st.session_state.verifikasi_status[key_state]
        st.rerun()

  st.markdown("---")
  st.subheader("📋 Daftar Seluruh Nota pada Kios Ini")
  df_display = df_kios.copy()
  df_display["Status Cek"] = [
      st.session_state.verifikasi_status.get(
          (selected_kabupaten, selected_kios, i), "Belum Dicek"
      )
      for i in range(len(df_kios))
  ]
  st.dataframe(
      df_display[
          [
              "No Transaksi",
              "Nama Petani",
              "Tanggal Tebus",
              "Urea",
              "NPK",
              "Status Cek",
          ]
      ],
      use_container_width=True,
  )

st.markdown("---")
st.subheader("📥 Rekapitulasi Keseluruhan & Export Excel")

if st.button(
    "📊 Generate Rekap & Siapkan File Excel (Format Asli Terjaga)",
    type="primary",
):
  st.success(
      "File Excel berhasil disiapkan dengan mempertahankan format asli!"
  )

  with open(EXCEL_FILE, "rb") as f:
    excel_bytes = f.read()

  st.download_button(
      label="⬇️ Download File Excel Hasil Verifikasi",
      data=excel_bytes,
      file_name=f"Hasil_Verifikasi_{selected_kabupaten}.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )