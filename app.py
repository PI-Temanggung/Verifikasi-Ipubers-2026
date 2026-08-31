import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Monitoring & Verifikasi Nota IPubers",
    page_icon="📊",
    layout="wide",
)

EXCEL_FILE = "IPUBERS-AGUSTUS.xlsx"


# Fungsi untuk mengambil daftar sheet tanpa error cache (menggunakan st.cache_data pada list string)
@st.cache_data
def get_sheet_names():
  xls = pd.ExcelFile(EXCEL_FILE)
  sheet_names = [s for s in xls.sheet_names if s != "Monitoring"]
  return sheet_names


try:
  kabupaten_list = get_sheet_names()
except Exception as e:
  st.error(
      f"Gagal membaca file Excel '{EXCEL_FILE}'. Pastikan file tersebut sudah"
      " di-upload ke GitHub sejajar dengan app.py.\n\nDetail Error: "
      f"{e}"
  )
  st.stop()

st.title("🔍 Aplikasi Verifikasi & Pengecekan Nota IPubers")
st.markdown(
    "Filter kecamatan, pilih kios, verifikasi nota, dan unduh rekap dalam"
    " format Excel."
)

selected_kabupaten = st.sidebar.selectbox(
    "Pilih Kabupaten / Wilayah", kabupaten_list
)


# Fungsi untuk memuat data sheet tertentu secara aman
@st.cache_data
def load_sheet_data(kab_name):
  df_raw = pd.read_excel(EXCEL_FILE, sheet_name=kab_name)
  header_row = 0
  for idx, row in df_raw.head(15).iterrows():
    if (
        "Kecamatan" in row.values
        or "Nama Kios" in row.values
        or "URL Bukti" in row.values
        or any(str(v).strip().lower() == "kecamatan" for v in row.values)
    ):
      header_row = idx
      break
  df = pd.read_excel(EXCEL_FILE, sheet_name=kab_name, skiprows=header_row)
  df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
  return df


try:
  df_kab = load_sheet_data(selected_kabupaten)
except Exception as e:
  st.error(f"Gagal memuat data sheet {selected_kabupaten}: {e}")
  st.stop()


# Mencari nama kolom secara fleksibel
def find_col(df, keywords):
  for col in df.columns:
    for kw in keywords:
      if kw.lower() in str(col).lower():
        return col
  return None


col_kec = find_col(df_kab, ["kecamatan"])
col_kios = find_col(df_kab, ["nama kios", "kios"])
col_trx = find_col(df_kab, ["no transaksi", "kode trx"])
col_petani = find_col(df_kab, ["nama petani", "petani"])

if not col_kec or not col_kios:
  st.error(
      "Kolom 'Kecamatan' atau 'Nama Kios' tidak ditemukan di sheet ini. Kolom"
      f" tersedia: {list(df_kab.columns)}"
  )
  st.stop()

# Filter Kecamatan
kecamatan_list = sorted(df_kab[col_kec].dropna().unique().tolist())
selected_kecamatan = st.sidebar.selectbox(
    "Pilih Kecamatan", ["-- Semua Kecamatan --"] + kecamatan_list
)

if selected_kecamatan != "-- Semua Kecamatan --":
  df_filtered = df_kab[df_kab[col_kec] == selected_kecamatan]
else:
  df_filtered = df_kab

# Filter Kios
kios_list = sorted(df_filtered[col_kios].dropna().unique().tolist())
selected_kios = st.sidebar.selectbox("Pilih Kios", ["-- Pilih Kios --"] + kios_list)

if "verifikasi_status" not in st.session_state:
  st.session_state.verifikasi_status = {}

if selected_kios != "-- Pilih Kios --":
  df_kios = df_filtered[df_filtered[col_kios] == selected_kios].reset_index(
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

    c_n1, c_n2, c_n3 = st.columns([1, 2, 1])
    with c_n1:
      if st.button("⬅️ Sebelumnya", use_container_width=True):
        if st.session_state.current_index > 0:
          st.session_state.current_index -= 1
          st.rerun()
    with c_n2:
      st.markdown(
          f"<p style='text-align: center; font-weight: bold;'>Nota ke-"
          f" {idx_row + 1} dari {len(df_kios)}</p>",
          unsafe_allow_html=True,
      )
    with c_n3:
      if st.button("Selanjutnya ➡️", use_container_width=True):
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
          st.rerun()

    st.markdown("### 📄 Detail Transaksi Nota")
    d1, d2 = st.columns(2)
    with d1:
      trx_val = row_data.get(col_trx, "-") if col_trx else "-"
      petani_val = row_data.get(col_petani, "-") if col_petani else "-"
      nik_val = row_data.get("NIK", "-")
      tgl_val = row_data.get("Tanggal Tebus", "-")
      st.info(
          f"**No Transaksi:** {trx_val}\n\n**Nama Petani:**"
          f" {petani_val}\n\n**NIK:** {nik_val}\n\n**Tanggal Tebus:** {tgl_val}"
      )
    with d2:
      url_col = find_col(df_kios, ["url", "bukti"])
      nota_url = row_data.get(url_col, "#") if url_col else "#"
      st.markdown("**Alokasi Pupuk:**")
      pupuk_info = []
      for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
        if p in df_kios.columns and pd.notna(row_data.get(p)):
          pupuk_info.append(f"- {p}: **{row_data.get(p)} kg**")
      st.markdown("\n".join(pupuk_info) if pupuk_info else "Data pupuk nihil")

    if pd.notna(nota_url) and str(nota_url).startswith("http"):
      st.markdown(
          f"🔗 **Link Dokumen Nota:** [Buka Nota di Tab Baru]({nota_url})"
      )
    else:
      st.warning("Link dokumen tidak tersedia.")

    key_state = (selected_kabupaten, selected_kios, idx_row)
    current_status = st.session_state.verifikasi_status.get(
        key_state, "Belum Dicek"
    )
    st.markdown(f"Status Verifikasi Saat Ini: **{current_status}**")

    b1, b2, b3 = st.columns(3)
    with b1:
      if st.button("✅ TERIMA", type="primary", use_container_width=True):
        st.session_state.verifikasi_status[key_state] = "Diterima"
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
        st.rerun()
    with b2:
      if st.button("❌ TOLAK", use_container_width=True):
        st.session_state.verifikasi_status[key_state] = "Ditolak"
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
        st.rerun()
    with b3:
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
  cols_to_show = [
      c
      for c in [col_trx, col_petani, "Tanggal Tebus", "Urea", "NPK", "Status Cek"]
      if c and c in df_display.columns
  ]
  st.dataframe(df_display[cols_to_show], use_container_width=True)

st.markdown("---")
st.subheader("📥 Download File Excel")
if st.button("📊 Siapkan File Excel", type="primary"):
  st.success("File siap diunduh!")
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
