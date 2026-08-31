import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Verifikasi Nota IPubers", page_icon="🔍", layout="wide"
)

EXCEL_FILE = "IPUBERS-AGUSTUS.xlsx"


# Memuat data dengan cache otomatis
@st.cache_data(ttl=60)
def load_excel_data():
  xls = pd.ExcelFile(EXCEL_FILE)
  sheet_name = xls.sheet_names[0]
  df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
  return df, sheet_name


try:
  df_original, active_sheet = load_excel_data()
except Exception as e:
  st.error(
      f"Gagal membaca file Excel '{EXCEL_FILE}'. Pastikan file sudah di-upload"
      f" ke GitHub sejajar dengan app.py.\n\nError: {e}"
  )
  st.stop()

st.title("🔍 Panel Verifikasi Nota Kios IPubers")
st.markdown(
    "Pilih Kecamatan dan Kode Kios di bawah untuk memverifikasi nota. Status"
    " verifikasi aman dan tidak akan hilang meskipun Anda menambah data baru di"
    " Excel."
)

# --- PENCARIAN NAMA KOLOM SECARA FLEKSIBEL ---
cols = list(df_original.columns)


def find_col(keywords):
  for c in cols:
    for kw in keywords:
      if kw.lower() in str(c).lower():
        return c
  return None


col_kec = find_col(["kecamatan"])
col_kios_code = find_col(["kode kios"])
col_kios_name = find_col(["nama kios", "kios"])
col_trx = find_col(["no transaksi", "kode trx"])
col_petani = find_col(["nama petani", "petani"])
col_url = find_col(["url bukti", "link", "url"])

if not col_kec or not col_kios_code or not col_trx:
  st.error(
      "Kolom penting ('Kecamatan', 'Kode Kios', atau 'No Transaksi') tidak"
      f" lengkap di dalam file Excel Anda. Kolom yang terdeteksi: {cols}"
  )
  st.stop()

# --- FILTER UTAMA DI BAGIAN ATAS ---
f_col1, f_col2 = st.columns(2)

with f_col1:
  kecamatan_list = sorted(df_original[col_kec].dropna().unique().tolist())
  selected_kecamatan = st.selectbox(
      "1. Pilih Kecamatan", ["-- Pilih Kecamatan --"] + kecamatan_list
  )

if selected_kecamatan != "-- Pilih Kecamatan --":
  df_filtered = df_original[df_original[col_kec] == selected_kecamatan]
else:
  df_filtered = df_original

with f_col2:
  df_filtered["Display_Kios"] = (
      df_filtered[col_kios_code].astype(str)
      + " - "
      + df_filtered[col_kios_name].astype(str)
  )
  kios_list = sorted(df_filtered["Display_Kios"].dropna().unique().tolist())
  selected_display_kios = st.selectbox(
      "2. Pilih Kode Kios", ["-- Pilih Kios --"] + kios_list
  )

st.markdown("---")

# Inisialisasi Session State menggunakan Dictionary berbasis No Transaksi (Unik)
if "verifikasi_dict" not in st.session_state:
  st.session_state.verifikasi_dict = {}

if selected_display_kios != "-- Pilih Kios --":
  selected_kode_kios = selected_display_kios.split(" - ")[0]

  df_kios_all = df_filtered[
      df_filtered[col_kios_code].astype(str) == selected_kode_kios
  ]

  # --- FILTER STATUS VERIFIKASI ---
  st.markdown("#### 🔎 Filter Berdasarkan Status Pengecekan:")
  status_filter_options = [
      "Semua Nota",
      "Belum Dicek",
      "TERIMA",
      "TOLAK",
  ]
  selected_status_filter = st.radio(
      "Tampilkan nota dengan status:",
      status_filter_options,
      horizontal=True,
  )

  # Menyaring baris berdasarkan status verifikasi menggunakan No Transaksi sebagai kunci
  filtered_indices = []
  for idx, row in df_kios_all.iterrows():
    trx_key = str(row[col_trx])
    status = st.session_state.verifikasi_dict.get(trx_key, "Belum Dicek")

    if selected_status_filter == "Belum Dicek" and status == "Belum Dicek":
      filtered_indices.append(idx)
    elif selected_status_filter == "TERIMA" and status == "TERIMA":
      filtered_indices.append(idx)
    elif selected_status_filter == "TOLAK" and status == "TOLAK":
      filtered_indices.append(idx)
    elif selected_status_filter == "Semua Nota":
      filtered_indices.append(idx)

  if len(filtered_indices) == 0:
    st.warning(
        f"Tidak ada data nota dengan status '{selected_status_filter}' untuk"
        " kios ini."
    )
  else:
    if "current_pos" not in st.session_state:
      st.session_state.current_pos = 0
    if (
        "last_kios" not in st.session_state
        or st.session_state.last_kios != selected_kode_kios
    ):
      st.session_state.current_pos = 0
      st.session_state.last_kios = selected_kode_kios

    if st.session_state.current_pos >= len(filtered_indices):
      st.session_state.current_pos = 0

    pos = st.session_state.current_pos
    row_idx = filtered_indices[pos]
    row_data = df_original.loc[row_idx]
    current_trx_key = str(row_data[col_trx])

    # Statistik Ringkas Kios
    total_nota_kios = len(df_kios_all)
    sudah_cek = sum(
        1
        for _, r in df_kios_all.iterrows()
        if str(r[col_trx]) in st.session_state.verifikasi_dict
    )
    diterima = sum(
        1
        for _, r in df_kios_all.iterrows()
        if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TERIMA"
    )
    ditolak = sum(
        1
        for _, r in df_kios_all.iterrows()
        if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TOLAK"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nota Kios", total_nota_kios)
    m2.metric("Sudah Diverifikasi", sudah_cek)
    m3.metric("Diterima", diterima)
    m4.metric("Ditolak", ditolak)

    st.markdown("---")

    # Layout Utama: Kolom Kiri (Detail & Aksi), Kolom Kanan (Preview Nota)
    col_kiri, col_kanan = st.columns([1, 2], gap="large")

    with col_kiri:
      st.subheader("📄 Detail Transaksi")
      trx_val = row_data.get(col_trx, "-")
      petani_val = row_data.get(col_petani, "-") if col_petani else "-"
      nik_val = row_data.get("NIK", "-")
      tgl_val = row_data.get("Tanggal Tebus", "-")

      st.markdown(f"**No Transaksi:**\n`{trx_val}`")
      st.markdown(f"**Nama Petani:**\n{petani_val}")
      st.markdown(f"**NIK:**\n{nik_val}")
      st.markdown(f"**Tanggal Tebus:**\n{tgl_val}")

      pupuk_info = []
      for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
        if p in df_original.columns and pd.notna(row_data.get(p)):
          pupuk_info.append(f"{p}: **{row_data.get(p)} kg**")
      if pupuk_info:
        st.markdown(f"🌾 **Alokasi Pupuk:**\n" + "\n".join(pupuk_info))

      current_status = st.session_state.verifikasi_dict.get(
          current_trx_key, "Belum Dicek"
      )
      status_color = (
          "green"
          if current_status == "TERIMA"
          else ("red" if current_status == "TOLAK" else "orange")
      )
      st.markdown(
          f"Status: <span"
          f" style='color:{status_color}; font-weight:bold;'>{current_status}</span>",
          unsafe_allow_html=True,
      )

      st.markdown("---")
      st.markdown("#### Aksi Verifikasi:")
      if st.button("✅ TERIMA", type="primary", key=f"terima_{row_idx}"):
        st.session_state.verifikasi_dict[current_trx_key] = "TERIMA"
        if pos < len(filtered_indices) - 1:
          st.session_state.current_pos += 1
        st.rerun()

      if st.button("❌ TOLAK", key=f"tolak_{row_idx}"):
        st.session_state.verifikasi_dict[current_trx_key] = "TOLAK"
        if pos < len(filtered_indices) - 1:
          st.session_state.current_pos += 1
        st.rerun()

      if st.button("🔄 Reset Status", key=f"reset_{row_idx}"):
        if current_trx_key in st.session_state.verifikasi_dict:
          del st.session_state.verifikasi_dict[current_trx_key]
        st.rerun()

      st.markdown("---")
      st.markdown("#### Navigasi Nota:")
      if st.button("⬅️ Sebelumnya", key=f"prev_{row_idx}"):
        if pos > 0:
          st.session_state.current_pos -= 1
          st.rerun()

      st.markdown(
          f"<p style='text-align: center; font-weight: bold; margin: 5px"
          f" 0;'>Nota ke-{pos + 1} dari {len(filtered_indices)} (Kategori"
          f" {selected_status_filter})</p>",
          unsafe_allow_html=True,
      )

      if st.button("Selanjutnya ➡️", key=f"next_{row_idx}"):
        if pos < len(filtered_indices) - 1:
          st.session_state.current_pos += 1
          st.rerun()

    with col_kanan:
      st.subheader("🖼️ Preview Nota (Diperbesar)")
      nota_url = row_data.get(col_url, None) if col_url else None

      if pd.notna(nota_url) and str(nota_url).startswith("http"):
        st.markdown(
            f'<iframe src="{nota_url}" width="100%" height="750px"'
            ' style="border: 2px solid #0055ff; border-radius: 8px;'
            ' background-color: white;"></iframe>',
            unsafe_allow_html=True,
        )
        st.markdown(f"🔗 [Buka Link Asli di Tab Baru]({nota_url})")
      else:
        st.warning(
            "Link atau URL bukti nota tidak tersedia pada baris data ini."
        )

# --- PANEL DOWNLOAD HASIL ---
st.markdown("---")
st.subheader("📥 Download File Excel Hasil Pengecekan")
st.markdown(
    "Pilih jenis file yang ingin di-download. Status verifikasi dicocokkan"
    " berdasarkan No Transaksi secara akurat."
)

dl_col1, dl_col2 = st.columns(2)

with dl_col1:
  st.markdown("#### 1. Download Seluruh Data (Semua File)")
  if st.button("📊 Download Semua Data Excel", type="primary"):
    df_export_all = df_original.copy()
    status_list_all = [
        st.session_state.verifikasi_dict.get(str(row[col_trx]), "Belum Dicek")
        for _, row in df_export_all.iterrows()
    ]
    df_export_all["Status_Verifikasi"] = status_list_all

    output_all = io.BytesIO()
    with pd.ExcelWriter(output_all, engine="openpyxl") as writer:
      df_export_all.to_excel(writer, sheet_name=active_sheet, index=False)

    st.download_button(
        label="⬇️ Simpan File (Semua Data)",
        data=output_all.getvalue(),
        file_name="Hasil_Verifikasi_Semua_Data.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

with dl_col2:
  st.markdown("#### 2. Download Data Kios/Kecamatan Terpilih")
  if st.button("📊 Download Data Terpilih Saja"):
    if selected_display_kios != "-- Pilih Kios --":
      df_export_filtered = df_kios_all.copy()
    elif selected_kecamatan != "-- Pilih Kecamatan --":
      df_export_filtered = df_filtered.copy()
    else:
      df_export_filtered = df_original.copy()

    status_list_filtered = [
        st.session_state.verifikasi_dict.get(str(row[col_trx]), "Belum Dicek")
        for _, row in df_export_filtered.iterrows()
    ]
    df_export_filtered["Status_Verifikasi"] = status_list_filtered

    if "Display_Kios" in df_export_filtered.columns:
      df_export_filtered = df_export_filtered.drop(columns=["Display_Kios"])

    output_filtered = io.BytesIO()
    with pd.ExcelWriter(output_filtered, engine="openpyxl") as writer:
      df_export_filtered.to_excel(writer, sheet_name=active_sheet, index=False)

    st.download_button(
        label="⬇️ Simpan File (Data Terpilih)",
        data=output_filtered.getvalue(),
        file_name="Hasil_Verifikasi_Terpilih.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
