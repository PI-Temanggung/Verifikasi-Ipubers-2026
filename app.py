import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Verifikasi Nota IPubers", page_icon="🔍", layout="wide"
)

EXCEL_FILE = "IPUBERS-AGUSTUS.xlsx"


# Mengambil daftar sheet (Kabupaten) secara aman
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

st.title("🔍 Panel Verifikasi Nota Kios IPubers")
st.markdown(
    "Pilih Kabupaten, Kecamatan, dan Kode Kios di bawah, lalu lakukan"
    " verifikasi nota langsung pada halaman ini."
)

# --- FILTER UTAMA DI BAGIAN ATAS ---
col_f1, col_f2, col_f3 = st.columns(3)

with col_f1:
  selected_kabupaten = st.selectbox("1. Pilih Kabupaten", kabupaten_list)


# Memuat data sheet terpilih dan membersihkannya dari baris kosong
@st.cache_data
def load_sheet_data(kab_name):
  df_raw = pd.read_excel(EXCEL_FILE, sheet_name=kab_name)
  header_row = 0
  for idx, row in df_raw.head(15).iterrows():
    if (
        "Kecamatan" in row.values
        or "Nama Kios" in row.values
        or "Kode Kios" in row.values
        or any(str(v).strip().lower() == "kecamatan" for v in row.values)
    ):
      header_row = idx
      break
  df = pd.read_excel(EXCEL_FILE, sheet_name=kab_name, skiprows=header_row)
  df.columns = df.iloc[0].astype(str)
  df = df[1:].reset_index(drop=True)
  df = df.loc[
      :, ~pd.Series(df.columns.astype(str)).str.startswith("Unnamed").values
  ]
  if "Kode Kios" in df.columns:
    df = df[df["Kode Kios"].notna() & (df["Kode Kios"] != "Kode Kios")]
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
col_kios_code = find_col(df_kab, ["kode kios"])
col_kios_name = find_col(df_kab, ["nama kios", "kios"])
col_trx = find_col(df_kab, ["no transaksi", "kode trx"])
col_petani = find_col(df_kab, ["nama petani", "petani"])

if not col_kec or not col_kios_code:
  st.error(
      "Kolom 'Kecamatan' atau 'Kode Kios' tidak ditemukan di sheet ini. Kolom"
      f" tersedia: {list(df_kab.columns)}"
  )
  st.stop()

with col_f2:
  kecamatan_list = sorted(df_kab[col_kec].dropna().unique().tolist())
  selected_kecamatan = st.selectbox(
      "2. Pilih Kecamatan", ["-- Pilih Kecamatan --"] + kecamatan_list
  )

if selected_kecamatan != "-- Pilih Kecamatan --":
  df_filtered = df_kab[df_kab[col_kec] == selected_kecamatan]
else:
  df_filtered = df_kab

with col_f3:
  df_filtered["Display_Kios"] = (
      df_filtered[col_kios_code].astype(str)
      + " - "
      + df_filtered[col_kios_name].astype(str)
  )
  kios_list = sorted(df_filtered["Display_Kios"].dropna().unique().tolist())
  selected_display_kios = st.selectbox(
      "3. Pilih Kode Kios", ["-- Pilih Kios --"] + kios_list
  )

st.markdown("---")

# Inisialisasi Session State untuk status verifikasi
if "verifikasi_status" not in st.session_state:
  st.session_state.verifikasi_status = {}

if selected_display_kios != "-- Pilih Kios --":
  selected_kode_kios = selected_display_kios.split(" - ")[0]
  df_kios = df_filtered[
      df_filtered[col_kios_code].astype(str) == selected_kode_kios
  ].reset_index(drop=True)

  if len(df_kios) == 0:
    st.warning("Tidak ada data transaksi untuk kios ini.")
  else:
    kios_indices = df_kios.index.tolist()
    cek_count = sum(
        1
        for idx in kios_indices
        if (selected_kabupaten, selected_kode_kios, idx)
        in st.session_state.verifikasi_status
    )
    diterima_count = sum(
        1
        for idx in kios_indices
        if st.session_state.verifikasi_status.get(
            (selected_kabupaten, selected_kode_kios, idx)
        )
        == "Diterima"
    )
    ditolak_count = sum(
        1
        for idx in kios_indices
        if st.session_state.verifikasi_status.get(
            (selected_kabupaten, selected_kode_kios, idx)
        )
        == "Ditolak"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nota di Kios", len(df_kios))
    m2.metric("Sudah Diverifikasi", cek_count)
    m3.metric("Diterima", diterima_count)
    m4.metric("Ditolak", ditolak_count)

    st.markdown("### 📄 Daftar Nota & Eksekusi Verifikasi")

    if "current_index" not in st.session_state:
      st.session_state.current_index = 0
    if (
        "last_kios_selected" not in st.session_state
        or st.session_state.last_kios_selected != selected_kode_kios
    ):
      st.session_state.current_index = 0
      st.session_state.last_kios_selected = selected_kode_kios

    if st.session_state.current_index >= len(df_kios):
      st.session_state.current_index = len(df_kios) - 1

    idx_row = st.session_state.current_index
    row_data = df_kios.iloc[idx_row]
    key_state = (selected_kabupaten, selected_kode_kios, idx_row)
    current_status = st.session_state.verifikasi_status.get(
        key_state, "Belum Dicek"
    )

    # Tombol Navigasi Atas
    c_nav1, c_nav2, c_nav3 = st.columns([1, 2, 1])
    with c_nav1:
      if st.button("⬅️ Sebelumnya"):
        if st.session_state.current_index > 0:
          st.session_state.current_index -= 1
          st.rerun()
    with c_nav2:
      st.markdown(
          f"<p style='text-align: center; font-weight: bold;'>Menampilkan Nota ke-"
          f" {idx_row + 1} dari {len(df_kios)}</p>",
          unsafe_allow_html=True,
      )
    with c_nav3:
      if st.button("Selanjutnya ➡️"):
        if st.session_state.current_index < len(df_kios) - 1:
          st.session_state.current_index += 1
          st.rerun()

    st.markdown("---")

    col_det, col_prev = st.columns([1, 1])

    with col_det:
      trx_val = row_data.get(col_trx, "-") if col_trx else "-"
      petani_val = row_data.get(col_petani, "-") if col_petani else "-"
      nik_val = row_data.get("NIK", "-")
      tgl_val = row_data.get("Tanggal Tebus", "-")

      st.markdown(f"**No Transaksi:** `{trx_val}`")
      st.markdown(f"**Nama Petani:** {petani_val}")
      st.markdown(f"**NIK:** {nik_val}")
      st.markdown(f"**Tanggal Tebus:** {tgl_val}")

      pupuk_info = []
      for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
        if p in df_kios.columns and pd.notna(row_data.get(p)):
          pupuk_info.append(f"{p}: **{row_data.get(p)} kg**")
      if pupuk_info:
        st.markdown(f"🌾 **Alokasi:** {' | '.join(pupuk_info)}")

      st.markdown(f"Status Saat Ini: **{current_status}**")

      st.markdown("#### Tombol Aksi Verifikasi:")
      b_1, b_2, b_3 = st.columns(3)
      with b_1:
        if st.button("✅ TERIMA", type="primary"):
          st.session_state.verifikasi_status[key_state] = "Diterima"
          if st.session_state.current_index < len(df_kios) - 1:
            st.session_state.current_index += 1
          st.rerun()
      with b_2:
        if st.button("❌ TOLAK"):
          st.session_state.verifikasi_status[key_state] = "Ditolak"
          if st.session_state.current_index < len(df_kios) - 1:
            st.session_state.current_index += 1
          st.rerun()
      with b_3:
        if st.button("🔄 Reset"):
          if key_state in st.session_state.verifikasi_status:
            del st.session_state.verifikasi_status[key_state]
          st.rerun()

    with col_prev:
      url_col = find_col(df_kios, ["url", "bukti"])
      nota_url = row_data.get(url_col, "#") if url_col else "#"

      st.markdown("#### 🖼️ Preview Bukti Nota:")
      if pd.notna(nota_url) and str(nota_url).startswith("http"):
        try:
          st.iframe(nota_url, height=450, scrolling=True)
          st.markdown(
              f"🔗 [Buka Link Asli di Tab Baru]({nota_url}) (Jika preview"
              " diblokir browser)"
          )
        except Exception:
          st.markdown(f"🔗 **[Buka Dokumen Nota]({nota_url})**")
      else:
        st.warning("Link dokumen / gambar nota tidak tersedia.")

  st.markdown("---")
  st.subheader("📥 Download Hasil Verifikasi")
  if st.button("📊 Siapkan File Excel Hasil Verifikasi", type="primary"):
    with open(EXCEL_FILE, "rb") as f:
      excel_bytes = f.read()
    st.download_button(
        label="⬇️ Download File Excel",
        data=excel_bytes,
        file_name=f"Hasil_Verifikasi_{selected_kabupaten}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
else:
  st.info(
      "👆 Silakan pilih Kabupaten, Kecamatan, dan Kode Kios di atas untuk"
      " mulai memverifikasi nota."
  )
