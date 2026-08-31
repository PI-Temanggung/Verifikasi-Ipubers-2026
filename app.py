import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Verifikasi Nota IPubers", page_icon="📝", layout="wide"
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

st.title("📝 Panel Verifikasi Nota Cepat IPubers")
st.markdown(
    "Pilih wilayah dan kios di bawah, lalu langsung verifikasi nota satu per"
    " satu dengan sekali klik."
)

# --- FILTER DI BAGIAN ATAS (UTAMA) ---
col_f1, col_f2, col_f3 = st.columns(3)

with col_f1:
  selected_kabupaten = st.selectbox(
      "1. Pilih Kabupaten / Wilayah", kabupaten_list
  )


# Memuat data sheet terpilih
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

with col_f2:
  kecamatan_list = sorted(df_kab[col_kec].dropna().unique().tolist())
  selected_kecamatan = st.selectbox(
      "2. Pilih Kecamatan", ["-- Semua Kecamatan --"] + kecamatan_list
  )

if selected_kecamatan != "-- Semua Kecamatan --":
  df_filtered = df_kab[df_kab[col_kec] == selected_kecamatan]
else:
  df_filtered = df_kab

with col_f3:
  kios_list = sorted(df_filtered[col_kios].dropna().unique().tolist())
  selected_kios = st.selectbox(
      "3. Pilih Nama Kios", ["-- Pilih Kios --"] + kios_list
  )

st.markdown("---")

# Inisialisasi Session State untuk status verifikasi
if "verifikasi_status" not in st.session_state:
  st.session_state.verifikasi_status = {}

if selected_kios != "-- Pilih Kios --":
  df_kios = df_filtered[df_filtered[col_kios] == selected_kios].reset_index(
      drop=True
  )

  if len(df_kios) == 0:
    st.warning("Tidak ada data transaksi untuk kios ini.")
  else:
    # Metrik ringkas
    kios_indices = df_kios.index.tolist()
    cek_count = sum(
        1
        for idx in kios_indices
        if (selected_kabupaten, selected_kios, idx)
        in st.session_state.verifikasi_status
    )
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

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nota di Kios", len(df_kios))
    m2.metric("Sudah Diverifikasi", cek_count)
    m3.metric("Diterima", diterima_count)
    m4.metric("Ditolak", ditolak_count)

    st.markdown("### 🔍 Daftar Nota & Verifikasi Langsung")
    st.info(
        "Setiap baris di bawah langsung menampilkan detail nota, link dokumen,"
        " dan tombol aksi."
    )

    # Looping menampilkan setiap nota secara langsung tanpa navigasi halaman
    for idx, row_data in df_kios.iterrows():
      key_state = (selected_kabupaten, selected_kios, idx)
      current_status = st.session_state.verifikasi_status.get(
          key_state, "Belum Dicek"
      )

      # Warna border card berdasarkan status
      border_color = "#ccc"
      if current_status == "Diterima":
        border_color = "#28a745"
      elif current_status == "Ditolak":
        border_color = "#dc3545"

      with st.container():
        st.markdown(
            f"""
                <div style="padding: 15px; border: 2px solid {border_color}; border-radius: 8px; margin-bottom: 15px; background-color: #f9f9f9;">
                    <h4 style="margin: 0 0 10px 0;">Nota #{idx + 1} — Status: <b>{current_status}</b></h4>
                </div>
                """,
            unsafe_allow_html=True,
        )

        col_info, col_action = st.columns([2, 1])

        with col_info:
          trx_val = row_data.get(col_trx, "-") if col_trx else "-"
          petani_val = row_data.get(col_petani, "-") if col_petani else "-"
          nik_val = row_data.get("NIK", "-")
          tgl_val = row_data.get("Tanggal Tebus", "-")

          st.markdown(f"**No Transaksi:** `{trx_val}`")
          st.markdown(f"**Nama Petani:** {petani_val} (NIK: {nik_val})")
          st.markdown(f"**Tanggal Tebus:** {tgl_val}")

          # Info Pupuk
          pupuk_info = []
          for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
            if p in df_kios.columns and pd.notna(row_data.get(p)):
              pupuk_info.append(f"{p}: **{row_data.get(p)} kg**")
          if pupuk_info:
            st.markdown(f"🌾 Alokasi: {' | '.join(pupuk_info)}")

          # Link / Bukti Nota
          url_col = find_col(df_kios, ["url", "bukti"])
          nota_url = row_data.get(url_col, "#") if url_col else "#"
          if pd.notna(nota_url) and str(nota_url).startswith("http"):
            st.markdown(
                f"🔗 **[Buka Dokumen Nota / Bukti di Tab Baru]({nota_url})**"
            )
          else:
            st.caption("Link dokumen nota tidak tersedia.")

        with col_action:
          st.markdown("**Aksi Verifikasi:**")
          btn_col1, btn_col2 = st.columns(2)

          with btn_col1:
            if st.button("✅ Terima", key=f"terima_{selected_kios}_{idx}"):
              st.session_state.verifikasi_status[key_state] = "Diterima"
              st.rerun()

          with btn_col2:
            if st.button("❌ Tolak", key=f"tolak_{selected_kios}_{idx}"):
              st.session_state.verifikasi_status[key_state] = "Ditolak"
              st.rerun()

          if current_status != "Belum Dicek":
            if st.button("🔄 Reset", key=f"reset_{selected_kios}_{idx}"):
              if key_state in st.session_state.verifikasi_status:
                del st.session_state.verifikasi_status[key_state]
              st.rerun()

        st.markdown("---")

  # Tombol Download Rekap
  st.subheader("📥 Download Rekap Hasil Verifikasi")
  if st.button("📊 Siapkan File Download", type="primary"):
    st.success("Data siap diunduh!")
    with open(EXCEL_FILE, "rb") as f:
      excel_bytes = f.read()
    st.download_button(
        label="⬇️ Download Excel Hasil Verifikasi",
        data=excel_bytes,
        file_name=f"Hasil_Verifikasi_{selected_kabupaten}_{selected_kios}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
else:
  st.info("👆 Silakan pilih Kabupaten, Kecamatan, dan Kios di atas terlebih dahulu.")
