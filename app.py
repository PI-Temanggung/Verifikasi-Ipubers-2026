import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Verifikasi Nota IPubers", page_icon="🔍", layout="wide"
)

EXCEL_FILE = "IPUBERS-AGUSTUS.xlsx"


# Memuat data secara utuh agar format asli tidak berubah saat di-download nanti
@st.cache_data
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
    "Pilih Kecamatan dan Kode Kios di bawah untuk memverifikasi nota satu"
    " per satu. Hasil verifikasi akan direkam tanpa mengubah format asli"
    " file Excel."
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

if not col_kec or not col_kios_code:
  st.error(
      "Kolom 'Kecamatan' atau 'Kode Kios' tidak ditemukan di dalam file Excel"
      f" Anda. Kolom yang ada: {cols}"
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

# Inisialisasi Session State untuk menyimpan status verifikasi
if "verifikasi_dict" not in st.session_state:
  st.session_state.verifikasi_dict = {}

if selected_display_kios != "-- Pilih Kios --":
  selected_kode_kios = selected_display_kios.split(" - ")[0]

  df_kios = df_filtered[
      df_filtered[col_kios_code].astype(str) == selected_kode_kios
  ]

  if len(df_kios) == 0:
    st.warning("Tidak ada data transaksi untuk kios ini.")
  else:
    indices = df_kios.index.tolist()

    if "current_pos" not in st.session_state:
      st.session_state.current_pos = 0
    if (
        "last_kios" not in st.session_state
        or st.session_state.last_kios != selected_kode_kios
    ):
      st.session_state.current_pos = 0
      st.session_state.last_kios = selected_kode_kios

    if st.session_state.current_pos >= len(indices):
      st.session_state.current_pos = len(indices) - 1

    pos = st.session_state.current_pos
    row_idx = indices[pos]
    row_data = df_original.loc[row_idx]

    # Statistik Ringkas
    total_nota = len(indices)
    sudah_cek = sum(
        1 for idx in indices if idx in st.session_state.verifikasi_dict
    )
    diterima = sum(
        1
        for idx in indices
        if st.session_state.verifikasi_dict.get(idx) == "TERIMA"
    )
    ditolak = sum(
        1
        for idx in indices
        if st.session_state.verifikasi_dict.get(idx) == "TOLAK"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nota Kios", total_nota)
    m2.metric("Sudah Diverifikasi", sudah_cek)
    m3.metric("Diterima", diterima)
    m4.metric("Ditolak", ditolak)

    st.markdown("---")

    # Layout Utama: Kiri (Detail & Aksi), Kanan (Preview Gambar Nota Langsung)
    col_kiri, col_kanan = st.columns([1, 1], gap="large")

    with col_kiri:
      st.subheader("📄 Detail Transaksi Nota")
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
        if p in df_original.columns and pd.notna(row_data.get(p)):
          pupuk_info.append(f"{p}: **{row_data.get(p)} kg**")
      if pupuk_info:
        st.markdown(f"🌾 **Alokasi Pupuk:** {' | '.join(pupuk_info)}")

      current_status = st.session_state.verifikasi_dict.get(
          row_idx, "Belum Dicek"
      )
      status_color = (
          "green"
          if current_status == "TERIMA"
          else ("red" if current_status == "TOLAK" else "orange")
      )
      st.markdown(
          f"Status Verifikasi Saat Ini: <span"
          f" style='color:{status_color}; font-weight:bold;'>{current_status}</span>",
          unsafe_allow_html=True,
      )

      st.markdown("#### Aksi Verifikasi:")
      b1, b2, b3 = st.columns(3)
      with b1:
        if st.button("✅ TERIMA", type="primary", use_container_width=True):
          st.session_state.verifikasi_dict[row_idx] = "TERIMA"
          if pos < len(indices) - 1:
            st.session_state.current_pos += 1
          st.rerun()
      with b2:
        if st.button("❌ TOLAK", use_container_width=True):
          st.session_state.verifikasi_dict[row_idx] = "TOLAK"
          if pos < len(indices) - 1:
            st.session_state.current_pos += 1
          st.rerun()
      with b3:
        if st.button("🔄 Reset", use_container_width=True):
          if row_idx in st.session_state.verifikasi_dict:
            del st.session_state.verifikasi_dict[row_idx]
          st.rerun()

      st.markdown("---")
      st.markdown("#### Navigasi Nota:")
      nav_prev, nav_info, nav_next = st.columns([1, 2, 1])
      with nav_prev:
        if st.button("⬅️ Sebelumnya", use_container_width=True):
          if pos > 0:
            st.session_state.current_pos -= 1
            st.rerun()
      with nav_info:
        st.markdown(
            f"<p style='text-align: center; font-weight: bold; margin-top:"
            f" 8px;'>Nota ke-{pos + 1} dari {len(indices)}</p>",
            unsafe_allow_html=True,
        )
      with nav_next:
        if st.button("Selanjutnya ➡️", use_container_width=True):
          if pos < len(indices) - 1:
            st.session_state.current_pos += 1
            st.rerun()

    with col_kanan:
      st.subheader("🖼️ Preview Gambar Nota")
      nota_url = row_data.get(col_url, None) if col_url else None

      if pd.notna(nota_url) and str(nota_url).startswith("http"):
        # Menampilkan gambar langsung di halaman menggunakan tag HTML img & st.image
        try:
          st.image(
              str(nota_url),
              caption=f"Nota Transaksi: {trx_val}",
              use_container_width=True,
          )
        except Exception:
          # Fallback jika st.image gagal memuat URL langsung
          st.markdown(
              f'<img src="{nota_url}" width="100%"'
              ' style="border-radius:8px; border:1px solid #ddd;" />',
              unsafe_allow_html=True,
          )

        st.markdown(
            f"🔗 [Buka Link Asli di Tab Baru]({nota_url}) (Jika gambar gagal"
            " dimuat browser)",
            help="Link eksternal asli",
        )
      else:
        st.warning(
            "Link atau URL bukti nota tidak tersedia pada baris data ini."
        )

  # --- TOMBOL DOWNLOAD HASIL ---
  st.markdown("---")
  st.subheader("📥 Download File Excel Hasil Pengecekan")
  st.markdown(
      "File hasil download mempertahankan seluruh format asli Excel Anda dan"
      " hanya menambahkan kolom **Status_Verifikasi** di bagian paling ujung."
  )

  if st.button("📊 Siapkan File Excel untuk Di-download", type="primary"):
    df_export = df_original.copy()

    status_list = []
    for idx, _ in df_export.iterrows():
      if idx in st.session_state.verifikasi_dict:
        status_list.append(st.session_state.verifikasi_dict[idx])
      else:
        status_list.append("Belum Dicek")

    df_export["Status_Verifikasi"] = status_list

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_export.to_excel(writer, sheet_name=active_sheet, index=False)
    processed_data = output.getvalue()

    st.download_button(
        label="⬇️ Download Excel Hasil Verifikasi",
        data=processed_data,
        file_name=f"Hasil_Verifikasi_{selected_kecamatan}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
else:
  st.info(
      "👆 Silakan pilih Kecamatan dan Kode Kios di atas untuk mulai melakukan"
      " pengecekan nota."
  )
