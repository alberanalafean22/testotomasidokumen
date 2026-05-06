
import streamlit as st
import pandas as pd
from docx import Document
import io
import requests

st.set_page_config(page_title="Generator SK BPS", layout="centered")

st.title("📝 Generator Dokumen SK BPS Kota Solok")
st.write("Aplikasi ini akan mengotomasi pengisian dokumen. Template DOK.docx diambil langsung dari GitHub.")

# --- KONFIGURASI GITHUB ---
# Ganti URL di bawah ini dengan URL "Raw" dari file DOK.docx di GitHub Anda
GITHUB_RAW_URL = "https://raw.githubusercontent.com/alberanalafean22/testotomasidokumen/main/DOK.docx"

# 1. Input Data Umum
st.subheader("1. Informasi Umum Dokumen")
col1, col2 = st.columns(2)

with col1:
    nomor = st.text_input("Nomor Dokumen", placeholder="Contoh: 123/BPS/2026")
    tanggal = st.text_input("Tanggal Ditetapkan", placeholder="Contoh: 15 Januari 2026")
    petugas = st.text_input("Petugas / Nama Tim", placeholder="Contoh: Tim Pengolah Data")

with col2:
    tentang = st.text_input("Tentang", placeholder="Contoh: TIM PENGOLAH DATA")
    pelaksanaan = st.text_input("Pelaksanaan Kegiatan", placeholder="Contoh: kegiatan Survei Sosial Ekonomi Nasional")

# 2. Input Data Tabel (Dinamis)
st.subheader("2. Data Petugas (Lampiran)")
st.write("Anda bisa menambah, menghapus, atau mengedit baris tabel di bawah ini. Tabel di Word akan otomatis menyesuaikan jumlah datanya.")

default_data = pd.DataFrame({
    "Nama/Jabatan": ["Alber Analafean"],
    "NIP/Golongan": ["199801012025011001 / III/a"],
    "Posisi": ["Ketua Tim"],
    "Honor": ["500.000"]
})

# Tabel interaktif
edited_df = st.data_editor(default_data, num_rows="dynamic", use_container_width=True)

# 3. Tombol Eksekusi
if st.button("Buat Dokumen", type="primary"):
    with st.spinner('Mengunduh template DOK.docx dari GitHub dan memproses dokumen...'):
        try:
            # Mengunduh file dari GitHub
            response = requests.get(GITHUB_RAW_URL)
            
            # Memastikan file berhasil diunduh (Status 200 OK)
            if response.status_code != 200:
                st.error(f"Gagal mengunduh template dari GitHub (Status Code: {response.status_code}). Pastikan URL Raw benar dan repositori bersifat Publik.")
            else:
                # Membaca template dokumen dari memori (hasil unduhan)
                template_file = io.BytesIO(response.content)
                doc = Document(template_file)
                
                # Dictionary untuk replace teks umum
                replacements = {
                    "{tentang}": tentang,
                    "{pelaksanaan}": pelaksanaan,
                    "{pelaksanan}": pelaksanaan, # Menghindari typo
                    "{petugas}": petugas,
                    "{tanggal}": tanggal,
                    "{nomor}": nomor
                }
                
                # Fungsi untuk me-replace teks di paragraf
                def replace_text_in_paragraphs(paragraphs):
                    for p in paragraphs:
                        for key, value in replacements.items():
                            if key in p.text:
                                p.text = p.text.replace(key, value)

                # Replace di paragraf utama
                replace_text_in_paragraphs(doc.paragraphs)
                
                # Replace di dalam tabel (kecuali baris data lampiran)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            replace_text_in_paragraphs(cell.paragraphs)

                # --- LOGIK KHUSUS UNTUK TABEL LAMPIRAN ---
                target_table = None
                template_row_idx = -1
                
                # Mencari tabel mana yang mengandung tag {nama}
                for table in doc.tables:
                    for i, row in enumerate(table.rows):
                        for cell in row.cells:
                            if "{nama}" in cell.text:
                                target_table = table
                                template_row_idx = i
                                break
                        if target_table: break
                    if target_table: break
                
                # Jika tabel lampiran ditemukan, masukkan data dataframe
                if target_table and template_row_idx != -1:
                    for index, row_data in edited_df.iterrows():
                        if index == 0:
                            target_row = target_table.rows[template_row_idx]
                        else:
                            target_row = target_table.add_row()
                        
                        # Mengisi sel
                        target_row.cells[0].text = f"{index + 1}."
                        target_row.cells[1].text = str(row_data["Nama/Jabatan"])
                        target_row.cells[2].text = str(row_data["NIP/Golongan"])
                        target_row.cells[3].text = str(row_data["Posisi"])
                        target_row.cells[4].text = f"Rp{row_data['Honor']}"

                # Simpan dokumen ke memory
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.success("✅ Dokumen berhasil dibuat!")
                
                # Tombol unduh
                st.download_button(
                    label="⬇️ Unduh Dokumen Hasil",
                    data=doc_io,
                    file_name=f"SK_{tentang.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except requests.exceptions.RequestException as e:
            st.error(f"Koneksi ke GitHub gagal: {e}")
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses dokumen: {e}")
