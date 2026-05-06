import streamlit as st
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Docx Generator", layout="centered")

st.title("Isi Surat Otomatis")

# Input Nama dan Nomor
nama_user = st.text_input("Nama Lengkap")
nomor_surat = st.text_input("Nomor Surat")

st.divider()

# Bagian Lampiran Dinamis
st.subheader("Lampiran")
jumlah_lampiran = st.number_input("Banyaknya Lampiran", min_value=1, value=1, step=1)

data_lampiran = []
for i in range(int(jumlah_lampiran)):
    isi_ket = st.text_input(f"Isi Lampiran ke-{i+1}", key=f"input_{i}")
    data_lampiran.append({'l_ket': isi_ket})

if st.button("Proses dan Unduh Docx"):
    if not nama_user or not nomor_surat:
        st.warning("Mohon isi Nama dan Nomor terlebih dahulu.")
    else:
        try:
            # Panggil file test.docx
            doc = DocxTemplate("test.docx")
            
            # Data yang dikirim ke Word
            context = {
                'nama': nama_user,
                'nomor': nomor_surat,
                'lampiran_table': data_lampiran
            }
            
            # Proses pengisian
            doc.render(context)
            
            # Simpan ke memori
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.success("Selesai! Klik tombol di bawah untuk mengunduh.")
            st.download_button(
                label="Download File .docx",
                data=bio,
                file_name=f"Surat_{nama_user}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Gagal memproses dokumen: {e}")
