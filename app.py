import streamlit as st
from docxtpl import DocxTemplate
import io

st.title("Generator Surat Otomatis")

# Input Nama dan Nomor
nama = st.text_input("Masukkan Nama")
nomor = st.text_input("Masukkan Nomor")

st.subheader("Isi Lampiran")
# Input jumlah lampiran
jumlah_lampiran = st.number_input("Jumlah Lampiran", min_value=1, value=1, step=1)

# Membuat form dinamis untuk lampiran
list_lampiran = []
for i in range(int(jumlah_lampiran)):
    col1, col2 = st.columns([1, 4])
    with col1:
        no_lamp = st.text_input(f"No {i+1}", value=str(i+1), key=f"no_{i}")
    with col2:
        ket_lamp = st.text_input(f"Keterangan {i+1}", key=f"ket_{i}")
    list_lampiran.append({'l_no': no_lamp, 'l_ket': ket_lamp})

if st.button("Generate Surat"):
    try:
        # Memuat template
        doc = DocxTemplate("test.docx")
        
        # Mapping data ke placeholder
        context = {
            'nama': nama,
            'nomor': nomor,
            'lampiran_table': list_lampiran # Pastikan di template tabel menggunakan perulangan Jinja2
        }
        
        # Render dokumen
        doc.render(context)
        
        # Simpan ke memory agar bisa diunduh
        target_stream = io.BytesIO()
        doc.save(target_stream)
        target_stream.seek(0)
        
        st.success("Dokumen berhasil dibuat!")
        st.download_button(
            label="Unduh Surat (.docx)",
            data=target_stream,
            file_name=f"Surat_{nama}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
