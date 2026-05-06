import streamlit as st
from docxtpl import DocxTemplate
import io

st.title("Generator Surat Otomatis")

# Input Nama dan Nomor[cite: 1]
nama = st.text_input("Nama:")
nomor = st.text_input("Nomor:")

st.subheader("Daftar Lampiran")
# Input jumlah lampiran agar baris bertambah ke bawah[cite: 1]
jumlah = st.number_input("Jumlah baris lampiran", min_value=1, value=1)

items = []
for i in range(int(jumlah)):
    ket = st.text_input(f"Isi Lampiran {i+1}", key=f"lp_{i}")
    items.append({'l_ket': ket})

if st.button("Buat Docx"):
    try:
        doc = DocxTemplate("test.docx")
        
        # Data yang dikirim harus sama dengan tag di {{ }} dan {% %}
        context = {
            'nama': nama,
            'nomor': nomor,
            'lampiran_table': items
        }
        
        doc.render(context)
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("Berhasil!")
        st.download_button(
            label="Download File",
            data=buffer,
            file_name="hasil_surat.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Gagal: {e}")
