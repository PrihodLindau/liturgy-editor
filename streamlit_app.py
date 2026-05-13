import streamlit as st
from docx import Document
import io

st.title("Литургический редактор")

uploaded = st.file_uploader("Загрузите DOCX", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    st.success("Готово")
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
    st.download_button("Скачать", output, "result.docx")
