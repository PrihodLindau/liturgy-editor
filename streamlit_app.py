import streamlit as st
from docx import Document
import io
import re

st.title("Литургический редактор - Шаг 2")

uploaded = st.file_uploader("Загрузите DOCX", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    
    # ШАГ 1: Перенос блока от даты до литургии вниз
    date_index = None
    liturgy_index = None
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower()
        if re.search(r'\d{1,2}[\.\s]\d{1,2}[\.\s]\d{2,4}', text) or any(month in text for month in ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря']):
            if date_index is None:
                date_index = i
        if "божественная литургия" in text:
            liturgy_index = i
            break
    
    if date_index is not None and liturgy_index is not None and date_index < liturgy_index:
        block_to_move = []
        for i in range(date_index, liturgy_index):
            block_to_move.append(doc.paragraphs[i].text)
            p = doc.paragraphs[i]
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
        
        for text in block_to_move:
            doc.add_paragraph(text)
    
    # ШАГ 2: Оставляем только нужные разделы
    keep_sections = ["тропар", "кондак", "прокимен", "аллилуиа", "причаст"]
    sections_to_delete = ["апостол", "евангелие", "часы", "изобразительны", "отпуст"]
    
    keep = False
    to_remove = []
    
    for para in doc.paragraphs:
        text = para.text.lower()
        
        # Включаем отображение при нужных разделах
        if any(section in text for section in keep_sections):
            keep = True
        
        # Выключаем при служебных разделах
        if any(section in text for section in sections_to_delete):
            keep = False
        
        # Помечаем на удаление
        if not keep and text.strip():
            to_remove.append(para)
    
    for para in to_remove:
        if para._element.getparent() is not None:
            para._element.getparent().remove(para._element)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
    st.success("✅ Шаг 2 выполнен: оставлены только тропари, кондаки, прокимны, аллилуйя, причастный стих")
    st.download_button("📥 Скачать", output, "step2.docx")
