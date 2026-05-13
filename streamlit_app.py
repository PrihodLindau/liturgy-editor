import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="Литургический редактор", layout="wide")
st.title("📖 Литургический редактор")

def process_document(doc):
    """Полная обработка документа"""
    
    # 1. Обрезаем до Божественной литургии
    for i, para in enumerate(doc.paragraphs):
        if "божественная литургия" in para.text.lower():
            for j in range(i-1, -1, -1):
                p = doc.paragraphs[j]
                if p._element.getparent() is not None:
                    p._element.getparent().remove(p._element)
            break
    
    # 2. Собираем стихиры из утрени
    in_matins = False
    stichera = []
    
    for para in doc.paragraphs:
        text = para.text.lower()
        
        if "утреня" in text:
            in_matins = True
            continue
        
        if in_matins and "канон" in text:
            in_matins = False
        
        if in_matins and "стихир" in text:
            if not any(x in text for x in ["воскресн", "воскресная"]):
                stichera.append(para)
    
    # 3. Оставляем нужные разделы
    keep_sections = ["тропари", "кондаки", "прокимны", "аллилуиа", "причастный"]
    keep = False
    to_remove = []
    
    for para in doc.paragraphs:
        text = para.text.lower()
        
        if any(s in text for s in keep_sections):
            keep = True
        
        if any(s in text for s in ["апостол", "евангелие", "отпуст"]):
            keep = False
        
        if not keep and text.strip():
            to_remove.append(para)
    
    for para in to_remove:
        if para._element.getparent() is not None:
            para._element.getparent().remove(para._element)
    
    # 4. Вставляем запричастные
    if stichera:
        for i, para in enumerate(doc.paragraphs):
            if "причастный" in para.text.lower():
                # Добавляем заголовок
                title = doc.paragraphs[i].insert_paragraph_before()
                title.add_run("ЗАПРИЧАСТНЫЕ (стихиры праздника)").bold = True
                
                # Добавляем стихиры
                for stih in stichera:
                    new_p = doc.paragraphs[i+1].insert_paragraph_before()
                    new_p.text = stih.text
                break
    
    return doc

uploaded = st.file_uploader("Загрузите DOCX файл", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    
    with st.spinner("Обработка..."):
        doc = process_document(doc)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
    st.success("✅ Обработка завершена!")
    st.download_button("📥 Скачать", output, "processed.docx", use_container_width=True)
