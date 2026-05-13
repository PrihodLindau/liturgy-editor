import streamlit as st
from docx import Document
import io

st.title("Литургический редактор - Шаг 1")

uploaded = st.file_uploader("Загрузите DOCX", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    
    # Находим заголовок с датой и "Божественную литургию"
    date_index = None
    liturgy_index = None
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower()
        # Ищем дату (цифры и точки, например "23.02.2025" или "25 марта")
        if re.search(r'\d{1,2}[\.\s]\d{1,2}[\.\s]\d{2,4}', text) or any(month in text for month in ['января','февраля','марта','апреля']):
            if date_index is None:
                date_index = i
        if "божественная литургия" in text:
            liturgy_index = i
            break
    
    if date_index is not None and liturgy_index is not None and date_index < liturgy_index:
        # Вырезаем блок между датой и литургией
        block_to_move = []
        for i in range(date_index, liturgy_index):
            block_to_move.append(doc.paragraphs[i].text)
            # Удаляем оригинал
            p = doc.paragraphs[i]
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
        
        # Вставляем в конец документа
        for text in block_to_move:
            doc.add_paragraph(text)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
    st.success("Шаг 1 выполнен: блок перенесён вниз")
    st.download_button("Скачать", output, "step1.docx")
