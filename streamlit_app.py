import streamlit as st
from docx import Document
import io

st.title("Литургический редактор")

uploaded = st.file_uploader("Загрузите DOCX", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    
    # Показываем текст для выбора
    st.write("### Найдите в тексте:")
    
    # Отображаем все параграфы для выбора
    texts = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            texts.append(f"{i}: {text[:100]}...")
    
    start_idx = st.number_input("Введите номер параграфа с ДАТОЙ", min_value=0, max_value=len(texts)-1, step=1)
    end_idx = st.number_input("Введите номер параграфа с 'Божественная литургия'", min_value=0, max_value=len(texts)-1, step=1)
    
    if st.button("Перенести блок вниз"):
        # Вырезаем блок
        block = []
        for i in range(start_idx, end_idx + 1):
            block.append(doc.paragraphs[i].text)
        
        # Удаляем с конца в начало (чтобы не сбивались индексы)
        for i in range(end_idx, start_idx - 1, -1):
            p = doc.paragraphs[i]
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
        
        # Вставляем в конец
        for text in block:
            doc.add_paragraph(text)
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        st.success("Готово!")
        st.download_button("Скачать", output, "result.docx")
    
    # Показываем текст для удобства
    with st.expander("Показать весь текст с номерами"):
        for i, text in enumerate(texts):
            st.write(f"**{i}:** {text}")
