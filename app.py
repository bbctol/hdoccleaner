import streamlit as st
from docx import Document
import re
from io import BytesIO
import os

def clean_docx(file):
    doc = Document(file)

    # Patterns to remove
    tag_patterns = [
        r'\[H\d+\]\s*',
        r'\[hed\]\s*',
        r'\[dek\]\s*'
    ]

    def clean_text(text):
        for pattern in tag_patterns:
            text = re.sub(pattern, '', text, flags=re.IGNORECASE)
        return text

    # Clean paragraphs (process full text, not per-run)
    for para in doc.paragraphs:
        full_text = para.text
        cleaned_text = clean_text(full_text)
        if cleaned_text != full_text:
            # Remove existing runs
            for _ in range(len(para.runs)):
                para.runs[0]._element.getparent().remove(para.runs[0]._element)
            para.add_run(cleaned_text)

    # Clean table content similarly
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full_text = para.text
                    cleaned_text = clean_text(full_text)
                    if cleaned_text != full_text:
                        for _ in range(len(para.runs)):
                            para.runs[0]._element.getparent().remove(para.runs[0]._element)
                        para.add_run(cleaned_text)

    # Remove comments
    doc_element = doc._element
    for comment in doc_element.xpath('//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference'):
        comment.getparent().remove(comment)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Streamlit interface
st.title("🧼 DOCX Cleaner")

uploaded_file = st.file_uploader("Upload a .docx file", type="docx")

if uploaded_file:
    cleaned_file = clean_docx(uploaded_file)
    
    base_filename = os.path.splitext(uploaded_file.name)[0]
    output_filename = f"{base_filename}_cleaned.docx"

    st.success("Cleaning complete!")
    st.download_button(
        label="Download cleaned .docx",
        data=cleaned_file,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
