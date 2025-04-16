import streamlit as st
from docx import Document
import re
from io import BytesIO
import os

def clean_docx(file):
    doc = Document(file)

    # Define all tag patterns to remove
    tag_patterns = [r'\[H\d+\]', r'\[hed\]', r'\[dek\]']

    # Clean paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            for pattern in tag_patterns:
                run.text = re.sub(pattern, '', run.text, flags=re.IGNORECASE)

    # Clean tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for pattern in tag_patterns:
                            run.text = re.sub(pattern, '', run.text, flags=re.IGNORECASE)

    # Remove comments (most visible ones)
    doc_element = doc._element
    for comment in doc_element.xpath('//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference'):
        comment.getparent().remove(comment)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

st.title("🧼 DOCX Cleaner")

uploaded_file = st.file_uploader("Upload a .docx file", type="docx")

if uploaded_file:
    cleaned_file = clean_docx(uploaded_file)
    
    # Get original filename without extension
    base_filename = os.path.splitext(uploaded_file.name)[0]
    output_filename = f"{base_filename}_cleaned.docx"

    st.success("Cleaning complete!")
    st.download_button(
        label="Download cleaned .docx",
        data=cleaned_file,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
