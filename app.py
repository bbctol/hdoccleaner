import streamlit as st
from docx import Document
import re
from io import BytesIO

def clean_docx(file):
    doc = Document(file)

    # Remove [H<number>] from paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            run.text = re.sub(r'\[H\d+\]', '', run.text)

    # Remove [H<number>] from tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.text = re.sub(r'\[H\d+\]', '', run.text)

    # Remove comment elements (works for most cases)
    doc_element = doc._element
    for comment in doc_element.xpath('//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference'):
        comment.getparent().remove(comment)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

st.title("🧼 Clean Your .docx File")

uploaded_file = st.file_uploader("Upload a .docx file", type="docx")

if uploaded_file:
    cleaned_file = clean_docx(uploaded_file)
    st.success("Cleaning complete!")
    st.download_button("Download cleaned .docx", cleaned_file, file_name="cleaned_output.docx")
