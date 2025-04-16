import streamlit as st
from docx import Document
import re
from io import BytesIO
import os

def clean_docx(file):
    from docx.text.run import Run

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

    def rebuild_runs(paragraph):
        # Build full text with formatting info per run
        parts = [(run.text, run.bold, run.italic, run.underline, run.style) for run in paragraph.runs]
        combined_text = ''.join(text for text, *_ in parts)
        cleaned_text = clean_text(combined_text)

        # Remove existing runs
        for _ in range(len(paragraph.runs)):
            paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)

        # Rebuild runs by mapping cleaned text back to original styles
        i = 0
        for text, bold, italic, underline, style in parts:
            for char in text:
                if i >= len(cleaned_text):
                    break
                if cleaned_text[i] == char:
                    new_run = paragraph.add_run(char)
                    new_run.bold = bold
                    new_run.italic = italic
                    new_run.underline = underline
                    new_run.style = style
                    i += 1
                else:
                    # Skip characters that were removed
                    while i < len(cleaned_text) and cleaned_text[i] != char:
                        i += 1

    # Clean paragraphs
    for para in doc.paragraphs:
        rebuild_runs(para)

    # Clean tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    rebuild_runs(para)

    # Remove comment indicators
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
