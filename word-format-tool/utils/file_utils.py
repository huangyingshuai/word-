from docx import Document
from io import BytesIO

def get_doc_from_uploaded(uploaded_file):
    return Document(BytesIO(uploaded_file.getvalue()))
