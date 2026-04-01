from docx import Document
from io import BytesIO

def get_doc_from_uploaded(uploaded_file):
    """读取上传的Word文件"""
    return Document(BytesIO(uploaded_file.getvalue()))
