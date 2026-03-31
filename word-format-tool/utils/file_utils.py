import os
import tempfile
from docx import Document

def get_doc_from_uploaded(uploaded_file):
    """从上传的文件获取Document对象"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    doc = Document(tmp_path)
    os.unlink(tmp_path)
    return doc

def safe_remove_file(file_path):
    """安全删除文件"""
    if file_path and os.path.exists(file_path):
        try:
            os.unlink(file_path)
        except Exception:
            pass