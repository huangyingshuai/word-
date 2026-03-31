from docx.shared import Pt
from docx.oxml.ns import qn

def set_run_font(run, font_name, font_size, bold=None):
    """中文字体100%生效"""
    try:
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

def set_en_number_font(run, font_name, font_size, bold=None):
    """数字/英文字体单独设置"""
    try:
        if font_name == "和正文一致":
            return
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run._element.rPr.rFonts.set(qn('w:cs'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass