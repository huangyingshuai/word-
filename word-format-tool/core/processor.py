import os
import tempfile
import time
import gc
import copy
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

from config.constants import (
    ALIGN_MAP, LINE_TYPE_MAP, FONT_SIZE_NUM, NUMBER_EN_PATTERN
)
from core.title_recognizer import get_title_level
from utils.font_utils import set_run_font, set_en_number_font
from utils.file_utils import safe_remove_file

def is_protected_para(para):
    """绝对保护机制：包含图片/分页符的段落完全不修改"""
    if not para:
        return True
    try:
        if para.paragraph_format.page_break_before:
            return True
        if para._element.find('.//w:sectPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        
        for run in para.runs:
            if run.contains_page_break:
                return True
            if run._element.xpath('.//w:drawing | .//w:pict | .//w:shape | .//w:oleObject', namespaces=run._element.nsmap):
                return True
        
        return False
    except Exception:
        return True

def process_number_in_para(para, body_font, body_size, number_config):
    """正文数字/英文单独设置，保留原有格式"""
    if not number_config["enable"]:
        for run in para.runs:
            set_run_font(run, body_font, body_size)
        return
    
    number_size = FONT_SIZE_NUM[number_config["size"]] if not number_config["size_same_as_body"] else body_size
    number_font = number_config["font"]
    number_bold = number_config["bold"]
    
    for run in para.runs:
        run_text = run.text
        if not run_text:
            continue
        
        if NUMBER_EN_PATTERN.fullmatch(run_text):
            set_en_number_font(run, number_font, number_size, number_bold)
        elif NUMBER_EN_PATTERN.search(run_text):
            original_format = {
                "font": run.font.name,
                "size": run.font.size.pt if run.font.size else body_size,
                "bold": run.font.bold,
                "italic": run.font.italic,
                "underline": run.font.underline
            }
            
            run.text = ""
            parts = []
            last_end = 0
            
            for match in NUMBER_EN_PATTERN.finditer(run_text):
                start, end = match.span()
                if start > last_end:
                    parts.append(("text", run_text[last_end:start]))
                parts.append(("number", run_text[start:end]))
                last_end = end
            if last_end < len(run_text):
                parts.append(("text", run_text[last_end:]))
            
            for part_type, part_text in parts:
                new_run = para.add_run(part_text)
                new_run.font.italic = original_format["italic"]
                new_run.font.underline = original_format["underline"]
                if part_type == "text":
                    set_run_font(new_run, body_font, body_size, original_format["bold"])
                else:
                    set_en_number_font(new_run, number_font, number_size, number_bold if number_bold is not None else original_format["bold"])
        else:
            set_run_font(run, body_font, body_size)

def process_doc(uploaded_file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    """核心排版逻辑，全链路异常防护"""
    start_time = time.time()
    tmp_path = None
    output_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            doc = Document(tmp_path)
            if len(doc.paragraphs) == 0:
                raise Exception("文档为空，没有可处理的内容")
        except Exception as e:
            raise Exception(f"文档打开失败，可能是文件损坏或格式不支持：{str(e)}")
        
        stats = {"一级标题":0,"二级标题":0,"三级标题":0,"正文":0,"表格":0,"图片":0}
        title_records = []

        original_image_count = 0
        for para in doc.paragraphs:
            try:
                original_image_count += len(para._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
                original_image_count += len(para._element.findall('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except:
                pass
        stats["图片"] = original_image_count

        last_levels = [0, 0, 0]
        for para_idx, para in enumerate(doc.paragraphs):
            if is_protected_para(para):
                continue

            text = para.text.strip()
            if not text and not clear_blank:
                continue

            try:
                level = get_title_level(para, enable_title_regex, last_levels)
            except Exception:
                level = "正文"
            
            stats[level] += 1
            title_records.append({
                "段落序号": para_idx,
                "识别结果": level,
                "文本内容": text[:50] + "..." if len(text) > 50 else text
            })

            if level == "一级标题":
                last_levels = [last_levels[0] + 1, 0, 0]
            elif level == "二级标题":
                last_levels[1] += 1
                last_levels[2] = 0
            elif level == "三级标题":
                last_levels[2] += 1

            if force_style:
                try:
                    if level == "一级标题":
                        para.style = doc.styles["标题 1"] if "标题 1" in doc.styles else doc.styles["Heading 1"]
                    elif level == "二级标题":
                        para.style = doc.styles["标题 2"] if "标题 2" in doc.styles else doc.styles["Heading 2"]
                    elif level == "三级标题":
                        para.style = doc.styles["标题 3"] if "标题 3" in doc.styles else doc.styles["Heading 3"]
                    else:
                        para.style = doc.styles["正文"] if "正文" in doc.styles else doc.styles["Normal"]
                except Exception:
                    pass

            cfg = config[level]
            font_size = FONT_SIZE_NUM[cfg["size"]]

            try:
                if ALIGN_MAP[cfg["align"]] is not None:
                    para.alignment = ALIGN_MAP[cfg["align"]]
                para.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                if cfg["line_type"] == "多倍行距":
                    para.paragraph_format.line_spacing = cfg["line_value"]
                elif cfg["line_type"] == "固定值":
                    para.paragraph_format.line_spacing = Pt(cfg["line_value"])
                if not keep_spacing:
                    para.paragraph_format.space_before = Pt(cfg.get("space_before", 0))
                    para.paragraph_format.space_after = Pt(cfg.get("space_after", 0))
                if level == "正文" and cfg["indent"] > 0:
                    para.paragraph_format.first_line_indent = Cm(cfg["indent"] * 0.37)
            except Exception:
                continue

            try:
                if level == "正文":
                    process_number_in_para(para, cfg["font"], font_size, number_config)
                else:
                    for run in para.runs:
                        set_run_font(run, cfg["font"], font_size, cfg["bold"])
            except Exception:
                continue

        for table in doc.tables:
            stats["表格"] += 1
            cfg = config["表格"]
            font_size = FONT_SIZE_NUM[cfg["size"]]
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if is_protected_para(p):
                            continue
                        try:
                            if ALIGN_MAP[cfg["align"]] is not None:
                                p.alignment = ALIGN_MAP[cfg["align"]]
                            p.paragraph_format.line_spacing_rule = LINE_TYPE_MAP[cfg["line_type"]]
                            if cfg["line_type"] == "多倍行距":
                                p.paragraph_format.line_spacing = cfg["line_value"]
                            elif cfg["line_type"] == "固定值":
                                p.paragraph_format.line_spacing = Pt(cfg["line_value"])
                        except Exception:
                            continue
                        try:
                            for run in p.runs:
                                set_run_font(run, cfg["font"], font_size, cfg["bold"])
                        except Exception:
                            continue

        if clear_blank:
            paras = list(doc.paragraphs)
            blank_count = 0
            for p in reversed(paras):
                if is_protected_para(p):
                    blank_count = 0
                    continue
                if not p.text.strip():
                    blank_count +=1
                    if blank_count > max_blank:
                        p._element.getparent().remove(p._element)
                else:
                    blank_count =0

        final_image_count = 0
        for para in doc.paragraphs:
            try:
                final_image_count += len(para._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
                final_image_count += len(para._element.findall('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
            except:
                pass

        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)
        with open(output_path, "rb") as f:
            file_bytes = f.read()
        
        end_time = time.time()
        process_time = end_time - start_time
        
        return file_bytes, stats, process_time, title_records

    except Exception as e:
        raise Exception(f"文档处理失败：{str(e)}")
    finally:
        safe_remove_file(tmp_path)
        safe_remove_file(output_path)
        gc.collect()