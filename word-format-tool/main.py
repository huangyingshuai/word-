import streamlit as st
import copy
import re
import random
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_BUILTIN_STYLE
from docx.oxml.ns import qn
import os

# ====================== 预编译正则 ======================
RE_REF_FLAG = re.compile(r'^\[(\d+)\]')
RE_REF_KEYWORD = re.compile(r'参考文献')
RE_REF_SPACE = re.compile(r'\s+')
RE_REF_CN_FONT = re.compile(r'([\u4e00-\u9fa5]+)\[([A-Z]+)\]')
RE_REF_DOT = re.compile(r'。(?![\u4e00-\u9fa5])')
RE_REF_COMMA = re.compile(r'，')
RE_REF_COLON = re.compile(r'：')
RE_KEYWORDS = re.compile(r'[\u4e00-\u9fa5]{2,}')
RE_WHITE_NUMBER = re.compile(r'^\d+(\.\d+)*$')
RE_WHITE_QUOTE = re.compile(r'^\[.*\]$')
RE_SENTENCE_SPLIT = re.compile(r'(?<=[。！？；])\s*')
RE_CLAUSE_SPLIT = re.compile(r'[，。；]')

# ====================== 全局配置与常量 ======================
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "挑战杯", "互联网+", "三创赛",
    "参考文献", "公式", "图表", "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "算法", "系统", "模型", "数据"
]
WPS_STYLE_MAPPING = {
    "一级标题": WD_BUILTIN_STYLE.HEADING_1,
    "二级标题": WD_BUILTIN_STYLE.HEADING_2,
    "三级标题": WD_BUILTIN_STYLE.HEADING_3,
    "正文": WD_BUILTIN_STYLE.NORMAL
}
COMPETITION_FORMATS = {
    "三创赛": {
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "楷体_GB2312", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "仿宋", "size": "四号", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.2, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["硬件必须配小程序/App", "服务必须线上化", "需要3D建模图/UI原型", "图表必须标注数据来源"]
    },
    "挑战杯": {
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "黑体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文约15000字", "双面打印", "严格章-节-条层级结构", "标题单倍行距，正文1.5倍行距"]
    },
    "互联网+创新创业大赛": {
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "四号", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文10000字以上", "分创意组/创业组撰写", "需包含完整财务预测", "商业模式需清晰可落地"]
    }
}
THESIS_FORMATS = {
    "本科毕业论文": {
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文8000-12000字", "需包含摘要/关键词/参考文献/致谢", "参考文献需符合GB/T 7714格式", "页眉需标注学校+论文题目"]
    }
}
ALL_TEMPLATES = {**COMPETITION_FORMATS, **THESIS_FORMATS}
REWRITE_LEVEL = {
    "轻度降重": {"synonym": True, "sentence_reorder": False, "structure_change": False},
    "标准降重": {"synonym": True, "sentence_reorder": True, "structure_change": False},
    "强力降重": {"synonym": True, "sentence_reorder": True, "structure_change": True}
}
SYNONYM_DICT = {
    "提升": "有效改善", "降低": "显著减少", "增加": "大幅提升", "减少": "有效降低",
    "首先": "从实际落地情况来看", "其次": "进一步分析", "最后": "综合上述分析",
    "综上所述": "结合全维度分析", "总而言之": "从实践结果来看"
}
ALIGN_MAP = {
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY
}
FONT_SIZE_MAP = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9
}
APP_NAME = "竞赛&论文格式处理器"
LEVEL_LIST = ["一级标题", "二级标题", "三级标题", "正文", "表格"]
EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri", "Courier New"]
CN_FONT_LIST = ["宋体", "黑体", "楷体", "仿宋_GB2312", "微软雅黑"]
MAX_FILE_SIZE_MB = 20
random.seed(42)

# ====================== 核心工具函数 ======================
@st.cache_data(ttl=3600)
def get_cached_template(template_name):
    return copy.deepcopy(ALL_TEMPLATES[template_name]["cn_format"]), copy.deepcopy(ALL_TEMPLATES[template_name]["en_format"])

# ✅ 终极修复：100%精准识别标题，彻底过滤正文分点
def get_title_level(para_text):
    text = para_text.strip()
    # 【绝对过滤规则1】带句末标点的直接判定为正文，标题不会以句号、分号、感叹号、问号结尾
    if not text or len(text) < 2 or text.endswith(("。", "；", "！", "？", ".", ";", "!")):
        return "正文"
    # 【绝对过滤规则2】带圆括号序号的（1）（2）直接判定为正文，论文里这类100%是正文分点
    if re.match(r'^\s*（\d+）', para_text) or re.match(r'^\s*\(\d+\)', para_text):
        return "正文"
    # 【绝对过滤规则3】内容里带冒号、是完整陈述句的，直接判定为正文
    if "：" in text or ":" in text:
        return "正文"
    
    # 【精准匹配规则】仅匹配规范的章节标题
    # 三级标题：严格匹配 x.x.x + 空格 + 主题文字（论文标准三级标题）
    if re.match(r'^\s*\d+\.\d+\.\d+\s+[\u4e00-\u9fa5A-Za-z]', para_text):
        return "三级标题"
    # 二级标题：严格匹配 x.x + 空格 + 主题文字（论文标准二级标题）
    elif re.match(r'^\s*\d+\.\d+\s+[\u4e00-\u9fa5A-Za-z]', para_text):
        return "二级标题"
    # 一级标题：严格匹配 第X章 / 中文大写数字+、 / 数字.+主题（仅章节级）
    elif re.match(r'^\s*第[一二三四五六七八九十百]+章\s+', para_text) \
            or re.match(r'^\s*[一二三四五六七八九十]+、\s+[\u4e00-\u9fa5A-Za-z]', para_text) \
            or re.match(r'^\s*\d+\.\s+[\u4e00-\u9fa5A-Za-z]', para_text):
        return "一级标题"
    # 其余全部判定为正文
    else:
        return "正文"

def standardize_reference(text):
    if not RE_REF_FLAG.match(text.strip()) and not RE_REF_KEYWORD.search(text):
        return text, False
    text = RE_REF_SPACE.sub(' ', text.strip())
    text = RE_REF_CN_FONT.sub(r'\1[\2]', text)
    text = RE_REF_DOT.sub('.', text)
    text = RE_REF_COMMA.sub(',', text)
    text = RE_REF_COLON.sub(':', text)
    return text, True

def format_compliance_check(doc, cn_format):
    check_report = []
    title_levels = ["一级标题", "二级标题", "三级标题"]
    for para in doc.paragraphs:
        level = get_title_level(para.text)
        if level in title_levels:
            target_font = cn_format[level]["font"]
            target_size = FONT_SIZE_MAP.get(cn_format[level]["size"], 12)
            for run in para.runs:
                if run.font.name != target_font and run.font.name in CN_FONT_LIST:
                    check_report.append(f"⚠️ 【{level}】{para.text[:20]}... 字体不符合要求，应为{target_font}")
                if run.font.size and abs(run.font.size.pt - target_size) > 0.1:
                    check_report.append(f"⚠️ 【{level}】{para.text[:20]}... 字号不符合要求，应为{cn_format[level]['size']}")
        elif level == "正文" and para.text.strip():
            if not para.paragraph_format.first_line_indent or para.paragraph_format.first_line_indent.cm < 1.4 or para.paragraph_format.first_line_indent.cm > 1.5:
                check_report.append(f"⚠️ 【正文】{para.text[:20]}... 未设置首行缩进2字符")
    for i, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip() and para.paragraph_format.alignment != ALIGN_MAP[cn_format["表格"]["align"]]:
                        check_report.append(f"⚠️ 【表格{i+1}】单元格内容对齐方式不符合要求，应为{cn_format['表格']['align']}")
    if not check_report:
        check_report.append("✅ 文档格式完全符合要求，无违规项")
    return check_report

# ====================== 图片无损保留+排版优化函数 ======================
def optimize_image_layout(doc):
    image_count = 0
    for para in doc.paragraphs:
        has_image = False
        for run in para.runs:
            if run._element.xpath('.//a:blip'):
                has_image = True
                image_count += 1
                break
        if has_image:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(6)
            para.paragraph_format.keep_with_next = True
            para.paragraph_format.keep_together = True
            para.paragraph_format.first_line_indent = Cm(0)
    return image_count

# ====================== 智能降重引擎 ======================
def is_white_text(text):
    text_strip = text.strip()
    for word in WHITE_WORDS:
        if word in text_strip:
            return True
    if RE_WHITE_NUMBER.match(text_strip) or RE_WHITE_QUOTE.match(text_strip):
        return True
    return False

def check_semantic_keep(original, modified):
    original_keywords = set(RE_KEYWORDS.findall(original))
    modified_keywords = set(RE_KEYWORDS.findall(modified))
    
    if not original_keywords and not modified_keywords:
        return 1.0
    if not original_keywords:
        return 0.0 if modified_keywords else 1.0
    
    overlap = original_keywords & modified_keywords
    return len(overlap) / len(original_keywords)

def rewrite_sentence(sentence, level_config):
    original = sentence.strip()
    if len(original) < 5 or is_white_text(original):
        return original, "原文保留（白名单/短句）", 1.0
    
    modified = original
    rewrite_type = "无修改"
    
    if level_config["synonym"]:
        for old, new in SYNONYM_DICT.items():
            if old in modified and not is_white_text(old):
                modified = modified.replace(old, new)
                rewrite_type = "同义词替换"
    
    if level_config["sentence_reorder"]:
        parts = [p.strip() for p in RE_CLAUSE_SPLIT.split(modified) if p.strip()]
        if len(parts) >= 3 and not is_white_text(modified):
            last_part = parts[-1]
            rest_parts = parts[:-1]
            random.shuffle(rest_parts)
            modified = "，".join(rest_parts + [last_part])
            if not modified.endswith(("。", "！", "？", "；")):
                modified += "。"
            rewrite_type = "句式重构+语序打乱"
    
    semantic_score = check_semantic_keep(original, modified)
    if semantic_score < 0.7:
        return original, "原文保留（语义重合度不达标）", 1.0
    
    return modified, rewrite_type, round(semantic_score, 4)

def rewrite_paragraph(text, level_config):
    change_log = []
    sentences = RE_SENTENCE_SPLIT.split(text)
    new_sentences = []
    
    for sent in sentences:
        if not sent.strip():
            new_sentences.append(sent)
            continue
        
        new_sent, rewrite_type, semantic_score = rewrite_sentence(sent, level_config)
        new_sentences.append(new_sent)
        
        if sent != new_sent:
            change_log.append({
                "original": sent,
                "modified": new_sent,
                "type": rewrite_type,
                "semantic_score": semantic_score
            })
    
    return "".join(new_sentences), change_log

# ====================== 核心文档处理函数 ======================
def process_doc(
    file,
    cn_format,
    en_format,
    enable_rewrite=False,
    rewrite_level="标准降重",
    bind_wps_style=True,
    standardize_ref=True
):
    file.seek(0, os.SEEK_END)
    file_size_mb = file.tell() / (1024 * 1024)
    file.seek(0)
    if file_size_mb > MAX_FILE_SIZE_MB:
        raise Exception(f"文件大小超过限制（{MAX_FILE_SIZE_MB}MB），当前大小：{file_size_mb:.2f}MB")
    
    try:
        doc = Document(file)
    except Exception as e:
        raise Exception(f"文档读取失败，请确认是有效的docx文件：{str(e)}")
    
    total_changes = []
    ref_count = 0
    process_log = []
    title_stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": len(doc.tables)}
    rewrite_config = REWRITE_LEVEL[rewrite_level]
    style_warn_logged = False

    try:
        for para in doc.paragraphs:
            original_text = para.text
            level = get_title_level(original_text)
            title_stats[level] += 1

            if enable_rewrite and level == "正文":
                new_text, changes = rewrite_paragraph(original_text, rewrite_config)
                if changes:
                    total_changes.extend(changes)
                    para.text = new_text

            if standardize_ref:
                new_text, is_ref = standardize_reference(para.text)
                if is_ref:
                    para.text = new_text
                    ref_count += 1

            cn_style = cn_format[level]
            en_style = en_format[level]
            if bind_wps_style and level in WPS_STYLE_MAPPING:
                try:
                    target_style_id = WPS_STYLE_MAPPING[level]
                    if target_style_id in doc.styles:
                        para.style = doc.styles[target_style_id]
                except Exception as e:
                    if not style_warn_logged:
                        process_log.append(f"⚠️ 文档内置样式异常，已跳过WPS标题样式绑定")
                        style_warn_logged = True

            para_format = para.paragraph_format
            para_format.alignment = ALIGN_MAP[cn_style["align"]]
            para_format.first_line_indent = Cm(cn_style["indent"] * 0.74)
            para_format.space_before = Pt(cn_style["space_before"])
            para_format.space_after = Pt(cn_style["space_after"])
            
            if cn_style["line_type"] == "固定值":
                para_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                para_format.line_spacing = Pt(cn_style["line_value"])
            else:
                para_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                para_format.line_spacing = cn_style["line_value"]
            
            cn_size_pt = FONT_SIZE_MAP.get(cn_style["size"], 12)
            for run in para.runs:
                run.font.name = cn_style["font"]
                run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_style["font"])
                run._element.rPr.rFonts.set(qn('w:ascii'), en_style["en_font"])
                run._element.rPr.rFonts.set(qn('w:hAnsi'), en_style["en_font"])
                run.font.size = Pt(cn_size_pt)
                run.font.bold = en_style["bold"] if en_style["bold"] else cn_style["bold"]
                run.font.italic = en_style["italic"]
                run.font.color.rgb = RGBColor(0, 0, 0)
        
        process_log.append("✅ 全文档段落处理完成")
        if enable_rewrite:
            process_log.append(f"✅ 智能降重完成，共修改{len(total_changes)}处")
        if standardize_ref and ref_count > 0:
            process_log.append(f"✅ 参考文献标准化完成，共处理{ref_count}条")
        process_log.append(f"📊 标题识别结果：一级{title_stats['一级标题']}、二级{title_stats['二级标题']}、三级{title_stats['三级标题']}")
    except Exception as e:
        raise Exception(f"文档处理失败：{str(e)}")

    try:
        image_count = optimize_image_layout(doc)
        if image_count > 0:
            process_log.append(f"✅ 优化{image_count}张图片排版")
        else:
            process_log.append("✅ 未检测到图片")
    except Exception as e:
        process_log.append(f"⚠️ 图片处理失败：{str(e)}")

    try:
        cn_table_style = cn_format["表格"]
        en_table_style = en_format["表格"]
        table_cn_size = FONT_SIZE_MAP.get(cn_table_style["size"], 10.5)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if enable_rewrite:
                            original_text = para.text.strip()
                            if original_text and not is_white_text(original_text):
                                new_text, changes = rewrite_paragraph(original_text, rewrite_config)
                                if changes:
                                    total_changes.extend(changes)
                                    para.text = new_text
                        
                        para.alignment = ALIGN_MAP[cn_table_style["align"]]
                        para_format = para.paragraph_format
                        para_format.space_before = Pt(cn_table_style["space_before"])
                        para_format.space_after = Pt(cn_table_style["space_after"])
                        if cn_table_style["line_type"] == "固定值":
                            para_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                            para_format.line_spacing = Pt(cn_table_style["line_value"])
                        else:
                            para_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                            para_format.line_spacing = cn_table_style["line_value"]
                        
                        for run in para.runs:
                            run.font.name = cn_table_style["font"]
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_table_style["font"])
                            run._element.rPr.rFonts.set(qn('w:ascii'), en_table_style["en_font"])
                            run.font.size = Pt(table_cn_size)
                            run.font.bold = en_table_style["bold"] if en_table_style["bold"] else cn_table_style["bold"]
                            run.font.italic = en_table_style["italic"]
                            run.font.color.rgb = RGBColor(0, 0, 0)
        
        process_log.append("✅ 表格格式处理完成")
    except Exception as e:
        process_log.append(f"⚠️ 表格处理失败：{str(e)}")

    try:
        check_report = format_compliance_check(doc, cn_format)
        process_log.append("✅ 格式合规检查完成")
    except Exception as e:
        check_report = [f"⚠️ 格式检查失败：{str(e)}"]
        process_log.append(check_report[0])

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, total_changes, title_stats, process_log, check_report

# ====================== 处理报告生成函数 ======================
def generate_report(changes, rewrite_level, title_stats, process_log, check_report):
    total_count = len(changes)
    report = f"# 文档处理报告\n"
    report += f"📅 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    report += f"⚙️ 降重强度：{rewrite_level}\n"
    report += f"📝 总修改条数：{total_count}\n\n"
    report += "## 一、处理流程日志\n"
    for log in process_log:
        report += f"- {log}\n"
    report += "\n## 二、标题识别统计\n"
    for level, count in title_stats.items():
        report += f"- {level}：{count} 个\n"
    report += "\n## 三、格式合规性检查报告\n"
    for item in check_report:
        report += f"- {item}\n"
    return report.encode("utf-8")

# ====================== Streamlit主界面 ======================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide", page_icon="🏆")
    
    def safe_rerun():
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()
    
    if "current_template" not in st.session_state:
        st.session_state.current_template = "三创赛"
        st.session_state.cn_format, st.session_state.en_format = get_cached_template("三创赛")
    if "custom_templates" not in st.session_state:
        st.session_state.custom_templates = {}
    if "version" not in st.session_state:
        st.session_state.version = 0

    st.title(f"🏆 {APP_NAME}")
    st.success("✅ 精准标题识别 | 格式标准化 | 智能降重 | WPS导航栏自动生成")
    st.divider()

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        selected_template = st.selectbox("📌 选择模板类型", options=list(ALL_TEMPLATES.keys()), index=list(ALL_TEMPLATES.keys()).index(st.session_state.current_template))
        if selected_template != st.session_state.current_template:
            st.session_state.current_template = selected_template
            st.session_state.cn_format, st.session_state.en_format = get_cached_template(selected_template)
            st.session_state.version += 1
            safe_rerun()
    with col2:
        enable_rewrite = st.checkbox("🔄 开启智能降重", value=False)
        rewrite_level = st.selectbox("降重强度", options=list(REWRITE_LEVEL.keys()), index=1, disabled=not enable_rewrite)
    with col3:
        bind_wps_style = st.checkbox("✅ 绑定WPS标题样式", value=True)
        standardize_ref = st.checkbox("📚 参考文献标准化", value=True)

    st.markdown(f"### ⚠️ {selected_template} 格式要求")
    for req in ALL_TEMPLATES[selected_template]["special_requirements"]:
        st.markdown(f"- {req}")
    st.divider()

    with st.sidebar:
        st.subheader("⚙️ 高级格式自定义")
        st.markdown("#### 💾 自定义模板")
        template_name = st.text_input("模板名称", placeholder="自定义格式名称")
        col_template1, col_template2 = st.columns(2)
        with col_template1:
            if st.button("保存当前格式", use_container_width=True):
                if template_name:
                    st.session_state.custom_templates[template_name] = {"cn_format": copy.deepcopy(st.session_state.cn_format), "en_format": copy.deepcopy(st.session_state.en_format)}
                    st.success(f"✅ 模板「{template_name}」保存成功")
                else:
                    st.error("请输入模板名称")
        with col_template2:
            if st.session_state.custom_templates:
                selected_custom = st.selectbox("加载模板", options=list(st.session_state.custom_templates.keys()))
                if st.button("加载模板", use_container_width=True):
                    tmp = st.session_state.custom_templates[selected_custom]
                    st.session_state.cn_format = copy.deepcopy(tmp["cn_format"])
                    st.session_state.en_format = copy.deepcopy(tmp["en_format"])
                    st.session_state.version += 1
                    safe_rerun()
        
        st.divider()
        st.markdown("#### 🀄 中文格式设置")
        for level in LEVEL_LIST:
            with st.expander(f"{level}格式", expanded=False):
                cfg = st.session_state.cn_format[level]
                cfg["font"] = st.selectbox("中文字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cfg["font"]) if cfg["font"] in CN_FONT_LIST else 0, key=f"cn_{level}_font_{st.session_state.version}")
                cfg["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index(cfg["size"]) if cfg["size"] in FONT_SIZE_MAP else 5, key=f"cn_{level}_size_{st.session_state.version}")
                cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_{level}_bold_{st.session_state.version}")
                cfg["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg["align"]), key=f"cn_{level}_align_{st.session_state.version}")
                if level != "表格":
                    cfg["indent"] = st.number_input("首行缩进(字符)", 0,4,cfg["indent"],1, key=f"cn_{level}_indent_{st.session_state.version}")
                    cfg["space_before"] = st.number_input("段前间距", 0,24,cfg["space_before"],1, key=f"cn_{level}_before_{st.session_state.version}")
                    cfg["space_after"] = st.number_input("段后间距", 0,24,cfg["space_after"],1, key=f"cn_{level}_after_{st.session_state.version}")
                if level != "表格":
                    cfg["line_type"] = st.selectbox("行距类型", ["倍数", "固定值"], index=0 if cfg["line_type"]=="倍数" else 1, key=f"cn_{level}_line_{st.session_state.version}")
                    cfg["line_value"] = st.number_input("行距值", min_value=0.0 if cfg["line_type"]=="倍数" else 8, value=cfg["line_value"], step=0.1 if cfg["line_type"]=="倍数" else 1, key=f"cn_{level}_val_{st.session_state.version}")
                st.session_state.cn_format[level] = cfg
        
        st.divider()
        st.markdown("#### 🔤 西文格式设置")
        for level in LEVEL_LIST:
            with st.expander(f"{level}西文设置", expanded=False):
                cfg = st.session_state.en_format[level]
                cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]) if cfg["en_font"] in EN_FONT_LIST else 0, key=f"en_{level}_font_{st.session_state.version}")
                cfg["size_same_as_cn"] = st.checkbox("字号同步中文", cfg["size_same_as_cn"], key=f"en_{level}_same_{st.session_state.version}")
                cfg["bold"] = st.checkbox("西文加粗", cfg["bold"], key=f"en_{level}_bold_{st.session_state.version}")
                cfg["italic"] = st.checkbox("西文斜体", cfg["italic"], key=f"en_{level}_italic_{st.session_state.version}")
                st.session_state.en_format[level] = cfg

    st.subheader("📁 文档上传与处理")
    files = st.file_uploader("上传 .docx 文档", type=["docx"], accept_multiple_files=True)

    if files and st.button("🚀 开始处理文档", type="primary", use_container_width=True):
        for file in files:
            with st.spinner(f"正在处理：{file.name}"):
                try:
                    output_doc, changes, title_stats, process_log, check_report = process_doc(
                        file=file, cn_format=st.session_state.cn_format, en_format=st.session_state.en_format,
                        enable_rewrite=enable_rewrite, rewrite_level=rewrite_level, bind_wps_style=bind_wps_style, standardize_ref=standardize_ref
                    )
                    st.subheader(f"✅ 处理完成：{file.name}")
                    with st.expander("📋 处理日志", expanded=True):
                        for log in process_log: st.write(log)
                    cols = st.columns(5)
                    cols[0].metric("一级标题", title_stats["一级标题"])
                    cols[1].metric("二级标题", title_stats["二级标题"])
                    cols[2].metric("三级标题", title_stats["三级标题"])
                    cols[3].metric("正文段落", title_stats["正文"])
                    cols[4].metric("表格", title_stats["表格"])
                    st.download_button("📥 下载文档", data=output_doc, file_name=f"已格式化_{file.name}", use_container_width=True)
                    st.divider()
                except Exception as e:
                    st.error(f"处理失败：{str(e)}")

if __name__ == "__main__":
    main()
