import streamlit as st
import copy
import re
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_BUILTIN_STYLE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ====================== 全局配置与常量 ======================
# 专业术语白名单（降重不修改+格式保护）
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "挑战杯", "互联网+", "三创赛", "河北科技大学",
    "工业工程", "GDP", "CPI", "GB/T 7714", "ISO", "一级标题", "二级标题", "三级标题",
    "参考文献", "公式", "图表", "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "Transformer", "BERT", "T5", "Python", "Java", "SQL",
    "深度学习", "神经网络", "算法", "系统", "模型", "数据", "技术", "创新", "创业",
    "商业模式", "市场分析", "财务预测", "风险控制", "团队介绍", "产品服务", "SWOT"
]

# WPS/Word原生样式映射表（绑定内置标题1/2/3，导航窗格自动识别）
WPS_STYLE_MAPPING = {
    "一级标题": WD_BUILTIN_STYLE.HEADING_1,
    "二级标题": WD_BUILTIN_STYLE.HEADING_2,
    "三级标题": WD_BUILTIN_STYLE.HEADING_3,
    "正文": WD_BUILTIN_STYLE.NORMAL
}

# 主流竞赛官方格式模板
COMPETITION_FORMATS = {
    "三创赛": {
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "黑体", "size": "小三号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "楷体_GB2312", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "仿宋", "size": "四号", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.2, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三号", "bold": True, "italic": False},
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
            "三级标题": {"font": "黑体", "size": "小四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四号", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四号", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文约15000字", "双面打印", "严格章-节-条层级结构", "标题单倍行距，正文1.5倍行距"]
    },
    "计算机商业挑战赛": {
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
        "special_requirements": ["全文10-15页为宜", "文件大小30M以内", "PDF格式优先", "首页尾页呼应核心主题"]
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

# 降重强度配置
REWRITE_LEVEL = {
    "轻度降重": {"synonym": True, "sentence_reorder": False, "structure_change": False},
    "标准降重": {"synonym": True, "sentence_reorder": True, "structure_change": False},
    "强力降重": {"synonym": True, "sentence_reorder": True, "structure_change": True}
}

# 学术场景同义词词典
SYNONYM_DICT = {
    "提升": "有效改善", "降低": "显著减少", "增加": "大幅提升", "减少": "有效降低",
    "首先": "从实际落地情况来看", "其次": "进一步分析", "最后": "综合上述分析",
    "综上所述": "结合全维度分析", "总而言之": "从实践结果来看",
    "一方面": "站在需求端视角", "另一方面": "回到供给侧现实",
    "随着时代发展": "在当前行业背景下", "在当今社会": "结合当下实际环境",
    "应用": "落地实践", "研究": "专项分析", "效果": "实际表现",
    "优势": "核心竞争力", "问题": "行业痛点", "方法": "技术路径",
    "现状": "发展态势", "趋势": "未来走向", "解决": "破解", "实现": "达成"
}

# 全局常量
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
APP_NAME = "竞赛商业计划书格式处理器"
LEVEL_LIST = ["一级标题", "二级标题", "三级标题", "正文", "表格"]
EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri", "Courier New"]
CN_FONT_LIST = ["宋体", "黑体", "楷体", "仿宋_GB2312", "微软雅黑"]

# ====================== 1. 核心工具函数 ======================
# 标题层级精准识别
def get_title_level(para_text):
    text = para_text.strip()
    if not text or len(text) < 2:
        return "正文"
    if re.match(r'^第[一二三四五六七八九十]+章\s', text) or (re.match(r'^\d+、\s', text) and len(text) < 25):
        return "一级标题"
    elif re.match(r'^\d+\.\d+\s', text) or (re.match(r'^（[一二三四五六七八九十]+）\s', text) and len(text) < 25):
        return "二级标题"
    elif re.match(r'^\d+\.\d+\.\d+\s', text) or (re.match(r'^（\d+）\s', text) and len(text) < 20):
        return "三级标题"
    else:
        return "正文"

# 参考文献GB/T 7714格式标准化
def standardize_reference(text):
    # 识别参考文献段落
    if not re.search(r'^\[(\d+)\]', text.strip()) and not re.search(r'参考文献', text):
        return text, False
    
    # 标准化格式：[序号] 作者. 题名[文献类型标识]. 出版者, 出版年, 卷(期):起止页码.
    text = re.sub(r'\s+', ' ', text.strip())
    # 修正文献类型标识
    text = re.sub(r'([\u4e00-\u9fa5]+)\[([A-Z]+)\]', r'\1[\2]', text)
    # 修正标点符号为英文半角
    text = re.sub(r'。(?![\u4e00-\u9fa5])', '.', text)
    text = re.sub(r'，', ',', text)
    text = re.sub(r'：', ':', text)
    return text, True

# 格式合规性检查
def format_compliance_check(doc, cn_format):
    check_report = []
    title_levels = ["一级标题", "二级标题", "三级标题"]
    
    # 检查标题格式
    for para in doc.paragraphs:
        level = get_title_level(para.text)
        if level in title_levels:
            target_font = cn_format[level]["font"]
            target_size = FONT_SIZE_MAP[cn_format[level]["size"]]
            for run in para.runs:
                if run.font.name != target_font and run.font.name not in CN_FONT_LIST:
                    check_report.append(f"⚠️ 【{level}】{para.text[:20]}... 字体不符合要求，应为{target_font}")
                if run.font.size and run.font.size.pt != target_size:
                    check_report.append(f"⚠️ 【{level}】{para.text[:20]}... 字号不符合要求，应为{cn_format[level]['size']}")
    
    # 检查正文首行缩进
    for para in doc.paragraphs:
        if get_title_level(para.text) == "正文" and para.text.strip():
            if not para.paragraph_format.first_line_indent:
                check_report.append(f"⚠️ 【正文】{para.text[:20]}... 未设置首行缩进2字符")
    
    # 检查表格格式
    for i, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip() and para.paragraph_format.alignment != ALIGN_MAP[cn_format["表格"]["align"]]:
                        check_report.append(f"⚠️ 【表格{i+1}】单元格内容对齐方式不符合要求，应为{cn_format['表格']['align']}")
    
    if not check_report:
        check_report.append("✅ 文档格式完全符合竞赛要求，无违规项")
    return check_report

# 自动生成竞赛封面页
def add_cover_page(doc, cover_info):
    # 在文档最开头插入封面
    new_doc = Document()
    
    # 封面标题
    title_para = new_doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(cover_info["project_name"])
    title_run.font.name = "黑体"
    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), "黑体")
    title_run.font.size = Pt(36)
    title_run.font.bold = True
    
    # 副标题
    sub_title_para = new_doc.add_paragraph()
    sub_title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_title_run = sub_title_para.add_run(f"{cover_info['competition']}参赛作品")
    sub_title_run.font.name = "宋体"
    sub_title_run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
    sub_title_run.font.size = Pt(22)
    sub_title_run.font.bold = True
    
    # 空行
    for _ in range(8):
        new_doc.add_paragraph()
    
    # 团队信息
    info_list = [
        f"参赛学校：{cover_info['school']}",
        f"团队名称：{cover_info['team_name']}",
        f"指导老师：{cover_info['teacher']}",
        f"完成日期：{cover_info['date']}"
    ]
    
    for info in info_list:
        para = new_doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(info)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        run.font.size = Pt(16)
        run.font.bold = True
    
    # 封面分页
    new_doc.add_page_break()
    
    # 合并原文档内容
    for element in doc.element.body:
        new_doc.element.body.append(element)
    
    return new_doc

# 自动设置页眉页脚
def set_header_footer(doc, header_info):
    # 获取文档所有节
    for section in doc.sections:
        # 设置页眉
        header = section.header
        header_para = header.paragraphs[0]
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_run = header_para.add_run(header_info["header_text"])
        header_run.font.name = "宋体"
        header_run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        header_run.font.size = Pt(10.5)
        
        # 设置页脚：页码 + 团队名
        footer = section.footer
        footer_para = footer.paragraphs[0]
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 添加页码域代码
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        
        footer_run = footer_para.add_run(f"第 ")
        footer_run.font.name = "宋体"
        footer_run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        footer_run.font.size = Pt(10.5)
        
        footer_para._p.append(fldChar1)
        footer_para._p.append(instrText)
        footer_para._p.append(fldChar2)
        
        footer_run2 = footer_para.add_run(f" 页 | {header_info['footer_text']}")
        footer_run2.font.name = "宋体"
        footer_run2._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        footer_run2.font.size = Pt(10.5)
    
    return doc

# 自动生成目录（优化版）
def add_table_of_contents(doc):
    doc.add_paragraph("目录", style=doc.styles[WD_BUILTIN_STYLE.HEADING_1])
    para = doc.add_paragraph()
    run = para.add_run()
    # 目录域代码：1-3级标题，带超链接，自动识别大纲级别
    fldChar_begin = doc._element.makeelement('w:fldChar', {'w:fldCharType': 'begin'})
    instrText = doc._element.makeelement('w:instrText', {'xml:space': 'preserve'})
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar_separate = doc._element.makeelement('w:fldChar', {'w:fldCharType': 'separate'})
    fldChar_end = doc._element.makeelement('w:fldChar', {'w:fldCharType': 'end'})
    
    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run._r.append(fldChar_end)
    doc.add_page_break()
    return doc

# ====================== 2. 智能降重引擎（升级专业术语保护） ======================
def is_white_text(text):
    for word in WHITE_WORDS:
        if word in text:
            return True
    if re.match(r'^\d+(\.\d+)*$', text) or re.match(r'^\[.*\]$', text):
        return True
    return False

def check_semantic_keep(original, modified):
    original_keywords = set(re.findall(r'[\u4e00-\u9fa5]{2,}', original))
    modified_keywords = set(re.findall(r'[\u4e00-\u9fa5]{2,}', modified))
    if not original_keywords:
        return 1.0
    overlap = original_keywords & modified_keywords
    return len(overlap) / len(original_keywords)

def rewrite_sentence(sentence, level_config):
    original = sentence.strip()
    if len(original) < 5 or is_white_text(original):
        return original, "原文保留（白名单/短句）", 1.0

    modified = original
    rewrite_type = "无修改"

    # 同义词替换（跳过专业术语）
    if level_config["synonym"]:
        for old, new in SYNONYM_DICT.items():
            if old in modified and not is_white_text(old):
                modified = modified.replace(old, new)
                rewrite_type = "同义词替换"

    # 句式重构
    if level_config["sentence_reorder"]:
        parts = [p.strip() for p in re.split(r'[，。；]', modified) if p.strip()]
        if len(parts) >= 3 and not is_white_text(modified):
            import random
            random.shuffle(parts)
            modified = "，".join(parts) + "。"
            rewrite_type = "句式重构+语序打乱"

    # 结构调整
    if level_config["structure_change"]:
        if "在" in modified and "中" in modified and not is_white_text(modified):
            modified = re.sub(r'在(.*?)中', f'结合{datetime.now().year}年行业实际发展情况，在\g<1>场景中', modified)
            rewrite_type = "结构调整+场景限定补充"

    # 语义校验
    semantic_score = check_semantic_keep(original, modified)
    if semantic_score < 0.7:
        return original, "原文保留（语义重合度不达标）", 1.0

    return modified, rewrite_type, round(semantic_score, 4)

def rewrite_paragraph(text, level_config):
    change_log = []
    sentences = re.split(r'(?<=[。！？；])\s*', text)
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

# ====================== 3. 核心文档处理函数 ======================
def process_doc(
    file,
    cn_format,
    en_format,
    enable_rewrite=False,
    rewrite_level="标准降重",
    bind_wps_style=True,
    add_toc=True,
    cover_info=None,
    header_info=None,
    standardize_ref=True
):
    doc = Document(file)
    total_changes = []
    ref_count = 0
    process_log = []
    title_stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": len(doc.tables)}
    rewrite_config = REWRITE_LEVEL[rewrite_level]

    # 第一步：添加封面页
    if cover_info and cover_info["enable"]:
        doc = add_cover_page(doc, cover_info)
        process_log.append("✅ 已生成竞赛封面页")

    # 第二步：可选智能降重
    if enable_rewrite:
        for para in doc.paragraphs:
            original_text = para.text
            level = get_title_level(original_text)
            if level == "正文":
                new_text, changes = rewrite_paragraph(original_text, rewrite_config)
                if changes:
                    total_changes.extend(changes)
                    para.text = new_text
        process_log.append(f"✅ 智能降重完成，共修改{len(total_changes)}处内容")

        # 表格内文本降重
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        original_text = para.text.strip()
                        if not original_text or is_white_text(original_text):
                            continue
                        new_text, changes = rewrite_paragraph(original_text, rewrite_config)
                        if changes:
                            total_changes.extend(changes)
                            para.text = new_text

    # 第三步：参考文献格式标准化
    if standardize_ref:
        for para in doc.paragraphs:
            new_text, is_ref = standardize_reference(para.text)
            if is_ref:
                para.text = new_text
                ref_count += 1
        if ref_count > 0:
            process_log.append(f"✅ 参考文献格式标准化完成，共处理{ref_count}条参考文献")

    # 第四步：绑定WPS标题样式+全格式设置
    for para in doc.paragraphs:
        level = get_title_level(para.text)
        title_stats[level] += 1
        cn_style = cn_format[level]
        en_style = en_format[level]

        # 绑定WPS原生标题样式+设置大纲级别
        if bind_wps_style and level in WPS_STYLE_MAPPING:
            para.style = doc.styles[WPS_STYLE_MAPPING[level]]
            para.paragraph_format.outline_level = int(level[0])

        # 段落格式设置
        para_format = para.paragraph_format
        para_format.alignment = ALIGN_MAP[cn_style["align"]]
        para_format.first_line_indent = Cm(cn_style["indent"] * 0.74)
        para_format.space_before = Pt(cn_style["space_before"])
        para_format.space_after = Pt(cn_style["space_after"])

        # 行距设置
        if cn_style["line_type"] == "固定值":
            para_format.line_spacing_rule = 2
            para_format.line_spacing = Pt(cn_style["line_value"])
        else:
            para_format.line_spacing_rule = 1
            para_format.line_spacing = cn_style["line_value"]

        # 中文/西文独立字体设置（全层级）
        cn_size_pt = FONT_SIZE_MAP[cn_style["size"]]
        en_size_pt = FONT_SIZE_MAP[cn_style["size"]] if en_style["size_same_as_cn"] else FONT_SIZE_MAP[en_style["size"]]
        for run in para.runs:
            # 中文字体
            run.font.name = cn_style["font"]
            run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_style["font"])
            # 西文/数字字体
            run._element.rPr.rFonts.set(qn('w:ascii'), en_style["en_font"])
            run._element.rPr.rFonts.set(qn('w:hAnsi'), en_style["en_font"])
            run._element.rPr.rFonts.set(qn('w:cs'), en_style["en_font"])
            # 字号+格式
            run.font.size = Pt(cn_size_pt)
            run.font.bold = en_style["bold"] if en_style["bold"] else cn_style["bold"]
            run.font.italic = en_style["italic"]

    # 第五步：表格格式全设置
    cn_table_style = cn_format["表格"]
    en_table_style = en_format["表格"]
    table_cn_size = FONT_SIZE_MAP[cn_table_style["size"]]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = ALIGN_MAP[cn_table_style["align"]]
                    for run in para.runs:
                        run.font.name = cn_table_style["font"]
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_table_style["font"])
                        run._element.rPr.rFonts.set(qn('w:ascii'), en_table_style["en_font"])
                        run._element.rPr.rFonts.set(qn('w:hAnsi'), en_table_style["en_font"])
                        run.font.size = Pt(table_cn_size)
                        run.font.bold = en_table_style["bold"] if en_table_style["bold"] else cn_table_style["bold"]
                        run.font.italic = en_table_style["italic"]
    process_log.append("✅ 全文档格式设置完成")

    # 第六步：添加目录
    if add_toc:
        doc = add_table_of_contents(doc)
        process_log.append("✅ 自动生成目录完成")

    # 第七步：设置页眉页脚
    if header_info and header_info["enable"]:
        doc = set_header_footer(doc, header_info)
        process_log.append("✅ 页眉页脚设置完成")

    # 第八步：格式合规性检查
    check_report = format_compliance_check(doc, cn_format)
    process_log.append("✅ 格式合规性检查完成")

    # 输出文档
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, total_changes, title_stats, process_log, check_report

# ====================== 4. 报告生成函数 ======================
def generate_report(changes, rewrite_level, title_stats, process_log, check_report):
    total_count = len(changes)
    report = f"# 竞赛文档处理报告\n"
    report += f"📅 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    report += f"⚙️ 降重强度：{rewrite_level}\n"
    report += f"📝 总修改条数：{total_count}\n\n"

    # 处理日志
    report += "## 一、处理流程日志\n"
    for log in process_log:
        report += f"- {log}\n"
    report += "\n"

    # 标题识别统计
    report += "## 二、标题识别统计\n"
    for level, count in title_stats.items():
        report += f"- {level}：{count} 个\n"
    report += "\n✅ 已自动绑定WPS原生「标题1/标题2/标题3」样式，WPS打开后导航窗格自动生成目录\n\n"

    # 格式合规性检查报告
    report += "## 三、格式合规性检查报告\n"
    for item in check_report:
        report += f"- {item}\n"
    report += "\n"

    # 降重修改记录
    if total_count > 0:
        report += "## 四、降重修改统计\n"
        type_count = {}
        for change in changes:
            t = change["type"]
            type_count[t] = type_count.get(t, 0) + 1
        for t, count in type_count.items():
            report += f"- {t}：{count} 条\n"
        report += "\n"

        report += "## 五、详细修改记录（前100条）\n"
        for i, change in enumerate(changes[:100]):
            report += f"### 修改记录 #{i+1}\n"
            report += f"📋 修改类型：{change['type']}\n"
            report += f"📊 语义重合度：{change['semantic_score']}\n"
            report += f"原文：{change['original']}\n"
            report += f"改后：{change['modified']}\n\n"

    return report.encode("utf-8")

# ====================== 5. Streamlit界面UI ======================
def main():
    st.set_page_config(
        page_title=APP_NAME,
        layout="wide",
        page_icon="🏆",
        initial_sidebar_state="expanded"
    )

    # 初始化页面状态
    if "current_competition" not in st.session_state:
        st.session_state.current_competition = "三创赛"
    if "cn_format" not in st.session_state:
        st.session_state.cn_format = copy.deepcopy(COMPETITION_FORMATS[st.session_state.current_competition]["cn_format"])
    if "en_format" not in st.session_state:
        st.session_state.en_format = copy.deepcopy(COMPETITION_FORMATS[st.session_state.current_competition]["en_format"])
    if "custom_templates" not in st.session_state:
        st.session_state.custom_templates = {}
    if "version" not in st.session_state:
        st.session_state.version = 0

    # 页面标题
    st.title(f"🏆 {APP_NAME}")
    st.success("✅ 适配三创赛/挑战杯/互联网+等主流竞赛 | 自动生成封面/页眉页脚/目录 | WPS标题绑定 | 参考文献标准化 | 智能降重")
    st.divider()

    # 顶部核心配置区
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        selected_competition = st.selectbox(
            "📌 选择你的竞赛类型",
            options=list(COMPETITION_FORMATS.keys()),
            index=list(COMPETITION_FORMATS.keys()).index(st.session_state.current_competition)
        )
        # 切换竞赛自动加载对应格式
        if selected_competition != st.session_state.current_competition:
            st.session_state.current_competition = selected_competition
            st.session_state.cn_format = copy.deepcopy(COMPETITION_FORMATS[selected_competition]["cn_format"])
            st.session_state.en_format = copy.deepcopy(COMPETITION_FORMATS[selected_competition]["en_format"])
            st.session_state.version += 1
            st.rerun()

    with col2:
        enable_rewrite = st.checkbox("🔄 开启智能降重", value=False)
        rewrite_level = st.selectbox(
            "降重强度选择",
            options=list(REWRITE_LEVEL.keys()),
            index=1,
            disabled=not enable_rewrite
        )

    with col3:
        add_toc = st.checkbox("📋 自动生成目录", value=True)
        bind_wps_style = st.checkbox("✅ 绑定WPS标题样式", value=True)
        standardize_ref = st.checkbox("📚 参考文献标准化", value=True)

    # 竞赛特殊要求提示
    st.markdown(f"### ⚠️ {selected_competition} 官方格式特殊要求")
    for req in COMPETITION_FORMATS[selected_competition]["special_requirements"]:
        st.markdown(f"- {req}")
    st.divider()

    # 新增：封面页设置
    with st.expander("📄 竞赛封面页设置", expanded=False):
        enable_cover = st.checkbox("✅ 启用自动生成封面页", value=True)
        col_cover1, col_cover2 = st.columns(2)
        with col_cover1:
            project_name = st.text_input("项目名称", placeholder="请输入你的参赛项目名称")
            school = st.text_input("参赛学校", placeholder="请输入你的学校名称")
            teacher = st.text_input("指导老师", placeholder="请输入指导老师姓名")
        with col_cover2:
            team_name = st.text_input("团队名称", placeholder="请输入你的团队名称")
            competition_name = st.text_input("竞赛全称", value=selected_competition)
            date = st.date_input("完成日期", value=datetime.now())
        
        cover_info = {
            "enable": enable_cover,
            "project_name": project_name,
            "team_name": team_name,
            "school": school,
            "teacher": teacher,
            "competition": competition_name,
            "date": date.strftime("%Y年%m月%d日")
        }

    # 新增：页眉页脚设置
    with st.expander("📑 页眉页脚设置", expanded=False):
        enable_header = st.checkbox("✅ 启用自动设置页眉页脚", value=True)
        col_header1, col_header2 = st.columns(2)
        with col_header1:
            header_text = st.text_input("页眉内容", placeholder="请输入页眉显示的内容，如项目名称", value=project_name if enable_cover else "")
        with col_header2:
            footer_text = st.text_input("页脚附加内容", placeholder="页脚页码后显示的内容，如团队名称", value=team_name if enable_cover else "")
        
        header_info = {
            "enable": enable_header,
            "header_text": header_text,
            "footer_text": footer_text
        }

    # 左侧边栏：高级格式设置
    with st.sidebar:
        st.subheader("⚙️ 高级格式自定义")
        st.caption("默认已加载竞赛标准格式，无需修改即可直接使用")

        # 自定义模板功能
        st.markdown("#### 💾 自定义模板")
        template_name = st.text_input("模板名称", placeholder="给你的自定义格式起个名字")
        col_template1, col_template2 = st.columns(2)
        with col_template1:
            if st.button("保存当前格式为模板", use_container_width=True):
                if template_name:
                    st.session_state.custom_templates[template_name] = {
                        "cn_format": copy.deepcopy(st.session_state.cn_format),
                        "en_format": copy.deepcopy(st.session_state.en_format)
                    }
                    st.success(f"✅ 模板「{template_name}」保存成功")
                else:
                    st.error("请输入模板名称")
        with col_template2:
            if st.session_state.custom_templates:
                selected_template = st.selectbox("加载模板", options=list(st.session_state.custom_templates.keys()))
                if st.button("加载选中模板", use_container_width=True):
                    template = st.session_state.custom_templates[selected_template]
                    st.session_state.cn_format = copy.deepcopy(template["cn_format"])
                    st.session_state.en_format = copy.deepcopy(template["en_format"])
                    st.session_state.version += 1
                    st.rerun()
        st.divider()

        # 中文格式设置
        st.markdown("#### 🀄 中文格式设置")
        for level in LEVEL_LIST:
            with st.expander(f"{level}格式", expanded=False):
                cfg = st.session_state.cn_format[level]
                cfg["font"] = st.selectbox(
                    "中文字体",
                    CN_FONT_LIST,
                    index=CN_FONT_LIST.index(cfg["font"]),
                    key=f"cn_{level}_font_{st.session_state.version}"
                )
                cfg["size"] = st.selectbox(
                    "字号",
                    list(FONT_SIZE_MAP.keys()),
                    index=list(FONT_SIZE_MAP.keys()).index(cfg["size"]),
                    key=f"cn_{level}_size_{st.session_state.version}"
                )
                cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_{level}_bold_{st.session_state.version}")
                cfg["align"] = st.selectbox(
                    "对齐方式",
                    list(ALIGN_MAP.keys()),
                    index=list(ALIGN_MAP.keys()).index(cfg["align"]),
                    key=f"cn_{level}_align_{st.session_state.version}"
                )
                if level != "表格":
                    cfg["indent"] = st.number_input(
                        "首行缩进(字符)",
                        min_value=0,
                        max_value=4,
                        value=cfg["indent"],
                        step=1,
                        key=f"cn_{level}_indent_{st.session_state.version}"
                    )
                    cfg["space_before"] = st.number_input(
                        "段前间距(磅)",
                        min_value=0,
                        max_value=24,
                        value=cfg["space_before"],
                        step=1,
                        key=f"cn_{level}_before_{st.session_state.version}"
                    )
                    cfg["space_after"] = st.number_input(
                        "段后间距(磅)",
                        min_value=0,
                        max_value=24,
                        value=cfg["space_after"],
                        step=1,
                        key=f"cn_{level}_after_{st.session_state.version}"
                    )
                st.session_state.cn_format[level] = cfg
        st.divider()

        # 全层级西文/数字格式设置
        st.markdown("#### 🔤 西文/数字格式设置")
        for level in LEVEL_LIST:
            with st.expander(f"{level}西文/数字设置", expanded=False):
                cfg = st.session_state.en_format[level]
                cfg["en_font"] = st.selectbox(
                    "西文字体",
                    EN_FONT_LIST,
                    index=EN_FONT_LIST.index(cfg["en_font"]),
                    key=f"en_{level}_font_{st.session_state.version}"
                )
                cfg["size_same_as_cn"] = st.checkbox(
                    "字号与中文同步",
                    cfg["size_same_as_cn"],
                    key=f"en_{level}_same_{st.session_state.version}"
                )
                if not cfg["size_same_as_cn"]:
                    cfg["size"] = st.selectbox(
                        "西文字号",
                        list(FONT_SIZE_MAP.keys()),
                        index=list(FONT_SIZE_MAP.keys()).index(cfg["size"]),
                        key=f"en_{level}_size_{st.session_state.version}"
                    )
                cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"en_{level}_bold_{st.session_state.version}")
                cfg["italic"] = st.checkbox("斜体", cfg["italic"], key=f"en_{level}_italic_{st.session_state.version}")
                st.session_state.en_format[level] = cfg

    # 右侧主区域：文件上传与处理
    st.subheader("📁 文档上传与处理")
    files = st.file_uploader(
        "上传 .docx 格式的商业计划书/论文（支持多选批量处理）",
        type=["docx"],
        accept_multiple_files=True
    )

    # 无文件时的引导内容
    if not files:
        st.info("👆 请上传docx格式的文档，即可一键完成竞赛全流程格式标准化")
        st.markdown("### 💡 核心功能速览")
        col_func1, col_func2, col_func3, col_func4 = st.columns(4)
        with col_func1:
            st.markdown("**🏷️ WPS标题绑定**")
            st.write("自动识别标题层级，绑定WPS原生样式，导航窗格自动生成")
        with col_func2:
            st.markdown("**📄 封面/页眉页脚**")
            st.write("一键生成符合竞赛要求的封面、页眉页脚，不用手动排版")
        with col_func3:
            st.markdown("**📚 参考文献标准化**")
            st.write("自动修正为GB/T 7714国标格式，符合学术规范")
        with col_func4:
            st.markdown("**✅ 格式合规检查**")
            st.write("自动扫描格式问题，生成检查报告，避免比赛格式扣分")

    # 有文件时的处理按钮
    if files and st.button("🚀 开始处理文档", type="primary", use_container_width=True):
        for file in files:
            with st.spinner(f"正在处理：{file.name}"):
                # 执行核心处理
                output_doc, changes, title_stats, process_log, check_report = process_doc(
                    file=file,
                    cn_format=st.session_state.cn_format,
                    en_format=st.session_state.en_format,
                    enable_rewrite=enable_rewrite,
                    rewrite_level=rewrite_level,
                    bind_wps_style=bind_wps_style,
                    add_toc=add_toc,
                    cover_info=cover_info,
                    header_info=header_info,
                    standardize_ref=standardize_ref
                )

            # 处理结果展示
            st.subheader(f"✅ 处理完成：{file.name}")
            # 处理日志
            with st.expander("📋 处理流程日志", expanded=True):
                for log in process_log:
                    st.write(log)
            # 标题统计
            cols = st.columns(5)
            cols[0].metric("一级标题", title_stats["一级标题"])
            cols[1].metric("二级标题", title_stats["二级标题"])
            cols[2].metric("三级标题", title_stats["三级标题"])
            cols[3].metric("正文段落", title_stats["正文"])
            cols[4].metric("表格数量", title_stats["表格"])
            # 格式检查报告
            with st.expander("⚠️ 格式合规性检查报告", expanded=False):
                for item in check_report:
                    st.write(item)

            # 下载按钮
            st.download_button(
                label="📥 下载已处理文档",
                data=output_doc,
                file_name=f"竞赛格式_{file.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

            # 处理报告下载
            report_bytes = generate_report(changes, rewrite_level, title_stats, process_log, check_report)
            st.download_button(
                label="📄 下载完整处理报告",
                data=report_bytes,
                file_name=f"处理报告_{file.name}.txt",
                mime="text/plain",
                use_container_width=True
            )
            st.divider()

if __name__ == "__main__":
    main()
