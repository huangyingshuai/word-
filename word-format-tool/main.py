import streamlit as st
import copy
import re
import random
import json
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_BUILTIN_STYLE
from docx.oxml.ns import qn
import os
import requests
import pandas as pd

# ====================== 全局样式注入（核心排版优化） ======================
def set_global_style():
    st.markdown("""
    <style>
    /* 全局主题色与字体 */
    :root {
        --primary-color: #2563eb;
        --secondary-color: #1e40af;
        --light-bg: #f1f5f9;
        --card-bg: #ffffff;
        --text-primary: #0f172a;
        --text-secondary: #475569;
        --border-radius: 8px;
        --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
        --shadow-md: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }
    html, body, [class*="css"] {
        font-family: "Inter", "Source Han Sans SC", "Microsoft YaHei", sans-serif;
        color: var(--text-primary);
    }
    /* 主标题样式 */
    h1 {
        font-size: 2rem;
        font-weight: 700;
        color: var(--primary-color);
        margin-bottom: 0.5rem;
    }
    h2 {
        font-size: 1.5rem;
        font-weight: 600;
        color: var(--text-primary);
        margin-top: 1rem;
        margin-bottom: 0.75rem;
    }
    h3 {
        font-size: 1.15rem;
        font-weight: 600;
        color: var(--text-primary);
        margin-top: 0.75rem;
        margin-bottom: 0.5rem;
    }
    /* 卡片容器 */
    .card {
        background: var(--card-bg);
        border-radius: var(--border-radius);
        padding: 1.25rem;
        box-shadow: var(--shadow-sm);
        border: 1px solid #e2e8f0;
        margin-bottom: 1rem;
    }
    /* 按钮美化 */
    .stButton>button {
        width: 100%;
        border-radius: var(--border-radius);
        font-weight: 500;
        padding: 0.5rem 1rem;
        border: none;
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: var(--shadow-md);
    }
    /* 侧边栏优化 */
    [data-testid="stSidebar"] {
        background-color: var(--light-bg);
        padding-top: 2rem;
    }
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: var(--secondary-color);
    }
    /* 上传组件美化 */
    [data-testid="stFileUploader"] {
        background: var(--card-bg);
        border-radius: var(--border-radius);
        padding: 1rem;
        border: 1px dashed #94a3b8;
    }
    /* 指标卡片美化 */
    [data-testid="stMetric"] {
        background: var(--light-bg);
        border-radius: var(--border-radius);
        padding: 1rem;
        box-shadow: var(--shadow-sm);
        text-align: center;
    }
    [data-testid="stMetricLabel"] {
        font-weight: 500;
        color: var(--text-secondary);
    }
    [data-testid="stMetricValue"] {
        font-weight: 700;
        color: var(--primary-color);
        font-size: 1.5rem;
    }
    /* 标签页美化 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        margin-bottom: 1rem;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: var(--border-radius);
        padding: 0.5rem 1.25rem;
        font-weight: 500;
        background-color: var(--light-bg);
    }
    .stTabs [aria-selected="true"] {
        background-color: var(--primary-color) !important;
        color: white !important;
    }
    /* 展开栏美化 */
    [data-testid="stExpander"] {
        border-radius: var(--border-radius);
        border: 1px solid #e2e8f0;
        box-shadow: var(--shadow-sm);
        margin-bottom: 0.75rem;
    }
    /* 分割线优化 */
    hr {
        border: none;
        height: 1px;
        background: #e2e8f0;
        margin: 1.5rem 0;
    }
    /* 成功/错误提示优化 */
    .stSuccess, .stInfo, .stWarning, .stError {
        border-radius: var(--border-radius);
        border: none;
        box-shadow: var(--shadow-sm);
    }
    </style>
    """, unsafe_allow_html=True)

# ====================== 预编译正则 ======================
RE_REF_FLAG = re.compile(r'^\[(\d+)\]')
RE_REF_KEYWORD = re.compile(r'参考文献|参考资料|References')
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
RE_RED_HIGHLIGHT = re.compile(r'<font color="red">(.*?)</font>', re.DOTALL)

# ====================== 全局配置与常量 ======================
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "PaperYY", "PaperFree", "挑战杯", "互联网+", "三创赛",
    "参考文献", "公式", "图表", "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "算法", "系统", "模型", "数据"
]
WPS_STYLE_MAPPING = {
    "一级标题": WD_BUILTIN_STYLE.HEADING_1,
    "二级标题": WD_BUILTIN_STYLE.HEADING_2,
    "三级标题": WD_BUILTIN_STYLE.HEADING_3,
    "正文": WD_BUILTIN_STYLE.NORMAL
}

# 竞赛模板
COMPETITION_FORMATS = {
    "三创赛-全国大学生电子商务创新创意及创业挑战赛": {
        "update_time": "2024-01-15",
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
    "挑战杯-全国大学生课外学术科技作品竞赛": {
        "update_time": "2024-02-20",
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
    "互联网+大学生创新创业大赛": {
        "update_time": "2024-03-10",
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

# 高校论文模板
UNIVERSITY_FORMATS = {
    "清华大学本科毕业论文模板": {
        "update_time": "2024-04-01",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 24, "space_after": 18},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["全文8000-15000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注清华大学本科毕业论文"]
    },
    "北京大学本科毕业论文模板": {
        "update_time": "2024-04-02",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文10000-20000字", "需包含摘要/关键词/参考文献", "参考文献需符合GB/T 7714", "页眉标注北京大学本科毕业论文"]
    }
}

THESIS_FORMATS = {
    "本科毕业论文-通用模板": {
        "update_time": "2024-04-01",
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

ALL_TEMPLATES = {**COMPETITION_FORMATS, **UNIVERSITY_FORMATS, **THESIS_FORMATS}
REWRITE_LEVEL = {
    "轻度润色": {"synonym": True, "sentence_reorder": False, "structure_change": False},
    "标准润色": {"synonym": True, "sentence_reorder": True, "structure_change": False},
    "深度润色": {"synonym": True, "sentence_reorder": True, "structure_change": True}
}
SYNONYM_DICT = {
    "提升": "有效改善", "降低": "显著减少", "增加": "大幅提升", "减少": "有效降低",
    "首先": "从实际落地情况来看", "其次": "进一步分析", "最后": "综合上述分析",
    "综上所述": "结合全维度分析", "总而言之": "从实践结果来看",
    "研究": "调研分析", "实验": "测试验证", "分析": "剖析", "结果": "结论",
    "方法": "方案", "系统": "平台", "模型": "架构", "数据": "信息"
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
EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri", "Courier New"]
CN_FONT_LIST = ["宋体", "黑体", "楷体", "仿宋_GB2312", "微软雅黑"]
MAX_FILE_SIZE_MB = 50
random.seed(42)

# ====================== 核心工具函数（完全保留原功能） ======================
@st.cache_data(ttl=3600)
def get_cached_template(template_name):
    return copy.deepcopy(ALL_TEMPLATES[template_name]["cn_format"]), copy.deepcopy(ALL_TEMPLATES[template_name]["en_format"])

def get_title_level(para_text):
    text = para_text.strip()
    if not text or len(text) < 2:
        return "正文"
    if re.match(r'^\s*（\d+）', para_text) or re.match(r'^\s*\(\d+\)', para_text):
        return "正文"
    end_with_punct = text.endswith(("。", "；", "！", "？", ".", ";", "!", "?"))
    is_single_number_start = re.match(r'^\s*\d+[、.]\s*', para_text)
    if end_with_punct and is_single_number_start:
        return "正文"
    if re.match(r'^\s*\d+\.\d+\.\d+[.、]?\s*', para_text):
        return "三级标题"
    elif re.match(r'^\s*\d+\.\d+[.、]?\s*', para_text):
        return "二级标题"
    elif re.match(r'^\s*第[一二三四五六七八九十百]+章\s+', para_text) \
            or re.match(r'^\s*[一二三四五六七八九十]+、\s*', para_text) \
            or (is_single_number_start and not end_with_punct):
        return "一级标题"
    else:
        return "正文"

def read_uploaded_doc_content(file):
    """读取上传的文档内容，支持docx和txt"""
    if not file:
        return "", None
    try:
        if file.name.endswith('.docx'):
            doc = Document(file)
            text = "\n".join([p.text for p in doc.paragraphs])
            return text, None
        elif file.name.endswith('.txt'):
            text = file.read().decode('utf-8')
            return text, None
        else:
            return "", "仅支持 .docx 和 .txt 格式文档"
    except Exception as e:
        return "", f"文档读取失败：{str(e)}"

def standardize_cnki_reference(text):
    if not text.strip():
        return text, False
    text = RE_REF_SPACE.sub(' ', text.strip())
    text = RE_REF_CN_FONT.sub(r'\1[\2]', text)
    text = RE_REF_DOT.sub('.', text)
    text = RE_REF_COMMA.sub(',', text)
    text = RE_REF_COLON.sub(':', text)
    if RE_REF_FLAG.match(text) or RE_REF_KEYWORD.search(text):
        return text, True
    return text, False

def parse_plagiarism_report(file):
    try:
        content = file.read().decode('utf-8', errors='ignore')
        red_parts = RE_RED_HIGHLIGHT.findall(content)
        plain_text = RE_RED_HIGHLIGHT.sub(r'\1', content)
        return red_parts, plain_text, None
    except Exception as e:
        return None, None, str(e)

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
            if para.paragraph_format.line_spacing:
                target_line = cn_format["正文"]["line_value"]
                if abs(para.paragraph_format.line_spacing - target_line) > 0.1:
                    check_report.append(f"⚠️ 【正文】{para.text[:20]}... 行间距不符合要求，应为{target_line}倍")
    for i, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip() and para.paragraph_format.alignment != ALIGN_MAP[cn_format["表格"]["align"]]:
                        check_report.append(f"⚠️ 【表格{i+1}】单元格内容对齐方式不符合要求，应为{cn_format['表格']['align']}")
    if not check_report:
        check_report.append("✅ 文档格式完全符合要求，无违规项")
    return check_report

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

def is_white_text(text):
    text_strip = text.strip()
    for word in WHITE_WORDS:
        if word in text_strip:
            return True
    if RE_WHITE_NUMBER.match(text_strip):
        return True
    if RE_WHITE_QUOTE.match(text_strip):
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

def call_doubao_api(text, api_key, prompt):
    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": "doubao-pro",
            "messages": [
                {"role": "system", "content": prompt},
                {"role": "user", "content": text}
            ]
        }
        response = requests.post("https://ark.cn-beijing.volces.com/api/v3/chat/completions", headers=headers, json=payload, timeout=30)
        if response.status_code == 200:
            return response.json()["choices"][0]["message"]["content"].strip(), None
        else:
            return None, f"API调用失败: {response.text}"
    except Exception as e:
        return None, str(e)

def rewrite_sentence(sentence, level_config, api_key=None, forbidden_text=None):
    original = sentence.strip()
    if len(original) < 5 or is_white_text(original):
        return original, "原文保留（白名单/短句）", 1.0
    modified = original
    rewrite_type = "无修改"
    if forbidden_text and original in forbidden_text:
        if api_key:
            result, error = call_doubao_api(original, api_key, "你是一个论文润色专家，请润色这段文本，保持原意，让它不重复，优化表达")
            if not error:
                modified = result
                rewrite_type = "AI针对性润色(规避查重)"
        else:
            parts = [p.strip() for p in RE_CLAUSE_SPLIT.split(modified) if p.strip()]
            if len(parts) >= 3:
                last_part = parts[-1]
                rest_parts = parts[:-1]
                random.shuffle(rest_parts)
                modified = "，".join(rest_parts + [last_part])
                if not modified.endswith(("。", "！", "？", "；")):
                    modified += "。"
                rewrite_type = "针对性语序调整(规避查重)"
    elif api_key:
        result, error = call_doubao_api(original, api_key, "你是一个论文润色专家，请润色这段学术文本，保持原意，优化表达")
        if not error:
            modified = result
            rewrite_type = "AI智能润色"
    if not api_key or rewrite_type == "无修改":
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

def rewrite_paragraph(text, level_config, api_key=None, forbidden_text=None):
    change_log = []
    sentences = RE_SENTENCE_SPLIT.split(text)
    new_sentences = []
    for sent in sentences:
        if not sent.strip():
            new_sentences.append(sent)
            continue
        new_sent, rewrite_type, semantic_score = rewrite_sentence(sent, level_config, api_key, forbidden_text)
        new_sentences.append(new_sent)
        if sent != new_sent:
            change_log.append({
                "original": sent,
                "modified": new_sent,
                "type": rewrite_type,
                "semantic_score": semantic_score
            })
    return "".join(new_sentences), change_log

def process_doc(
    file,
    cn_format,
    en_format,
    enable_rewrite=False,
    rewrite_level="标准润色",
    bind_wps_style=True,
    standardize_ref=True,
    api_key=None,
    forbidden_text=None
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
                new_text, changes = rewrite_paragraph(original_text, rewrite_config, api_key, forbidden_text)
                if changes:
                    total_changes.extend(changes)
                    para.text = new_text
            if standardize_ref:
                new_text, is_ref = standardize_cnki_reference(para.text)
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
            process_log.append(f"✅ 智能润色完成，共修改{len(total_changes)}处")
        if standardize_ref and ref_count > 0:
            process_log.append(f"✅ 知网参考文献标准化完成，共处理{ref_count}条")
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
                                new_text, changes = rewrite_paragraph(original_text, rewrite_config, api_key, forbidden_text)
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
                            run._element.rPr.rFonts.set(qn('w:hAnsi'), en_table_style["en_font"])
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

def generate_report(changes, rewrite_level, title_stats, process_log, check_report):
    total_count = len(changes)
    report = f"# 文档处理报告\n"
    report += f"📅 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    report += f"⚙️ 润色强度：{rewrite_level}\n"
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
    if changes:
        report += "\n## 四、润色修改详情\n"
        for i, change in enumerate(changes[:20]):
            report += f"\n### 修改 {i+1}\n"
            report += f"- **原句**: {change['original']}\n"
            report += f"- **修改**: {change['modified']}\n"
            report += f"- **类型**: {change['type']}\n"
            report += f"- **语义保留**: {change['semantic_score']*100:.1f}%\n"
    return report.encode("utf-8")

# ====================== 兼容rerun函数 ======================
def safe_rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

# ====================== 主程序 ======================
def main():
    # 页面基础配置
    st.set_page_config(
        page_title="智能论文格式处理中心",
        layout="wide",
        page_icon="📝",
        initial_sidebar_state="expanded"
    )
    # 注入全局样式
    set_global_style()

    # 初始化session_state
    if "current_template" not in st.session_state:
        st.session_state.current_template = "三创赛-全国大学生电子商务创新创意及创业挑战赛"
        st.session_state.cn_format, st.session_state.en_format = get_cached_template(st.session_state.current_template)
    if "custom_templates" not in st.session_state:
        st.session_state.custom_templates = {}
    if "version" not in st.session_state:
        st.session_state.version = 0
    if "learned_forbidden" not in st.session_state:
        st.session_state.learned_forbidden = None
    if "learn_history" not in st.session_state:
        st.session_state.learn_history = []
    if "api_key" not in st.session_state:
        st.session_state.api_key = ""

    # ====================== 左侧边栏：配置中心（重构优化） ======================
    with st.sidebar:
        st.markdown("# 📝 配置中心")
        st.markdown("---")
        
        # 1. 核心模板选择（高频操作前置）
        st.markdown("## 1. 模板选择")
        template_options = list(ALL_TEMPLATES.keys()) + list(st.session_state.custom_templates.keys())
        selected_template = st.selectbox(
            "选择目标格式模板",
            options=template_options, 
            index=template_options.index(st.session_state.current_template) if st.session_state.current_template in template_options else 0
        )
        # 模板切换逻辑
        if selected_template != st.session_state.current_template:
            st.session_state.current_template = selected_template
            if selected_template in ALL_TEMPLATES:
                st.session_state.cn_format, st.session_state.en_format = get_cached_template(selected_template)
            else:
                tmp = st.session_state.custom_templates[selected_template]
                st.session_state.cn_format = copy.deepcopy(tmp["cn_format"])
                st.session_state.en_format = copy.deepcopy(tmp["en_format"])
            st.session_state.version += 1
            safe_rerun()
        
        # 模板特殊要求展示
        if selected_template in ALL_TEMPLATES:
            with st.expander("📋 模板官方格式要求", expanded=False):
                for req in ALL_TEMPLATES[selected_template]["special_requirements"]:
                    st.markdown(f"- {req}")
        st.markdown("---")
        
        # 2. 处理参数配置
        st.markdown("## 2. 处理参数配置")
        enable_rewrite = st.checkbox("开启智能润色", value=False)
        rewrite_level = st.selectbox(
            "润色强度",
            options=list(REWRITE_LEVEL.keys()),
            index=1,
            disabled=not enable_rewrite
        )
        bind_wps_style = st.checkbox("绑定WPS标题样式", value=True)
        standardize_ref = st.checkbox("知网参考文献标准化", value=True)
        st.markdown("---")
        
        # 3. 全局API配置（合并重复入口）
        st.markdown("## 3. API配置（可选）")
        st.session_state.api_key = st.text_input(
            "豆包API Key",
            type="password",
            help="用于AI智能润色和查重，全局生效",
            value=st.session_state.api_key
        )
        st.markdown("---")
        
        # 4. 自定义模板管理（低频功能折叠）
        with st.expander("💾 自定义模板管理", expanded=False):
            st.markdown("### 保存当前配置为新模板")
            template_name = st.text_input("新模板名称", placeholder="输入模板名称")
            if st.button("保存为自定义模板", use_container_width=True):
                if template_name:
                    st.session_state.custom_templates[template_name] = {
                        "cn_format": copy.deepcopy(st.session_state.cn_format), 
                        "en_format": copy.deepcopy(st.session_state.en_format),
                        "update_time": datetime.now().strftime('%Y-%m-%d')
                    }
                    st.success(f"✅ 模板「{template_name}」已保存")
                    safe_rerun()
                else:
                    st.error("请输入模板名称")
            
            # 自定义模板管理
            if st.session_state.custom_templates:
                st.markdown("### 已保存的自定义模板")
                for name, info in st.session_state.custom_templates.items():
                    st.markdown(f"- **{name}** | 更新时间：{info['update_time']}")

    # ====================== 主界面：标签页结构重构 ======================
    st.title("📝 智能论文&竞赛格式处理中心")
    st.markdown("一站式完成文档格式标准化、智能润色、查重分析与模板管理")
    st.markdown("---")

    # 核心功能标签页
    tab1, tab2, tab3 = st.tabs([
        "📄 文档批量格式处理",
        "🔍 查重与智能润色",
        "⚙️ 模板参数自定义"
    ])

    # ====================== 标签页1：文档批量格式处理（核心功能） ======================
    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("📁 文档上传")
        files = st.file_uploader(
            "上传 .docx 格式文档（支持批量上传）",
            type=["docx"],
            accept_multiple_files=True
        )
        # 查重学习状态提示
        if st.session_state.learned_forbidden:
            st.info(f"✅ 已学习查重报告，本次润色将自动规避{len(st.session_state.learned_forbidden)}处重复内容")
        st.markdown('</div>', unsafe_allow_html=True)

        # 处理按钮
        if files:
            process_btn = st.button(
                "🚀 开始批量处理",
                type="primary",
                use_container_width=True
            )
            # 处理逻辑
            if process_btn:
                for file in files:
                    with st.spinner(f"正在处理：{file.name}"):
                        try:
                            output_doc, changes, title_stats, process_log, check_report = process_doc(
                                file=file,
                                cn_format=st.session_state.cn_format,
                                en_format=st.session_state.en_format,
                                enable_rewrite=enable_rewrite,
                                rewrite_level=rewrite_level,
                                bind_wps_style=bind_wps_style,
                                standardize_ref=standardize_ref,
                                api_key=st.session_state.api_key,
                                forbidden_text=st.session_state.learned_forbidden
                            )
                            
                            # 处理结果卡片
                            st.markdown('---')
                            st.markdown(f'<div class="card">', unsafe_allow_html=True)
                            st.subheader(f"✅ 处理完成：{file.name}")
                            
                            # 数据看板（优化列数）
                            st.markdown("### 📊 文档统计")
                            m1, m2, m3, m4, m5 = st.columns(5)
                            m1.metric("一级标题", title_stats["一级标题"])
                            m2.metric("二级标题", title_stats["二级标题"])
                            m3.metric("三级标题", title_stats["三级标题"])
                            m4.metric("正文段落", title_stats["正文"])
                            m5.metric("表格数量", title_stats["表格"])
                            
                            # 处理日志与检查报告（分栏展示）
                            log_col, check_col = st.columns(2)
                            with log_col:
                                with st.expander("📋 处理流程日志", expanded=True):
                                    for log in process_log:
                                        st.write(log)
                            with check_col:
                                with st.expander("✅ 格式合规检查报告", expanded=True):
                                    for item in check_report:
                                        st.write(item)
                            
                            # 下载按钮（优化排版）
                            st.markdown("### 📥 下载文件")
                            col_d1, col_d2, col_space = st.columns([2, 2, 2])
                            with col_d1:
                                st.download_button(
                                    label="📥 下载标准格式文档",
                                    data=output_doc,
                                    file_name=f"标准格式_{file.name}",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                            with col_d2:
                                report = generate_report(changes, rewrite_level, title_stats, process_log, check_report)
                                st.download_button(
                                    label="📋 下载详细处理报告",
                                    data=report,
                                    file_name=f"处理报告_{file.name}.txt",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            st.markdown('</div>', unsafe_allow_html=True)

                        except Exception as e:
                            st.error(f"处理失败：{file.name} | 错误信息：{str(e)}")

    # ====================== 标签页2：查重与智能润色 ======================
    with tab2:
        st.markdown("支持在线智能查重、本地查重报告解析、针对性降重润色")
        st.markdown("---")
        # 左右分栏优化
        col_check, col_rewrite = st.columns(2)
        
        # 左侧：在线查重
        with col_check:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("🌐 在线智能查重")
            if not st.session_state.api_key:
                st.warning("请先在左侧配置中心填写豆包API Key")
            # 文档上传
            check_doc_file = st.file_uploader(
                "上传待查重文档（支持 .docx / .txt）",
                type=["docx", "txt"],
                key="check_doc_uploader"
            )
            doc_content = ""
            if check_doc_file:
                doc_content, read_error = read_uploaded_doc_content(check_doc_file)
                if read_error:
                    st.error(read_error)
                else:
                    st.success(f"✅ 文档读取成功，共 {len(doc_content)} 字符")
            # 文本输入框
            check_text = st.text_area(
                "输入待查重文本（上传文档将自动填充）",
                value=doc_content if doc_content else "",
                height=220,
                key="check_text_area"
            )
            # 查重按钮
            if st.button("开始智能查重", use_container_width=True, disabled=not st.session_state.api_key):
                if check_text and st.session_state.api_key:
                    with st.spinner("正在查重分析中..."):
                        result, error = call_doubao_api(
                            check_text,
                            st.session_state.api_key,
                            "你是一个专业的论文查重专家，请分析这段文本的重复情况，给出重复率评估、重复片段定位和降重建议"
                        )
                        if not error:
                            st.success("查重分析完成！")
                            st.markdown(result)
                        else:
                            st.error(f"查重失败：{error}")
                elif not check_text:
                    st.warning("请上传文档或输入待查重文本")
            st.markdown('</div>', unsafe_allow_html=True)
        
        # 右侧：查重报告解析与针对性润色
        with col_rewrite:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("📄 查重报告解析与降重")
            st.info("上传HTML/TXT格式查重报告，系统自动学习标红内容，润色时自动规避重复")
            report_file = st.file_uploader(
                "上传查重报告",
                type=["html", "txt"],
                key="plag_report_uploader"
            )
            if report_file:
                red_parts, plain_text, error = parse_plagiarism_report(report_file)
                if error:
                    st.error(f"解析失败：{error}")
                elif red_parts:
                    st.success(f"解析完成！发现 {len(red_parts)} 处标红重复内容")
                    st.session_state.learned_forbidden = red_parts
                    st.session_state.learn_history.append({
                        "time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        "forbidden_count": len(red_parts)
                    })
                    # 标红内容预览
                    with st.expander("标红内容预览", expanded=False):
                        for i, part in enumerate(red_parts[:10]):
                            st.text(f"{i+1}. {part[:100]}...")
                    # 针对性润色按钮
                    if st.button("针对标红部分一键降重润色", use_container_width=True):
                        with st.spinner("正在针对性降重润色..."):
                            new_text = plain_text
                            change_count = 0
                            for part in red_parts:
                                if len(part) > 10:
                                    modified, _ = rewrite_paragraph(
                                        part,
                                        REWRITE_LEVEL["标准润色"],
                                        st.session_state.api_key,
                                        red_parts
                                    )
                                    new_text = new_text.replace(part, modified)
                                    change_count += 1
                            st.text_area("润色降重后的文本", new_text, height=220)
                            st.success(f"润色完成！共优化 {change_count} 处重复内容")
                            st.download_button(
                                label="📥 下载润色后文本",
                                data=new_text.encode("utf-8"),
                                file_name="降重润色后文本.txt",
                                mime="text/plain",
                                use_container_width=True
                            )
            st.markdown('</div>', unsafe_allow_html=True)

    # ====================== 标签页3：模板参数自定义 ======================
    with tab3:
        st.markdown("自定义当前模板的格式参数，调整后可保存为自定义模板")
        st.markdown("---")
        # 中文字体格式配置
        st.subheader("📝 中文字体格式配置")
        cn_format = st.session_state.cn_format
        # 分栏配置
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("#### 一级标题")
            cn_format["一级标题"]["font"] = st.selectbox("字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cn_format["一级标题"]["font"]), key="h1_font")
            cn_format["一级标题"]["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index(cn_format["一级标题"]["size"]), key="h1_size")
            cn_format["一级标题"]["bold"] = st.checkbox("加粗", value=cn_format["一级标题"]["bold"], key="h1_bold")
            cn_format["一级标题"]["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cn_format["一级标题"]["align"]), key="h1_align")
        with col2:
            st.markdown("#### 二级标题")
            cn_format["二级标题"]["font"] = st.selectbox("字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cn_format["二级标题"]["font"]), key="h2_font")
            cn_format["二级标题"]["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index(cn_format["二级标题"]["size"]), key="h2_size")
            cn_format["二级标题"]["bold"] = st.checkbox("加粗", value=cn_format["二级标题"]["bold"], key="h2_bold")
            cn_format["二级标题"]["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cn_format["二级标题"]["align"]), key="h2_align")
        with col3:
            st.markdown("#### 三级标题")
            cn_format["三级标题"]["font"] = st.selectbox("字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cn_format["三级标题"]["font"]), key="h3_font")
            cn_format["三级标题"]["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index(cn_format["三级标题"]["size"]), key="h3_size")
            cn_format["三级标题"]["bold"] = st.checkbox("加粗", value=cn_format["三级标题"]["bold"], key="h3_bold")
            cn_format["三级标题"]["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cn_format["三级标题"]["align"]), key="h3_align")
        
        # 正文与表格配置
        st.markdown("---")
        col4, col5 = st.columns(2)
        with col4:
            st.markdown("#### 正文")
            cn_format["正文"]["font"] = st.selectbox("字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cn_format["正文"]["font"]), key="body_font")
            cn_format["正文"]["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index(cn_format["正文"]["size"]), key="body_size")
            cn_format["正文"]["bold"] = st.checkbox("加粗", value=cn_format["正文"]["bold"], key="body_bold")
            cn_format["正文"]["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cn_format["正文"]["align"]), key="body_align")
            cn_format["正文"]["line_type"] = st.selectbox("行距类型", ["倍数", "固定值"], index=0 if cn_format["正文"]["line_type"] == "倍数" else 1, key="body_line_type")
            cn_format["正文"]["line_value"] = st.number_input("行距值", value=cn_format["正文"]["line_value"], step=0.1, key="body_line_value")
            cn_format["正文"]["indent"] = st.number_input("首行缩进(字符)", value=cn_format["正文"]["indent"], step=0.5, key="body_indent")
        with col5:
            st.markdown("#### 表格")
            cn_format["表格"]["font"] = st.selectbox("字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cn_format["表格"]["font"]), key="table_font")
            cn_format["表格"]["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), index=list(FONT_SIZE_MAP.keys()).index(cn_format["表格"]["size"]), key="table_size")
            cn_format["表格"]["bold"] = st.checkbox("加粗", value=cn_format["表格"]["bold"], key="table_bold")
            cn_format["表格"]["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cn_format["表格"]["align"]), key="table_align")
        
        # 保存修改
        if st.button("💾 保存当前参数修改", type="primary", use_container_width=True):
            st.session_state.cn_format = cn_format
            st.success("✅ 模板参数已更新，可在左侧保存为自定义模板")
            safe_rerun()

if __name__ == "__main__":
    main()
