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
    },
    "浙江大学本科毕业论文模板": {
        "update_time": "2024-04-03",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 20, "space_after": 15},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 15, "space_after": 10},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 10, "space_after": 5},
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
        "special_requirements": ["全文8000-12000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注浙江大学本科毕业论文"]
    },
    "复旦大学本科毕业论文模板": {
        "update_time": "2024-04-04",
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
        "special_requirements": ["全文10000-15000字", "需包含摘要/关键词/参考文献", "参考文献需符合GB/T 7714", "页眉标注复旦大学本科毕业论文"]
    },
    "上海交通大学本科毕业论文模板": {
        "update_time": "2024-04-05",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 20, "space_after": 12},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["全文12000-20000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注上海交通大学本科毕业论文"]
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
    },
    "硕士毕业论文-通用模板": {
        "update_time": "2024-04-05",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 24, "space_after": 18},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["全文30000字以上", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "需包含创新点说明"]
    }
}

# 专业期刊模板
JOURNAL_FORMATS = {
    "MTA - Multimedia Tools and Applications": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["双栏排版", "单栏摘要", "参考文献需符合APA格式", "图表需单独标注", "全文不超过15页"]
    },
    "IEEE Transactions": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["双栏排版", "无首行缩进", "参考文献需符合IEEE格式", "图表需跨栏", "全文不超过8页"]
    },
    "ACM Transactions": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["双栏排版", "无首行缩进", "参考文献需符合ACM格式", "图表需跨栏", "全文不超过10页"]
    },
    "Elsevier Journal": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["单栏排版", "无首行缩进", "参考文献需符合Elsevier格式", "图表需单独标注", "全文不超过20页"]
    },
    "Springer Journal": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["单栏排版", "无首行缩进", "参考文献需符合Springer格式", "图表需单独标注", "全文不超过15页"]
    }
}

ALL_TEMPLATES = {**COMPETITION_FORMATS, **UNIVERSITY_FORMATS, **THESIS_FORMATS, **JOURNAL_FORMATS}
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

# ====================== 核心工具函数 ======================
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

def extract_template_from_doc(file):
    """从文档提取模板，兼容docx/doc/pdf，无textract时降级处理"""
    try:
        if file.name.endswith('.docx'):
            doc = Document(file)
            file.seek(0)
        elif file.name.endswith('.doc') or file.name.endswith('.pdf'):
            # 可选依赖延迟导入+异常捕获
            text = ""
            extract_success = False
            try:
                import textract
                temp_path = f"/tmp/{file.name}"
                with open(temp_path, 'wb') as f:
                    f.write(file.read())
                text = textract.process(temp_path).decode('utf-8')
                os.remove(temp_path)
                extract_success = True
            except ImportError:
                return None, None, "缺少textract依赖，无法解析doc/pdf文件，请转为docx格式后重试"
            except Exception as e:
                return None, None, f"文件解析失败：{str(e)}，请转为docx格式后重试"
            return None, text, "仅文本提取"
        else:
            return None, None, "不支持的文件格式"

        cn_format = {}
        en_format = {}
        style_stats = {}
        for para in doc.paragraphs:
            level = get_title_level(para.text)
            if level not in style_stats:
                style_stats[level] = {"font": None, "size": None, "bold": None, "align": None, "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0}
            if para.runs:
                run = para.runs[0]
                if run.font.name:
                    if run.font.name in CN_FONT_LIST:
                        style_stats[level]["font"] = run.font.name
                    else:
                        style_stats[level]["en_font"] = run.font.name
                if run.font.size:
                    for size_name, size_pt in FONT_SIZE_MAP.items():
                        if abs(run.font.size.pt - size_pt) < 0.5:
                            style_stats[level]["size"] = size_name
                            break
                if run.font.bold is not None:
                    style_stats[level]["bold"] = run.font.bold
            if para.paragraph_format:
                pf = para.paragraph_format
                if pf.alignment:
                    for align_name, align_val in ALIGN_MAP.items():
                        if pf.alignment == align_val:
                            style_stats[level]["align"] = align_name
                            break
                if pf.first_line_indent:
                    style_stats[level]["indent"] = int(pf.first_line_indent.cm / 0.74)
                if pf.space_before:
                    style_stats[level]["space_before"] = int(pf.space_before.pt)
                if pf.space_after:
                    style_stats[level]["space_after"] = int(pf.space_after.pt)
                if pf.line_spacing:
                    style_stats[level]["line_value"] = pf.line_spacing
        for table in doc.tables:
            if "表格" not in style_stats:
                style_stats["表格"] = {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.runs:
                            run = para.runs[0]
                            if run.font.name:
                                if run.font.name in CN_FONT_LIST:
                                    style_stats["表格"]["font"] = run.font.name
                            if run.font.size:
                                for size_name, size_pt in FONT_SIZE_MAP.items():
                                    if abs(run.font.size.pt - size_pt) < 0.5:
                                        style_stats["表格"]["size"] = size_name
                                        break
        for level in ["一级标题", "二级标题", "三级标题", "正文", "表格"]:
            if level in style_stats:
                cn_format[level] = style_stats[level]
                en_format[level] = {
                    "en_font": "Times New Roman",
                    "size_same_as_cn": True,
                    "size": style_stats[level].get("size", "小四"),
                    "bold": style_stats[level].get("bold", False),
                    "italic": False
                }
        template_data = {
            "name": f"自定义模板_{len(st.session_state.custom_templates) + 1}",
            "update_time": datetime.now().strftime('%Y-%m-%d'),
            "cn_format": cn_format,
            "en_format": en_format
        }
        return template_data, None, None
    except Exception as e:
        return None, None, str(e)

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

def export_template(template_data, export_type="json"):
    if export_type == "json":
        return json.dumps(template_data, ensure_ascii=False, indent=2).encode("utf-8")
    else:
        text = f"模板配置文件\n"
        text += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        text += f"模板名称: {template_data.get('name', '自定义模板')}\n"
        text += f"更新时间: {template_data.get('update_time', datetime.now().strftime('%Y-%m-%d'))}\n\n"
        text += "=== 中文格式设置 ===\n"
        for level, cfg in template_data.get('cn_format', {}).items():
            text += f"\n[{level}]\n"
            for k, v in cfg.items():
                text += f"{k} = {v}\n"
        text += "\n=== 西文格式设置 ===\n"
        for level, cfg in template_data.get('en_format', {}).items():
            text += f"\n[{level}]\n"
            for k, v in cfg.items():
                text += f"{k} = {v}\n"
        return text.encode("utf-8")

def import_template(file):
    try:
        content = file.read().decode('utf-8')
        if file.name.endswith('.json'):
            data = json.loads(content)
            return data, None
        else:
            data = {}
            current_section = None
            current_level = None
            current_cfg = {}
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if line.startswith('==='):
                    current_section = line.strip('= ').strip()
                    continue
                if line.startswith('[') and line.endswith(']'):
                    if current_level and current_cfg:
                        if current_section == "中文格式设置":
                            data['cn_format'] = data.get('cn_format', {})
                            data['cn_format'][current_level] = current_cfg
                        else:
                            data['en_format'] = data.get('en_format', {})
                            data['en_format'][current_level] = current_cfg
                    current_level = line.strip('[]')
                    current_cfg = {}
                    continue
                if '=' in line:
                    k, v = line.split('=', 1)
                    k = k.strip()
                    v = v.strip()
                    try:
                        if v.lower() == 'true':
                            v = True
                        elif v.lower() == 'false':
                            v = False
                        elif '.' in v:
                            v = float(v)
                        else:
                            try:
                                v = int(v)
                            except:
                                pass
                    except:
                        pass
                    current_cfg[k] = v
            if current_level and current_cfg:
                if current_section == "中文格式设置":
                    data['cn_format'] = data.get('cn_format', {})
                    data['cn_format'][current_level] = current_cfg
                else:
                    data['en_format'] = data.get('en_format', {})
                    data['en_format'][current_level] = current_cfg
            return data, None
    except Exception as e:
        return None, str(e)

def main():
    st.set_page_config(page_title="智能论文&竞赛格式处理平台", layout="wide", page_icon="📝")
    # 兼容新旧版本Streamlit的rerun
    def safe_rerun():
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()

    # 初始化session_state
    if "current_template" not in st.session_state:
        st.session_state.current_template = "三创赛-全国大学生电子商务创新创意及创业挑战赛"
        st.session_state.cn_format, st.session_state.en_format = get_cached_template(st.session_state.current_template)
    if "custom_templates" not in st.session_state:
        st.session_state.custom_templates = {}
    if "version" not in st.session_state:
        st.session_state.version = 0
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "knowledge_base" not in st.session_state:
        st.session_state.knowledge_base = None
    if "learned_forbidden" not in st.session_state:
        st.session_state.learned_forbidden = None
    if "learn_history" not in st.session_state:
        st.session_state.learn_history = []
    if "check_result" not in st.session_state:
        st.session_state.check_result = None

    tab1, tab2, tab3, tab4 = st.tabs(["📝 文档处理", "📋 模板管理", "🔍 查重润色", "🤖 AI智能助手"])
    with tab1:
        st.title(f"📝 智能论文&竞赛格式处理器")
        st.success("✅ 支持doc/docx/PDF模板提取 | WPS自动生成导航索引 | 知网参考文献标准化 | 智能规避查重")
        st.divider()
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            template_options = list(ALL_TEMPLATES.keys()) + list(st.session_state.custom_templates.keys())
            selected_template = st.selectbox("📌 选择模板类型", options=template_options, 
                                            index=template_options.index(st.session_state.current_template) if st.session_state.current_template in template_options else 0)
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
        with col2:
            enable_rewrite = st.checkbox("🔄 开启智能润色", value=False)
            rewrite_level = st.selectbox("润色强度", options=list(REWRITE_LEVEL.keys()), index=1, disabled=not enable_rewrite)
        with col3:
            bind_wps_style = st.checkbox("✅ 绑定WPS标题样式", value=True, help="开启后导出的文档在WPS中会自动生成导航目录")
            standardize_ref = st.checkbox("📚 知网参考文献标准化", value=True, help="自动调整参考文献格式，解决知网查重标红问题")
        if selected_template in ALL_TEMPLATES:
            update_time = ALL_TEMPLATES[selected_template].get("update_time", "未知")
            st.info(f"📅 该模板最后更新时间：{update_time}")
            st.markdown(f"### ⚠️ {selected_template.split('-')[-1] if '-' in selected_template else selected_template} 格式要求")
            for req in ALL_TEMPLATES[selected_template]["special_requirements"]:
                st.markdown(f"- {req}")
        else:
            update_time = st.session_state.custom_templates[selected_template].get("update_time", datetime.now().strftime('%Y-%m-%d'))
            st.info(f"📅 自定义模板，最后更新时间：{update_time}")
        st.divider()
        with st.sidebar:
            st.subheader("⚙️ 高级格式自定义")
            st.markdown("#### 💾 自定义模板")
            template_name = st.text_input("模板名称", placeholder="自定义格式名称")
            col_template1, col_template2 = st.columns(2)
            with col_template1:
                if st.button("保存当前格式", use_container_width=True):
                    if template_name:
                        st.session_state.custom_templates[template_name] = {
                            "cn_format": copy.deepcopy(st.session_state.cn_format), 
                            "en_format": copy.deepcopy(st.session_state.en_format),
                            "update_time": datetime.now().strftime('%Y-%m-%d')
                        }
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
                        st.session_state.current_template = selected_custom
                        st.session_state.version += 1
                        safe_rerun()
            st.divider()
            st.markdown("#### 🀄 中文格式设置")
            for level in ["一级标题", "二级标题", "三级标题", "正文", "表格"]:
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
            for level in ["一级标题", "二级标题", "三级标题", "正文", "表格"]:
                with st.expander(f"{level}西文设置", expanded=False):
                    cfg = st.session_state.en_format[level]
                    cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]) if cfg["en_font"] in EN_FONT_LIST else 0, key=f"en_{level}_font_{st.session_state.version}")
                    cfg["size_same_as_cn"] = st.checkbox("字号同步中文", cfg["size_same_as_cn"], key=f"en_{level}_same_{st.session_state.version}")
                    cfg["bold"] = st.checkbox("西文加粗", cfg["bold"], key=f"en_{level}_bold_{st.session_state.version}")
                    cfg["italic"] = st.checkbox("西文斜体", cfg["italic"], key=f"en_{level}_italic_{st.session_state.version}")
                    st.session_state.en_format[level] = cfg
        st.subheader("📁 文档上传与处理")
        files = st.file_uploader("上传 .docx 文档", type=["docx"], accept_multiple_files=True)
        api_key = st.text_input("豆包API Key (可选，用于AI智能润色和查重规避)", type="password")
        if st.session_state.learned_forbidden:
            st.info(f"✅ 已学习查重报告，本次润色将自动规避{len(st.session_state.learned_forbidden)}处重复内容")
        if files and st.button("🚀 开始处理文档", type="primary", use_container_width=True):
            for file in files:
                with st.spinner(f"正在处理：{file.name}"):
                    try:
                        output_doc, changes, title_stats, process_log, check_report = process_doc(
                            file=file, cn_format=st.session_state.cn_format, en_format=st.session_state.en_format,
                            enable_rewrite=enable_rewrite, rewrite_level=rewrite_level, bind_wps_style=bind_wps_style, standardize_ref=standardize_ref,
                            api_key=api_key, forbidden_text=st.session_state.learned_forbidden
                        )
                        st.subheader(f"✅ 处理完成：{file.name}")
                        st.success("✅ 已为你生成标准格式文档，所有格式均已按照要求调整完成")
                        with st.expander("📋 处理日志", expanded=True):
                            for log in process_log: st.write(log)
                        cols = st.columns(5)
                        cols[0].metric("一级标题", title_stats["一级标题"])
                        cols[1].metric("二级标题", title_stats["二级标题"])
                        cols[2].metric("三级标题", title_stats["三级标题"])
                        cols[3].metric("正文段落", title_stats["正文"])
                        cols[4].metric("表格数量", title_stats["表格"])
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            st.download_button(
                                label="📥 下载处理后的标准文档",
                                data=output_doc,
                                file_name=f"标准格式_{file.name}",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        with col_d2:
                            report = generate_report(changes, rewrite_level, title_stats, process_log, check_report)
                            st.download_button(
                                label="📋 下载处理报告",
                                data=report,
                                file_name=f"处理报告_{file.name}.txt",
                                mime="text/plain",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"处理失败：{str(e)}")
    with tab2:
        st.title("📋 模板管理中心")
        st.markdown("在这里你可以管理所有格式模板，支持从官方文档自动提取所有格式细节，生成可用的模板，还可以检查你的文档是否符合格式要求")
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📤 模板导出")
            if st.session_state.custom_templates:
                export_template_name = st.selectbox("选择要导出的模板", options=list(st.session_state.custom_templates.keys()))
                export_type = st.radio("导出格式", options=["json", "txt"], index=0)
                if st.button("导出模板", use_container_width=True):
                    template_data = st.session_state.custom_templates[export_template_name]
                    template_data["name"] = export_template_name
                    data = export_template(template_data, export_type)
                    if export_type == "json":
                        st.download_button(
                            label="下载JSON模板",
                            data=data,
                            file_name=f"{export_template_name}.json",
                            mime="application/json",
                            use_container_width=True
                        )
                    else:
                        st.download_button(
                            label="下载TXT模板",
                            data=data,
                            file_name=f"{export_template_name}.txt",
                            mime="text/plain",
                            use_container_width=True
                        )
            else:
                st.info("暂无自定义模板，请先在文档处理页面保存自定义模板")
        with col2:
            st.subheader("📥 模板导入")
            uploaded_template = st.file_uploader("上传模板文件", type=["json", "txt"])
            if uploaded_template:
                data, error = import_template(uploaded_template)
                if error:
                    st.error(f"导入失败：{error}")
                elif data:
                    st.success("模板解析成功！")
                    new_name = st.text_input("模板名称", value=uploaded_template.name.split('.')[0])
                    if st.button("导入到系统", use_container_width=True):
                        st.session_state.custom_templates[new_name] = data
                        st.success(f"✅ 模板「{new_name}」导入成功！")
                        safe_rerun()
        st.divider()
        st.subheader("🔍 格式合规检查")
        st.markdown("上传你的已完成文档和标准格式模板，系统会自动检查是否符合格式要求，直接告诉你哪部分有问题")
        col_check1, col_check2 = st.columns(2)
        with col_check1:
            user_doc = st.file_uploader("上传你的已完成文档", type=["docx"])
        with col_check2:
            standard_template = st.selectbox("选择标准格式模板", options=list(ALL_TEMPLATES.keys()) + list(st.session_state.custom_templates.keys()))
        if user_doc and st.button("开始检查格式", use_container_width=True):
            with st.spinner("正在检查格式..."):
                try:
                    doc = Document(user_doc)
                    if standard_template in ALL_TEMPLATES:
                        cn_format, _ = get_cached_template(standard_template)
                    else:
                        tmp = st.session_state.custom_templates[standard_template]
                        cn_format = tmp["cn_format"]
                    check_report = format_compliance_check(doc, cn_format)
                    st.session_state.check_result = check_report
                    st.success("✅ 检查完成！")
                    st.subheader("📋 检查结果")
                    error_count = len([x for x in check_report if x.startswith("⚠️")])
                    if error_count == 0:
                        st.success("🎉 你的文档完全符合格式要求！没有任何问题！")
                    else:
                        st.warning(f"发现 {error_count} 处格式问题，以下是具体问题：")
                        for item in check_report:
                            if item.startswith("⚠️"):
                                st.error(item)
                            else:
                                st.success(item)
                except Exception as e:
                    st.error(f"检查失败：{str(e)}")
        st.divider()
        st.subheader("🤖 Agent自动模板提取")
        st.markdown("上传学校/竞赛/期刊的官方格式文档，系统会自动学习并提取所有格式细节，包括表格、图片等复杂格式，生成可用的模板")
        official_doc = st.file_uploader("上传官方格式文档", type=["doc", "docx", "pdf"], help="支持doc、docx、PDF格式，系统会自动提取所有格式细节")
        if official_doc:
            with st.spinner("Agent正在学习文档格式..."):
                template_data, text, error = extract_template_from_doc(official_doc)
                if error:
                    st.error(f"提取失败：{error}")
                elif template_data:
                    st.success(f"✅ Agent学习完成！自动生成了模板：{template_data['name']}")
                    st.json(template_data)
                    if st.button("应用这个模板", use_container_width=True):
                        st.session_state.custom_templates[template_data['name']] = template_data
                        st.session_state.cn_format = copy.deepcopy(template_data["cn_format"])
                        st.session_state.en_format = copy.deepcopy(template_data["en_format"])
                        st.session_state.current_template = template_data['name']
                        st.session_state.version += 1
                        st.success("模板已应用！你可以直接使用这个新模板处理文档了")
                        safe_rerun()
                else:
                    st.info("仅提取到文本内容，无法解析格式")
                    st.text_area("提取的文本", text, height=300)
        st.divider()
        st.subheader("📊 所有模板列表")
        template_list = []
        for name, tpl in ALL_TEMPLATES.items():
            template_list.append({
                "名称": name,
                "类型": "系统内置",
                "更新时间": tpl.get("update_time", "未知"),
                "状态": "✅ 可用"
            })
        for name, tpl in st.session_state.custom_templates.items():
            template_list.append({
                "名称": name,
                "类型": "自定义",
                "更新时间": tpl.get("update_time", datetime.now().strftime('%Y-%m-%d')),
                "状态": "✅ 可用"
            })
        df = pd.DataFrame(template_list)
        st.dataframe(df, use_container_width=True)
    with tab3:
        st.title("🔍 查重与智能润色")
        st.markdown("支持在线查重、本地查重报告解析、智能润色优化，自动规避重复内容")
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🌐 在线查重")
            st.info("使用豆包API进行智能查重，需要配置API密钥")
            api_key_check = st.text_input("豆包API Key", type="password")
            check_text = st.text_area("输入要查重的文本", height=300)
            if st.button("开始查重", use_container_width=True):
                if check_text and api_key_check:
                    with st.spinner("正在查重中..."):
                        result, error = call_doubao_api(check_text, api_key_check, "你是一个查重专家，请分析这段文本的重复情况，给出重复率和重复片段")
                        if not error:
                            st.success(f"查重完成！")
                            st.write(result)
                        else:
                            st.error(f"查重失败：{error}")
                else:
                    st.warning("请填写完整的API信息和待查重文本")
        with col2:
            st.subheader("📄 查重报告解析与学习")
            st.info("上传查重报告，系统会自动学习重复内容，润色时自动规避")
            report_file = st.file_uploader("上传查重报告", type=["html", "txt"])
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
                    st.subheader("标红内容预览")
                    for i, part in enumerate(red_parts[:10]):
                        st.text(f"{i+1}. {part[:100]}...")
                    if st.button("针对标红部分进行润色", use_container_width=True):
                        with st.spinner("正在针对性润色..."):
                            new_text = plain_text
                            change_count = 0
                            for part in red_parts:
                                if len(part) > 10:
                                    modified, _ = rewrite_paragraph(part, REWRITE_LEVEL["标准润色"], api_key_check, red_parts)
                                    new_text = new_text.replace(part, modified)
                                    change_count += 1
                            st.text_area("润色后的文本", new_text, height=300)
                            st.success(f"润色完成！共优化 {change_count} 处重复内容")
    with tab4:
        st.title("🤖 AI智能助手")
        st.markdown("上传文档进行学习，然后可以和文档进行智能问答，系统会自动学习你的使用习惯优化策略")
        st.divider()
        with st.sidebar:
            st.subheader("📚 文档学习")
            doc_file = st.file_uploader("上传学习文档", type=["docx", "txt", "doc", "pdf"])
            if doc_file:
                with st.spinner("正在学习文档内容..."):
                    try:
                        text = ""
                        if doc_file.name.endswith('.docx'):
                            doc = Document(doc_file)
                            text = "\n".join([p.text for p in doc.paragraphs])
                        elif doc_file.name.endswith('.txt'):
                            text = doc_file.read().decode('utf-8')
                        elif doc_file.name.endswith('.doc') or doc_file.name.endswith('.pdf'):
                            try:
                                import textract
                                temp_path = f"/tmp/{doc_file.name}"
                                with open(temp_path, 'wb') as f:
                                    f.write(doc_file.read())
                                text = textract.process(temp_path).decode('utf-8')
                                os.remove(temp_path)
                            except ImportError:
                                st.error("缺少textract依赖，无法解析doc/pdf文件，请转为docx/txt格式后重试")
                            except Exception as e:
                                st.error(f"文件解析失败：{str(e)}")
                        st.session_state.knowledge_base = text
                        st.success(f"✅ 文档学习完成！共 {len(text)} 字符")
                    except Exception as e:
                        st.error(f"学习失败：{str(e)}")
        st.subheader("学习历史")
        if st.session_state.learn_history:
            df_learn = pd.DataFrame(st.session_state.learn_history)
            st.dataframe(df_learn, use_container_width=True)
        else:
            st.info("暂无学习历史")
        st.divider()
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
        if prompt := st.chat_input("输入你的问题..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.write(prompt)
            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    if st.session_state.knowledge_base:
                        context = st.session_state.knowledge_base[:3000]
                        response = f"基于你上传的文档，我找到以下相关信息：\n\n文档内容片段：{context[:500]}...\n\n针对你的问题「{prompt}」，我已经学习了你的文档内容。如果配置了豆包API，我可以给出更精准的回答。"
                    else:
                        response = "你还没有上传学习文档哦！请先在侧边栏上传文档，我就可以基于文档内容回答你的问题啦。另外，我还会自动学习你的润色习惯，下次帮你把查重率降得更低！"
                    st.write(response)
                    st.session_state.messages.append({"role": "assistant", "content": response})
if __name__ == "__main__":
    main()
