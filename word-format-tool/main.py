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
# ====================== 预编译正则（核心逻辑完整保留）======================
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
# ====================== 全局配置与常量（完整保留+新增扩展项）======================
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
# 全量模板库（完整保留原模板+新增图片图注、参考文献格式项）
COMPETITION_FORMATS = {
    "三创赛-全国大学生电子商务创新创意及创业挑战赛": {
        "update_time": "2024-01-15",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "楷体_GB2312", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "仿宋", "size": "四号", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.2, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["硬件必须配小程序/App", "服务必须线上化", "需要3D建模图/UI原型", "图表必须标注数据来源"]
    },
    "挑战杯-全国大学生课外学术科技作品竞赛": {
        "update_time": "2024-02-20",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文约15000字", "双面打印", "严格章-节-条层级结构", "标题单倍行距，正文1.5倍行距"]
    },
    "互联网+大学生创新创业大赛": {
        "update_time": "2024-03-10",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "四号", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文10000字以上", "分创意组/创业组撰写", "需包含完整财务预测", "商业模式需清晰可落地"]
    }
}
UNIVERSITY_FORMATS = {
    "清华大学本科毕业论文模板": {
        "update_time": "2024-04-01",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 24, "space_after": 18, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文8000-15000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注清华大学本科毕业论文"]
    },
    "北京大学本科毕业论文模板": {
        "update_time": "2024-04-02",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文10000-20000字", "需包含摘要/关键词/参考文献", "参考文献需符合GB/T 7714", "页眉标注北京大学本科毕业论文"]
    },
    "浙江大学本科毕业论文模板": {
        "update_time": "2024-04-03",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 20, "space_after": 15, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 15, "space_after": 10, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 10, "space_after": 5, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文8000-12000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注浙江大学本科毕业论文"]
    },
    "复旦大学本科毕业论文模板": {
        "update_time": "2024-04-04",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文10000-15000字", "需包含摘要/关键词/参考文献", "参考文献需符合GB/T 7714", "页眉标注复旦大学本科毕业论文"]
    },
    "上海交通大学本科毕业论文模板": {
        "update_time": "2024-04-05",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 20, "space_after": 12, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文12000-20000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注上海交通大学本科毕业论文"]
    }
}
THESIS_FORMATS = {
    "本科毕业论文-通用模板": {
        "update_time": "2024-04-01",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "三号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文8000-12000字", "需包含摘要/关键词/参考文献/致谢", "参考文献需符合GB/T 7714格式", "页眉需标注学校+论文题目"]
    },
    "硕士毕业论文-通用模板": {
        "update_time": "2024-04-05",
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 24, "space_after": 18, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 2, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文30000字以上", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "需包含创新点说明"]
    }
}
JOURNAL_FORMATS = {
    "MTA - Multimedia Tools and Applications": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["双栏排版", "单栏摘要", "参考文献需符合APA格式", "图表需单独标注", "全文不超过15页"]
    },
    "IEEE Transactions": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["双栏排版", "无首行缩进", "参考文献需符合IEEE格式", "图表需跨栏", "全文不超过8页"]
    },
    "ACM Transactions": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["双栏排版", "无首行缩进", "参考文献需符合ACM格式", "图表需跨栏", "全文不超过10页"]
    },
    "Elsevier Journal": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["单栏排版", "无首行缩进", "参考文献需符合Elsevier格式", "图表需单独标注", "全文不超过20页"]
    },
    "Springer Journal": {
        "update_time": "2024-04-10",
        "cn_format": {
            "一级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "宋体", "size": "小四", "bold": True, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "两端对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "图片与图注": {"font": "宋体", "size": "小五", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 6, "char_spacing": 0},
            "参考文献": {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 3, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False},
            "图片与图注": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False},
            "参考文献": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["单栏排版", "无首行缩进", "参考文献需符合Springer格式", "图表需单独标注", "全文不超过15页"]
    }
}
ALL_TEMPLATES = {**COMPETITION_FORMATS, **UNIVERSITY_FORMATS, **THESIS_FORMATS, **JOURNAL_FORMATS}
# 润色等级配置
REWRITE_LEVEL = {
    "轻度润色": {"synonym": True, "sentence_reorder": False, "structure_change": False, "desc": "仅同义词替换，保留原文句式，语义保留度100%"},
    "标准润色": {"synonym": True, "sentence_reorder": True, "structure_change": False, "desc": "同义词替换+句式重构，语义保留度≥90%"},
    "深度润色": {"synonym": True, "sentence_reorder": True, "structure_change": True, "desc": "句式重构+段落结构优化+AI重写，语义保留度≥70%"}
}
SYNONYM_DICT = {
    "提升": "有效改善", "降低": "显著减少", "增加": "大幅提升", "减少": "有效降低",
    "首先": "从实际落地情况来看", "其次": "进一步分析", "最后": "综合上述分析",
    "综上所述": "结合全维度分析", "总而言之": "从实践结果来看",
    "研究": "调研分析", "实验": "测试验证", "分析": "剖析", "结果": "结论",
    "方法": "方案", "系统": "平台", "模型": "架构", "数据": "信息"
}
# 格式映射常量
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
CN_FONT_LIST = ["宋体", "黑体", "楷体_GB2312", "仿宋_GB2312", "微软雅黑"]
MAX_FILE_SIZE_MB = 200
random.seed(42)
# ====================== 核心工具函数（100%保留原功能）======================
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
    elif RE_REF_KEYWORD.search(text):
        return "参考文献"
    elif re.match(r'^图\s*\d+', text) or re.match(r'^表\s*\d+', text):
        return "图片与图注"
    else:
        return "正文"
def extract_template_from_doc(file):
    try:
        if file.name.endswith('.docx'):
            doc = Document(file)
            file.seek(0)
        elif file.name.endswith('.doc') or file.name.endswith('.pdf'):
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
                style_stats[level] = {"font": "宋体", "size": "小四", "bold": False, "italic": False, "color": "#000000", "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0}
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
                if run.font.italic is not None:
                    style_stats[level]["italic"] = run.font.italic
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
                style_stats["表格"] = {"font": "宋体", "size": "五号", "bold": False, "italic": False, "color": "#000000", "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0}
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
        for level in ["一级标题", "二级标题", "三级标题", "正文", "表格", "图片与图注", "参考文献"]:
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
def optimize_image_layout(doc, img_format):
    image_count = 0
    for para in doc.paragraphs:
        has_image = False
        for run in para.runs:
            if run._element.xpath('.//a:blip'):
                has_image = True
                image_count += 1
                break
        if has_image or get_title_level(para.text) == "图片与图注":
            para.alignment = ALIGN_MAP[img_format["align"]]
            para_format = para.paragraph_format
            para_format.space_before = Pt(img_format["space_before"])
            para_format.space_after = Pt(img_format["space_after"])
            para_format.keep_with_next = True
            para_format.keep_together = True
            para_format.first_line_indent = Cm(img_format["indent"] * 0.74)
            if img_format["line_type"] == "固定值":
                para_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                para_format.line_spacing = Pt(img_format["line_value"])
            else:
                para_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                para_format.line_spacing = img_format["line_value"]
            cn_size_pt = FONT_SIZE_MAP.get(img_format["size"], 9)
            for run in para.runs:
                run.font.name = img_format["font"]
                run._element.rPr.rFonts.set(qn('w:eastAsia'), img_format["font"])
                run.font.size = Pt(cn_size_pt)
                run.font.bold = img_format["bold"]
                run.font.italic = img_format["italic"]
                run.font.color.rgb = RGBColor.from_string(img_format["color"].lstrip('#'))
                if img_format.get("char_spacing", 0) > 0:
                    run.font.spacing = Pt(img_format["char_spacing"])
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
def simulate_check_rate(text):
    """模拟查重率计算，可替换为真实API"""
    words = RE_KEYWORDS.findall(text)
    if not words:
        return 10.0
    repeat_count = sum(1 for w in words if w in WHITE_WORDS)
    rate = min(40, max(5, repeat_count / len(words) * 100))
    return round(rate, 1)
def process_doc(
    file,
    cn_format,
    en_format,
    enable_rewrite=False,
    rewrite_level="标准润色",
    bind_wps_style=True,
    standardize_ref=True,
    optimize_image=True,
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
    title_stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": len(doc.tables), "图片与图注": 0, "参考文献": 0}
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
            if standardize_ref and level == "参考文献":
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
            en_size_pt = FONT_SIZE_MAP.get(en_style["size"], 12) if not en_style["size_same_as_cn"] else cn_size_pt
            font_color = RGBColor.from_string(cn_style["color"].lstrip('#'))
            for run in para.runs:
                run.font.name = cn_style["font"]
                run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_style["font"])
                run._element.rPr.rFonts.set(qn('w:ascii'), en_style["en_font"])
                run._element.rPr.rFonts.set(qn('w:hAnsi'), en_style["en_font"])
                run.font.size = Pt(cn_size_pt)
                run.font.bold = en_style["bold"] if en_style["bold"] is not None else cn_style["bold"]
                run.font.italic = en_style["italic"] if en_style["italic"] is not None else cn_style["italic"]
                run.font.color.rgb = font_color
                if cn_style.get("char_spacing", 0) > 0:
                    run.font.spacing = Pt(cn_style["char_spacing"])
        process_log.append("✅ 全文档段落格式处理完成")
        if enable_rewrite:
            process_log.append(f"✅ 智能润色完成，共修改{len(total_changes)}处")
        if standardize_ref and ref_count > 0:
            process_log.append(f"✅ 知网参考文献标准化完成，共处理{ref_count}条")
        process_log.append(f"📊 标题识别结果：一级{title_stats['一级标题']}、二级{title_stats['二级标题']}、三级{title_stats['三级标题']}、参考文献{title_stats['参考文献']}条")
    except Exception as e:
        raise Exception(f"文档处理失败：{str(e)}")
    try:
        if optimize_image:
            image_count = optimize_image_layout(doc, cn_format["图片与图注"])
            title_stats["图片与图注"] = image_count
            if image_count > 0:
                process_log.append(f"✅ 优化{image_count}张图片与图注排版")
            else:
                process_log.append("✅ 未检测到图片")
    except Exception as e:
        process_log.append(f"⚠️ 图片处理失败：{str(e)}")
    try:
        cn_table_style = cn_format["表格"]
        en_table_style = en_format["表格"]
        table_cn_size = FONT_SIZE_MAP.get(cn_table_style["size"], 10.5)
        table_en_size = FONT_SIZE_MAP.get(en_table_style["size"], 10.5) if not en_table_style["size_same_as_cn"] else table_cn_size
        table_color = RGBColor.from_string(cn_table_style["color"].lstrip('#'))
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
                            run.font.bold = en_table_style["bold"] if en_table_style["bold"] is not None else cn_table_style["bold"]
                            run.font.italic = en_table_style["italic"] if en_table_style["italic"] is not None else cn_table_style["italic"]
                            run.font.color.rgb = table_color
                            if cn_table_style.get("char_spacing", 0) > 0:
                                run.font.spacing = Pt(cn_table_style["char_spacing"])
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
    full_text = "\n".join([p.text for p in doc.paragraphs])
    return output, total_changes, title_stats, process_log, check_report, full_text
def generate_report(changes, rewrite_level, title_stats, process_log, check_report):
    total_count = len(changes)
    report = f"# 文档处理报告\n"
    report += f"📅 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    report += f"⚙️ 润色强度：{rewrite_level}\n"
    report += f"📝 总修改条数：{total_count}\n\n"
    report += "## 一、处理流程日志\n"
    for log in process_log:
        report += f"- {log}\n"
    report += "\n## 二、文档结构统计\n"
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
def safe_rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()
# ====================== Session状态初始化 ======================
def init_session_state():
    default_cn_format, default_en_format = get_cached_template("本科毕业论文-通用模板")
    if "custom_cn_format" not in st.session_state:
        st.session_state.custom_cn_format = copy.deepcopy(default_cn_format)
    if "custom_en_format" not in st.session_state:
        st.session_state.custom_en_format = copy.deepcopy(default_en_format)
    if "custom_templates" not in st.session_state:
        st.session_state.custom_templates = {}
    if "format_version" not in st.session_state:
        st.session_state.format_version = 0
    if "need_polish" not in st.session_state:
        st.session_state.need_polish = False
    if "learned_forbidden" not in st.session_state:
        st.session_state.learned_forbidden = None
    if "formatted_doc" not in st.session_state:
        st.session_state.formatted_doc = None
    if "formatted_report" not in st.session_state:
        st.session_state.formatted_report = None
    if "polish_doc" not in st.session_state:
        st.session_state.polish_doc = None
    if "polish_report" not in st.session_state:
        st.session_state.polish_report = None
    if "check_rate" not in st.session_state:
        st.session_state.check_rate = None
    if "original_check_rate" not in st.session_state:
        st.session_state.original_check_rate = None
    if "doc_full_text" not in st.session_state:
        st.session_state.doc_full_text = ""
    if "process_timestamp" not in st.session_state:
        st.session_state.process_timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    if "selected_template" not in st.session_state:
        st.session_state.selected_template = "本科毕业论文-通用模板"
# ====================== 主应用UI（核心修复布局与主题适配）======================
def main():
    # 页面基础配置
    st.set_page_config(
        page_title="智能论文&竞赛格式处理平台",
        layout="wide",
        page_icon="📝",
        initial_sidebar_state="collapsed"
    )
    # ========== 核心修复：全局CSS适配深色模式+左侧栏无滚动布局 ==========
    st.markdown("""
    <style>
    /* 全局适配，消除横向滚动，适配主题 */
    .stApp {
        min-width: 1200px;
        overflow-x: hidden;
        background-color: var(--background-color);
    }
    /* 左侧栏：flex垂直布局，无全局滚动，适配全视口高度 */
    .left-column {
        position: sticky;
        top: 0;
        height: 100vh;
        display: flex;
        flex-direction: column;
        gap: 0.5rem;
        padding-right: 1rem;
        border-right: 1px solid var(--border-color, #e5e7eb);
        overflow: hidden; /* 禁止整个左侧栏滚动 */
    }
    /* 左侧栏三个模块的flex分配 */
    .left-top-block {
        flex: 0 0 auto; /* 高度自适应内容，不伸缩 */
        overflow: visible;
    }
    .left-middle-block {
        flex: 1 1 auto; /* 占满剩余空间，内容溢出时内部滚动 */
        overflow-y: auto;
        overflow-x: hidden;
        padding-right: 0.25rem;
    }
    .left-bottom-block {
        flex: 0 0 auto; /* 高度自适应内容，不伸缩 */
        overflow: visible;
    }
    /* 右侧栏流式布局 */
    .right-column {
        overflow-y: auto;
        padding-left: 1rem;
    }
    /* 固定顶部标题栏，适配主题色，解决黑块问题 */
    .fixed-header {
        position: sticky;
        top: 0;
        background-color: var(--background-color);
        z-index: 999;
        padding: 1rem 0;
        border-bottom: 1px solid var(--border-color, #e5e7eb);
        margin-bottom: 1rem;
    }
    /* 模块分割线，适配主题 */
    .module-divider-green {
        margin: 1rem 0;
        border: none;
        border-top: 3px solid #10b981;
    }
    .module-divider-blue {
        margin: 1rem 0;
        border: none;
        border-top: 3px solid #3b82f6;
    }
    .module-divider-gray {
        margin: 1rem 0;
        border: none;
        border-top: 3px solid var(--border-color, #6b7280);
    }
    /* 组件样式统一，适配主题 */
    .stButton>button {
        border-radius: 8px;
        width: 100%;
        font-weight: 500;
    }
    .stFileUploader>div {
        border-radius: 8px;
        border: 1px dashed var(--border-color, #4b5563);
        background-color: var(--secondary-background-color, #f9fafb);
    }
    .stSelectbox>div>div {
        border-radius: 8px;
    }
    .stTextInput>div>div {
        border-radius: 8px;
    }
    .stExpander {
        border-radius: 8px;
        border: 1px solid var(--border-color, #e5e7eb);
    }
    /* 进度条样式 */
    .stProgress>div>div {
        background-color: #10b981;
    }
    /* 滚动条美化，适配深色模式 */
    ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
    }
    ::-webkit-scrollbar-track {
        background: var(--secondary-background-color, #f1f1f1);
        border-radius: 3px;
    }
    ::-webkit-scrollbar-thumb {
        background: var(--border-color, #c1c1c1);
        border-radius: 3px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: var(--text-color, #888);
    }
    </style>
    """, unsafe_allow_html=True)
    # 初始化状态
    init_session_state()
    # 全局分栏：左1右4 严格遵循
    left_col, right_col = st.columns([1, 4])
    # ============== 左栏：纯自定义模板生成工作台（无全局滚动）==============
    with left_col:
        st.markdown('<div class="left-column">', unsafe_allow_html=True)
        
        # 上模块：模板命名与保存区（flex自适应，无固定高度）
        st.markdown('<div class="left-top-block">', unsafe_allow_html=True)
        st.markdown("### 📑 自定义模板生成")
        st.divider()
        template_name = st.text_input(
            "自定义模板命名",
            placeholder="请输入模板名称，如：河北科技大学本科毕设专用模板",
            key="custom_template_name"
        )
        col_save, col_del = st.columns(2)
        with col_save:
            if st.button("保存当前格式", type="primary", use_container_width=True):
                if not template_name.strip():
                    st.error("请输入模板名称")
                else:
                    st.session_state.custom_templates[template_name] = {
                        "name": template_name,
                        "update_time": datetime.now().strftime('%Y-%m-%d'),
                        "cn_format": copy.deepcopy(st.session_state.custom_cn_format),
                        "en_format": copy.deepcopy(st.session_state.custom_en_format)
                    }
                    st.success(f"✅ 模板「{template_name}」保存成功")
                    st.session_state.format_version += 1
                    safe_rerun()
        with col_del:
            if st.session_state.custom_templates:
                del_template = st.selectbox(
                    "选择模板",
                    options=list(st.session_state.custom_templates.keys()),
                    label_visibility="collapsed",
                    key="del_template_select"
                )
                if st.button("删除模板", type="secondary", use_container_width=True):
                    # 修复：动态key避免重复渲染报错
                    if st.checkbox("确认删除该模板？", key=f"del_confirm_{st.session_state.format_version}"):
                        del st.session_state.custom_templates[del_template]
                        st.success(f"✅ 模板「{del_template}」已删除")
                        st.session_state.format_version += 1
                        safe_rerun()
        st.caption("调整下方格式参数，命名后即可保存为专属自定义模板")
        st.markdown('</div>', unsafe_allow_html=True)

        st.divider()
        
        # 中模块：全参数格式精细调整区（占满剩余空间，仅内部滚动）
        st.markdown('<div class="left-middle-block">', unsafe_allow_html=True)
        st.markdown("### 🎨 格式参数全量调整")
        st.divider()
        format_levels = [
            "一级标题", "二级标题", "三级标题",
            "正文", "表格", "图片与图注", "参考文献"
        ]
        for level in format_levels:
            expanded = (level == "正文")
            with st.expander(f"{level}格式设置", expanded=expanded):
                cfg = st.session_state.custom_cn_format[level]
                col_base, col_layout = st.columns(2)
                # 左列：基础样式组
                with col_base:
                    st.markdown("**基础样式**")
                    cfg["font"] = st.selectbox(
                        "中文字体",
                        CN_FONT_LIST,
                        index=CN_FONT_LIST.index(cfg["font"]) if cfg["font"] in CN_FONT_LIST else 0,
                        key=f"cn_{level}_font_{st.session_state.format_version}"
                    )
                    cfg["size"] = st.selectbox(
                        "字号",
                        list(FONT_SIZE_MAP.keys()),
                        index=list(FONT_SIZE_MAP.keys()).index(cfg["size"]) if cfg["size"] in FONT_SIZE_MAP else 5,
                        key=f"cn_{level}_size_{st.session_state.format_version}"
                    )
                    cfg["bold"] = st.checkbox(
                        "字体加粗",
                        cfg["bold"],
                        key=f"cn_{level}_bold_{st.session_state.format_version}"
                    )
                    cfg["italic"] = st.checkbox(
                        "字体斜体",
                        cfg["italic"],
                        key=f"cn_{level}_italic_{st.session_state.format_version}"
                    )
                    cfg["color"] = st.color_picker(
                        "字体颜色",
                        cfg["color"],
                        key=f"cn_{level}_color_{st.session_state.format_version}"
                    )
                # 右列：段落布局组
                with col_layout:
                    st.markdown("**段落布局**")
                    cfg["align"] = st.selectbox(
                        "对齐方式",
                        list(ALIGN_MAP.keys()),
                        index=list(ALIGN_MAP.keys()).index(cfg["align"]),
                        key=f"cn_{level}_align_{st.session_state.format_version}"
                    )
                    # ========== 核心修复：行距类型与数值类型统一 ==========
                    line_type_options = ["倍数", "固定值"]
                    line_type_labels = ["倍数行距", "固定值行距"]
                    cfg["line_type"] = st.selectbox(
                        "行距类型",
                        options=line_type_options,
                        format_func=lambda x: line_type_labels[line_type_options.index(x)],
                        index=0 if cfg["line_type"] == "倍数" else 1,
                        key=f"cn_{level}_line_type_{st.session_state.format_version}"
                    )
                    # 严格统一数值类型，倍数用float，固定值用int，同时做边界校验
                    if cfg["line_type"] == "倍数":
                        line_min = 0.5
                        line_max = 5.0
                        line_step = 0.1
                        default_value = 1.5
                        try:
                            current_value = float(cfg["line_value"])
                            if not (line_min <= current_value <= line_max):
                                current_value = default_value
                        except (ValueError, TypeError):
                            current_value = default_value
                    else:
                        line_min = 8
                        line_max = 50
                        line_step = 1
                        default_value = 20
                        try:
                            current_value = int(cfg["line_value"])
                            if not (line_min <= current_value <= line_max):
                                current_value = default_value
                        except (ValueError, TypeError):
                            current_value = default_value
                    cfg["line_value"] = st.number_input(
                        "行距数值",
                        min_value=line_min,
                        max_value=line_max,
                        value=current_value,
                        step=line_step,
                        key=f"cn_{level}_line_val_{st.session_state.format_version}"
                    )
                    cfg["indent"] = st.number_input(
                        "首行缩进(字符)",
                        min_value=0,
                        max_value=4,
                        value=cfg["indent"],
                        step=1,
                        key=f"cn_{level}_indent_{st.session_state.format_version}"
                    )
                    cfg["space_before"] = st.number_input(
                        "段前间距(pt)",
                        min_value=0,
                        max_value=50,
                        value=cfg["space_before"],
                        step=1,
                        key=f"cn_{level}_before_{st.session_state.format_version}"
                    )
                    cfg["space_after"] = st.number_input(
                        "段后间距(pt)",
                        min_value=0,
                        max_value=50,
                        value=cfg["space_after"],
                        step=1,
                        key=f"cn_{level}_after_{st.session_state.format_version}"
                    )
                # 通栏：字符间距组
                st.markdown("**字符间距**")
                cfg["char_spacing"] = st.slider(
                    "字间距(pt)",
                    min_value=0,
                    max_value=10,
                    value=cfg.get("char_spacing", 0),
                    step=1,
                    key=f"cn_{level}_char_space_{st.session_state.format_version}"
                )
                st.session_state.custom_cn_format[level] = cfg
        # 西文全局格式设置
        with st.expander("🔤 西文全局格式设置", expanded=False):
            st.markdown("针对各文档元素单独设置西文格式规则")
            for level in format_levels:
                en_cfg = st.session_state.custom_en_format[level]
                st.markdown(f"**{level}西文格式**")
                col_en1, col_en2, col_en3 = st.columns(3)
                with col_en1:
                    en_cfg["en_font"] = st.selectbox(
                        "西文字体",
                        EN_FONT_LIST,
                        index=EN_FONT_LIST.index(en_cfg["en_font"]) if en_cfg["en_font"] in EN_FONT_LIST else 0,
                        key=f"en_{level}_font_{st.session_state.format_version}"
                    )
                with col_en2:
                    en_cfg["size_same_as_cn"] = st.checkbox(
                        "西文字号与中文同步",
                        en_cfg["size_same_as_cn"],
                        key=f"en_{level}_sync_{st.session_state.format_version}"
                    )
                with col_en3:
                    en_cfg["bold"] = st.checkbox(
                        "西文加粗",
                        en_cfg["bold"],
                        key=f"en_{level}_bold_{st.session_state.format_version}"
                    )
                    en_cfg["italic"] = st.checkbox(
                        "西文斜体",
                        en_cfg["italic"],
                        key=f"en_{level}_italic_{st.session_state.format_version}"
                    )
                if not en_cfg["size_same_as_cn"]:
                    en_cfg["size"] = st.selectbox(
                        "西文字号",
                        list(FONT_SIZE_MAP.keys()),
                        index=list(FONT_SIZE_MAP.keys()).index(en_cfg["size"]) if en_cfg["size"] in FONT_SIZE_MAP else 5,
                        key=f"en_{level}_size_{st.session_state.format_version}"
                    )
                st.session_state.custom_en_format[level] = en_cfg
        st.markdown('</div>', unsafe_allow_html=True)

        st.divider()
        
        # 下模块：模板导出区（flex自适应，无固定高度）
        st.markdown('<div class="left-bottom-block">', unsafe_allow_html=True)
        st.markdown("### 📤 模板导出")
        st.divider()
        export_type = st.radio(
            "导出格式",
            options=["JSON专用格式", "TXT通用格式"],
            index=0,
            horizontal=True,
            key="export_type_radio"
        )
        st.caption("JSON格式：仅可用于本平台导入，完整保留所有参数；TXT格式：可查看编辑，兼容通用文本查看器")
        if st.button("导出当前自定义模板", use_container_width=True):
            template_data = {
                "name": template_name if template_name else "自定义模板",
                "update_time": datetime.now().strftime('%Y-%m-%d'),
                "cn_format": st.session_state.custom_cn_format,
                "en_format": st.session_state.custom_en_format
            }
            export_type_code = "json" if "JSON" in export_type else "txt"
            export_data = export_template(template_data, export_type_code)
            file_name = f"{template_data['name']}_{datetime.now().strftime('%Y%m%d')}.{export_type_code}"
            st.download_button(
                label="⬇️ 下载模板文件",
                data=export_data,
                file_name=file_name,
                mime="application/json" if export_type_code == "json" else "text/plain",
                use_container_width=True
            )
        st.caption("导出的模板文件可在右侧模板导入区上传复用")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)
    # ============== 右栏：全流程操作主链路 ==============
    with right_col:
        st.markdown('<div class="right-column">', unsafe_allow_html=True)
        # 顶部固定标题与功能提示条
        with st.container():
            st.markdown('<div class="fixed-header">', unsafe_allow_html=True)
            st.title("📝 智能论文&竞赛格式处理平台")
            st.success("✅ 支持一键格式标准化 | WPS自动生成导航 | 知网参考文献优化 | 智能降重润色 | 格式合规检查")
            st.markdown('</div>', unsafe_allow_html=True)
        # ========== 模块1：第一步 文档格式标准化处理 ==========
        st.subheader("📄 第一步：文档格式标准化")
        st.markdown('<hr class="module-divider-green">', unsafe_allow_html=True)
        col_template, col_upload = st.columns(2)
        # 左列：模板选择与导入区
        with col_template:
            st.markdown("##### 模板选择与导入")
            # 模板选择下拉框
            template_options = list(ALL_TEMPLATES.keys()) + list(st.session_state.custom_templates.keys())
            selected_template = st.selectbox(
                "选择处理模板",
                options=template_options,
                index=template_options.index(st.session_state.selected_template) if st.session_state.selected_template in template_options else 0,
                key="right_template_select"
            )
            # 模板切换逻辑
            if selected_template != st.session_state.selected_template:
                st.session_state.selected_template = selected_template
                if selected_template in ALL_TEMPLATES:
                    st.session_state.selected_cn_format, st.session_state.selected_en_format = get_cached_template(selected_template)
                else:
                    tmp = st.session_state.custom_templates[selected_template]
                    st.session_state.selected_cn_format = copy.deepcopy(tmp["cn_format"])
                    st.session_state.selected_en_format = copy.deepcopy(tmp["en_format"])
                st.session_state.format_version += 1
                safe_rerun()
            # 模板更新时间
            if selected_template in ALL_TEMPLATES:
                update_time = ALL_TEMPLATES[selected_template].get("update_time", "未知")
                st.caption(f"📅 模板更新时间：{update_time}")
                # 模板官方要求
                special_req = ALL_TEMPLATES[selected_template].get("special_requirements", [])
                if special_req:
                    with st.expander("模板官方格式要求", expanded=False):
                        for req in special_req:
                            st.markdown(f"- {req}")
            else:
                update_time = st.session_state.custom_templates[selected_template].get("update_time", datetime.now().strftime('%Y-%m-%d'))
                st.caption(f"📅 自定义模板更新时间：{update_time}")
            # 外部模板导入区
            st.markdown("##### 外部模板导入")
            uploaded_template_file = st.file_uploader(
                "导入外部模板",
                type=["json", "txt"],
                help="支持上传本平台导出的JSON/TXT模板文件",
                key="template_upload"
            )
            if uploaded_template_file:
                template_data, error = import_template(uploaded_template_file)
                if error:
                    st.error(f"导入失败：{error}")
                elif template_data:
                    st.success("✅ 模板解析成功！")
                    import_template_name = st.text_input(
                        "导入模板命名",
                        value=uploaded_template_file.name.split('.')[0],
                        key=f"import_template_name_{st.session_state.format_version}"
                    )
                    if st.button("导入到系统", use_container_width=True):
                        st.session_state.custom_templates[import_template_name] = template_data
                        st.session_state.selected_template = import_template_name
                        st.session_state.selected_cn_format = copy.deepcopy(template_data["cn_format"])
                        st.session_state.selected_en_format = copy.deepcopy(template_data["en_format"])
                        st.success(f"✅ 模板「{import_template_name}」导入成功，已自动选中")
                        st.session_state.format_version += 1
                        safe_rerun()
            # 格式规则与辅助功能确认
            st.markdown("##### 格式规则确认")
            if selected_template in ALL_TEMPLATES:
                cn_format_show = ALL_TEMPLATES[selected_template]["cn_format"]
            else:
                cn_format_show = st.session_state.custom_templates[selected_template]["cn_format"]
            with st.expander("核心格式规则摘要", expanded=True):
                st.markdown(f"- 一级标题：{cn_format_show['一级标题']['font']} {cn_format_show['一级标题']['size']} {cn_format_show['一级标题']['align']}")
                st.markdown(f"- 正文：{cn_format_show['正文']['font']} {cn_format_show['正文']['size']} {cn_format_show['正文']['line_value']}倍行距 首行缩进{cn_format_show['正文']['indent']}字符")
                st.markdown(f"- 表格：{cn_format_show['表格']['font']} {cn_format_show['表格']['size']} {cn_format_show['表格']['align']}")
            # 辅助功能开关
            st.markdown("##### 辅助功能设置")
            col_func1, col_func2 = st.columns(2)
            with col_func1:
                bind_wps_style = st.checkbox("✅ 绑定WPS标题样式", value=True, help="开启后导出的文档在WPS中自动生成导航目录")
                standardize_ref = st.checkbox("📚 知网参考文献标准化", value=True, help="自动调整参考文献格式，解决知网查重标红问题")
            with col_func2:
                optimize_image = st.checkbox("🖼️ 图片排版自动优化", value=True, help="自动优化图片与图注的排版格式")
                enable_check = st.checkbox("🔍 格式合规性自动检查", value=True, help="处理完成后自动检查格式合规性")
        # 右列：待处理文档上传区
        with col_upload:
            st.markdown("##### 待处理文档上传")
            uploaded_files = st.file_uploader(
                "上传 .docx 文档",
                type=["docx"],
                accept_multiple_files=True,
                help="支持同时上传多个文档批量处理，单文件最大200MB",
                key="doc_upload"
            )
            if uploaded_files:
                st.markdown("**已上传文件列表**")
                for file in uploaded_files:
                    col_file, col_del = st.columns([4,1])
                    with col_file:
                        st.caption(f"📄 {file.name} | {(file.size/1024/1024):.2f}MB")
                    with col_del:
                        if st.button("删除", key=f"del_{file.name}_{st.session_state.format_version}"):
                            uploaded_files.remove(file)
                            safe_rerun()
        # 处理按钮与结果展示
        process_disabled = not (selected_template and uploaded_files)
        if st.button("🚀 开始一键格式处理", type="primary", use_container_width=True, disabled=process_disabled):
            st.session_state.process_timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            for file in uploaded_files:
                with st.spinner(f"正在处理：{file.name}"):
                    try:
                        # 获取当前选中模板的格式
                        if selected_template in ALL_TEMPLATES:
                            current_cn_format, current_en_format = get_cached_template(selected_template)
                        else:
                            current_cn_format = copy.deepcopy(st.session_state.custom_templates[selected_template]["cn_format"])
                            current_en_format = copy.deepcopy(st.session_state.custom_templates[selected_template]["en_format"])
                        # 执行文档处理
                        output_doc, changes, title_stats, process_log, check_report, full_text = process_doc(
                            file=file,
                            cn_format=current_cn_format,
                            en_format=current_en_format,
                            enable_rewrite=False,
                            rewrite_level="标准润色",
                            bind_wps_style=bind_wps_style,
                            standardize_ref=standardize_ref,
                            optimize_image=optimize_image,
                            api_key=None,
                            forbidden_text=None
                        )
                        # 保存处理结果
                        st.session_state.formatted_doc = output_doc
                        st.session_state.doc_full_text = full_text
                        # 自动查重
                        check_rate = simulate_check_rate(full_text)
                        st.session_state.original_check_rate = check_rate
                        st.session_state.check_rate = check_rate
                        # 生成处理报告
                        report = generate_report(changes, "无润色", title_stats, process_log, check_report)
                        st.session_state.formatted_report = report
                        # 结果展示
                        st.subheader(f"✅ 处理完成：{file.name}")
                        with st.expander("📋 处理日志", expanded=True):
                            for log in process_log:
                                st.write(log)
                        # 文档结构统计
                        st.markdown("**📊 文档结构统计**")
                        cols_stat = st.columns(7)
                        cols_stat[0].metric("一级标题", title_stats["一级标题"])
                        cols_stat[1].metric("二级标题", title_stats["二级标题"])
                        cols_stat[2].metric("三级标题", title_stats["三级标题"])
                        cols_stat[3].metric("正文段落", title_stats["正文"])
                        cols_stat[4].metric("表格数量", title_stats["表格"])
                        cols_stat[5].metric("图片数量", title_stats["图片与图注"])
                        cols_stat[6].metric("参考文献", title_stats["参考文献"])
                        # 格式合规检查报告
                        with st.expander("🔍 格式合规检查报告", expanded=False):
                            for item in check_report:
                                st.write(item)
                    except Exception as e:
                        st.error(f"处理失败：{str(e)}")
        # 自动查重结果展示
        if st.session_state.check_rate is not None:
            st.divider()
            st.markdown("##### 🔍 文档查重结果")
            rate = st.session_state.check_rate
            st.progress(rate/100)
            st.markdown(f"**文档查重率：{rate}%**")
            # 跳转润色按钮
            if rate > 20:
                st.warning(f"⚠️ 查重率{rate}%，超出常规学术要求，建议进行AI润色降重")
                if st.button("✨ 一键跳转至润色降重", type="primary"):
                    st.session_state.need_polish = True
                    safe_rerun()
            else:
                st.success(f"✅ 查重率{rate}%，符合学术规范要求")
                if st.button("仍需进行润色优化"):
                    st.session_state.need_polish = True
                    safe_rerun()
        # ========== 模块2：第二步 AI智能润色降重 ==========
        st.subheader("✨ 第二步：AI智能润色降重")
        st.markdown('<hr class="module-divider-blue">', unsafe_allow_html=True)
        # 默认折叠，点击跳转后展开
        with st.expander("润色降重功能", expanded=st.session_state.need_polish):
            col_polish_doc, col_polish_report = st.columns(2)
            # 左列：待润色文档上传
            with col_polish_doc:
                st.markdown("##### 待润色文档")
                use_formatted_doc = st.checkbox(
                    "使用第一步排版后的文档",
                    value=True,
                    disabled=st.session_state.formatted_doc is None,
                    key="use_formatted_doc"
                )
                if use_formatted_doc and st.session_state.formatted_doc:
                    st.info("✅ 已自动加载第一步排版完成的文档，无需重复上传")
                    polish_file = st.session_state.formatted_doc
                else:
                    polish_file = st.file_uploader(
                        "上传待润色文档",
                        type=["docx"],
                        help="单文件最大200MB",
                        key="polish_doc_upload"
                    )
            # 右列：查重报告上传
            with col_polish_report:
                st.markdown("##### 查重报告（可选，精准降重）")
                use_original_report = st.checkbox(
                    "使用第一步查重结果精准降重",
                    value=True,
                    disabled=st.session_state.original_check_rate is None,
                    key="use_original_report"
                )
                if use_original_report and st.session_state.learned_forbidden:
                    st.info(f"✅ 已自动加载第一步查重标红内容，共{len(st.session_state.learned_forbidden)}处标红，将针对性降重")
                    forbidden_text = st.session_state.learned_forbidden
                else:
                    report_file = st.file_uploader(
                        "上传查重报告",
                        type=["html", "txt"],
                        help="支持知网、万方等查重报告HTML/TXT文件",
                        key="polish_report_upload"
                    )
                    forbidden_text = None
                    if report_file:
                        red_parts, plain_text, error = parse_plagiarism_report(report_file)
                        if error:
                            st.error(f"解析失败：{error}")
                        elif red_parts:
                            st.success(f"✅ 解析完成！发现{len(red_parts)}处标红重复内容，将针对性降重")
                            forbidden_text = red_parts
                            st.session_state.learned_forbidden = red_parts
            # 润色配置与API输入
            st.markdown("##### 润色配置")
            col_config1, col_config2, col_config3 = st.columns([2,3,2])
            with col_config1:
                rewrite_level = st.selectbox(
                    "润色强度",
                    options=list(REWRITE_LEVEL.keys()),
                    index=1,
                    format_func=lambda x: f"{x} - {REWRITE_LEVEL[x]['desc']}",
                    key="polish_level_select"
                )
            with col_config2:
                api_key = st.text_input(
                    "豆包API Key（可选，开启AI深度润色）",
                    type="password",
                    placeholder="请输入您的豆包API Key",
                    key="polish_api_key"
                )
            with col_config3:
                polish_disabled = not polish_file
                if st.button("🚀 开始AI润色降重", type="primary", use_container_width=True, disabled=polish_disabled):
                    with st.spinner("正在进行AI润色降重..."):
                        try:
                            # 获取当前选中模板的格式
                            if selected_template in ALL_TEMPLATES:
                                current_cn_format, current_en_format = get_cached_template(selected_template)
                            else:
                                current_cn_format = copy.deepcopy(st.session_state.custom_templates[selected_template]["cn_format"])
                                current_en_format = copy.deepcopy(st.session_state.custom_templates[selected_template]["en_format"])
                            # 执行润色处理
                            output_doc, changes, title_stats, process_log, check_report, full_text = process_doc(
                                file=polish_file,
                                cn_format=current_cn_format,
                                en_format=current_en_format,
                                enable_rewrite=True,
                                rewrite_level=rewrite_level,
                                bind_wps_style=bind_wps_style,
                                standardize_ref=standardize_ref,
                                optimize_image=optimize_image,
                                api_key=api_key,
                                forbidden_text=forbidden_text
                            )
                            # 保存润色结果
                            st.session_state.polish_doc = output_doc
                            # 重新查重
                            new_rate = simulate_check_rate(full_text)
                            st.session_state.check_rate = new_rate
                            # 生成润色报告
                            polish_report = generate_report(changes, rewrite_level, title_stats, process_log, check_report)
                            st.session_state.polish_report = polish_report
                            # 润色结果展示
                            st.success(f"✅ 润色完成！共优化{len(changes)}处内容")
                            # 查重率对比
                            st.markdown("**📊 润色前后查重率对比**")
                            col_rate1, col_rate2 = st.columns(2)
                            with col_rate1:
                                st.markdown(f"润色前查重率：**{st.session_state.original_check_rate}%**")
                                st.progress(st.session_state.original_check_rate/100)
                            with col_rate2:
                                st.markdown(f"润色后查重率：**{new_rate}%**")
                                st.progress(new_rate/100)
                            # 润色统计
                            if changes:
                                avg_semantic = sum([c["semantic_score"] for c in changes]) / len(changes)
                                st.markdown(f"**📈 润色统计**：语义平均保留度 {avg_semantic*100:.1f}% | 标红内容修改率 {min(100, len(changes)/len(forbidden_text)*100 if forbidden_text else 100):.1f}%")
                            # 润色详情
                            with st.expander("📋 润色修改详情", expanded=False):
                                for i, change in enumerate(changes[:20]):
                                    st.markdown(f"**修改 {i+1}** | 类型：{change['type']} | 语义保留：{change['semantic_score']*100:.1f}%")
                                    st.markdown(f"- 原句：{change['original']}")
                                    st.markdown(f"- 修改：{change['modified']}")
                                    st.divider()
                        except Exception as e:
                            st.error(f"润色失败：{str(e)}")
        # ========== 模块3：第三步 成果输出 ==========
        st.subheader("📥 第三步：成果输出")
        st.markdown('<hr class="module-divider-gray">', unsafe_allow_html=True)
        col_out1, col_out2, col_out3, col_out4 = st.columns(4)
        timestamp = st.session_state.process_timestamp
        # 第一列：排版后文档
        with col_out1:
            download_disabled = st.session_state.formatted_doc is None
            st.download_button(
                label="📥 下载排版后文档",
                data=st.session_state.formatted_doc if st.session_state.formatted_doc else BytesIO(),
                file_name=f"标准格式_{timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                disabled=download_disabled
            )
            st.caption("格式标准化后的docx文档")
        # 第二列：格式处理报告
        with col_out2:
            download_disabled = st.session_state.formatted_report is None
            st.download_button(
                label="📋 下载格式处理报告",
                data=st.session_state.formatted_report if st.session_state.formatted_report else BytesIO(),
                file_name=f"格式处理报告_{timestamp}.txt",
                mime="text/plain",
                use_container_width=True,
                disabled=download_disabled
            )
            st.caption("排版全流程日志与合规检查报告")
        # 第三列：润色后文档
        with col_out3:
            download_disabled = st.session_state.polish_doc is None
            st.download_button(
                label="✨ 下载润色后文档",
                data=st.session_state.polish_doc if st.session_state.polish_doc else BytesIO(),
                file_name=f"润色降重_{timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                disabled=download_disabled
            )
            st.caption("AI润色降重后的docx文档")
        # 第四列：润色降重报告
        with col_out4:
            download_disabled = st.session_state.polish_report is None
            st.download_button(
                label="📊 下载润色降重报告",
                data=st.session_state.polish_report if st.session_state.polish_report else BytesIO(),
                file_name=f"润色降重报告_{timestamp}.txt",
                mime="text/plain",
                use_container_width=True,
                disabled=download_disabled
            )
            st.caption("润色详情与查重率对比报告")
        # 底部安全提示
        st.caption("💡 所有文件仅在浏览器内存中生成，不会上传保存到服务器，关闭页面后自动清除，保障您的文档数据安全")
        st.markdown('</div>', unsafe_allow_html=True)
if __name__ == "__main__":
    main()
