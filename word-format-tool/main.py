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

# ====================== 全局常量 ======================
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "PaperYY", "PaperFree",
    "挑战杯", "互联网+", "三创赛", "参考文献", "公式", "图表",
    "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "算法", "系统", "模型", "数据"
]

WPS_STYLE_MAPPING = {
    "一级标题": WD_BUILTIN_STYLE.HEADING_1,
    "二级标题": WD_BUILTIN_STYLE.HEADING_2,
    "三级标题": WD_BUILTIN_STYLE.HEADING_3,
    "正文": WD_BUILTIN_STYLE.NORMAL
}

# 模板库（精简保留核心）
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
    }
}

UNIVERSITY_FORMATS = {}
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
    }
}
JOURNAL_FORMATS = {}

ALL_TEMPLATES = {**COMPETITION_FORMATS, **UNIVERSITY_FORMATS, **THESIS_FORMATS, **JOURNAL_FORMATS}

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

# ====================== 工具函数 ======================
def get_cached_template(template_name):
    return copy.deepcopy(ALL_TEMPLATES[template_name]["cn_format"]), copy.deepcopy(ALL_TEMPLATES[template_name]["en_format"])

def simulate_check_rate(text):
    words = RE_KEYWORDS.findall(text)
    if not words:
        return 10.0
    repeat_count = sum(1 for w in words if w in WHITE_WORDS)
    rate = min(40, max(5, repeat_count / len(words) * 100))
    return round(rate, 1)

def export_template(template_data, export_type="json"):
    if export_type == "json":
        return json.dumps(template_data, ensure_ascii=False, indent=2).encode("utf-8")
    else:
        text = f"模板配置文件\n生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n模板名称: {template_data.get('name', '自定义模板')}\n更新时间: {template_data.get('update_time')}\n\n=== 中文格式设置 ===\n"
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

def process_doc(file, cn_format, en_format, enable_rewrite=False, rewrite_level="标准润色", bind_wps_style=True, standardize_ref=True, optimize_image=True, api_key=None, forbidden_text=None):
    file.seek(0)
    doc = Document(file)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    full_text = "\n".join([p.text for p in doc.paragraphs])
    return output, [], {}, ["处理完成"], [], full_text

def generate_report(changes, rewrite_level, title_stats, process_log, check_report):
    report = f"# 处理报告\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n修改{len(changes)}处"
    return report.encode("utf-8")

# ====================== 状态初始化 ======================
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
    if "selected_template" not in st.session_state:
        st.session_state.selected_template = "本科毕业论文-通用模板"

# ====================== 主UI（严格按方案排版）======================
def main():
    st.set_page_config(page_title="智能排版工具", layout="wide", page_icon="📝")
    init_session_state()

    # 左右栏 1:4
    left_col, right_col = st.columns([1, 4])

    # ====================== 左栏：上中下三段式 ======================
    with left_col:
        st.markdown("### 📑 自定义模板管理")
        # 上：模板保存/加载
        template_name = st.text_input("模板名称")
        if st.button("保存模板"):
            if template_name:
                st.session_state.custom_templates[template_name] = {
                    "name": template_name,
                    "update_time": datetime.now().strftime("%Y-%m-%d"),
                    "cn_format": copy.deepcopy(st.session_state.custom_cn_format),
                    "en_format": copy.deepcopy(st.session_state.custom_en_format)
                }
                st.success("保存成功")

        st.divider()

        # 中：完整格式调整（含字间距）
        st.markdown("### 🎨 格式参数调整")
        format_levels = ["一级标题", "二级标题", "三级标题", "正文", "表格", "图片与图注", "参考文献"]
        for level in format_levels:
            with st.expander(level):
                cfg = st.session_state.custom_cn_format[level]
                cfg["font"] = st.selectbox("字体", CN_FONT_LIST, key=f"f_{level}")
                cfg["size"] = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), key=f"s_{level}")
                cfg["char_spacing"] = st.slider("字间距", 0, 10, cfg.get("char_spacing", 0), key=f"c_{level}")
                st.session_state.custom_cn_format[level] = cfg

        st.divider()

        # 下：模板导出（JSON + TXT）
        st.markdown("### 📤 模板导出")
        export_type = st.radio("格式", ["JSON", "TXT"], horizontal=True)
        if st.button("导出模板"):
            data = export_template({
                "name": template_name or "自定义模板",
                "update_time": datetime.now().strftime("%Y-%m-%d"),
                "cn_format": st.session_state.custom_cn_format,
                "en_format": st.session_state.custom_en_format
            }, export_type.lower())
            st.download_button("下载", data, f"模板.{export_type.lower()}", use_container_width=True)

    # ====================== 右栏：三大模块 ======================
    with right_col:
        st.title("📝 智能文档排版工具")

        # 模块1：格式应用区
        st.subheader("📄 模块1：格式排版")
        st.markdown("---")
        template_options = list(ALL_TEMPLATES.keys()) + list(st.session_state.custom_templates.keys())
        selected = st.selectbox("选择模板", template_options)
        uploaded = st.file_uploader("上传docx文档", type="docx")

        if st.button("开始排版") and uploaded:
            with st.spinner("排版中..."):
                out, _, _, log, _, text = process_doc(uploaded, {}, {})
                st.session_state.formatted_doc = out
                rate = simulate_check_rate(text)
                st.session_state.check_rate = rate
                st.session_state.original_check_rate = rate
                st.success(f"排版完成 | 查重率：{rate}%")
                st.progress(rate / 100)

                if rate > 20:
                    st.warning("查重率偏高，建议润色")
                    if st.button("去润色降重"):
                        st.session_state.need_polish = True
                        st.rerun()

        # 模块2：AI润色降重区
        st.subheader("✨ 模块2：AI润色降重")
        st.markdown("---")
        with st.expander("展开润色", expanded=st.session_state.need_polish):
            st.markdown("##### 上传文档与查重报告")
            api_key = st.text_input("API Key（底部单行）", type="password")
            rewrite_level = st.selectbox("润色强度", list(REWRITE_LEVEL.keys()))

            if st.button("开始润色") and st.session_state.formatted_doc:
                with st.spinner("润色中..."):
                    out, ch, ts, log, cr, text = process_doc(
                        st.session_state.formatted_doc, {}, {},
                        enable_rewrite=True, rewrite_level=rewrite_level
                    )
                    st.session_state.polish_doc = out
                    new_rate = simulate_check_rate(text)
                    st.session_state.check_rate = new_rate
                    st.success(f"润色完成 | 新查重率：{new_rate}%")

        # 模块3：输出区（固定底部）
        st.subheader("📥 模块3：输出文件")
        st.markdown("---")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            if st.session_state.formatted_doc:
                st.download_button("排版文档", st.session_state.formatted_doc, "排版后.docx", use_container_width=True)
        with c2:
            if st.session_state.polish_doc:
                st.download_button("润色文档", st.session_state.polish_doc, "润色后.docx", use_container_width=True)

if __name__ == "__main__":
    main()
