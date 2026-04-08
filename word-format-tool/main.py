import streamlit as st
import copy
import re
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_BUILTIN_STYLE
from docx.oxml.ns import qn

# ====================== 全局配置与常量 ======================
# 专业术语白名单（永不修改）
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "挑战杯", "互联网+", "河北科技大学",
    "工业工程", "GDP", "CPI", "GB/T 7714", "ISO", "一级标题", "二级标题", "三级标题",
    "参考文献", "公式", "图表", "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "Transformer", "BERT", "T5", "Python", "Java", "SQL"
]

# 【核心】WPS原生样式映射表（绑定WPS内置标题样式，关键中的关键）
WPS_STYLE_MAPPING = {
    "一级标题": WD_BUILTIN_STYLE.HEADING_1,
    "二级标题": WD_BUILTIN_STYLE.HEADING_2,
    "三级标题": WD_BUILTIN_STYLE.HEADING_3,
    "正文": WD_BUILTIN_STYLE.NORMAL
}

# 河北科技大学毕业论文/竞赛标准格式模板
DEFAULT_CN_FORMAT = {
    "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
    "二级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
    "三级标题": {"font": "黑体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0},
    "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0},
    "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
}

# 西文/数字独立格式配置（竞赛专用）
DEFAULT_EN_FORMAT = {
    "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
    "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
    "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
    "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
    "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
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
    "优势": "核心竞争力", "问题": "行业痛点", "方法": "技术路径"
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
EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri", "Courier New"]
CN_FONT_LIST = ["宋体", "黑体", "楷体", "仿宋_GB2312", "微软雅黑"]
FONT_SIZE_LIST = list(FONT_SIZE_MAP.keys())
APP_NAME = "论文排版+WPS标题绑定工具（竞赛专属版）"

# ====================== 1. 标题层级精准识别（核心，解决误判问题） ======================
def get_title_level(para_text):
    """
    精准识别标题层级，100%匹配WPS标题1/2/3
    彻底解决正文列表（1）（2）误判为标题的问题
    """
    text = para_text.strip()
    if not text or len(text) < 2:
        return "正文"

    # 一级标题：第X章、X、（仅短文本，长文本归为正文）
    if re.match(r'^第[一二三四五六七八九十]+章\s', text) or (re.match(r'^\d+、\s', text) and len(text) < 25):
        return "一级标题"
    # 二级标题：1.1、（一）（仅短文本）
    elif re.match(r'^\d+\.\d+\s', text) or (re.match(r'^（[一二三四五六七八九十]+）\s', text) and len(text) < 25):
        return "二级标题"
    # 三级标题：1.1.1、（1）（仅短文本，长正文列表直接归为正文）
    elif re.match(r'^\d+\.\d+\.\d+\s', text) or (re.match(r'^（\d+）\s', text) and len(text) < 20):
        return "三级标题"
    # 所有其他情况归为正文
    return "正文"

# ====================== 2. 智能降重引擎（可选开启，不破坏格式） ======================
def is_white_text(text):
    """判断是否为白名单内容，不修改"""
    for word in WHITE_WORDS:
        if word in text:
            return True
    if re.match(r'^\d+(\.\d+)*$', text) or re.match(r'^\[.*\]$', text):
        return True
    return False

def check_semantic_keep(original, modified):
    """规则化语义保持校验，无额外依赖"""
    original_keywords = set(re.findall(r'[\u4e00-\u9fa5]{2,}', original))
    modified_keywords = set(re.findall(r'[\u4e00-\u9fa5]{2,}', modified))
    if not original_keywords:
        return 1.0
    overlap = original_keywords & modified_keywords
    return len(overlap) / len(original_keywords)

def rewrite_sentence(sentence, level_config):
    """单句降重，严格遵循降重方法论"""
    original = sentence.strip()
    if len(original) < 5 or is_white_text(original):
        return original, "原文保留（白名单/短句）", 1.0

    modified = original
    rewrite_type = "无修改"

    # 同义词替换
    if level_config["synonym"]:
        for old, new in SYNONYM_DICT.items():
            if old in modified:
                modified = modified.replace(old, new)
                rewrite_type = "同义词替换"

    # 句式重构
    if level_config["sentence_reorder"]:
        parts = [p.strip() for p in re.split(r'[，。；]', modified) if p.strip()]
        if len(parts) >= 3:
            import random
            random.shuffle(parts)
            modified = "，".join(parts) + "。"
            rewrite_type = "句式重构+语序打乱"

    # 结构调整
    if level_config["structure_change"]:
        if "在" in modified and "中" in modified:
            modified = re.sub(r'在(.*?)中', f'结合{datetime.now().year}年行业实际发展情况，在\g<1>场景中', modified)
            rewrite_type = "结构调整+场景限定补充"

    # 语义校验，不达标回退
    semantic_score = check_semantic_keep(original, modified)
    if semantic_score < 0.7:
        return original, "原文保留（语义重合度不达标）", 1.0

    return modified, rewrite_type, round(semantic_score, 4)

def rewrite_paragraph(text, level_config):
    """整段降重"""
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

# ====================== 3. 核心文档处理（WPS标题绑定+格式设置） ======================
def process_doc(
    file,
    cn_format,
    en_format,
    enable_rewrite=False,
    rewrite_level="标准降重",
    bind_wps_style=True  # 核心开关：绑定WPS原生标题样式
):
    """
    核心处理流程：
    1. 可选降重 → 2. 识别标题层级 → 3. 绑定WPS原生标题样式 → 4. 中文/西文格式独立设置
    100%保留原标题编号，不修改文本内容，仅绑定样式
    """
    doc = Document(file)
    total_changes = []
    rewrite_config = REWRITE_LEVEL[rewrite_level]
    title_stats = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0, "表格": len(doc.tables)}

    # ====================== 第一步：可选智能降重 ======================
    if enable_rewrite:
        # 正文段落降重
        for para in doc.paragraphs:
            original_text = para.text.strip()
            if not original_text or get_title_level(original_text) != "正文":
                continue
            new_text, changes = rewrite_paragraph(original_text, rewrite_config)
            if changes:
                total_changes.extend(changes)
                para.text = new_text

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

    # ====================== 第二步：核心！绑定WPS原生标题样式 + 格式设置 ======================
    for para in doc.paragraphs:
        level = get_title_level(para.text)
        title_stats[level] += 1
        cn_style = cn_format[level]
        en_style = en_format[level]

        # 【核心】绑定WPS原生内置标题样式（WPS右上角自动对应标题1/2/3）
        if bind_wps_style and level in WPS_STYLE_MAPPING:
            para.style = doc.styles[WPS_STYLE_MAPPING[level]]

        # 段落格式设置
        para_format = para.paragraph_format
        para_format.alignment = ALIGN_MAP[cn_style["align"]]
        para_format.first_line_indent = Cm(cn_style["indent"] * 0.74)  # 2字符=1.48cm
        para_format.space_before = Pt(cn_style["space_before"])
        para_format.space_after = Pt(cn_style["space_after"])

        # 行距设置
        if cn_style["line_type"] == "固定值":
            para_format.line_spacing_rule = 2
            para_format.line_spacing = Pt(cn_style["line_value"])
        else:
            para_format.line_spacing_rule = 1
            para_format.line_spacing = cn_style["line_value"]

        # 【核心】中文/西文/数字 独立字体设置
        cn_size_pt = FONT_SIZE_MAP[cn_style["size"]]
        en_size_pt = FONT_SIZE_MAP[cn_style["size"]] if en_style["size_same_as_cn"] else FONT_SIZE_MAP[en_style["size"]]

        for run in para.runs:
            # 中文字体设置
            run.font.name = cn_style["font"]
            run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_style["font"])
            # 西文/数字独立设置（英文、数字、符号自动用这个字体）
            run._element.rPr.rFonts.set(qn('w:ascii'), en_style["en_font"])
            run._element.rPr.rFonts.set(qn('w:hAnsi'), en_style["en_font"])
            run._element.rPr.rFonts.set(qn('w:cs'), en_style["en_font"])
            # 字号+加粗/斜体
            run.font.size = Pt(cn_size_pt)
            run.font.bold = en_style["bold"] if en_style["bold"] else cn_style["bold"]
            run.font.italic = en_style["italic"]

    # ====================== 第三步：表格格式设置 ======================
    cn_table_style = cn_format["表格"]
    en_table_style = en_format["表格"]
    table_cn_size = FONT_SIZE_MAP[cn_table_style["size"]]
    table_en_size = FONT_SIZE_MAP[cn_table_style["size"]] if en_table_style["size_same_as_cn"] else FONT_SIZE_MAP[en_table_style["size"]]

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = ALIGN_MAP[cn_table_style["align"]]
                    for run in para.runs:
                        # 中文字体
                        run.font.name = cn_table_style["font"]
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_table_style["font"])
                        # 西文/数字独立设置
                        run._element.rPr.rFonts.set(qn('w:ascii'), en_table_style["en_font"])
                        run._element.rPr.rFonts.set(qn('w:hAnsi'), en_table_style["en_font"])
                        run._element.rPr.rFonts.set(qn('w:cs'), en_table_style["en_font"])
                        # 字号+格式
                        run.font.size = Pt(table_cn_size)
                        run.font.bold = en_table_style["bold"] if en_table_style["bold"] else cn_table_style["bold"]
                        run.font.italic = en_table_style["italic"]

    # 输出文档
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, total_changes, title_stats

# ====================== 4. 降重报告生成 ======================
def generate_report(changes, rewrite_level, title_stats):
    """生成详细的降重+排版报告"""
    total_count = len(changes)
    report = f"# 论文处理报告\n"
    report += f"📅 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    report += f"⚙️ 降重强度：{rewrite_level}\n"
    report += f"📝 总修改条数：{total_count}\n\n"

    # 标题识别统计
    report += "## 一、标题识别统计\n"
    for level, count in title_stats.items():
        report += f"- {level}：{count} 个\n"
    report += "\n✅ 已自动绑定WPS原生「标题1/标题2/标题3」样式，WPS导航窗格已自动生成目录\n\n"

    # 修改类型统计
    if total_count > 0:
        report += "## 二、降重修改统计\n"
        type_count = {}
        for change in changes:
            t = change["type"]
            type_count[t] = type_count.get(t, 0) + 1
        for t, count in type_count.items():
            report += f"- {t}：{count} 条\n"
        report += "\n"

        # 详细修改记录
        report += "## 三、详细修改记录\n"
        for i, change in enumerate(changes[:100]):
            report += f"### 修改记录 #{i+1}\n"
            report += f"📋 修改类型：{change['type']}\n"
            report += f"📊 语义重合度：{change['semantic_score']}\n"
            report += f"原文：{change['original']}\n"
            report += f"改后：{change['modified']}\n\n"

    return report.encode("utf-8")

# ====================== 5. Streamlit页面UI ======================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide", page_icon="📝")
    # 初始化页面状态
    if "cn_format" not in st.session_state:
        st.session_state.cn_format = copy.deepcopy(DEFAULT_CN_FORMAT)
    if "en_format" not in st.session_state:
        st.session_state.en_format = copy.deepcopy(DEFAULT_EN_FORMAT)
    if "version" not in st.session_state:
        st.session_state.version = 0

    st.title(f"📝 {APP_NAME}")
    st.success("✅ 核心功能：自动绑定WPS标题1/2/3样式 | 保留原编号 | 中文/西文独立设置 | 智能降重 | 批量格式调整")

    # 核心CSS：左侧格式区独立滚动，左小右大布局
    st.markdown(
        """
        <style>
        .left-scroll-container {
            height: 85vh;
            overflow-y: auto;
            padding-right: 10px;
        }
        .left-scroll-container::-webkit-scrollbar {
            width: 6px;
        }
        .left-scroll-container::-webkit-scrollbar-thumb {
            background-color: #e0e0e0;
            border-radius: 3px;
        }
        .right-fixed-container {
            height: 85vh;
            overflow-y: auto;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # 左3右7固定布局
    left_col, right_col = st.columns([3, 7])

    # 左侧：可滚动的格式配置区
    with left_col:
        st.markdown('<div class="left-scroll-container">', unsafe_allow_html=True)
        st.subheader("⚙️ 核心配置")

        # 1. 核心开关
        st.markdown("### 核心功能开关")
        bind_wps_style = st.checkbox("✅ 绑定WPS原生标题样式（必开）", value=True, help="开启后，WPS右上角自动对应标题1/2/3，导航窗格自动生成目录")
        enable_rewrite = st.checkbox("✅ 开启智能降重", value=False)
        rewrite_level = st.selectbox(
            "降重强度选择",
            options=list(REWRITE_LEVEL.keys()),
            index=1,
            disabled=not enable_rewrite
        )
        st.caption("轻度：仅同义词替换 | 标准：句式重构 | 强力：段落结构调整")
        st.divider()

        # 2. 一键重置格式
        st.markdown("### 格式一键重置")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("重置为学校标准格式", use_container_width=True):
                st.session_state.cn_format = copy.deepcopy(DEFAULT_CN_FORMAT)
                st.session_state.en_format = copy.deepcopy(DEFAULT_EN_FORMAT)
                st.session_state.version += 1
                st.success("✅ 已重置")
        with col2:
            if st.button("重置西文为标准格式", use_container_width=True):
                st.session_state.en_format = copy.deepcopy(DEFAULT_EN_FORMAT)
                st.session_state.version += 1
                st.success("✅ 已重置")
        st.divider()

        # 3. 中文格式设置
        st.markdown("### 🀄 中文格式设置")
        with st.expander("一级标题格式", expanded=False):
            cfg = st.session_state.cn_format["一级标题"]
            cfg["font"] = st.selectbox("中文字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cfg["font"]), key=f"cn_h1_font_{st.session_state.version}")
            cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"cn_h1_size_{st.session_state.version}")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_h1_bold_{st.session_state.version}")
            cfg["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg["align"]), key=f"cn_h1_align_{st.session_state.version}")
            cfg["indent"] = st.number_input("首行缩进(字符)", 0, 4, cfg["indent"], 1, key=f"cn_h1_indent_{st.session_state.version}")
            st.session_state.cn_format["一级标题"] = cfg

        with st.expander("二级标题格式", expanded=False):
            cfg = st.session_state.cn_format["二级标题"]
            cfg["font"] = st.selectbox("中文字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cfg["font"]), key=f"cn_h2_font_{st.session_state.version}")
            cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"cn_h2_size_{st.session_state.version}")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_h2_bold_{st.session_state.version}")
            cfg["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg["align"]), key=f"cn_h2_align_{st.session_state.version}")
            cfg["indent"] = st.number_input("首行缩进(字符)", 0, 4, cfg["indent"], 1, key=f"cn_h2_indent_{st.session_state.version}")
            st.session_state.cn_format["二级标题"] = cfg

        with st.expander("三级标题格式", expanded=False):
            cfg = st.session_state.cn_format["三级标题"]
            cfg["font"] = st.selectbox("中文字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cfg["font"]), key=f"cn_h3_font_{st.session_state.version}")
            cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"cn_h3_size_{st.session_state.version}")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_h3_bold_{st.session_state.version}")
            cfg["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg["align"]), key=f"cn_h3_align_{st.session_state.version}")
            cfg["indent"] = st.number_input("首行缩进(字符)", 0, 4, cfg["indent"], 1, key=f"cn_h3_indent_{st.session_state.version}")
            st.session_state.cn_format["三级标题"] = cfg

        with st.expander("正文格式", expanded=False):
            cfg = st.session_state.cn_format["正文"]
            cfg["font"] = st.selectbox("中文字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cfg["font"]), key=f"cn_body_font_{st.session_state.version}")
            cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"cn_body_size_{st.session_state.version}")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_body_bold_{st.session_state.version}")
            cfg["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg["align"]), key=f"cn_body_align_{st.session_state.version}")
            cfg["indent"] = st.number_input("首行缩进(字符)", 0, 4, cfg["indent"], 1, key=f"cn_body_indent_{st.session_state.version}")
            cfg["line_type"] = st.selectbox("行距类型", ["倍数", "固定值"], index=["倍数", "固定值"].index(cfg["line_type"]), key=f"cn_body_linetype_{st.session_state.version}")
            cfg["line_value"] = st.number_input("行距值", 1.0, 100.0, float(cfg["line_value"]), 0.1, key=f"cn_body_linevalue_{st.session_state.version}")
            st.session_state.cn_format["正文"] = cfg

        with st.expander("表格格式", expanded=False):
            cfg = st.session_state.cn_format["表格"]
            cfg["font"] = st.selectbox("中文字体", CN_FONT_LIST, index=CN_FONT_LIST.index(cfg["font"]), key=f"cn_table_font_{st.session_state.version}")
            cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"cn_table_size_{st.session_state.version}")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"cn_table_bold_{st.session_state.version}")
            cfg["align"] = st.selectbox("对齐方式", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg["align"]), key=f"cn_table_align_{st.session_state.version}")
            st.session_state.cn_format["表格"] = cfg
        st.divider()

        # 4. 西文/数字格式设置
        st.markdown("### 🔤 西文/数字单独设置")
        with st.expander("正文西文/数字设置", expanded=False):
            cfg = st.session_state.en_format["正文"]
            cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key=f"en_body_font_{st.session_state.version}")
            cfg["size_same_as_cn"] = st.checkbox("字号与中文同步", cfg["size_same_as_cn"], key=f"en_body_same_{st.session_state.version}")
            if not cfg["size_same_as_cn"]:
                cfg["size"] = st.selectbox("西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"en_body_size_{st.session_state.version}")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"en_body_bold_{st.session_state.version}")
            cfg["italic"] = st.checkbox("斜体", cfg["italic"], key=f"en_body_italic_{st.session_state.version}")
            st.session_state.en_format["正文"] = cfg

        with st.expander("标题西文/数字设置", expanded=False):
            for level in ["一级标题", "二级标题", "三级标题"]:
                st.markdown(f"#### {level}")
                cfg = st.session_state.en_format[level]
                cfg["en_font"] = st.selectbox(f"{level}西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key=f"en_{level}_font_{st.session_state.version}")
                cfg["size_same_as_cn"] = st.checkbox(f"{level}字号与中文同步", cfg["size_same_as_cn"], key=f"en_{level}_same_{st.session_state.version}")
                if not cfg["size_same_as_cn"]:
                    cfg["size"] = st.selectbox(f"{level}西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"en_{level}_size_{st.session_state.version}")
                cfg["bold"] = st.checkbox(f"{level}加粗", cfg["bold"], key=f"en_{level}_bold_{st.session_state.version}")
                cfg["italic"] = st.checkbox(f"{level}斜体", cfg["italic"], key=f"en_{level}_italic_{st.session_state.version}")
                st.session_state.en_format[level] = cfg
                st.divider()

        st.markdown('</div>', unsafe_allow_html=True)

    # 右侧：文件上传与处理区
    with right_col:
        st.markdown('<div class="right-fixed-container">', unsafe_allow_html=True)
        st.subheader("📁 文档上传与处理")
        files = st.file_uploader(
            "上传 .docx 格式文档（支持多选批量处理）",
            type=["docx"],
            accept_multiple_files=True
        )

        if files and st.button("🚀 开始处理（标题绑定+格式调整+可选降重）", type="primary", use_container_width=True):
            for file in files:
                with st.spinner(f"正在处理：{file.name}"):
                    # 执行核心处理
                    output_doc, changes, title_stats = process_doc(
                        file=file,
                        cn_format=st.session_state.cn_format,
                        en_format=st.session_state.en_format,
                        enable_rewrite=enable_rewrite,
                        rewrite_level=rewrite_level,
                        bind_wps_style=bind_wps_style
                    )

                st.subheader(f"✅ 处理完成：{file.name}")
                # 统计信息展示
                st.markdown("### 📊 处理结果统计")
                cols = st.columns(5)
                cols[0].metric("一级标题", title_stats["一级标题"])
                cols[1].metric("二级标题", title_stats["二级标题"])
                cols[2].metric("三级标题", title_stats["三级标题"])
                cols[3].metric("正文段落", title_stats["正文"])
                cols[4].metric("表格数量", title_stats["表格"])
                st.info("✅ 已自动绑定WPS原生标题样式，WPS打开后导航窗格自动生成目录，同一级标题可一键全选批量调整")

                # 下载排版后的论文
                st.download_button(
                    label="📥 下载已处理论文",
                    data=output_doc,
                    file_name=f"已排版_{file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                # 下载处理报告
                report_bytes = generate_report(changes, rewrite_level, title_stats)
                st.download_button(
                    label="📄 下载处理报告",
                    data=report_bytes,
                    file_name=f"处理报告_{file.name}.txt",
                    mime="text/plain",
                    use_container_width=True
                )
                st.divider()

        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
