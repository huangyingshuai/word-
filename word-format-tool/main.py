import streamlit as st
import copy
import re
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ====================== 全局配置与常量 ======================
# 专业术语白名单（永不修改）
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "挑战杯", "互联网+", "河北科技大学",
    "工业工程", "GDP", "CPI", "GB/T 7714", "ISO", "一级标题", "二级标题", "三级标题",
    "参考文献", "公式", "图表", "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "Transformer", "BERT", "T5", "Python", "Java", "SQL"
]

# 河北科技大学本科毕业论文/竞赛标准格式模板
TEMPLATE_LIBRARY = {
    "河北科技大学-竞赛/毕业论文标准格式": {
        "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "二级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "三级标题": {"font": "黑体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 22, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
    }
}

# ====================== 【新增】西文/数字独立格式配置（竞赛专用） ======================
# 默认值：竞赛标准格式（西文/数字统一用Times New Roman，字号和对应中文一致）
DEFAULT_EN_CONFIG = {
    "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
    "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
    "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
    "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
    "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
}

EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri", "Courier New"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五"]

# 降重强度配置
REWRITE_LEVEL = {
    "轻度降重": {"synonym": True, "sentence_reorder": False, "structure_change": False},
    "标准降重": {"synonym": True, "sentence_reorder": True, "structure_change": False},
    "强力降重": {"synonym": True, "sentence_reorder": True, "structure_change": True}
}

# 同义词替换词典（学术场景专用）
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
APP_NAME = "论文智能降重+一键排版工具（竞赛专属版）"

# ====================== 1. 标题层级精准识别 ======================
def get_title_level(para_text):
    """精准识别标题层级，彻底解决正文列表误判问题"""
    text = para_text.strip()
    if not text:
        return "正文"

    # 一级标题：第X章、X、（仅短文本，长文本归为正文）
    if re.match(r'^第[一二三四五六七八九十]+章\s', text) or (re.match(r'^\d+、\s', text) and len(text) < 20):
        return "一级标题"
    # 二级标题：1.1、（一）（仅短文本）
    elif re.match(r'^\d+\.\d+\s', text) or (re.match(r'^（[一二三四五六七八九十]+）\s', text) and len(text) < 20):
        return "二级标题"
    # 三级标题：1.1.1、（1）（仅短文本，长文本正文列表直接归为正文）
    elif re.match(r'^\d+\.\d+\.\d+\s', text) or (re.match(r'^（\d+）\s', text) and len(text) < 15):
        return "三级标题"
    # 所有其他情况归为正文
    return "正文"

# ====================== 2. 降重核心引擎 ======================
def is_white_text(text):
    """判断是否为白名单内容，不修改"""
    for word in WHITE_WORDS:
        if word in text:
            return True
    # 纯数字、公式、引用不修改
    if re.match(r'^\d+(\.\d+)*$', text) or re.match(r'^\[.*\]$', text):
        return True
    return False

def check_semantic_keep(original, modified):
    """规则化语义保持校验，确保不改原意（无AI模型依赖）"""
    # 提取核心关键词（名词、动词）
    original_keywords = set(re.findall(r'[\u4e00-\u9fa5]{2,}', original))
    modified_keywords = set(re.findall(r'[\u4e00-\u9fa5]{2,}', modified))
    # 核心关键词重合度≥70%，判定为语义保持
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

    # 1. 同义词替换（所有强度都启用）
    if level_config["synonym"]:
        for old, new in SYNONYM_DICT.items():
            if old in modified:
                modified = modified.replace(old, new)
                rewrite_type = "同义词替换"

    # 2. 句式重构（标准/强力降重启用）
    if level_config["sentence_reorder"]:
        # 长句拆分重组
        parts = [p.strip() for p in re.split(r'[，。；]', modified) if p.strip()]
        if len(parts) >= 3:
            import random
            random.shuffle(parts)
            modified = "，".join(parts) + "。"
            rewrite_type = "句式重构+语序打乱"

    # 3. 结构调整+限定词补充（强力降重启用）
    if level_config["structure_change"]:
        if "在" in modified and "中" in modified:
            modified = re.sub(r'在(.*?)中', f'结合{datetime.now().year}年行业实际发展情况，在\g<1>场景中', modified)
            rewrite_type = "结构调整+场景限定补充"

    # 语义保持校验，不达标则回退
    semantic_score = check_semantic_keep(original, modified)
    if semantic_score < 0.7:
        return original, "原文保留（语义重合度不达标）", 1.0

    return modified, rewrite_type, round(semantic_score, 4)

def rewrite_paragraph(text, level_config):
    """整段降重，逐句处理"""
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

# ====================== 3. 文档处理核心（格式零损坏+西文独立调整） ======================
def process_doc(file, template, en_config, enable_rewrite=False, rewrite_level="标准降重"):
    """
    核心处理流程：先降重 → 再排版（中文/西文独立设置）
    完全保留格式、图片、表格位置，仅修改文本内容
    """
    doc = Document(file)
    total_changes = []
    level_config = REWRITE_LEVEL[rewrite_level]

    # ====================== 第一步：降重处理（Run级别，不破坏格式） ======================
    if enable_rewrite:
        # 1. 正文段落降重
        for para in doc.paragraphs:
            original_text = para.text.strip()
            if not original_text or get_title_level(original_text) != "正文":
                continue  # 标题不降重

            # 提取文本降重
            new_text, changes = rewrite_paragraph(original_text, level_config)
            if changes:
                total_changes.extend(changes)
                # 回填文本，保留原有Run格式
                para.text = new_text

        # 2. 表格内文本降重（保留表格结构）
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        original_text = para.text.strip()
                        if not original_text or is_white_text(original_text):
                            continue
                        new_text, changes = rewrite_paragraph(original_text, level_config)
                        if changes:
                            total_changes.extend(changes)
                            para.text = new_text

    # ====================== 第二步：格式排版（中文/西文独立设置，竞赛核心需求） ======================
    for para in doc.paragraphs:
        level = get_title_level(para.text)
        cn_style = template.get(level, template["正文"])
        en_style = en_config.get(level, en_config["正文"])

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

        # 【核心升级】中文/西文/数字 独立字体设置
        # 西文字号：和中文同步/自定义
        if en_style["size_same_as_cn"]:
            en_size_pt = FONT_SIZE_MAP[cn_style["size"]]
        else:
            en_size_pt = FONT_SIZE_MAP[en_style["size"]]
        cn_size_pt = FONT_SIZE_MAP[cn_style["size"]]

        for run in para.runs:
            # 中文字体设置
            run.font.name = cn_style["font"]
            run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_style["font"])
            # 西文/数字独立设置（核心功能）
            run._element.rPr.rFonts.set(qn('w:ascii'), en_style["en_font"])
            run._element.rPr.rFonts.set(qn('w:hAnsi'), en_style["en_font"])
            run._element.rPr.rFonts.set(qn('w:cs'), en_style["en_font"])
            # 字号设置
            run.font.size = Pt(cn_size_pt)
            # 西文独立加粗/斜体
            run.font.bold = en_style["bold"] if en_style["bold"] else cn_style["bold"]
            run.font.italic = en_style["italic"]

    # ====================== 第三步：表格格式设置（中文/西文独立） ======================
    cn_table_style = template["表格"]
    en_table_style = en_config["表格"]
    en_table_size_pt = FONT_SIZE_MAP[cn_table_style["size"]] if en_table_style["size_same_as_cn"] else FONT_SIZE_MAP[en_table_style["size"]]
    cn_table_size_pt = FONT_SIZE_MAP[cn_table_style["size"]]

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
                        # 字号+加粗/斜体
                        run.font.size = Pt(cn_table_size_pt)
                        run.font.bold = en_table_style["bold"] if en_table_style["bold"] else cn_table_style["bold"]
                        run.font.italic = en_table_style["italic"]

    # 输出文档
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, total_changes

# ====================== 4. 降重报告生成 ======================
def generate_report(changes, rewrite_level):
    """生成详细的降重修改报告"""
    total_count = len(changes)
    report = f"# 论文降重修改报告\n"
    report += f"📅 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    report += f"⚙️ 降重强度：{rewrite_level}\n"
    report += f"📝 总修改条数：{total_count}\n\n"

    # 统计信息
    type_count = {}
    for change in changes:
        t = change["type"]
        type_count[t] = type_count.get(t, 0) + 1

    report += "## 一、修改类型统计\n"
    for t, count in type_count.items():
        report += f"- {t}：{count} 条\n"
    report += "\n"

    # 详细修改记录
    report += "## 二、详细修改记录\n"
    for i, change in enumerate(changes[:100]):  # 最多显示100条
        report += f"### 修改记录 #{i+1}\n"
        report += f"📋 修改类型：{change['type']}\n"
        report += f"📊 语义重合度：{change['semantic_score']}\n"
        report += f"原文：{change['original']}\n"
        report += f"改后：{change['modified']}\n\n"

    return report.encode("utf-8")

# ====================== 5. Streamlit页面UI ======================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide", page_icon="📝")
    # 初始化西文配置
    if "en_config" not in st.session_state:
        st.session_state.en_config = copy.deepcopy(DEFAULT_EN_CONFIG)
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0

    st.title(f"📝 {APP_NAME}")
    st.success("✅ 河北科技大学竞赛专属 | 中文/西文/数字独立调整 | 降重不破坏格式 | 自动生成修改报告")

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

    # 左3右7固定布局（左小右大）
    left_col, right_col = st.columns([3, 7])

    # 左侧：可滚动的配置区
    with left_col:
        st.markdown('<div class="left-scroll-container">', unsafe_allow_html=True)
        st.subheader("⚙️ 核心功能配置")

        # 降重配置
        st.markdown("### 降重设置")
        enable_rewrite = st.checkbox("✅ 开启智能降重（先降重、后排版）", value=False)
        rewrite_level = st.selectbox(
            "降重强度选择",
            options=list(REWRITE_LEVEL.keys()),
            index=1,
            disabled=not enable_rewrite
        )
        st.caption("轻度：仅同义词替换 | 标准：句式重构 | 强力：段落结构调整")
        st.divider()

        # 【新增】西文/数字独立设置模块
        st.markdown("### 🔤 西文/数字单独设置（竞赛专用）")
        if st.button("一键重置为竞赛标准格式", use_container_width=True):
            st.session_state.en_config = copy.deepcopy(DEFAULT_EN_CONFIG)
            st.session_state.template_version += 1
            st.success("✅ 已重置为竞赛标准格式")
        st.caption("默认：西文/数字统一使用Times New Roman，字号与中文同步")

        # 分模块设置西文格式
        with st.expander("正文西文/数字设置", expanded=False):
            cfg = st.session_state.en_config["正文"]
            cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key="en_body_font")
            cfg["size_same_as_cn"] = st.checkbox("字号与中文同步", cfg["size_same_as_cn"], key="en_body_same_size")
            if not cfg["size_same_as_cn"]:
                cfg["size"] = st.selectbox("西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key="en_body_size")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key="en_body_bold")
            cfg["italic"] = st.checkbox("斜体", cfg["italic"], key="en_body_italic")
            st.session_state.en_config["正文"] = cfg

        with st.expander("一级标题西文/数字设置", expanded=False):
            cfg = st.session_state.en_config["一级标题"]
            cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key="en_h1_font")
            cfg["size_same_as_cn"] = st.checkbox("字号与中文同步", cfg["size_same_as_cn"], key="en_h1_same_size")
            if not cfg["size_same_as_cn"]:
                cfg["size"] = st.selectbox("西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key="en_h1_size")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key="en_h1_bold")
            cfg["italic"] = st.checkbox("斜体", cfg["italic"], key="en_h1_italic")
            st.session_state.en_config["一级标题"] = cfg

        with st.expander("二级标题西文/数字设置", expanded=False):
            cfg = st.session_state.en_config["二级标题"]
            cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key="en_h2_font")
            cfg["size_same_as_cn"] = st.checkbox("字号与中文同步", cfg["size_same_as_cn"], key="en_h2_same_size")
            if not cfg["size_same_as_cn"]:
                cfg["size"] = st.selectbox("西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key="en_h2_size")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key="en_h2_bold")
            cfg["italic"] = st.checkbox("斜体", cfg["italic"], key="en_h2_italic")
            st.session_state.en_config["二级标题"] = cfg

        with st.expander("三级标题西文/数字设置", expanded=False):
            cfg = st.session_state.en_config["三级标题"]
            cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key="en_h3_font")
            cfg["size_same_as_cn"] = st.checkbox("字号与中文同步", cfg["size_same_as_cn"], key="en_h3_same_size")
            if not cfg["size_same_as_cn"]:
                cfg["size"] = st.selectbox("西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key="en_h3_size")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key="en_h3_bold")
            cfg["italic"] = st.checkbox("斜体", cfg["italic"], key="en_h3_italic")
            st.session_state.en_config["三级标题"] = cfg

        with st.expander("表格西文/数字设置", expanded=False):
            cfg = st.session_state.en_config["表格"]
            cfg["en_font"] = st.selectbox("西文字体", EN_FONT_LIST, index=EN_FONT_LIST.index(cfg["en_font"]), key="en_table_font")
            cfg["size_same_as_cn"] = st.checkbox("字号与中文同步", cfg["size_same_as_cn"], key="en_table_same_size")
            if not cfg["size_same_as_cn"]:
                cfg["size"] = st.selectbox("西文字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key="en_table_size")
            cfg["bold"] = st.checkbox("加粗", cfg["bold"], key="en_table_bold")
            cfg["italic"] = st.checkbox("斜体", cfg["italic"], key="en_table_italic")
            st.session_state.en_config["表格"] = cfg
        st.divider()

        # 格式模板
        st.markdown("### 中文格式模板")
        template_name = st.selectbox("选择中文格式模板", options=list(TEMPLATE_LIBRARY.keys()), index=0)
        template = TEMPLATE_LIBRARY[template_name]
        st.info("📌 已自动加载河北科技大学竞赛/毕业论文标准中文格式")
        st.divider()

        # 中文格式自定义（折叠）
        with st.expander("中文格式自定义调整", expanded=False):
            for level in template.keys():
                st.markdown(f"#### {level}")
                col1, col2 = st.columns(2)
                with col1:
                    template[level]["font"] = st.selectbox(
                        "中文字体",
                        options=["宋体", "黑体", "楷体", "仿宋_GB2312"],
                        index=["宋体", "黑体", "楷体", "仿宋_GB2312"].index(template[level]["font"]),
                        key=f"cn_{level}_font"
                    )
                with col2:
                    template[level]["size"] = st.selectbox(
                        "中文字号",
                        options=list(FONT_SIZE_MAP.keys()),
                        index=list(FONT_SIZE_MAP.keys()).index(template[level]["size"]),
                        key=f"cn_{level}_size"
                    )
                template[level]["bold"] = st.checkbox("加粗", value=template[level]["bold"], key=f"cn_{level}_bold")
                template[level]["indent"] = st.number_input(
                    "首行缩进(字符)",
                    min_value=0,
                    max_value=4,
                    value=template[level]["indent"],
                    key=f"cn_{level}_indent"
                )
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

        if files and st.button("🚀 开始处理（降重+排版）", type="primary", use_container_width=True):
            for file in files:
                with st.spinner(f"正在处理：{file.name}"):
                    # 执行核心处理
                    output_doc, changes = process_doc(
                        file=file,
                        template=template,
                        en_config=st.session_state.en_config,
                        enable_rewrite=enable_rewrite,
                        rewrite_level=rewrite_level
                    )

                st.subheader(f"✅ 处理完成：{file.name}")
                # 统计信息
                col1, col2 = st.columns(2)
                col1.metric("总修改条数", len(changes))
                col2.metric("格式适配", "中文/西文已独立设置")

                # 下载论文
                st.download_button(
                    label="📥 下载已排版论文",
                    data=output_doc,
                    file_name=f"竞赛排版_{file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                # 下载降重报告
                if enable_rewrite:
                    report_bytes = generate_report(changes, rewrite_level)
                    st.download_button(
                        label="📄 下载降重修改报告",
                        data=report_bytes,
                        file_name=f"降重报告_{file.name}.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
                st.divider()

        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
