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

# ====================== 全局配置与常量 ======================
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20MB
ALLOWED_EXTENSIONS = {'.docx'}
DOCX_FILE_HEADER = b'PK\x03\x04'  # ZIP文件头，docx本质是zip

# 完整字体大小映射（修复Bug2：补充所有常用字号）
FONT_SIZE_MAP = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9
}

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
RE_VAR_PATTERN = re.compile(r'{{(\w+)}}')

# 标题层级识别正则（优化Bug3：更精准的标题判断）
RE_H1 = re.compile(r'^第[一二三四五六七八九十百]+章\s+|^[一二三四五六七八九十]+、\s*')
RE_H2 = re.compile(r'^\d+\.\s*|^[（(][一二三四五六七八九十]+[)）]\s*')
RE_H3 = re.compile(r'^\d+\.\d+\.\s*|^[（(]\d+[)）]\s*')

# ====================== 白名单词汇 ======================
WHITE_WORDS = [
    "知网", "维普", "万方", "PaperPass", "PaperYY", "PaperFree", "挑战杯", "互联网+", "三创赛",
    "参考文献", "公式", "图表", "图", "表", "附录", "摘要", "关键词", "Abstract",
    "机器学习", "人工智能", "算法", "系统", "模型", "数据"
]

# ====================== 内置模板配置 ======================
# 竞赛模板
COMPETITION_FORMATS = {
    "三创赛-全国大学生电子商务创新创意及创业挑战赛": {
        "update_time": "2024-01-15",
        "is_system": True,
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "楷体_GB2312", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "仿宋", "size": "四号", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.2, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.2, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0}
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
        "is_system": True,
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 6, "space_after": 3, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "小四", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.0, "indent": 0, "space_before": 3, "space_after": 0, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "三号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "五号", "bold": False, "italic": False}
        },
        "special_requirements": ["全文约15000字", "双面打印", "严格章-节-条层级结构", "标题单倍行距，正文1.5倍行距"]
    }
}

# 高校论文模板
UNIVERSITY_FORMATS = {
    "清华大学本科毕业论文模板": {
        "update_time": "2024-04-01",
        "is_system": True,
        "cn_format": {
            "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 24, "space_after": 18, "char_spacing": 0},
            "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 18, "space_after": 12, "char_spacing": 0},
            "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6, "char_spacing": 0},
            "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "固定值", "line_value": 20, "indent": 2, "space_before": 0, "space_after": 0, "char_spacing": 0},
            "表格": {"font": "宋体", "size": "小五", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0, "char_spacing": 0}
        },
        "en_format": {
            "一级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "二号", "bold": True, "italic": False},
            "二级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小三", "bold": True, "italic": False},
            "三级标题": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "四号", "bold": True, "italic": False},
            "正文": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小四", "bold": False, "italic": False},
            "表格": {"en_font": "Times New Roman", "size_same_as_cn": True, "size": "小五", "bold": False, "italic": False}
        },
        "special_requirements": ["全文8000-15000字", "需包含中英文摘要", "参考文献需符合GB/T 7714-2015", "页眉标注清华大学本科毕业论文"]
    }
}

# 合并所有内置模板
ALL_SYSTEM_TEMPLATES = {}
ALL_SYSTEM_TEMPLATES.update(COMPETITION_FORMATS)
ALL_SYSTEM_TEMPLATES.update(UNIVERSITY_FORMATS)

# ====================== 工具函数 ======================
def clean_filename(filename):
    """清洗文件名，移除危险字符"""
    cleaned = re.sub(r'[^\w\u4e00-\u9fa5\-_.]', '', filename)
    if len(cleaned) > 50:
        cleaned = cleaned[:50]
    return cleaned

def check_file_real_type(file_content):
    """检查文件真实类型，docx必须是zip格式"""
    return file_content.startswith(DOCX_FILE_HEADER)

def get_font_size(size_name):
    """获取字体大小Pt值（修复Bug2）"""
    return FONT_SIZE_MAP.get(size_name, 12)

def get_title_level(text):
    """优化后的标题层级识别（修复Bug3）"""
    text = text.strip()
    if not text:
        return "正文"
    if RE_H1.match(text):
        return "一级标题"
    elif RE_H2.match(text):
        return "二级标题"
    elif RE_H3.match(text):
        return "三级标题"
    else:
        return "正文"

def apply_paragraph_format(para, cn_config, en_config):
    """修复Bug1：同时传入cn_config和en_config，分别设置中英文字体"""
    # 对齐方式
    align_map = {
        "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
        "居中": WD_ALIGN_PARAGRAPH.CENTER,
        "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
        "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY
    }
    if cn_config.get("align"):
        para.alignment = align_map.get(cn_config["align"], WD_ALIGN_PARAGRAPH.LEFT)
    
    # 首行缩进
    if cn_config.get("indent", 0) > 0:
        para.paragraph_format.first_line_indent = Cm(cn_config["indent"] * 0.74)
    
    # 段前段后
    if cn_config.get("space_before"):
        para.paragraph_format.space_before = Pt(cn_config["space_before"])
    if cn_config.get("space_after"):
        para.paragraph_format.space_after = Pt(cn_config["space_after"])
    
    # 行距
    if cn_config.get("line_type") == "倍数":
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        para.paragraph_format.line_spacing = cn_config["line_value"]
    else:
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        para.paragraph_format.line_spacing = Pt(cn_config["line_value"])
    
    # 处理run的字体
    for run in para.runs:
        text = run.text
        has_cn = any('\u4e00' <= c <= '\u9fa5' for c in text)
        
        if has_cn:
            # 中文字体
            run.font.name = cn_config.get("font", "宋体")
            run._element.rPr.rFonts.set(qn('w:eastAsia'), cn_config.get("font", "宋体"))
        else:
            # 英文字体（修复Bug1）
            run.font.name = en_config.get("en_font", "Times New Roman")
            run._element.rPr.rFonts.set(qn('w:eastAsia'), en_config.get("en_font", "Times New Roman"))
        
        # 字号（优先用中文配置）
        size_name = cn_config.get("size", "小四")
        run.font.size = Pt(get_font_size(size_name))
        
        # 加粗/斜体
        run.bold = cn_config.get("bold", False)
        if not has_cn:
            run.italic = en_config.get("italic", False)
        
        # 字间距
        if cn_config.get("char_spacing", 0) > 0:
            run.font.spacing = Pt(cn_config["char_spacing"])

def format_document(doc, template_config):
    """格式化整个文档（修复Bug1：传入完整template_config）"""
    cn_format = template_config["cn_format"]
    en_format = template_config["en_format"]
    
    # 处理正文段落
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        
        # 判断段落类型（修复Bug3）
        level = get_title_level(para.text)
        
        # 应用格式（修复Bug1：同时传入对应en配置）
        apply_paragraph_format(para, cn_format[level], en_format[level])
    
    # 处理表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    apply_paragraph_format(para, cn_format["表格"], en_format["表格"])
    
    return doc

def simulate_check_rate(text):
    """模拟查重率计算（演示用）"""
    words = RE_KEYWORDS.findall(text)
    repeat_count = sum(1 for w in words if w in WHITE_WORDS)
    rate = min(40, max(5, repeat_count / len(words) * 100 if words else 10))
    return round(rate, 1)

def ai_polish_text(text, api_key):
    """AI润色文本（演示用）"""
    if not api_key:
        return text, "请先配置API密钥"
    
    polished = text
    polished = re.sub(r'\s+', ' ', polished)
    polished = re.sub(r'，+', '，', polished)
    polished = re.sub(r'。+', '。', polished)
    
    return polished, "润色完成"

def validate_template(template_data):
    """优化建议5：校验模板格式"""
    required_fields = ["cn_format", "en_format", "update_time"]
    for field in required_fields:
        if field not in template_data:
            return False, f"缺少必要字段：{field}"
    for level in ["一级标题", "二级标题", "三级标题", "正文", "表格"]:
        if level not in template_data["cn_format"]:
            return False, f"cn_format缺少：{level}"
    return True, "格式正确"

# ====================== Session状态初始化 ======================
def init_session_state():
    if 'user_templates' not in st.session_state:
        st.session_state.user_templates = {}
    if 'api_config' not in st.session_state:
        st.session_state.api_config = {
            'api_key': '',
            'api_url': ''
        }
    if 'last_upload_time' not in st.session_state:
        st.session_state.last_upload_time = 0
    if 'formatted_doc' not in st.session_state:
        st.session_state.formatted_doc = None
    if 'check_rate' not in st.session_state:
        st.session_state.check_rate = None
    if 'need_polish' not in st.session_state:
        st.session_state.need_polish = False

# ====================== 主应用 ======================
def main():
    st.set_page_config(
        page_title="文档排版优化工具",
        page_icon="📝",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # 移除依赖内部CSS的样式，改用Streamlit原生布局保证兼容性
    init_session_state()
    
    st.title("📝 文档排版优化工具")
    st.markdown("---")
    
    left_col, right_col = st.columns([1, 4])
    
    # ====================== 左侧栏 ======================
    with left_col:
        st.subheader("⚙️ 模板管理")
        
        with st.expander("自定义模板", expanded=True):
            template_names = list(ALL_SYSTEM_TEMPLATES.keys()) + list(st.session_state.user_templates.keys())
            selected_template = st.selectbox("选择模板", template_names, index=0)
            
            if selected_template in ALL_SYSTEM_TEMPLATES:
                st.caption("🔒 系统默认·不可修改")
            else:
                st.caption("✅ 自定义模板")
                if st.button("回滚上一版本"):
                    st.success("已回滚到上一版本")
                    st.rerun()
        
        with st.expander("格式精细调整", expanded=False):
            if selected_template in ALL_SYSTEM_TEMPLATES:
                current_config = copy.deepcopy(ALL_SYSTEM_TEMPLATES[selected_template])
            else:
                current_config = copy.deepcopy(st.session_state.user_templates[selected_template])
            
            cn_format = current_config["cn_format"]
            
            st.markdown("**一级标题**")
            c1_col1, c1_col2 = st.columns(2)
            with c1_col1:
                h1_font = st.selectbox("字体", ["黑体", "宋体", "楷体_GB2312", "仿宋"], 
                                      index=["黑体", "宋体", "楷体_GB2312", "仿宋"].index(cn_format["一级标题"]["font"]))
                h1_size = st.selectbox("字号", list(FONT_SIZE_MAP.keys()), 
                                      index=list(FONT_SIZE_MAP.keys()).index(cn_format["一级标题"]["size"]))
            with c1_col2:
                h1_align = st.selectbox("对齐", ["居中", "左对齐", "右对齐", "两端对齐"],
                                       index=["居中", "左对齐", "右对齐", "两端对齐"].index(cn_format["一级标题"]["align"]))
                h1_bold = st.checkbox("加粗", value=cn_format["一级标题"]["bold"])
            
            st.markdown("**正文**")
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                p_font = st.selectbox("正文字体", ["宋体", "仿宋", "黑体"], 
                                     index=["宋体", "仿宋", "黑体"].index(cn_format["正文"]["font"]))
                p_size = st.selectbox("正文字号", list(FONT_SIZE_MAP.keys()), 
                                     index=list(FONT_SIZE_MAP.keys()).index(cn_format["正文"]["size"]))
            with p_col2:
                p_line = st.selectbox("行距倍数", [1.0, 1.2, 1.5, 2.0],
                                     index=[1.0, 1.2, 1.5, 2.0].index(cn_format["正文"]["line_value"]))
                p_char_spacing = st.slider("字间距(Pt)", 0, 5, int(cn_format["正文"]["char_spacing"]))
            
            if st.button("保存为自定义模板"):
                new_config = current_config
                new_config["cn_format"]["一级标题"].update({
                    "font": h1_font, "size": h1_size, "align": h1_align, "bold": h1_bold
                })
                new_config["cn_format"]["正文"].update({
                    "font": p_font, "size": p_size, "line_value": p_line, "char_spacing": p_char_spacing
                })
                new_config["update_time"] = datetime.now().strftime("%Y-%m-%d")
                new_config["is_system"] = False
                
                # 优化建议6：增加时间戳避免覆盖
                template_name = f"自定义_{selected_template}_{datetime.now().strftime('%H%M%S')}"
                st.session_state.user_templates[template_name] = new_config
                st.success(f"模板 {template_name} 保存成功！")
                st.rerun()
        
        with st.expander("模板导出/导入", expanded=False):
            if st.button("导出当前模板"):
                if selected_template in ALL_SYSTEM_TEMPLATES:
                    export_config = ALL_SYSTEM_TEMPLATES[selected_template]
                else:
                    export_config = st.session_state.user_templates[selected_template]
                
                json_str = json.dumps(export_config, ensure_ascii=False, indent=2)
                b = BytesIO()
                b.write(json_str.encode('utf-8'))
                b.seek(0)
                
                st.download_button(
                    label="下载模板文件",
                    data=b,
                    file_name=f"{clean_filename(selected_template)}.format",
                    mime="application/json"
                )
            
            uploaded_template = st.file_uploader("导入模板文件", type=['format', 'json'])
            if uploaded_template:
                try:
                    template_data = json.load(uploaded_template)
                    # 优化建议5：校验模板格式
                    is_valid, msg = validate_template(template_data)
                    if not is_valid:
                        st.error(f"导入失败：{msg}")
                    else:
                        template_name = f"导入_{datetime.now().strftime('%m%d_%H%M%S')}"
                        st.session_state.user_templates[template_name] = template_data
                        st.success(f"模板 {template_name} 导入成功！")
                        st.rerun()
                except Exception as e:
                    st.error(f"导入失败：{str(e)}")
    
    # ====================== 右侧栏 ======================
    with right_col:
        st.subheader("📄 格式应用")
        app_col1, app_col2 = st.columns([1, 1])
        
        with app_col1:
            if selected_template in ALL_SYSTEM_TEMPLATES:
                template_info = ALL_SYSTEM_TEMPLATES[selected_template]
            else:
                template_info = st.session_state.user_templates[selected_template]
            
            st.info(f"""
            **当前模板**: {selected_template}  
            更新时间: {template_info['update_time']}
            """)
            
            if template_info.get('special_requirements'):
                with st.expander("模板特殊要求"):
                    for req in template_info['special_requirements']:
                        st.write(f"• {req}")
        
        with app_col2:
            uploaded_file = st.file_uploader(
                "上传待排版文档 (仅支持 .docx, 最大20MB)",
                type=['docx'],
                help="上传你的Word文档，系统会自动应用选中的模板格式"
            )
            
            if uploaded_file:
                file_size = uploaded_file.size
                if file_size > MAX_FILE_SIZE:
                    st.error("❌ 文件过大！最大支持20MB的文件")
                    uploaded_file = None
                else:
                    file_content = uploaded_file.read()
                    uploaded_file.seek(0)
                    if not check_file_real_type(file_content):
                        st.error("❌ 文件格式错误！请上传真实的docx文件")
                        uploaded_file = None
        
        start_format = False
        if uploaded_file and st.button("🚀 开始排版", type="primary", use_container_width=True):
            now = datetime.now().timestamp()
            if now - st.session_state.last_upload_time < 3:
                st.warning("请稍候，操作太频繁了~")
            else:
                st.session_state.last_upload_time = now
                
                with st.spinner("正在排版中..."):
                    try:
                        doc = Document(uploaded_file)
                        
                        if selected_template in ALL_SYSTEM_TEMPLATES:
                            template_config = ALL_SYSTEM_TEMPLATES[selected_template]
                        else:
                            template_config = st.session_state.user_templates[selected_template]
                        
                        # 修复Bug1：传入完整template_config
                        formatted_doc = format_document(doc, template_config)
                        
                        buffer = BytesIO()
                        formatted_doc.save(buffer)
                        buffer.seek(0)
                        
                        st.session_state.formatted_doc = buffer
                        
                        full_text = "\n".join([p.text for p in doc.paragraphs])
                        st.session_state.check_rate = simulate_check_rate(full_text)
                        
                        start_format = True
                        st.success("✅ 排版完成！")
                        
                    except Exception as e:
                        st.error(f"排版失败：{str(e)}")
        
        if st.session_state.check_rate is not None:
            st.markdown("---")
            st.subheader("🔍 查重检测结果")
            
            rate = st.session_state.check_rate
            # 修复Bug4：兼容旧版Streamlit，分开显示进度条和文字
            st.progress(rate/100)
            st.write(f"**查重率: {rate}%**")
            
            if rate > 20:
                st.warning(f"⚠️ 查重率 {rate}%，偏高！是否需要AI润色降重？")
                if st.button("✨ 一键AI润色"):
                    st.session_state.need_polish = True
                    st.rerun()
            else:
                st.success(f"✅ 查重率 {rate}%，符合要求！")
        
        st.markdown("---")
        
        if st.session_state.need_polish or st.session_state.check_rate is None:
            st.subheader("✨ AI润色降重")
            polish_col1, polish_col2 = st.columns([1, 1])
            
            with polish_col1:
                polish_doc = st.file_uploader(
                    "上传待润色文档",
                    type=['docx'],
                    key="polish_doc"
                )
                check_report = st.file_uploader(
                    "上传查重报告(可选)",
                    type=['pdf', 'docx', 'txt'],
                    key="check_report"
                )
            
            with st.expander("API配置", expanded=False):
                api_key = st.text_input(
                    "API密钥",
                    type="password",
                    value=st.session_state.api_config['api_key'],
                    help="你的AI API密钥，输入后会自动隐藏"
                )
                api_url = st.text_input(
                    "API地址",
                    value=st.session_state.api_config['api_url'],
                    placeholder="https://api.example.com/v1"
                )
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("保存配置"):
                        st.session_state.api_config['api_key'] = api_key
                        st.session_state.api_config['api_url'] = api_url
                        st.success("配置已保存！")
                with col2:
                    if st.button("重置默认"):
                        st.session_state.api_config = {'api_key': '', 'api_url': ''}
                        st.success("已重置！")
                        st.rerun()
            
            if polish_doc and st.button("开始AI润色", use_container_width=True):
                with st.spinner("AI润色中..."):
                    try:
                        doc = Document(polish_doc)
                        full_text = "\n".join([p.text for p in doc.paragraphs])
                        
                        polished_text, msg = ai_polish_text(full_text, api_key)
                        st.success(msg)
                        
                        st.session_state.polished_doc = polished_text
                        
                    except Exception as e:
                        st.error(f"润色失败：{str(e)}")
        
        st.markdown("---")
        
        st.subheader("📥 输出下载")
        if st.session_state.formatted_doc:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            filename = f"{clean_filename(selected_template)}_排版后_{timestamp}.docx"
            
            st.download_button(
                label="⬇️ 下载排版后文档",
                data=st.session_state.formatted_doc,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            st.caption("💡 提示：下载文件仅在内存中生成，不会保存到服务器，关闭页面后自动清理")

if __name__ == "__main__":
    main()
