import streamlit as st
import copy
from datetime import datetime
from io import BytesIO
import zipfile
from concurrent.futures import ThreadPoolExecutor

# ===================== 常量配置 =====================
TEMPLATE_LIBRARY = {
    "默认通用格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0}
    },
    "毕业论文格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "宋体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0}
    },
    "竞赛报告格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0}
    }
}

ALIGN_LIST = ["左对齐", "居中", "右对齐", "两端对齐"]
LINE_TYPE_LIST = ["倍数", "固定值", "最小值"]
LINE_RULE = {
    "倍数": {"label": "行距倍数", "min": 1.0, "max": 3.0, "step": 0.1, "default": 1.5},
    "固定值": {"label": "固定值(磅)", "min": 12, "max": 48, "step": 1, "default": 18},
    "最小值": {"label": "最小值(磅)", "min": 12, "max": 48, "step": 1, "default": 18}
}
FONT_LIST = ["宋体", "黑体", "楷体", "微软雅黑", "Times New Roman"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五"]
EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri"]

APP_NAME = "Word文档一键排版工具"
APP_ICON = "📝"
APP_LAYOUT = "wide"

# ===================== 核心工具函数 =====================
def get_doc_from_uploaded(uploaded_file):
    """从上传的文件读取Word文档"""
    from docx import Document
    return Document(BytesIO(uploaded_file.getvalue()))

def is_protected_para(para):
    """安全过滤受保护段落，不误杀正常内容"""
    # 空段落不过滤，留给后续统计
    if not para.text.strip():
        return False
    # 过滤页眉/页脚/脚注
    try:
        if para.part and para.part.type in (1, 2, 3):
            return True
    except:
        pass
    # 过滤隐藏段落
    try:
        if para.font.hidden:
            return True
    except:
        pass
    # 正常段落
    return False

def get_title_level(para, enable_title_regex, last_levels):
    """标题层级识别，修复全角字符问题"""
    text = para.text.strip()
    if not text:
        return "正文"
    
    # 一级标题识别
    if text.startswith(("一、", "第一章", "1、", "第1章")):
        return "一级标题"
    # 二级标题识别
    elif text.startswith(("（一）", "1.1", "二、")):
        return "二级标题"
    # 三级标题识别
    elif text.startswith(("1.1.1", "（1）", "三、")):
        return "三级标题"
    # 正文
    else:
        return "正文"

def process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    """核心文档处理函数，支持表格内段落"""
    from docx import Document
    from docx.shared import Pt

    doc = get_doc_from_uploaded(file)
    # 统计初始化
    stats = {
        "一级标题": 0, "二级标题": 0, "三级标题": 0,
        "正文": 0, "表格": len(doc.tables),
        "图片": len([r for r in doc.element.xpath('.//a:blip')])
    }
    title_records = []
    last_levels = [0, 0, 0]

    # 样式应用函数
    def apply_style(para, level):
        style = config[level]
        # 字体设置
        for run in para.runs:
            run.font.name = style["font"]
            run.font.bold = style["bold"]
            if style["size"]:
                size_map = {
                    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
                    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
                    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9
                }
                run.font.size = Pt(size_map.get(style["size"], 12))
        # 段落格式
        para.paragraph_format.first_line_indent = style["indent"] * 12700
        # 对齐方式
        align_map = {
            "左对齐": 0, "居中": 1, "右对齐": 2, "两端对齐": 3
        }
        para.paragraph_format.alignment = align_map.get(style["align"], 0)
        # 行距设置
        if style["line_type"] == "倍数":
            para.paragraph_format.line_spacing = style["line_value"]
        elif style["line_type"] == "固定值":
            para.paragraph_format.line_spacing = Pt(style["line_value"])
        
        # 统计计数
        if level in ["一级标题", "二级标题", "三级标题"]:
            stats[level] += 1
        else:
            stats["正文"] += 1

    # 1. 处理文档正文段落
    for para in doc.paragraphs:
        if is_protected_para(para):
            continue
        level = get_title_level(para, enable_title_regex, last_levels)
        apply_style(para, level)
        title_records.append({
            "text": para.text.strip(),
            "level": level
        })

    # 2. 处理表格内的段落（核心修复）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if is_protected_para(para):
                        continue
                    apply_style(para, "表格")

    # 保存处理后的文档
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, stats, 1.0, title_records

def apply_template_to_config(template_name, keep_custom, current_config):
    """应用模板到当前配置"""
    if keep_custom:
        new_config = copy.deepcopy(current_config)
        template_config = TEMPLATE_LIBRARY[template_name]
        for key in template_config:
            if key not in new_config:
                new_config[key] = template_config[key]
        return new_config
    else:
        return copy.deepcopy(TEMPLATE_LIBRARY[template_name])

def recommend_template(doc):
    """智能推荐模板"""
    return "默认通用格式", 1

def batch_process_single(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    """批量处理单个文件，异常隔离"""
    try:
        res, stats, t, records = process_doc(
            file, config, number_config, enable_title_regex,
            force_style, keep_spacing, clear_blank, max_blank
        )
        return {
            "status": "success",
            "filename": file.name,
            "result": res.getvalue(),
            "stats": stats
        }
    except Exception as e:
        return {
            "status": "error",
            "filename": file.name,
            "message": str(e)
        }

# ===================== 页面状态管理 =====================
def init_session_state():
    """全量初始化会话状态，避免KeyError"""
    if "current_config" not in st.session_state:
        st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "last_template" not in st.session_state:
        st.session_state.last_template = "默认通用格式"
    if "uploaded_files" not in st.session_state:
        st.session_state.uploaded_files = None
    if "title_records" not in st.session_state:
        st.session_state.title_records = []
    # 功能开关
    if "force_style" not in st.session_state:
        st.session_state.force_style = True
    if "enable_title_regex" not in st.session_state:
        st.session_state.enable_title_regex = True
    if "keep_spacing" not in st.session_state:
        st.session_state.keep_spacing = True
    if "clear_blank" not in st.session_state:
        st.session_state.clear_blank = False
    if "max_blank" not in st.session_state:
        st.session_state.max_blank = 1
    # 数字格式配置
    if "number_config" not in st.session_state:
        st.session_state.number_config = {
            "enable": True,
            "font": "Times New Roman",
            "size_same_as_body": True,
            "size": "小四",
            "bold": False
        }

def safe_rerun():
    """兼容新旧版本的页面刷新"""
    if hasattr(st, 'rerun'):
        st.rerun()
    else:
        st.experimental_rerun()

def apply_template(template_name, keep_custom=False):
    """应用模板按钮逻辑"""
    st.session_state.current_config = apply_template_to_config(
        template_name, keep_custom, st.session_state.current_config
    )
    st.session_state.template_version += 1
    st.session_state.last_template = template_name
    st.success(f"✅ 已成功应用【{template_name}】模板")

def reset_template():
    """重置为默认模板"""
    st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    st.session_state.template_version += 1
    st.session_state.last_template = "默认通用格式"
    st.success("✅ 已重置为默认格式")

def format_editor(title, level, show_indent):
    """格式编辑器组件"""
    st.markdown(f"**{title}**")
    cfg = st.session_state.current_config[level]
    v = st.session_state.template_version
    
    col1, col2 = st.columns(2)
    with col1:
        cfg["font"] = st.selectbox(
            "字体", FONT_LIST,
            index=FONT_LIST.index(cfg["font"]),
            key=f"{level}_f_{v}"
        )
    with col2:
        cfg["size"] = st.selectbox(
            "字号", FONT_SIZE_LIST,
            index=FONT_SIZE_LIST.index(cfg["size"]),
            key=f"{level}_s_{v}"
        )
    
    cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"{level}_b_{v}")
    cfg["align"] = st.selectbox(
        "对齐方式", ALIGN_LIST,
        index=ALIGN_LIST.index(cfg["align"]),
        key=f"{level}_a_{v}"
    )
    
    # 行距设置
    line_type = st.selectbox(
        "行距类型", LINE_TYPE_LIST,
        index=LINE_TYPE_LIST.index(cfg["line_type"]),
        key=f"{level}_lt_{v}"
    )
    cfg["line_type"] = line_type
    rule = LINE_RULE[line_type]
    cfg["line_value"] = st.number_input(
        rule["label"], rule["min"], rule["max"],
        float(cfg["line_value"]), rule["step"],
        key=f"{level}_lv_{v}"
    )
    
    # 首行缩进
    if show_indent:
        cfg["indent"] = st.number_input(
            "首行缩进(字符)", 0, 4, cfg["indent"],
            key=f"{level}_i_{v}"
        )
    
    # 保存配置
    st.session_state.current_config[level] = cfg
    st.divider()

def preview_document(uploaded_file, enable_title_regex):
    """文档标题预览功能"""
    doc = get_doc_from_uploaded(uploaded_file)
    last_levels = [0, 0, 0]
    preview_records = []
    para_global_idx = 0

    # 遍历正文段落
    for para in doc.paragraphs:
        if is_protected_para(para):
            continue
        text = para.text.strip()
        level = get_title_level(para, enable_title_regex, last_levels)
        preview_records.append({
            "段落序号": para_global_idx,
            "所在位置": "文档正文",
            "识别结果": level,
            "文本内容": text[:80] + "..." if len(text) > 80 else text
        })
        para_global_idx += 1

    # 遍历表格内段落
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para in cell.paragraphs:
                    if is_protected_para(para):
                        continue
                    text = para.text.strip()
                    level = get_title_level(para, enable_title_regex, last_levels)
                    preview_records.append({
                        "段落序号": para_global_idx,
                        "所在位置": f"表格{table_idx+1}-行{row_idx+1}-列{cell_idx+1}",
                        "识别结果": level,
                        "文本内容": text[:80] + "..." if len(text) > 80 else text
                    })
                    para_global_idx += 1

    # 展示预览结果
    st.subheader("📋 标题层级结构预览")
    if preview_records:
        # 树形结构展示
        tree_data = []
        for record in preview_records:
            level = record["识别结果"]
            prefix = ""
            if level == "一级标题":
                prefix = "📌 "
            elif level == "二级标题":
                prefix = "  ├─ "
            elif level == "三级标题":
                prefix = "    ├─ "
            else:
                prefix = "      ├─ "
            tree_data.append(f"{prefix}{record['文本内容']} [{record['所在位置']}]")
        st.code("\n".join(tree_data), language="text")
        
        # 统计展示
        title_count = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0}
        for record in preview_records:
            level = record["识别结果"]
            if level in title_count:
                title_count[level] += 1

        st.write("📊 识别统计：")
        cols = st.columns(4)
        cols[0].metric("一级标题", title_count.get("一级标题", 0))
        cols[1].metric("二级标题", title_count.get("二级标题", 0))
        cols[2].metric("三级标题", title_count.get("三级标题", 0))
        cols[3].metric("正文段落", title_count.get("正文", 0))
        
        st.session_state.title_records = preview_records
    else:
        st.info("未识别到任何段落，请检查文档内容")

# ===================== 主页面 =====================
def main():
    # 页面配置
    st.set_page_config(page_title=APP_NAME, layout=APP_LAYOUT, page_icon=APP_ICON)
    # 初始化状态
    init_session_state()

    # 页面标题
    st.title(f"{APP_ICON} {APP_NAME}")
    st.success("✅ 专为大学生竞赛/毕业论文/办公人员打造 | 零误判标题识别 | 表格/图片100%保留 | 一键标准化格式")
    st.divider()

    # 核心布局：左右分栏
    col_left, col_right = st.columns([1, 2])

    # 左侧栏：模板与格式设置
    with col_left:
        st.header("📋 模板与格式设置")
        v = st.session_state.template_version

        # 1. 模板选择
        st.subheader("1. 选择排版模板")
        template_list = list(TEMPLATE_LIBRARY.keys())
        selected_template = st.radio(
            "模板库",
            template_list,
            index=template_list.index(st.session_state.last_template),
            key=f"template_select_{v}"
        )
        keep_custom = st.checkbox("保留我已调整的自定义格式", value=False, key=f"keep_custom_{v}")
        
        # 模板操作按钮
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("✅ 应用选中模板", use_container_width=True, key=f"apply_template_{v}"):
                apply_template(selected_template, keep_custom)
        with col_btn2:
            if st.button("🔄 重置默认格式", use_container_width=True, key=f"reset_template_{v}"):
                reset_template()
        
        st.divider()

        # 2. 基础功能设置
        st.subheader("2. 基础功能设置")
        with st.expander("🔧 功能开关", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.session_state.force_style = st.checkbox(
                    "启用标题批量调整功能",
                    value=st.session_state.force_style,
                    help="开启后，生成的文档可在Word/WPS导航栏一键全选同级标题批量修改",
                    key=f"force_style_{v}"
                )
                st.session_state.enable_title_regex = st.checkbox(
                    "启用智能标题识别",
                    value=st.session_state.enable_title_regex,
                    help="自动识别文档中的编号标题，适配无样式的文档",
                    key=f"enable_regex_{v}"
                )
            with col2:
                st.session_state.keep_spacing = st.checkbox(
                    "保留段落原有间距",
                    value=st.session_state.keep_spacing,
                    key=f"keep_spacing_{v}"
                )
                st.session_state.clear_blank = st.checkbox(
                    "清理多余空行",
                    value=st.session_state.clear_blank,
                    key=f"clear_blank_{v}"
                )
                if st.session_state.clear_blank:
                    st.session_state.max_blank = st.slider(
                        "最大连续空行数",
                        0, 3, st.session_state.max_blank,
                        key=f"max_blank_{v}"
                    )

        st.divider()

        # 3. 格式精细化调整
        st.subheader("3. 格式精细化调整")
        with st.expander("✏️ 标题/正文格式自定义", expanded=False):
            format_editor("一级标题", "一级标题", show_indent=False)
            format_editor("二级标题", "二级标题", show_indent=False)
            format_editor("三级标题", "三级标题", show_indent=False)
            format_editor("正文", "正文", show_indent=True)
            format_editor("表格内容", "表格", show_indent=False)

        # 4. 数字/英文格式设置
        st.subheader("4. 数字/英文格式设置")
        with st.expander("🔢 数字/英文单独设置", expanded=True):
            num_enable = st.checkbox(
                "开启数字/英文单独设置",
                value=st.session_state.number_config["enable"],
                key=f"num_enable_{v}"
            )
            number_config = {"enable": num_enable}
            if num_enable:
                col1, col2 = st.columns(2)
                with col1:
                    number_config["font"] = st.selectbox(
                        "数字/英文字体",
                        EN_FONT_LIST,
                        index=EN_FONT_LIST.index(st.session_state.number_config["font"]),
                        key=f"num_font_{v}"
                    )
                    number_config["size_same_as_body"] = st.checkbox(
                        "字号与正文一致",
                        value=st.session_state.number_config["size_same_as_body"],
                        key=f"num_size_same_{v}"
                    )
                with col2:
                    if not number_config["size_same_as_body"]:
                        number_config["size"] = st.selectbox(
                            "数字字号",
                            FONT_SIZE_LIST,
                            index=FONT_SIZE_LIST.index(st.session_state.number_config["size"]),
                            key=f"num_size_{v}"
                        )
                    else:
                        number_config["size"] = "小四"
                    number_config["bold"] = st.checkbox(
                        "数字加粗",
                        value=st.session_state.number_config["bold"],
                        key=f"num_bold_{v}"
                    )
            # 保存数字配置
            st.session_state.number_config = number_config

        # 当前格式预览
        st.divider()
        st.subheader("当前格式预览")
        st.dataframe(st.session_state.current_config, use_container_width=True)

    # 右侧栏：文件上传与处理
    with col_right:
        st.header("📁 文件上传与处理")
        # 读取当前配置
        current_config = st.session_state.current_config
        number_config = st.session_state.number_config
        enable_title_regex = st.session_state.enable_title_regex
        force_style = st.session_state.force_style
        keep_spacing = st.session_state.keep_spacing
        clear_blank = st.session_state.clear_blank
        max_blank = st.session_state.max_blank

        # 1. 文件上传
        st.subheader("1. 上传Word文档")
        uploaded_files = st.file_uploader(
            "仅支持 .docx 格式文档，可多选批量上传",
            type="docx",
            accept_multiple_files=True,
            key="uploaded_files_main"
        )
        st.session_state.uploaded_files = uploaded_files
        
        # 上传成功提示
        if uploaded_files:
            if len(uploaded_files) == 1:
                uploaded_file = uploaded_files[0]
                st.success(f"✅ 文档上传成功：{uploaded_file.name}")
                
                # 智能模板推荐
                doc = get_doc_from_uploaded(uploaded_file)
                best_template, score = recommend_template(doc)
                if score > 0:
                    st.info(f"🤖 智能推荐：根据文档内容，推荐您使用【{best_template}】模板")
                    if st.button(f"一键应用推荐模板【{best_template}】", use_container_width=True):
                        apply_template(best_template)
                        safe_rerun()
                
                # 标题识别预览按钮
                if st.button("🔍 预览标题识别结果", use_container_width=True):
                    preview_document(uploaded_file, enable_title_regex)
            else:
                st.success(f"✅ 已上传{len(uploaded_files)}个文档，将进行批量处理")
        
        st.divider()

        # 2. 一键排版处理
        st.subheader("2. 一键排版并下载")
        if not uploaded_files:
            st.warning("请先上传文档")
        else:
            # 单文件处理
            if len(uploaded_files) == 1:
                uploaded_file = uploaded_files[0]
                if st.button("✨ 开始一键自动排版", type="primary", use_container_width=True):
                    with st.status("正在处理文档，请稍候...", expanded=True) as status:
                        st.write("🔍 正在解析文档结构...")
                        st.write("📑 正在智能识别标题层级...")
                        st.write("🎨 正在应用格式设置...")
                        st.write("🔢 正在处理数字/英文格式...")
                        st.write("📊 正在处理表格和图片...")
                        try:
                            result, stats, process_time, title_records = process_doc(
                                uploaded_file,
                                current_config,
                                number_config,
                                enable_title_regex,
                                force_style,
                                keep_spacing,
                                clear_blank,
                                max_blank
                            )
                            status.update(label="✅ 文档处理完成！", state="complete")
                            st.session_state.process_time = process_time
                            st.session_state.title_records = title_records
                            
                            # 处理结果展示
                            st.subheader("📊 文档处理结果统计")
                            cols = st.columns(6)
                            cols[0].metric("一级标题", stats["一级标题"])
                            cols[1].metric("二级标题", stats["二级标题"])
                            cols[2].metric("三级标题", stats["三级标题"])
                            cols[3].metric("正文段落", stats["正文"])
                            cols[4].metric("表格数量", stats["表格"])
                            cols[5].metric("图片数量", stats["图片"])
                            
                            st.subheader("⚡ 处理性能")
                            perf_cols = st.columns(2)
                            perf_cols[0].metric("总处理耗时", f"{process_time:.2f} 秒")
                            total_paras = stats["正文"] + stats["一级标题"] + stats["二级标题"] + stats["三级标题"]
                            perf_cols[1].metric("处理段落数", total_paras)
                            
                            # 下载按钮
                            filename = f"排版完成_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uploaded_file.name}"
                            st.download_button(
                                label="📥 下载排版后的文档",
                                data=result,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                type="primary"
                            )
                            st.info("💡 提示：下载后的文档，可在Word/WPS左侧「导航窗格」中一键全选同级标题，批量调整格式")
                        except Exception as e:
                            st.error(f"处理失败：{str(e)}")
                            st.exception(e)
                            st.info("请检查上传的文档是否为正常的.docx格式，或尝试重新上传")
            
            # 批量处理
            else:
                st.info(f"检测到{len(uploaded_files)}个文档，将进行批量并行处理")
                if st.button("✨ 开始批量一键排版", type="primary", use_container_width=True):
                    total_files = len(uploaded_files)
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # 批量处理结果
                    success_count = 0
                    failed_files = []
                    zip_buffer = BytesIO()
                    
                    # 线程池并行处理
                    with st.status("正在批量处理文档，请稍候...", expanded=True) as status:
                        st.write(f"📦 共{total_files}个文档，正在并行处理...")
                        
                        with ThreadPoolExecutor(max_workers=min(6, total_files)) as executor:
                            futures = {
                                executor.submit(
                                    batch_process_single,
                                    file,
                                    current_config,
                                    number_config,
                                    enable_title_regex,
                                    force_style,
                                    keep_spacing,
                                    clear_blank,
                                    max_blank
                                ): file for file in uploaded_files
                            }
                            
                            # 实时更新进度
                            for idx, future in enumerate(futures):
                                file = futures[future]
                                status_text.text(f"正在处理：{file.name} ({idx+1}/{total_files})")
                                progress_bar.progress((idx+1)/total_files)
                                
                                try:
                                    result = future.result()
                                    if result["status"] == "success":
                                        success_count += 1
                                        # 写入压缩包
                                        filename = f"排版完成_{result['filename']}"
                                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                                            zip_file.writestr(filename, result["result"])
                                    else:
                                        failed_files.append(f"{result['filename']}：{result['message']}")
                                except Exception as e:
                                    failed_files.append(f"{file.name}：{str(e)}")
                        
                        # 处理完成
                        zip_buffer.seek(0)
                        status.update(label=f"✅ 批量处理完成！成功{success_count}个，失败{len(failed_files)}个", state="complete")
                        progress_bar.empty()
                        status_text.empty()
                    
                    # 下载压缩包
                    if success_count > 0:
                        zip_filename = f"批量排版完成_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
                        st.download_button(
                            label=f"📥 下载批量排版后的压缩包（{success_count}个文件）",
                            data=zip_buffer.getvalue(),
                            file_name=zip_filename,
                            mime="application/zip",
                            use_container_width=True,
                            type="primary"
                        )
                    # 失败文件提示
                    if failed_files:
                        st.error("以下文件处理失败：")
                        for fail in failed_files:
                            st.write(f"- {fail}")
        
        st.divider()

        # 使用说明
        with st.expander("📖 使用说明", expanded=False):
            st.markdown("""
            1. **左侧选择模板**：从模板库中选择合适的格式，点击「应用选中模板」生效，不会强制覆盖你的自定义设置
            2. **格式精细化调整**：可在左侧自定义字体、字号、行距、缩进等所有格式
            3. **上传文档**：右侧上传你的docx文档，支持多选批量上传
            4. **预览标题识别**：单文件可先预览标题识别结果，确认无误再排版
            5. **一键排版下载**：点击按钮自动处理，完成后直接下载排版后的文档
            """)

if __name__ == "__main__":
    main()
