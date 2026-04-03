import streamlit as st
import copy
from datetime import datetime
from io import BytesIO
import zipfile
from concurrent.futures import ThreadPoolExecutor

# ====================== 国家标准毕业论文/竞赛格式模板（GB/T 7714-2015）======================
TEMPLATE_LIBRARY = {
    "国家标准毕业论文格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 24, "space_after": 18},
        "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "通用办公格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "二级标题": {"font": "黑体", "size": "小三", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "三级标题": {"font": "楷体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "倍数", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 0},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "倍数", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "倍数", "line_value": 1.25, "indent": 0, "space_before": 0, "space_after": 0}
    }
}

# ====================== 全局常量配置 ======================
ALIGN_LIST = ["左对齐", "居中", "右对齐", "两端对齐"]
LINE_TYPE_LIST = ["倍数", "固定值", "最小值"]
LINE_RULE = {
    "倍数": {"label": "行距倍数", "min": 1.0, "max": 5.0, "step": 0.1, "default": 1.5},
    "固定值": {"label": "固定值(磅)", "min": 12, "max": 100, "step": 1, "default": 20},
    "最小值": {"label": "最小值(磅)", "min": 12, "max": 100, "step": 1, "default": 20}
}
FONT_LIST = ["宋体", "黑体", "楷体", "微软雅黑", "仿宋_GB2312", "Times New Roman"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五"]
EN_FONT = "Times New Roman"  # 固定西文/数字标准字体
APP_NAME = "Word文档一键排版工具（竞赛专用版）"
APP_ICON = "📝"
APP_LAYOUT = "wide"

# ====================== 工具函数导入 ======================
def get_doc_from_uploaded(uploaded_file):
    from utils.file_utils import get_doc_from_uploaded
    return get_doc_from_uploaded(uploaded_file)

def get_title_level(para_text, enable_title_regex=True, prev_para_text=None):
    from core.title_recognizer import get_title_level
    return get_title_level(para_text.strip(), enable_title_regex, prev_para_text)

def process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    from core.processor import process_doc
    return process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank)

# ====================== 页面状态初始化 ======================
def init_session_state():
    if "current_config" not in st.session_state:
        st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["国家标准毕业论文格式"])
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "last_template" not in st.session_state:
        st.session_state.last_template = "国家标准毕业论文格式"
    if "title_records" not in st.session_state:
        st.session_state.title_records = []
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
    if "number_config" not in st.session_state:
        st.session_state.number_config = {
            "enable": True,
            "auto_number": True,
            "font": EN_FONT,
            "size_same_as_body": True,
            "size": "小四",
            "bold": False
        }

def safe_rerun():
    st.rerun() if hasattr(st, 'rerun') else st.experimental_rerun()

# ====================== 模板操作函数 ======================
def apply_template(template_name):
    st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY[template_name])
    st.session_state.template_version += 1
    st.session_state.last_template = template_name
    st.success(f"✅ 已应用【{template_name}】")

def reset_template():
    st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["国家标准毕业论文格式"])
    st.session_state.template_version += 1
    st.session_state.last_template = "国家标准毕业论文格式"
    st.success("✅ 已重置为国家标准格式")

# ====================== 格式编辑器（补全首行缩进+段前段后设置）======================
def format_editor(title, level):
    st.markdown(f"### {title}")
    cfg = st.session_state.current_config[level]
    v = st.session_state.template_version
    col1, col2 = st.columns(2)
    
    # 基础字体设置
    with col1: 
        cfg["font"] = st.selectbox("中文字体", FONT_LIST, index=FONT_LIST.index(cfg["font"]), key=f"{level}_font_{v}")
    with col2: 
        cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"{level}_size_{v}")
    
    # 格式开关
    col3, col4 = st.columns(2)
    with col3:
        cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"{level}_bold_{v}")
    with col4:
        cfg["align"] = st.selectbox("对齐方式", ALIGN_LIST, index=ALIGN_LIST.index(cfg["align"]), key=f"{level}_align_{v}")
    
    # 行距设置
    line_type = st.selectbox("行距类型", LINE_TYPE_LIST, index=LINE_TYPE_LIST.index(cfg["line_type"]), key=f"{level}_linetype_{v}")
    cfg["line_type"] = line_type
    rule = LINE_RULE[line_type]
    cfg["line_value"] = st.number_input(rule["label"], rule["min"], rule["max"], float(cfg["line_value"]), rule["step"], key=f"{level}_linevalue_{v}")
    
    # 段前段后设置
    col5, col6 = st.columns(2)
    with col5:
        cfg["space_before"] = st.number_input("段前间距(磅)", 0, 100, cfg["space_before"], 1, key=f"{level}_before_{v}")
    with col6:
        cfg["space_after"] = st.number_input("段后间距(磅)", 0, 100, cfg["space_after"], 1, key=f"{level}_after_{v}")
    
    # 首行缩进设置（全格式补全）
    cfg["indent"] = st.number_input("首行缩进(字符)", 0, 4, cfg["indent"], 1, key=f"{level}_indent_{v}")
    st.caption("正文推荐2字符，标题推荐0字符")
    
    # 保存配置
    st.session_state.current_config[level] = cfg
    st.divider()

# ====================== 标题识别预览函数 ======================
def preview_document(uploaded_file, enable_title_regex):
    doc = get_doc_from_uploaded(uploaded_file)
    preview_records = []
    idx = 0
    prev_text = None  # 上下文校验用

    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        level = get_title_level(para.text, enable_title_regex, prev_text)
        preview_records.append({
            "序号": idx,
            "位置": "正文",
            "识别级别": level,
            "文本内容": para.text.strip()[:100]
        })
        prev_text = para.text.strip()
        idx += 1

    # 表格内内容预览
    for t_idx, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if not para.text.strip():
                        continue
                    level = get_title_level(para.text, enable_title_regex, prev_text)
                    preview_records.append({
                        "序号": idx,
                        "位置": f"表格{t_idx+1}",
                        "识别级别": level,
                        "文本内容": para.text.strip()[:100]
                    })
                    idx += 1

    # 统计结果展示
    st.subheader("📋 标题识别预览结果")
    count = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0}
    for r in preview_records:
        if r["识别级别"] in count:
            count[r["识别级别"]] += 1
    
    cols = st.columns(4)
    cols[0].metric("✅ 一级标题", count["一级标题"])
    cols[1].metric("✅ 二级标题", count["二级标题"])
    cols[2].metric("✅ 三级标题", count["三级标题"])
    cols[3].metric("📄 正文段落", count["正文"])
    
    # 详细列表
    st.dataframe(preview_records, use_container_width=True)
    st.session_state.title_records = preview_records

# ====================== 批量处理函数 ======================
def batch_process_single(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    try:
        res, stats, _, _ = process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank)
        return {"status": "success", "filename": file.name, "result": res.getvalue(), "stats": stats}
    except Exception as e:
        return {"status": "error", "filename": file.name, "message": str(e)}

# ====================== 主页面渲染 ======================
def main():
    st.set_page_config(page_title=APP_NAME, layout=APP_LAYOUT, page_icon=APP_ICON)
    init_session_state()
    
    # 核心CSS：左侧格式区独立滚动，固定高度，左小右大布局
    st.markdown(
        """
        <style>
        /* 左侧滚动容器：固定高度，超出自动滚动 */
        .left-scroll-container {
            height: 85vh;
            overflow-y: auto;
            padding-right: 10px;
        }
        /* 隐藏滚动条但保留功能 */
        .left-scroll-container::-webkit-scrollbar {
            width: 6px;
        }
        .left-scroll-container::-webkit-scrollbar-thumb {
            background-color: #e0e0e0;
            border-radius: 3px;
        }
        /* 右侧内容固定，不随左侧滚动 */
        .right-fixed-container {
            height: 85vh;
            overflow-y: auto;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # 页面标题
    st.title(f"{APP_ICON} {APP_NAME}")
    st.success("✅ 挑战杯/互联网+竞赛专用 | 国家标准格式 | WPS目录自动生成 | 中西文字体分离")
    st.divider()

    # 左3右7固定占比布局（左小右大）
    left_col, right_col = st.columns([3, 7])

    # 左侧：可独立滚动的格式设置区
    with left_col:
        with st.container():
            st.markdown('<div class="left-scroll-container">', unsafe_allow_html=True)
            
            # 模板选择区
            st.header("📋 模板选择")
            v = st.session_state.template_version
            templates = list(TEMPLATE_LIBRARY.keys())
            selected = st.radio(
                "选择格式模板",
                templates,
                index=templates.index(st.session_state.last_template),
                key=f"template_select_{v}"
            )
            col1, col2 = st.columns(2)
            with col1:
                if st.button("应用模板", use_container_width=True):
                    apply_template(selected)
            with col2:
                if st.button("重置格式", use_container_width=True):
                    reset_template()
            st.divider()

            # 功能开关区
            st.header("⚙️ 功能设置")
            st.session_state.force_style = st.checkbox("强制应用标题样式", True, key=f"force_style_{v}")
            st.session_state.enable_title_regex = st.checkbox("智能标题识别", True, key=f"title_regex_{v}")
            st.session_state.number_config["auto_number"] = st.checkbox("自动生成多级序号", True, key=f"auto_number_{v}")
            st.session_state.clear_blank = st.checkbox("清理多余空行", False, key=f"clear_blank_{v}")
            if st.session_state.clear_blank:
                st.session_state.max_blank = st.number_input("最大连续空行数", 1, 5, 1, 1, key=f"max_blank_{v}")
            st.divider()

            # 格式自定义区（全格式带首行缩进）
            st.header("✏️ 格式自定义")
            with st.expander("一级标题格式", expanded=False):
                format_editor("一级标题", "一级标题")
            with st.expander("二级标题格式", expanded=False):
                format_editor("二级标题", "二级标题")
            with st.expander("三级标题格式", expanded=False):
                format_editor("三级标题", "三级标题")
            with st.expander("正文格式", expanded=True):
                format_editor("正文", "正文")
            with st.expander("表格格式", expanded=False):
                format_editor("表格", "表格")
            
            st.markdown('</div>', unsafe_allow_html=True)

    # 右侧：文件上传与处理区
    with right_col:
        st.markdown('<div class="right-fixed-container">', unsafe_allow_html=True)
        st.header("📁 文档上传与处理")
        files = st.file_uploader(
            "上传Word文档（.docx格式，支持多选批量处理）",
            type="docx",
            accept_multiple_files=True,
            key="file_uploader"
        )

        if files:
            st.success(f"✅ 已上传 {len(files)} 个文档")
            # 单文档预览
            if len(files) == 1:
                if st.button("🔍 预览标题识别结果", use_container_width=True):
                    preview_document(files[0], st.session_state.enable_title_regex)
            
            # 一键排版按钮
            if st.button("✨ 一键自动排版", type="primary", use_container_width=True):
                # 单文档处理
                if len(files) == 1:
                    with st.status("正在排版处理中...", expanded=True) as status:
                        st.write("正在读取文档...")
                        st.write("正在识别标题层级...")
                        st.write("正在应用格式规范...")
                        st.write("正在生成自动序号...")
                        st.write("正在分离中西文字体...")
                        out, stats, _, _ = process_doc(
                            files[0],
                            st.session_state.current_config,
                            st.session_state.number_config,
                            st.session_state.enable_title_regex,
                            st.session_state.force_style,
                            st.session_state.keep_spacing,
                            st.session_state.clear_blank,
                            st.session_state.max_blank
                        )
                        status.update(label="✅ 排版完成！", state="complete", expanded=False)
                    
                    # 排版结果统计
                    st.subheader("📊 排版结果统计")
                    col1, col2, col3, col4, col5, col6 = st.columns(6)
                    col1.metric("一级标题", stats["一级标题"])
                    col2.metric("二级标题", stats["二级标题"])
                    col3.metric("三级标题", stats["三级标题"])
                    col4.metric("正文段落", stats["正文"])
                    col5.metric("表格数量", stats["表格"])
                    col6.metric("图片数量", stats["图片"])
                    
                    # 下载按钮
                    st.download_button(
                        "📥 下载排版后的文档",
                        out,
                        f"排版_{files[0].name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                # 多文档批量处理
                else:
                    progress_bar = st.progress(0)
                    zip_buf = BytesIO()
                    success_count = 0
                    failed_list = []
                    total = len(files)

                    with ThreadPoolExecutor(max_workers=5) as executor:
                        futures = [
                            executor.submit(
                                batch_process_single,
                                f,
                                st.session_state.current_config,
                                st.session_state.number_config,
                                st.session_state.enable_title_regex,
                                st.session_state.force_style,
                                st.session_state.keep_spacing,
                                st.session_state.clear_blank,
                                st.session_state.max_blank
                            ) for f in files
                        ]
                        for i, future in enumerate(futures):
                            progress_bar.progress((i+1)/total)
                            result = future.result()
                            if result["status"] == "success":
                                success_count += 1
                                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_file:
                                    zip_file.writestr(f"排版_{result['filename']}", result["result"])
                            else:
                                failed_list.append(result["filename"])
                    
                    zip_buf.seek(0)
                    st.success(f"✅ 批量处理完成：成功 {success_count} 个，失败 {len(failed_list)} 个")
                    if failed_list:
                        st.error(f"失败文档：{failed_list}")
                    
                    st.download_button(
                        "📥 下载全部排版文档（压缩包）",
                        zip_buf,
                        f"批量排版_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
