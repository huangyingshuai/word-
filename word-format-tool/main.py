import streamlit as st
import copy
from datetime import datetime
from io import BytesIO
import zipfile
from concurrent.futures import ThreadPoolExecutor

# 常量配置（论文/比赛标准模板）
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
    }
}
ALIGN_LIST = ["左对齐", "居中", "右对齐", "两端对齐"]
LINE_TYPE_LIST = ["倍数", "固定值", "最小值"]
LINE_RULE = {"倍数": {"label": "行距倍数", "min": 1.0, "max": 3.0, "step": 0.1, "default": 1.5}, "固定值": {"label": "固定值(磅)", "min": 12, "max": 48, "step": 1, "default": 18}, "最小值": {"label": "最小值(磅)", "min": 12, "max": 48, "step": 1, "default": 18}}
FONT_LIST = ["宋体", "黑体", "楷体", "微软雅黑", "Times New Roman"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五"]
EN_FONT_LIST = ["Times New Roman", "Arial", "Calibri"]
APP_NAME = "Word文档一键排版工具"
APP_ICON = "📝"
APP_LAYOUT = "wide"

# 工具导入
def get_doc_from_uploaded(uploaded_file):
    from utils.file_utils import get_doc_from_uploaded
    return get_doc_from_uploaded(uploaded_file)

def get_title_level(para, enable_title_regex, last_levels, prev_para_text=None):
    from core.title_recognizer import get_title_level
    return get_title_level(para.text.strip(), enable_title_regex, prev_para_text)

def process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    from core.processor import process_doc
    return process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank)

# 初始化状态
def init_session_state():
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
        st.session_state.number_config = {"enable":True,"font":"Times New Roman","size_same_as_body":True,"size":"小四","bold":False}

def safe_rerun():
    st.rerun() if hasattr(st,'rerun') else st.experimental_rerun()

def apply_template(template_name, keep_custom=False):
    st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY[template_name])
    st.session_state.template_version +=1
    st.session_state.last_template = template_name
    st.success(f"✅ 已应用【{template_name}】")

def reset_template():
    st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    st.session_state.template_version +=1
    st.session_state.last_template = "默认通用格式"
    st.success("✅ 已重置默认格式")

def format_editor(title, level, show_indent):
    st.markdown(f"**{title}**")
    cfg = st.session_state.current_config[level]
    v = st.session_state.template_version
    col1, col2 = st.columns(2)
    with col1: cfg["font"] = st.selectbox("字体", FONT_LIST, index=FONT_LIST.index(cfg["font"]), key=f"{level}_f_{v}")
    with col2: cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"{level}_s_{v}")
    cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"{level}_b_{v}")
    cfg["align"] = st.selectbox("对齐", ALIGN_LIST, index=ALIGN_LIST.index(cfg["align"]), key=f"{level}_a_{v}")
    line_type = st.selectbox("行距类型", LINE_TYPE_LIST, index=LINE_TYPE_LIST.index(cfg["line_type"]), key=f"{level}_lt_{v}")
    cfg["line_type"] = line_type
    rule = LINE_RULE[line_type]
    cfg["line_value"] = st.number_input(rule["label"], rule["min"], rule["max"], float(cfg["line_value"]), rule["step"], key=f"{level}_lv_{v}")
    if show_indent:
        cfg["indent"] = st.number_input("首行缩进", 0,4,cfg["indent"],key=f"{level}_i_{v}")
    st.session_state.current_config[level] = cfg
    st.divider()

def preview_document(uploaded_file, enable_title_regex):
    doc = get_doc_from_uploaded(uploaded_file)
    last_levels = [0,0,0]
    preview_records = []
    idx = 0
    # 新增：记录上一段文本，用于上下文校验
    prev_text = None

    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        # 传入上一段文本，进行上下文校验
        level = get_title_level(para, enable_title_regex, last_levels, prev_text)
        preview_records.append({"序号":idx,"位置":"正文","级别":level,"文本":para.text.strip()[:80]})
        # 更新上一段文本
        prev_text = para.text.strip()
        idx +=1

    for t_idx, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if not para.text.strip():
                        continue
                    level = get_title_level(para, enable_title_regex, last_levels, prev_text)
                    preview_records.append({"序号":idx,"position":f"表格{t_idx+1}","级别":level,"文本":para.text.strip()[:80]})
                    idx +=1

    st.subheader("📋 标题识别预览")
    count = {"一级标题":0,"二级标题":0,"三级标题":0,"正文":0}
    for r in preview_records:
        if r["级别"] in count: count[r["级别"]] +=1
    cols = st.columns(4)
    cols[0].metric("一级标题", count["一级标题"])
    cols[1].metric("二级标题", count["二级标题"])
    cols[2].metric("三级标题", count["三级标题"])
    cols[3].metric("正文", count["正文"])
    st.session_state.title_records = preview_records

def batch_process_single(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    try:
        res, stats, t, _ = process_doc(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank)
        return {"status":"success","filename":file.name,"result":res.getvalue(),"stats":stats}
    except Exception as e:
        return {"status":"error","filename":file.name,"message":str(e)}

def main():
    st.set_page_config(page_title=APP_NAME, layout=APP_LAYOUT, page_icon=APP_ICON)
    init_session_state()
    st.title(f"{APP_ICON} {APP_NAME}")
    st.success("✅ 大学生竞赛专用 | 标题层级精准 | 中西文字体分离 | 表格全兼容")
    st.divider()
    left, right = st.columns([1,2])

    with left:
        st.header("📋 模板设置")
        v = st.session_state.template_version
        templates = list(TEMPLATE_LIBRARY.keys())
        selected = st.radio("选择模板", templates, index=templates.index(st.session_state.last_template), key=f"tmp_{v}")
        c1,c2 = st.columns(2)
        with c1:
            if st.button("应用模板", use_container_width=True): apply_template(selected)
        with c2:
            if st.button("重置格式", use_container_width=True): reset_template()
        st.divider()

        st.subheader("功能开关")
        st.session_state.force_style = st.checkbox("批量调整标题", True, key=f"force_{v}")
        st.session_state.enable_title_regex = st.checkbox("智能标题识别", True, key=f"regex_{v}")
        st.session_state.clear_blank = st.checkbox("清理空行", False, key=f"blank_{v}")

        st.divider()
        st.subheader("格式自定义")
        with st.expander("调整格式"):
            format_editor("一级标题","一级标题",False)
            format_editor("二级标题","二级标题",False)
            format_editor("三级标题","三级标题",False)
            format_editor("正文","正文",True)
            format_editor("表格","表格",False)

    with right:
        st.header("📁 上传与处理")
        files = st.file_uploader("上传docx（可多选）", type="docx", accept_multiple_files=True)
        if files:
            if len(files)==1:
                st.success(f"✅ {files[0].name}")
                if st.button("🔍 预览标题识别", use_container_width=True):
                    preview_document(files[0], st.session_state.enable_title_regex)
            
            if st.button("✨ 一键排版", type="primary", use_container_width=True):
                if len(files)==1:
                    with st.status("排版处理中..."):
                        out, stats, t, _ = process_doc(files[0], st.session_state.current_config, st.session_state.number_config, st.session_state.enable_title_regex, st.session_state.force_style, st.session_state.keep_spacing, st.session_state.clear_blank, st.session_state.max_blank)
                    st.subheader("📊 排版结果统计")
                    c1,c2,c3,c4,c5,c6 = st.columns(6)
                    c1.metric("一级标题", stats["一级标题"])
                    c2.metric("二级标题", stats["二级标题"])
                    c3.metric("三级标题", stats["三级标题"])
                    c4.metric("正文", stats["正文"])
                    c5.metric("表格", stats["表格"])
                    c6.metric("图片", stats["图片"])
                    st.download_button("📥 下载排版文件", out, f"排版_{files[0].name}", use_container_width=True)
                else:
                    progress = st.progress(0)
                    zip_buf = BytesIO()
                    success = 0
                    failed = []
                    with ThreadPoolExecutor(max_workers=5) as executor:
                        futures = [executor.submit(batch_process_single, f, st.session_state.current_config, st.session_state.number_config, st.session_state.enable_title_regex, st.session_state.force_style, st.session_state.keep_spacing, st.session_state.clear_blank, st.session_state.max_blank) for f in files]
                        for i,fut in enumerate(futures):
                            progress.progress((i+1)/len(files))
                            res = fut.result()
                            if res["status"]=="success":
                                success +=1
                                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as z:
                                    z.writestr(f"排版_{res['filename']}", res["result"])
                            else:
                                failed.append(res["filename"])
                    zip_buf.seek(0)
                    st.success(f"✅ 批量完成：{success}个文件成功")
                    if failed: st.error(f"失败文件：{failed}")
                    st.download_button("📥 下载全部压缩包", zip_buf, f"批量排版_{datetime.now().strftime('%Y%m%d%H%M')}.zip", use_container_width=True)

if __name__ == "__main__":
    main()
