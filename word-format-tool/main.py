import streamlit as st
import copy
from datetime import datetime
from io import BytesIO
import zipfile
from concurrent.futures import ThreadPoolExecutor

from config.constants import (
    TEMPLATE_LIBRARY, ALIGN_LIST, LINE_TYPE_LIST, LINE_RULE,
    FONT_LIST, FONT_SIZE_LIST, EN_FONT_LIST, TEST_TITLE_CASES
)
from config.settings import APP_NAME, APP_ICON, APP_LAYOUT
from core.processor import process_doc, is_protected_para
from core.title_recognizer import get_title_level
from core.template_manager import apply_template_to_config, recommend_template
from utils.file_utils import get_doc_from_uploaded

def init_session_state():
    """初始化会话状态"""
    if "current_config" not in st.session_state:
        st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "title_records" not in st.session_state:
        st.session_state.title_records = []
    if "last_template" not in st.session_state:
        st.session_state.last_template = "默认通用格式"
    if "uploaded_files" not in st.session_state:
        st.session_state.uploaded_files = None
    if "process_time" not in st.session_state:
        st.session_state.process_time = 0

def safe_rerun():
    """兼容新旧版本的rerun"""
    if hasattr(st, 'rerun'):
        st.rerun()
    else:
        st.experimental_rerun()

def apply_template(template_name, keep_custom=False):
    """应用模板"""
    try:
        st.session_state.current_config = apply_template_to_config(
            template_name, keep_custom, st.session_state.current_config
        )
        st.session_state.template_version += 1
        st.session_state.last_template = template_name
        st.success(f"✅ 已成功应用【{template_name}】模板")
        safe_rerun()
    except Exception as e:
        st.error(f"应用模板失败：{str(e)}")

def reset_template():
    """重置为默认模板"""
    st.session_state.current_config = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])
    st.session_state.template_version += 1
    st.session_state.last_template = "默认通用格式"
    st.success("✅ 已重置为默认格式")
    safe_rerun()

def format_editor(title, level, show_indent):
    """格式编辑器组件"""
    st.markdown(f"**{title}**")
    cfg = st.session_state.current_config[level]
    v = st.session_state.template_version
    
    col1, col2 = st.columns(2)
    with col1: 
        cfg["font"] = st.selectbox("字体", FONT_LIST, index=FONT_LIST.index(cfg["font"]), key=f"{level}_f_{v}")
    with col2: 
        cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST, index=FONT_SIZE_LIST.index(cfg["size"]), key=f"{level}_s_{v}")
    
    cfg["bold"] = st.checkbox("加粗", cfg["bold"], key=f"{level}_b_{v}")
    cfg["align"] = st.selectbox("对齐方式", ALIGN_LIST, index=ALIGN_LIST.index(cfg["align"]), key=f"{level}_a_{v}")
    
    line_type = st.selectbox("行距类型", LINE_TYPE_LIST, index=LINE_TYPE_LIST.index(cfg["line_type"]), key=f"{level}_lt_{v}")
    if line_type != cfg["line_type"]:
        cfg["line_type"] = line_type
        cfg["line_value"] = LINE_RULE[line_type]["default"]
    rule = LINE_RULE[cfg["line_type"]]
    cfg["line_value"] = st.number_input(rule["label"], rule["min"], rule["max"], float(cfg["line_value"]), rule["step"], key=f"{level}_lv_{v}")
    
    if show_indent:
        cfg["indent"] = st.number_input("首行缩进(字符)", 0, 4, cfg["indent"], key=f"{level}_i_{v}")
    
    st.session_state.current_config[cfg] = cfg
    st.divider()

def preview_document(uploaded_file, enable_title_regex):
    """预览文档标题识别结果"""
    try:
        doc = get_doc_from_uploaded(uploaded_file)
        last_levels = [0,0,0]
        preview_records = []
        for para_idx, para in enumerate(doc.paragraphs):
            if is_protected_para(para):
                continue
            text = para.text.strip()
            if not text:
                continue
            level = get_title_level(para, enable_title_regex, last_levels)
            preview_records.append({
                "段落序号": para_idx,
                "识别结果": level,
                "文本内容": text[:80] + "..." if len(text) > 80 else text
            })
            if level == "一级标题":
                last_levels = [last_levels[0]+1,0,0]
            elif level == "二级标题":
                last_levels[1] +=1
            elif level == "三级标题":
                last_levels[2] +=1
        
        st.subheader("📋 标题层级结构预览")
        if preview_records:
            tree_data = []
            for record in preview_records:
                level = record["识别结果"]
                if level == "一级标题":
                    tree_data.append(f"📌 {record['文本内容']}")
                elif level == "二级标题":
                    tree_data.append(f"  ├─ {record['文本内容']}")
                elif level == "三级标题":
                    tree_data.append(f"    ├─ {record['文本内容']}")
            st.code("\n".join(tree_data), language="text")
            
            # 修复：正确缩进，无语法错误的标题统计
            title_count = {"一级标题": 0, "二级标题": 0, "三级标题": 0, "正文": 0}
            for record in preview_records:
                level = record["识别结果"]
                if level in title_count:
                    title_count[level] += 1

            st.write("📊 识别统计：")
            cols = st.columns(3)
            cols[0].metric("一级标题", title_count.get("一级标题", 0))
            cols[1].metric("二级标题", title_count.get("二级标题", 0))
            cols[2].metric("三级标题", title_count.get("三级标题", 0))
        else:
            st.info("未识别到标题")
    except Exception as e:
        st.error(f"预览失败：{str(e)}")

def show_process_result(result, stats, process_time, original_filename):
    """展示处理结果"""
    st.balloons()
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
    
    filename = f"排版完成_{datetime.now().strftime('%Y%m%d%H%M%S')}_{original_filename}"
    st.download_button(
        label="📥 下载排版后的文档", 
        data=result, 
        file_name=filename, 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    st.info("💡 提示：下载后的文档，可在Word/WPS左侧「导航窗格」中一键全选同级标题，批量调整格式")

def batch_process_single(file, config, number_config, enable_title_regex, force_style, keep_spacing, clear_blank, max_blank):
    """批量处理单个文件的包装函数"""
    try:
        result, stats, process_time, _ = process_doc(
            file, config, number_config, enable_title_regex, 
            force_style, keep_spacing, clear_blank, max_blank
        )
        return {
            "status": "success",
            "filename": file.name,
            "result": result,
            "stats": stats
        }
    except Exception as e:
        return {
            "status": "error",
            "filename": file.name,
            "message": str(e)
        }

def main():
    st.set_page_config(page_title=APP_NAME, layout=APP_LAYOUT, page_icon=APP_ICON)
    init_session_state()

    st.title(f"{APP_ICON} {APP_NAME}")
    st.success("✅ 专为大学生竞赛/毕业论文/办公人员打造 | 零误判标题识别 | 图片100%完整保留 | 一键标准化格式")

    step_tabs = st.tabs(["📋 Step1 选择排版模板", "⚙️ Step2 自定义格式设置", "📂 Step3 上传并预览文档", "✨ Step4 排版并下载"])
    
    with step_tabs[0]:
        st.subheader("选择适合你的排版模板")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("### 🎓 学生专用模板")
            if st.button("河北科技大学-本科毕业论文", use_container_width=True):
                apply_template("河北科技大学-本科毕业论文")
            if st.button("大学生竞赛报告通用模板", use_container_width=True):
                apply_template("大学生竞赛报告通用模板")
            if st.button("国标-本科毕业论文通用", use_container_width=True):
                apply_template("国标-本科毕业论文通用")
        
        with col2:
            st.markdown("### 💼 办公专用模板")
            if st.button("企业办公通用报告模板", use_container_width=True):
                apply_template("企业办公通用报告模板")
            if st.button("党政机关公文国标GB/T 9704-2012", use_container_width=True):
                apply_template("党政机关公文国标GB/T 9704-2012")
        
        with col3:
            st.markdown("### 🧩 基础模板")
            if st.button("默认通用格式", use_container_width=True):
                apply_template("默认通用格式")
            keep_custom = st.checkbox("保留我已调整的自定义格式", value=False)
            if st.button("🔄 重置为默认格式", use_container_width=True):
                reset_template()
        
        st.subheader("当前模板格式预览")
        st.dataframe(st.session_state.current_config, use_container_width=True)

    with step_tabs[1]:
        st.subheader("精细化格式自定义（可选，不修改则使用模板默认值）")
        cfg = st.session_state.current_config
        v = st.session_state.template_version

        with st.expander("🔧 基础功能设置", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                force_style = st.checkbox("启用标题批量调整功能", value=True, help="开启后，生成的文档可在Word/WPS导航栏一键全选同级标题批量修改", key=f"force_style_{v}")
                enable_title_regex = st.checkbox("启用智能标题识别", value=True, help="自动识别文档中的编号标题，适配无样式的文档", key=f"enable_regex_{v}")
            with col2:
                keep_spacing = st.checkbox("保留段落原有间距", value=True, key=f"keep_spacing_{v}")
                clear_blank = st.checkbox("清理多余空行", value=False, key=f"clear_blank_{v}")
                max_blank = st.slider("最大连续空行数", 0, 3, 1, key=f"max_blank_{v}") if clear_blank else 1

        with st.expander("✏️ 标题/正文格式自定义", expanded=False):
            format_editor("一级标题", "一级标题", show_indent=False)
            format_editor("二级标题", "二级标题", show_indent=False)
            format_editor("三级标题", "三级标题", show_indent=False)
            format_editor("正文", "正文", show_indent=True)
            format_editor("表格内容", "表格", show_indent=False)

        with st.expander("🔢 正文数字/英文格式设置", expanded=True):
            num_enable = st.checkbox("开启数字/英文单独设置", value=True, key=f"num_enable_{v}")
            number_config = {"enable": num_enable}
            if num_enable:
                col1, col2 = st.columns(2)
                with col1:
                    number_config["font"] = st.selectbox("数字/英文字体", EN_FONT_LIST, 1, key=f"num_font_{v}")
                    number_config["size_same_as_body"] = st.checkbox("字号与正文一致", value=True, key=f"num_size_same_{v}")
                with col2:
                    number_config["size"] = st.selectbox("数字字号", FONT_SIZE_LIST, 9, key=f"num_size_{v}") if not number_config["size_same_as_body"] else "小四"
                    number_config["bold"] = st.checkbox("数字加粗", False, key=f"num_bold_{v}")

    with step_tabs[2]:
        st.subheader("上传Word文档")
        uploaded_files = st.file_uploader("仅支持 .docx 格式文档，可多选批量上传", type="docx", accept_multiple_files=True, key="uploaded_files")
        st.session_state.uploaded_files = uploaded_files
        
        if uploaded_files:
            if len(uploaded_files) == 1:
                uploaded_file = uploaded_files[0]
                st.success(f"✅ 文档上传成功：{uploaded_file.name}")
                
                doc = get_doc_from_uploaded(uploaded_file)
                best_template, score = recommend_template(doc)
                if score > 0:
                    st.info(f"🤖 智能推荐：根据文档内容，推荐您使用【{best_template}】模板")
                    if st.button(f"一键应用推荐模板【{best_template}】", use_container_width=True):
                        apply_template(best_template)
                
                if st.button("🔍 预览标题识别结果", use_container_width=True):
                    preview_document(uploaded_file, enable_title_regex)
            else:
                st.success(f"✅ 已上传{len(uploaded_files)}个文档，将进行批量处理")

    with step_tabs[3]:
        st.subheader("一键排版并下载")
        uploaded_files = st.session_state.get("uploaded_files", [])
        if not uploaded_files:
            st.warning("请先在Step3上传文档")
        else:
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
                                st.session_state.current_config, 
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
                            show_process_result(result, stats, process_time, uploaded_file.name)
                        except Exception as e:
                            st.error(f"处理失败：{str(e)}")
                            st.info("请检查上传的文档是否为正常的.docx格式，或尝试重新上传")
            else:
                st.info(f"检测到{len(uploaded_files)}个文档，将进行批量处理")
                if st.button("✨ 开始批量一键排版", type="primary", use_container_width=True):
                    with st.status("正在批量处理文档，请稍候...", expanded=True) as status:
                        st.write(f"📦 共{len(uploaded_files)}个文档，正在逐个处理...")
                        
                        zip_buffer = BytesIO()
                        success_count = 0
                        failed_files = []
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for file in uploaded_files:
                                try:
                                    result, stats, _, _ = process_doc(
                                        file, 
                                        st.session_state.current_config, 
                                        number_config,
                                        enable_title_regex, 
                                        force_style, 
                                        keep_spacing,
                                        clear_blank, 
                                        max_blank
                                    )
                                    if result:
                                        filename = f"排版完成_{file.name}"
                                        zip_file.writestr(filename, result)
                                        success_count +=1
                                except Exception as e:
                                    failed_files.append(f"{file.name}：{str(e)}")
                        
                        zip_buffer.seek(0)
                        status.update(label=f"✅ 批量处理完成！成功{success_count}个，失败{len(failed_files)}个", state="complete")
                    
                    if success_count > 0:
                        filename = f"批量排版完成_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
                        st.download_button(
                            label="📥 下载批量排版后的压缩包", 
                            data=zip_buffer.getvalue(), 
                            file_name=filename, 
                            mime="application/zip",
                            use_container_width=True
                        )
                    if failed_files:
                        st.error("以下文件处理失败：")
                        for fail in failed_files:
                            st.write(f"- {fail}")

    st.divider()
    with st.expander("📖 使用说明", expanded=False):
        st.markdown("""
        1. **Step1 选模板**：根据你的使用场景，选择对应的模板，一键套用标准格式
        2. **Step2 自定义设置**：如果需要精细化调整，可以在这里修改字体、行距、缩进等
        3. **Step3 上传文档**：上传你的docx文档，可先预览标题识别结果，确认无误再排版
        4. **Step4 排版下载**：点击一键排版，处理完成后即可下载排版后的文档
        """)

if __name__ == "__main__":
    main()
