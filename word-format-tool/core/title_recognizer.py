import re

def get_title_level(para_text, enable_title_regex=True, prev_para_text=None):
    """
    精准标题分级：彻底解决正文列表误识别、三级标题被二级吞没
    三重校验：格式匹配 + 上下文区分 + 语义过滤
    """
    text = para_text.strip()
    if not text:
        return "正文"

    # ====================== 一级标题（严格匹配，不冲突）======================
    # 匹配：第X章、1、、1、（带顿号的一级标题）
    if re.match(r'^第[一二三四五六七八九十]+章', text) or re.match(r'^\d+、', text):
        return "一级标题"
    
    # ====================== 二级标题（严格匹配，不冲突）======================
    # 匹配：（一）、1.1（带点的二级标题，排除纯数字列表）
    elif re.match(r'^（[一二三四五六七八九十]）', text) or re.match(r'^\d+\.\d+\s', text):
        return "二级标题"
    
    # ====================== 三级标题（核心修复：区分标题和正文列表）======================
    # 1. 先匹配格式：（1）、1.1.1
    elif re.match(r'^（\d+）', text) or re.match(r'^\d+\.\d+\.\d+', text):
        # 2. 上下文校验：如果上一段是正文/空行，且当前段落是长文本（>15字），判定为正文列表
        if prev_para_text and len(text) > 15:
            # 3. 语义过滤：如果开头是「电脑硬件的科普」这种描述性内容，直接判定为正文
            if re.match(r'^（\d+）[a-zA-Z\u4e00-\u9fa5]{2,}', text):
                return "正文"
        # 4. 否则才判定为三级标题（真正的章节标题）
        return "三级标题"
    
    # 所有不匹配的，全部判定为正文
    return "正文"
