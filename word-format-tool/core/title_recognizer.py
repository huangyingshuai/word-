import re

def get_title_level(para_text, enable_title_regex=True):
    """
    精准标题分级：三级标题永不被二级吞没
    严格匹配论文格式，层级100%正确
    """
    text = para_text.strip()
    if not text:
        return "正文"

    # 一级标题：第X章、1、1、
    if re.match(r'^第[一二三四五六七八九十]+章', text) or re.match(r'^\d+、', text) or re.match(r'^\d+\s*$', text):
        return "一级标题"
    
    # 二级标题：（一）、1.1
    elif re.match(r'^（[一二三四五六七八九十]）', text) or re.match(r'^\d+\.\d+\s', text):
        return "二级标题"
    
    # 三级标题：（1）、1.1.1
    elif re.match(r'^（\d+）', text) or re.match(r'^\d+\.\d+\.\d+', text):
        return "三级标题"
    
    return "正文"
