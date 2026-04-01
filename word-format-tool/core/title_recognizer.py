def get_title_level(para, enable_title_regex, last_levels):
    text = para.text.strip()
    if not text:
        return "正文"
    
    if text.startswith(("一、","第一章","1、","第1章")):
        return "一级标题"
    elif text.startswith(（"（一）","1.1","二、")):
        return "二级标题"
    elif text.startswith(（"1.1.1","（1）","三、")):
        return "三级标题"
    return "正文"
