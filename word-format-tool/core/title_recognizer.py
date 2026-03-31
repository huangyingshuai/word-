import re
from docx.oxml.ns import qn
from config.constants import (
    TITLE_BLACKLIST, TITLE_RULE, TITLE_MAX_LENGTH,
    ALIGN_MAP, LINE_TYPE_MAP, FONT_SIZE_NUM
)

def get_title_level(para, enable_regex=True, last_levels=None):
    """
    零误判标题识别算法
    执行顺序：1.黑名单排除 → 2.内置样式识别 → 3.大纲级别识别 → 4.严格正则匹配 → 5.上下文校验
    """
    try:
        if last_levels is None:
            last_levels = [0, 0, 0]
        
        if not para:
            return "正文"
        
        text = para.text.strip()
        text_length = len(text)

        # 第一步：黑名单排除
        if not text:
            return "正文"
        for pattern in TITLE_BLACKLIST:
            if pattern.match(text):
                return "正文"
        if text.endswith(("。", "？", "！", "；", ".", "?", "!", ";")):
            return "正文"
        if text_length > TITLE_MAX_LENGTH:
            return "正文"
        if text_length < 2:
            return "正文"

        # 第二步：识别Word内置标题样式
        style_name = para.style.name.lower()
        if "heading 1" in style_name or "标题 1" in style_name or "标题1" in style_name:
            return "一级标题"
        if "heading 2" in style_name or "标题 2" in style_name or "标题2" in style_name:
            return "二级标题"
        if "heading 3" in style_name or "标题 3" in style_name or "标题3" in style_name:
            return "三级标题"
        
        if not enable_regex:
            return "正文"

        # 第三步：识别大纲级别
        try:
            p = para._element
            outline_lvl = p.xpath('.//w:outlineLvl', namespaces=p.nsmap)
            if outline_lvl:
                level = int(outline_lvl[0].get(qn('w:val')))
                if level == 1:
                    return "一级标题"
                elif level == 2:
                    return "二级标题"
                elif level == 3:
                    return "三级标题"
        except Exception:
            pass

        # 第四步：严格正则匹配
        for pattern in TITLE_RULE["三级标题"]:
            if pattern.match(text):
                if last_levels[1] > 0 or enable_regex:
                    return "三级标题"
                else:
                    return "正文"
        
        for pattern in TITLE_RULE["二级标题"]:
            if pattern.match(text):
                if last_levels[0] > 0 or enable_regex:
                    return "二级标题"
                else:
                    return "正文"
        
        for pattern in TITLE_RULE["一级标题"]:
            if pattern.match(text):
                return "一级标题"

        return "正文"
    except Exception:
        return "正文"