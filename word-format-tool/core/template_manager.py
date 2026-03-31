import copy
from config.constants import TEMPLATE_LIBRARY

def validate_template(template):
    """验证模板格式是否正确"""
    required_levels = ["一级标题", "二级标题", "三级标题", "正文", "表格"]
    required_properties = ["font", "size", "bold", "align", "line_type", "line_value"]
    
    for level in required_levels:
        if level not in template:
            return False, f"模板缺少 {level} 定义"
        for prop in required_properties:
            if prop not in template[level]:
                return False, f"{level} 缺少 {prop} 属性"
    
    return True, "模板格式正确"

def apply_template_to_config(template_name, keep_custom=False, current_config=None):
    """应用模板到配置"""
    if template_name not in TEMPLATE_LIBRARY:
        raise ValueError(f"模板 {template_name} 不存在")
    
    template = TEMPLATE_LIBRARY[template_name]
    valid, msg = validate_template(template)
    if not valid:
        raise ValueError(msg)
    
    if keep_custom and current_config is not None:
        new_config = copy.deepcopy(current_config)
        for level in template.keys():
            if level not in new_config:
                new_config[level] = copy.deepcopy(template[level])
            else:
                for key in template[level].keys():
                    if key not in new_config[level]:
                        new_config[level][key] = template[level][key]
        return new_config
    else:
        return copy.deepcopy(template)

def recommend_template(doc):
    """根据文档内容智能推荐模板"""
    from config.constants import DOC_TYPE_RULES
    full_text = "\n".join([para.text for para in doc.paragraphs]).lower()
    match_score = {}
    for template_name, keywords in DOC_TYPE_RULES.items():
        score = 0
        for keyword in keywords:
            if keyword.lower() in full_text:
                score +=1
        match_score[template_name] = score
    best_template = max(match_score, key=match_score.get)
    return best_template, match_score[best_template]