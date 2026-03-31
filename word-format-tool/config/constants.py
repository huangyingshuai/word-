import re
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

# ====================== 全局配置与常量 ======================
ALIGN_MAP = {
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "不修改": None
}
ALIGN_LIST = list(ALIGN_MAP.keys())

LINE_TYPE_MAP = {
    "单倍行距": WD_LINE_SPACING.SINGLE,
    "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
    "2倍行距": WD_LINE_SPACING.DOUBLE,
    "多倍行距": WD_LINE_SPACING.MULTIPLE,
    "固定值": WD_LINE_SPACING.EXACTLY
}
LINE_TYPE_LIST = list(LINE_TYPE_MAP.keys())

LINE_RULE = {
    "单倍行距": {"default": 1.0, "min": 1.0, "max": 1.0, "step": 1.0, "label": "行距倍数"},
    "1.5倍行距": {"default": 1.5, "min": 1.5, "max": 1.5, "step": 0.1, "label": "行距倍数"},
    "2倍行距": {"default": 2.0, "min": 2.0, "max": 2.0, "step": 0.1, "label": "行距倍数"},
    "多倍行距": {"default": 1.5, "min": 0.5, "max": 5.0, "step": 0.1, "label": "行距倍数"},
    "固定值": {"default": 20.0, "min": 1.0, "max": 100.0, "step": 0.1, "label": "固定值(磅)"}
}

FONT_LIST = ["宋体", "黑体", "微软雅黑", "楷体", "仿宋"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六"]
FONT_SIZE_NUM = {k:v for k,v in zip(FONT_SIZE_LIST, [42.0,36.0,26.0,24.0,22.0,18.0,16.0,15.0,14.0,12.0,10.5,9.0,7.5,6.5])}
EN_FONT_LIST = ["和正文一致", "Times New Roman", "Arial", "Calibri", "宋体", "黑体", "仿宋_GB2312"]

# ====================== 标题识别规则 ======================
TITLE_BLACKLIST = [
    re.compile(r"^图\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^表\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^figure\s*[0-9]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^table\s*[0-9]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^[（(]\s*[0-9]+[)）]\s*.*[。？！；;]$"),
    re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*.*[。？！；;]$"),
    re.compile(r"^注\s*[0-9]*[：:.]\s*"),
    re.compile(r"^参考文献\s*[:：]?$"),
    re.compile(r"^附录\s*[0-9A-Z]*[:：]?$"),
]

TITLE_RULE = {
    "一级标题": [
        re.compile(r"^第[一二三四五六七八九十0-9]+章\s*[^\s。？！；;]{2,40}$"),
        re.compile(r"^[一二三四五六七八九十]+、\s*[^\s。？！；;]{2,40}$"),
    ],
    "二级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\s*[^\s。？！；;]{2,50}$"),
        re.compile(r"^[（(][一二三四五六七八九十]+[)）]\s*[^\s。？！；;]{2,50}$"),
    ],
    "三级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+\s*[^\s。？！；;]{2,60}$"),
        re.compile(r"^[（(][0-9]+[)）]\s*[^\s。？！；;]{2,60}$"),
    ]
}

TITLE_MAX_LENGTH = 60
BODY_MIN_LENGTH = 100

# 预编译正则，避免循环重复编译
NUMBER_EN_PATTERN = re.compile(r"[a-zA-Z0-9\.\-%\+]+")

# ====================== 模板库 ======================
TEMPLATE_LIBRARY = {
    "默认通用格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "河北科技大学-本科毕业论文": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "国标-本科毕业论文通用": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "党政机关公文国标GB/T 9704-2012": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 6},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "大学生竞赛报告通用模板": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "企业办公通用报告模板": {
        "一级标题": {"font": "微软雅黑", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "微软雅黑", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "微软雅黑", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 3, "space_after": 3},
        "正文": {"font": "微软雅黑", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "微软雅黑", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0}
    }
}

# ====================== 智能模板推荐规则 ======================
DOC_TYPE_RULES = {
    "河北科技大学-本科毕业论文": ["河北科技大学", "毕业论文", "摘要", "关键词", "参考文献", "致谢"],
    "大学生竞赛报告通用模板": ["竞赛", "作品", "创新点", "技术实现", "测试", "应用前景"],
    "党政机关公文国标GB/T 9704-2012": ["国务院", "通知", "公告", "函", "请示", "批复", "印发"],
    "企业办公通用报告模板": ["公司", "部门", "工作总结", "工作计划", "汇报", "业绩", "项目"],
    "默认通用格式": []
}

# ====================== 标题识别测试用例 ======================
TEST_TITLE_CASES = [
    ("第1章 绪论", "一级标题", "标准第X章一级标题"),
    ("一、研究背景与意义", "一级标题", "中文序号一级标题"),
    ("1.1 国内外研究现状", "二级标题", "标准1.1二级标题"),
    ("（一）国内研究现状", "二级标题", "中文括号二级标题"),
    ("1.1.1 技术发展历程", "三级标题", "标准1.1.1三级标题"),
    ("（1）核心技术突破", "三级标题", "数字括号三级标题"),
    ("图1 系统架构图", "正文", "图注黑名单排除"),
    ("表2 性能对比表", "正文", "表注黑名单排除"),
    ("（1）本文提出的算法有效提升了准确率。", "正文", "带句号的列表项排除"),
    ("参考文献", "正文", "参考文献黑名单排除"),
]