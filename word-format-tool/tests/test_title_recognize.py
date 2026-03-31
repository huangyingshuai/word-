import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from core.title_recognizer import get_title_level
from config.constants import TEST_TITLE_CASES

def test_title_accuracy():
    """测试标题识别准确率"""
    total = len(TEST_TITLE_CASES)
    correct = 0
    error_cases = []
    
    doc = Document()
    for text, expect, desc in TEST_TITLE_CASES:
        para = doc.add_paragraph(text)
        result = get_title_level(para, enable_regex=True, last_levels=[1,1,0])
        if result == expect:
            correct +=1
        else:
            error_cases.append(f"文本：{text} | 预期：{expect} | 实际：{result} | 描述：{desc}")
    
    accuracy = (correct / total) * 100
    print(f"===== 标题识别测试报告 =====")
    print(f"总用例数：{total} | 正确数：{correct} | 准确率：{accuracy:.2f}%")
    if error_cases:
        print(f"错误用例：")
        for case in error_cases:
            print(f"- {case}")
    return accuracy, error_cases

if __name__ == "__main__":
    test_title_accuracy()