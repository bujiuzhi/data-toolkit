# utils/string_utils.py

import re

def calculate_display_width(s):
    """
    计算字符串的显示宽度，中文字符计为2，其他字符计为1。

    :param s: 输入字符串
    :return: 显示宽度
    """
    if not s:
        return 0
    width = 0
    for char in s:
        if re.match(r'[\u4e00-\u9fff]', char):
            width += 2
        else:
            width += 1
    return width
