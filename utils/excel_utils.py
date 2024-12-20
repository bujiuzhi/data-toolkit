# utils/excel_utils.py

from openpyxl.utils import get_column_letter
from .string_utils import calculate_display_width


def adjust_column_widths(worksheet, df, columns_info=None, max_width=100):
    """
    自动调整 Excel 工作表的列宽，参考列注释和数据内容。

    :param worksheet: openpyxl 的工作表对象
    :param df: pandas DataFrame
    :param columns_info: 列信息列表，每个元素为 (列名, 列类型, 列注释) 或 None
    :param max_width: 列宽的最大限制
    """
    if columns_info:
        # 创建一个字典，映射列名（中文注释）到列注释
        column_comment_dict = {col[2] if col[2] else col[0]: (col[2] if col[2] else col[0]) for col in columns_info}
    else:
        # 如果没有 columns_info，使用列名本身作为注释
        column_comment_dict = {col: col for col in df.columns}

    for idx, column in enumerate(df.columns, 1):  # 1-based index
        column_letter = get_column_letter(idx)

        # 获取列注释
        column_comment = column_comment_dict.get(column, column)

        # 计算列注释的显示宽度
        header_width = calculate_display_width(str(column_comment))

        # 计算列数据的最大显示宽度
        if not df[column].empty:
            data_max_length = df[column].astype(str).map(calculate_display_width).max()
        else:
            data_max_length = 0

        # 取最大值并加一些额外空间
        max_display_width = max(header_width, data_max_length) + 2  # 额外空间

        # 设置列宽，限制最大宽度
        adjusted_width = min(max_display_width, max_width)
        worksheet.column_dimensions[column_letter].width = adjusted_width
