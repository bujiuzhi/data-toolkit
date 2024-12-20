# utils/merge_excel.py

import os
import pandas as pd
from utils.excel_utils import adjust_column_widths  # 引入调整列宽的工具函数
from utils.string_utils import calculate_display_width  # 引入计算显示宽度的工具函数
import re
from collections import defaultdict


def merge_excel_sheets(
        input_folder,
        output_folder,
        file_prefix=None,
        sheet_prefix="",
        delete_source=True,
        source_sheet_indices=None
):
    """
    合并多个 Excel 文件中的多个 sheet 到一个或多个新的 Excel 文件中，
    根据规则命名 sheet，并自动调整列宽。合并完成后可选择删除源文件。
    支持指定多个 source_sheet_indices 以选择源文件中的多个 Sheet。
    支持按指定前缀或自动按共同前缀分组合并文件。

    :param input_folder: 输入文件夹路径，包含要合并的 Excel 文件
    :param output_folder: 输出文件夹路径，合并后的文件将保存在此目录
    :param file_prefix: 指定要合并的文件的前缀。如果为 None，则自动按共同前缀分组合并
    :param sheet_prefix: 合并后每个 sheet 名称的前缀（通常为空字符串）
    :param delete_source: 是否删除源文件，默认 True
    :param source_sheet_indices: 要合并的源文件中的 sheet 索引列表（1-based）。例如 [1, 3] 表示合并第1和第3个 Sheet。
                                 默认为 [1]，即仅合并第一个 Sheet。
    """
    if source_sheet_indices is None:
        source_sheet_indices = [1]  # 默认仅合并第一个 Sheet

    # 验证 source_sheet_indices
    if not isinstance(source_sheet_indices, list):
        raise TypeError("source_sheet_indices 应该是一个整数列表，例如 [1, 3]。")
    for idx in source_sheet_indices:
        if not isinstance(idx, int) or idx < 1:
            raise ValueError("source_sheet_indices 中的每个索引应该是大于等于1的整数。")

    # 检查输入文件夹是否存在
    if not os.path.exists(input_folder):
        raise FileNotFoundError(f"输入文件夹 '{input_folder}' 不存在。请确保该文件夹存在并包含 Excel 文件。")

    # 创建输出文件夹（如果不存在）
    os.makedirs(output_folder, exist_ok=True)

    # 获取所有 Excel 文件
    excel_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]
    total_files = len(excel_files)
    print(f"扫描到 {total_files} 个 Excel 文件需要合并。")

    if total_files == 0:
        print("没有找到需要合并的 Excel 文件。")
        return

    # 根据是否指定前缀，筛选文件或分组文件
    if file_prefix:
        # 仅合并指定前缀的文件
        selected_files = [f for f in excel_files if f.startswith(file_prefix)]
        if not selected_files:
            print(f"没有找到以前缀 '{file_prefix}' 开头的文件。")
            return
        groups = {file_prefix: selected_files}
    else:
        # 自动按共同前缀分组
        groups = group_files_by_common_prefix(excel_files)

    # 如果没有指定前缀且未能分组，则将所有文件作为一个组
    if not groups and not file_prefix:
        groups = {"合并结果": excel_files}

    # 进行每个组的合并
    for group_prefix, files in groups.items():
        print(f"\n----- 开始合并前缀 '{group_prefix}' 的 {len(files)} 个文件 -----")
        # 生成输出文件名
        output_file = generate_output_filename(output_folder, group_prefix)

        # 创建一个新的 Excel writer 对象
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 跟踪合并的工作表名称以避免重复
            existing_sheet_names = set()

            for file_index, excel_file in enumerate(files, 1):
                file_path = os.path.join(input_folder, excel_file)
                print(f"正在处理文件 {file_index}/{len(files)}: {file_path}")

                # 加载 Excel 文件中的所有 sheet
                try:
                    excel_data = pd.read_excel(file_path, sheet_name=None)  # sheet_name=None 会读取所有 sheet
                except Exception as e:
                    print(f"无法读取文件 {excel_file}: {e}")
                    continue

                sheet_names = list(excel_data.keys())
                num_sheets = len(sheet_names)

                for sheet_idx in source_sheet_indices:
                    if sheet_idx > num_sheets:
                        print(f"警告: 文件 {excel_file} 中没有第 {sheet_idx} 个 sheet，跳过该 Sheet。")
                        continue

                    selected_sheet_name = sheet_names[sheet_idx - 1]
                    df = excel_data[selected_sheet_name]

                    # 根据 sheet 的数量决定命名规则
                    if num_sheets == 1:
                        base_sheet_name = f"{sheet_prefix}{os.path.splitext(excel_file)[0]}"
                    else:
                        base_sheet_name = f"{sheet_prefix}{os.path.splitext(excel_file)[0]}_{selected_sheet_name}"

                    unique_sheet_name = get_unique_sheet_name(base_sheet_name, existing_sheet_names)
                    existing_sheet_names.add(unique_sheet_name)

                    print(f"正在将文件 '{excel_file}' 的 sheet '{selected_sheet_name}' 写入新的文件，sheet 名称为 '{unique_sheet_name}'")

                    # 将数据写入新的 Excel 文件
                    try:
                        df.to_excel(writer, sheet_name=unique_sheet_name, index=False)
                        worksheet = writer.sheets[unique_sheet_name]

                        # 调整列宽
                        adjust_column_widths(worksheet, df, columns_info=None)  # 使用列名作为依据

                    except Exception as e:
                        print(f"写入 sheet '{unique_sheet_name}' 时发生错误: {e}")
                        continue

                # 删除源文件
                if delete_source:
                    try:
                        os.remove(file_path)
                        print(f"已删除源文件: {file_path}")
                    except Exception as e:
                        print(f"无法删除文件 {file_path}: {e}")

        # 汇总报告
        print(f"已合并前缀 '{group_prefix}' 的文件到: {output_file}")
        print(f"生成的文件包含 {len(existing_sheet_names)} 个 sheet。")
        print("所有 sheet 已成功合并且列宽已调整。")
        print(f"----- 完成合并前缀 '{group_prefix}' 的文件 -----\n")


def group_files_by_common_prefix(files):
    """
    将文件按最长共同前缀分组。

    :param files: 文件名列表
    :return: 字典，键为前缀，值为对应的文件名列表
    """
    prefix_dict = defaultdict(list)
    for file in files:
        # 假设前缀以第一个下划线结束
        match = re.match(r'^[^_]+_', file)
        if match:
            prefix = match.group(0)
        else:
            prefix = '无前缀'
        prefix_dict[prefix].append(file)
    return prefix_dict


def generate_output_filename(output_folder, prefix):
    """
    根据前缀生成唯一的输出文件名。

    :param output_folder: 输出文件夹路径
    :param prefix: 前缀，用作文件名
    :return: 唯一的输出文件路径
    """
    sanitized_prefix = sanitize_filename(prefix) or "合并结果"
    output_file = os.path.join(output_folder, f"{sanitized_prefix}.xlsx")
    counter = 1
    base, ext = os.path.splitext(output_file)
    while os.path.exists(output_file):
        output_file = f"{base}_{counter}{ext}"
        counter += 1
    if counter > 1:
        print(f"警告: 文件名 '{os.path.basename(output_file)}' 已存在，使用新的文件名。")
    return output_file


def sanitize_filename(filename):
    """
    清理文件名，移除非法字符

    :param filename: 原始文件名
    :return: 清理后的文件名
    """
    sanitized = re.sub(r'[<>:"/\\|?*]', '', filename)
    return sanitized[:255]


def get_unique_sheet_name(base_name, existing_names):
    """
    生成一个唯一的 sheet 名称，避免与现有 sheet 重名。

    :param base_name: 基础 sheet 名称
    :param existing_names: 已存在的 sheet 名称集合
    :return: 唯一的 sheet 名称
    """
    if base_name not in existing_names:
        return base_name
    counter = 1
    while True:
        new_name = f"{base_name}_{counter}"
        if new_name not in existing_names:
            return new_name
        counter += 1


if __name__ == "__main__":
    # 设置输入和输出路径
    input_folder = "../input"  # 输入文件夹路径
    output_folder = "../output"  # 输出文件夹路径
    file_prefix = "教学活动"  # 设置为特定前缀，如 "prefix_", 或设置为 None 自动按共同前缀分组合并
    sheet_prefix = ""  # 设置 Sheet 名称前缀
    delete_source = False  # 是否删除源文件
    source_sheet_indices = [1, 3]  # 选择要合并的源文件中的第几个 sheet

    # 调用合并函数
    merge_excel_sheets(
        input_folder=input_folder,
        output_folder=output_folder,
        file_prefix=file_prefix,
        sheet_prefix=sheet_prefix,
        delete_source=delete_source,
        source_sheet_indices=source_sheet_indices
    )
