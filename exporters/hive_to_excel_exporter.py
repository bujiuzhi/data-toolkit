# exporter/hive_to_excel_exporter.py

import pandas as pd
import re
import os
from connectors.hive_connection import get_hive_connection  # 引入连接管理工具
from utils.excel_utils import adjust_column_widths
from utils.string_utils import calculate_display_width


class HiveToExcelExporter:
    def __init__(self, database, prefix, excel_output_dir):
        """
        初始化参数

        :param database: Hive 数据库名
        :param prefix: 表名的前缀
        :param excel_output_dir: 输出 Excel 文件的目录路径
        """
        self.database = database
        self.prefix = prefix
        self.excel_output_dir = excel_output_dir
        self.existing_filenames = {}  # 用于跟踪文件名以处理重复

    def get_tables_by_prefix(self, connection):
        """
        获取符合前缀的所有表名

        :param connection: Hive 连接对象
        :return: 表名列表
        """
        cursor = connection.cursor()
        cursor.execute(f"USE {self.database}")  # 切换到指定数据库
        cursor.execute(f"SHOW TABLES LIKE '{self.prefix}*'")  # 使用 LIKE 来匹配前缀
        tables = cursor.fetchall()

        print("----- 获取符合前缀的表名 -----")
        table_count = len(tables)
        print(f"找到 {table_count} 张符合前缀 '{self.prefix}' 的表。")
        if not tables:
            print(f"没有找到符合前缀 '{self.prefix}' 的表。")
        else:
            table_names = [table[0] for table in tables]
            print("找到的表名列表:")
            for table in table_names:
                print(f"  - {table}")

        print("----- 结束获取表名 -----\n")
        return [table[0] for table in tables]

    def get_table_comment(self, table_name, connection):
        """
        获取表的中文注释，使用 SHOW TBLPROPERTIES

        :param table_name: 表名
        :param connection: Hive 连接对象
        :return: 表的中文注释
        """
        cursor = connection.cursor()
        # 使用 SHOW TBLPROPERTIES 获取注释
        try:
            cursor.execute(f"SHOW TBLPROPERTIES {table_name}('comment')")
            result = cursor.fetchall()

            # 假设第一个元素就是所需的注释
            if result and len(result) > 0 and result[0][0]:
                table_comment = result[0][0].strip()
                return table_comment
            else:
                print(f"表 '{table_name}' 没有通过 SHOW TBLPROPERTIES 获取注释，使用表名作为注释。")
                return table_name
        except Exception as e:
            print(f"使用 SHOW TBLPROPERTIES 获取表 '{table_name}' 注释时发生错误: {e}")
            print(f"使用表名作为 '{table_name}' 的注释。")
            return table_name  # 出现错误时，使用表名作为默认

    def get_table_description(self, table_name, connection):
        """
        获取表的列信息和中文注释

        :param table_name: 表名
        :param connection: Hive 连接对象
        :return: 列信息列表，每个元素为 (列名, 列类型, 列注释)
        """
        cursor = connection.cursor()
        cursor.execute(f"DESCRIBE {table_name}")  # 使用 DESCRIBE 获取表结构

        columns = []
        for row in cursor.fetchall():
            column_name = row[0].strip() if row[0] else ''
            column_type = row[1].strip() if row[1] else ''
            column_comment = row[2].strip() if len(row) > 2 and row[2] else ""  # 获取列的注释
            columns.append((column_name, column_type, column_comment))

        return columns

    def fetch_table_data(self, table_name, columns, connection):
        """
        从 Hive 中获取数据

        :param table_name: 表名
        :param columns: 表的列信息
        :param connection: Hive 连接对象
        :return: pandas DataFrame
        """
        cursor = connection.cursor()
        cursor.execute(f"SELECT * FROM {table_name}")  # 获取表中的所有数据
        data = cursor.fetchall()

        # 使用中文注释作为列名，如果没有注释则使用原列名
        column_names = [col[2] if col[2] else col[0] for col in columns]
        df = pd.DataFrame(data, columns=column_names)
        return df

    def sanitize_filename(self, filename):
        """
        清理文件名，移除非法字符

        :param filename: 原始文件名
        :return: 清理后的文件名
        """
        # 移除非法字符
        sanitized = re.sub(r'[<>:"/\\|?*]', '', filename)
        # 截断到适合文件系统的长度（例如，255字符）
        return sanitized[:255]

    def get_unique_filename(self, base_filename):
        """
        获取唯一的文件名，如果文件名已存在，则在末尾追加序号。

        :param base_filename: 基础文件名（不包含扩展名）
        :return: 唯一的文件名（包含扩展名）
        """
        if base_filename not in self.existing_filenames:
            self.existing_filenames[base_filename] = 1
            return base_filename + ".xlsx"
        else:
            self.existing_filenames[base_filename] += 1
            unique_filename = f"{base_filename}_{self.existing_filenames[base_filename]}.xlsx"
            print(f"警告: 文件名 '{unique_filename}' 已存在，使用新的文件名。")
            return unique_filename

    def write_to_excel(self, df, table_comment, table_name, columns, table_index, total_tables):
        """
        将单个 DataFrame 写入 Excel 文件，并自动调整列宽，同时在 Sheet2 中写入表的详细信息

        :param df: pandas DataFrame
        :param table_comment: 表的中文注释，用作文件名
        :param table_name: 表名，用作辅助信息
        :param columns: 表的列信息，用于获取列注释
        :param table_index: 当前处理的表序号
        :param total_tables: 符合条件的总表数量
        """
        if not os.path.exists(self.excel_output_dir):
            os.makedirs(self.excel_output_dir)  # 如果目录不存在，创建目录

        valid_filename = self.sanitize_filename(table_comment) or "输出"
        unique_filename = self.get_unique_filename(valid_filename)
        file_path = os.path.join(self.excel_output_dir, unique_filename)

        # 检查文件是否存在（实际上通过 get_unique_filename 已经处理）
        if os.path.exists(file_path):
            print(f"警告: 文件 '{unique_filename}' 已存在，无法覆盖。")
            # 这里已通过 get_unique_filename 生成唯一文件名，通常不会到达这里
            # 但保留作为安全措施
        else:
            # 使用 openpyxl 引擎写入 Excel 文件
            try:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # 写入数据到 Sheet1
                    df.to_excel(writer, sheet_name="数据", index=False)

                    # 获取工作簿和工作表对象
                    worksheet_data = writer.sheets["数据"]

                    # 使用实用工具模块调整列宽
                    adjust_column_widths(worksheet_data, df, columns_info=None)

                    # 准备表的详细信息
                    table_info = [
                        ["表名", table_name],
                        ["表注释", table_comment],
                        ["数据量", len(df)]
                    ]

                    # 准备列的信息
                    column_info_headers = ["列名", "列类型", "列注释"]
                    column_info = [column_info_headers] + list(columns)

                    # 写入表信息到 Sheet2
                    df_table_info = pd.DataFrame(table_info, columns=["信息", "内容"])
                    df_table_info.to_excel(writer, sheet_name="表信息", index=False)

                    # 写入列信息到 Sheet3
                    df_column_info = pd.DataFrame(column_info[1:], columns=column_info_headers)
                    df_column_info.to_excel(writer, sheet_name="列信息", index=False)

                    # 获取工作表对象并调整列宽
                    worksheet_table_info = writer.sheets["表信息"]
                    adjust_column_widths(worksheet_table_info, df_table_info, columns_info=None)

                    worksheet_column_info = writer.sheets["列信息"]
                    adjust_column_widths(worksheet_column_info, df_column_info, columns_info=None)

                print(f"已导出表 {table_index}/{total_tables}: '{table_comment}' 到文件 '{unique_filename}'。")
            except Exception as e:
                print(f"导出表 '{table_comment}' 到文件 '{unique_filename}' 时发生错误: {e}")
                raise

    def export(self):
        """
        从 Hive 导出符合前缀的所有表数据并写入单独的 Excel 文件
        """
        with get_hive_connection() as connection:
            try:
                print("----- 开始导出过程 -----\n")
                print("连接到 Hive...")
                print("----- 连接成功 -----\n")

                # 获取符合前缀的所有表名
                tables = self.get_tables_by_prefix(connection)
                total_tables = len(tables)
                if not tables:
                    raise ValueError(f"未找到符合前缀 '{self.prefix}' 的任何表。")

                failed_tables = []
                exported_files = 0

                # 遍历每个表，获取数据并导出
                for idx, table in enumerate(tables, 1):
                    try:
                        print("----- 开始处理表 -----")
                        print(f"表名: {table} ({idx}/{total_tables})")

                        # 获取表的中文注释
                        table_comment = self.get_table_comment(table, connection)

                        # 打印表名和表注释
                        print(f"表注释: {table_comment}")

                        # 获取表结构和列信息
                        columns = self.get_table_description(table, connection)

                        # 打印列信息
                        print("列信息:")
                        for col in columns:
                            print(f"  列名: {col[0]}, 列注释: {col[2]}")

                        # 获取表数据
                        df = self.fetch_table_data(table, columns, connection)

                        # 将表数据写入单独的 Excel 文件
                        self.write_to_excel(df, table_comment, table, columns, idx, total_tables)
                        exported_files += 1

                        print("----- 结束处理表 -----\n")

                    except Exception as table_e:
                        print(f"导出表 '{table}' 失败: {table_e}")
                        failed_tables.append(table)

                # 汇总报告
                print("----- 导出过程汇总 -----")
                print(f"预期生成文件数量: {total_tables}")
                print(f"实际生成文件数量: {exported_files}")
                if failed_tables:
                    print(f"以下表导出失败 ({len(failed_tables)}):")
                    for failed_table in failed_tables:
                        print(f"  - {failed_table}")
                else:
                    print("所有表均已成功导出。")
                print("----- 导出过程结束 -----")

            except Exception as e:
                print(f"导出过程中发生错误: {e}")
                raise  # 抛出异常，方便调用方处理

    def test_exporter(self):
        """
        测试方法入口，用于测试 HiveToExcelExporter 类的功能。
        """
        try:
            self.export()
        except Exception as e:
            print(f"测试导出过程中发生错误: {e}")


if __name__ == "__main__":
    # 配置导出参数
    exporter = HiveToExcelExporter(
        database="dws_hainan_hospital_info",  # 替换为你的数据库名
        prefix="dws_teaching_activity_",      # 替换为你需要的表前缀
        excel_output_dir="../input"           # 替换为你的输出 Excel 文件目录
    )

    # 执行测试导出
    exporter.test_exporter()
