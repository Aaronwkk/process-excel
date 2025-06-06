import os
import pandas as pd
import sys # 确保导入 sys 模块
from dotenv import load_dotenv

load_dotenv()

def convert_xls_to_xlsx_mac(folder_path, output_folder):
    """
    在 Mac 上将指定文件夹及其子文件夹中的所有 .xls 文件转换为 .xlsx 格式。
    此方法使用 pandas 和 openpyxl/xlrd，不依赖于 Microsoft Excel 应用程序。

    Args:
        folder_path (str): 包含 .xls 文件的文件夹路径。
        output_folder (str): 转换后的 .xlsx 文件保存的文件夹路径。
                             如果不存在，脚本将尝试创建。
    """

    if not os.path.isdir(folder_path):
        print(f"错误：源文件夹 '{folder_path}' 不存在。")
        return False

    # 确保输出文件夹存在，如果不存在则创建
    if not os.path.exists(output_folder):
        print(f"输出文件夹 '{output_folder}' 不存在，正在创建...")
        try:
            os.makedirs(output_folder)
            print("输出文件夹创建成功。")
        except OSError as e:
            print(f"错误：无法创建输出文件夹 '{output_folder}': {e}")
            return False

    converted_count = 0
    skipped_count = 0
    error_count = 0

    print(f"开始扫描源文件夹：'{folder_path}'")
    print(f"转换后的文件将保存到：'{output_folder}'")

    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".xls"):
                xls_path = os.path.join(root, file)

                # 构建在输出文件夹中的完整 .xlsx 路径
                # 保持原始文件在源文件夹中的相对路径结构
                relative_path = os.path.relpath(xls_path, folder_path)
                xlsx_output_path = os.path.join(output_folder, os.path.splitext(relative_path)[0] + ".xlsx")

                # 确保输出文件的父目录存在
                os.makedirs(os.path.dirname(xlsx_output_path), exist_ok=True)

                if os.path.exists(xlsx_output_path):
                    print(f"跳过：'{xlsx_output_path}' 对应的 .xlsx 文件已存在于输出目录。")
                    skipped_count += 1
                    continue

                try:
                    print(f"正在转换：'{xls_path}' 到 '{xlsx_output_path}'...")
                    # 读取所有工作表
                    xls = pd.ExcelFile(xls_path)
                    writer = pd.ExcelWriter(xlsx_output_path, engine='openpyxl')

                    for sheet_name in xls.sheet_names:
                        df = xls.parse(sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False) # index=False 避免写入 DataFrame 索引

                    writer.close() # 确保关闭 ExcelWriter 来保存文件
                    print("转换成功。")
                    converted_count += 1
                except Exception as e:
                    print(f"转换 '{xls_path}' 时发生错误：{e}")
                    error_count += 1

    print("\n--- 转换摘要 ---")
    print(f"成功转换文件数：{converted_count}")
    print(f"已跳过文件数（.xlsx 已存在于输出目录）：{skipped_count}")
    print(f"转换失败文件数：{error_count}")

    return True

# --- 主函数 ---
def main():
    """
    脚本的入口点。
    根据命令行参数或当前目录执行 .xls 到 .xlsx 的转换。
    """

    # 您可以在这里修改这两个路径以适应您的需求
    # target_folder 是您要扫描的包含 .xls 文件的原始文件夹
    # output_folder 是您希望保存转换后的 .xlsx 文件的目标文件夹

    target_folder = os.getenv("CONVERT_FILE")
    output_folder = os.getenv("OUTPUT_FILE") # 存放Excel文件的目录

    print(target_folder, output_folder)

    print("\n--- 开始 .xls 到 .xlsx 转换 ---")
    success = convert_xls_to_xlsx_mac(target_folder, output_folder)

    if success:
        print("\n所有 .xls 文件转换完成（或跳过）。")
    else:
        print("\n转换过程中发生错误，请检查以上输出。")

if __name__ == "__main__":
    # 确保在运行脚本前已安装 pandas, openpyxl, xlrd
    # pip install pandas openpyxl xlrd
    main()