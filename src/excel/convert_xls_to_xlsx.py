import os
import pandas as pd
import sys # 确保导入 sys 模块

def convert_xls_to_xlsx_mac(folder_path):
    """
    在 Mac 上将指定文件夹及其子文件夹中的所有 .xls 文件转换为 .xlsx 格式。
    此方法使用 pandas 和 openpyxl/xlrd，不依赖于 Microsoft Excel 应用程序。
    
    Args:
        folder_path (str): 包含 .xls 文件的文件夹路径。
    """
    if not os.path.isdir(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在。")
        return False

    converted_count = 0
    skipped_count = 0
    error_count = 0

    print(f"开始扫描文件夹：'{folder_path}'")
    
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".xls"):
                xls_path = os.path.join(root, file)
                xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"

                if os.path.exists(xlsx_path):
                    print(f"跳过：'{xls_path}' 对应的 .xlsx 文件已存在。")
                    skipped_count += 1
                    continue

                try:
                    print(f"正在转换：'{xls_path}' 到 '{xlsx_path}'...")
                    # 读取所有工作表
                    xls = pd.ExcelFile(xls_path)
                    writer = pd.ExcelWriter(xlsx_path, engine='openpyxl')

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
    print(f"已跳过文件数（.xlsx 已存在）：{skipped_count}")
    print(f"转换失败文件数：{error_count}")
    
    return True

# --- 主函数 ---
def main():
    """
    脚本的入口点。
    根据命令行参数或当前目录执行 .xls 到 .xlsx 的转换。
    """
    if len(sys.argv) > 1:
        target_folder = sys.argv[1]
        print(f"使用命令行指定的文件夹路径：'{target_folder}'")
    else:
        target_folder = os.getcwd()
        print(f"未指定文件夹路径，将使用当前目录：'{target_folder}'")

    print("\n--- 开始 .xls 到 .xlsx 转换 ---")
    success = convert_xls_to_xlsx_mac(target_folder)
    
    if success:
        print("\n所有 .xls 文件转换完成（或跳过）。")
    else:
        print("\n转换过程中发生错误，请检查以上输出。")

if __name__ == "__main__":
    # 确保在运行脚本前已安装 pandas, openpyxl, xlrd
    # pip install pandas openpyxl xlrd
    main()
