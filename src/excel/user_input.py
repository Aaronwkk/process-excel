import os

def get_user_input():
    """交互式获取用户配置"""
    print("=== Excel理赔金额新增列计算工具 ===")

    default_folder_path = "/Users/a1/理赔文件"
    folder_path = input(f"请输入包含Excel文件的文件夹路径 (只支持.xlsx文件，默认为: {default_folder_path}，直接按回车使用默认值): ").strip()
    if not folder_path:
        folder_path = default_folder_path

    if not os.path.isdir(folder_path):
        print(f"错误: 路径 '{folder_path}' 不是一个有效的文件夹。请重新运行。")
        return None

    default_insurance_area_header = "投保面积"
    insurance_area_header = input(
        f"请输入“投保面积”列的表头名称（默认为: {default_insurance_area_header}，直接按回车使用默认值）: "
    ).strip()
    if not insurance_area_header:
        insurance_area_header = default_insurance_area_header

    default_compensation_factor = "23.0"
    compensation_factor_str = input(
        f"请输入自定义的“赔偿系数”（默认为: {default_compensation_factor}，直接按回车使用默认值）: "
    ).strip()
    if not compensation_factor_str:
        compensation_factor_str = default_compensation_factor

    try:
        compensation_factor = float(compensation_factor_str)
    except ValueError:
        print(f"错误: 无效的赔偿系数 '{compensation_factor_str}'。请输入一个数字。请重新运行。")
        return None

    default_output_column_header = "赔偿款"
    output_column_header = input(
        f"请输入要新增的“赔偿金额”列的表头名称（默认为: {default_output_column_header}，直接按回车使用默认值）。如果该列不存在将自动创建: "
    ).strip()
    if not output_column_header:
        output_column_header = default_output_column_header

    default_header_rows = "6,7"
    header_rows_input = input(
        f"请输入表头所在的行号（可以是一个或多个，用逗号分隔，例如: {default_header_rows}，直接按回车使用默认值）: "
    ).strip()

    if not header_rows_input:
        header_rows = [5, 6]
    else:
        try:
            header_rows = sorted(list(map(int, header_rows_input.split(','))))
            if not header_rows or any(row <= 0 for row in header_rows):
                raise ValueError
        except ValueError:
            print(f"错误: 无效的表头行输入。请用逗号分隔正整数。使用默认值 {default_header_rows}。")
            header_rows = [5, 6]

    return {
        "folder_path": folder_path,
        "insurance_area_header": insurance_area_header,
        "compensation_factor": compensation_factor,
        "output_column_header": output_column_header,
        "header_rows": header_rows
    }