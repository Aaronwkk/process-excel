import os
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.cell_range import CellRange

def clean_header_string(header_str):
    """
    清理表头字符串，移除首尾空格、换行符等。
    """
    if header_str is None:
        return None
    # 转换为字符串，移除所有换行符（包括\n和\r），然后移除首尾空格
    return str(header_str).replace('\n', '').replace('\r', '').strip()

def get_merged_cell_value(sheet, row, col):
    """
    获取单元格的真实值，考虑合并单元格的情况。
    如果单元格在合并区域内，返回合并区域左上角的值。
    """
    cell = sheet.cell(row=row, column=col)

    for merged_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value

    return cell.value

def apply_excel_styles(sheet, header_rows, output_col_idx):
    """
    为Excel工作表应用指定样式。
    - 第一、二、三行合并单元格并设置字体大小24，加粗。
    - 第四行合并单元格。
    - 第六行和第七行（如果存在）合并单元格，字体加粗。
    - 所有单元格内的文字居中。
    - 所有列的列宽设置为15。
    :param sheet: openpyxl工作表对象
    :param header_rows: 表头行列表
    :param output_col_idx: 新增列的索引
    """
    max_col = sheet.max_column

    # Style 1: Merge A1:max_col_at_row_3, font size 24, bold, center
    if sheet.max_row >= 3:
        # Check if cells A1 to current max_col are already merged or contain data before merging
        if not sheet.merged_cells.ranges:
            sheet.merge_cells(start_row=1, start_column=1, end_row=3, end_column=max_col)
            top_left_cell = sheet.cell(row=1, column=1)
            top_left_cell.font = Font(size=24, bold=True)
            top_left_cell.alignment = Alignment(horizontal='center', vertical='center')
        else:
            print(f"    警告: 工作表 {sheet.title} 包含现有合并单元格，可能无法应用第一行到第三行的合并样式。")


    # Style 2: Merge row 4 (across all columns), center
    if sheet.max_row >= 4:
        sheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=max_col)
        sheet.cell(row=4, column=1).alignment = Alignment(horizontal='center', vertical='center')

    # Style 3: Merge header rows (row 6 and 7 if applicable) for the new column, bold, center
    for h_row in header_rows:
        if h_row <= sheet.max_row:
            for col_idx in range(1, max_col + 1):
                cell = sheet.cell(row=h_row, column=col_idx)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Apply centering to all cells (overwrites previous alignments if applied)
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # New Style: Set column width for all columns to 15
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        sheet.column_dimensions[col_letter].width = 15


def batch_process_excel_add_column(folder_path, insurance_area_header, compensation_factor, output_column_header, header_rows=[5, 6]):
    """
    批量处理Excel文件，新增“赔偿金额”列并根据投保面积和自定义赔偿系数计算填充数据。
    不会修改表格内的其他原有内容。
    :param folder_path: 文件夹路径
    :param insurance_area_header: 投保面积的表头名称（例如 "投保面积"）
    :param compensation_factor: 自定义的赔偿系数（浮点数）
    :param output_column_header: 赔偿金额的表头名称（例如 "赔偿金额"）
    :param header_rows: 表头可能存在的行列表（例如 [5, 6]）
    """
    processed_files = 0

    # 对输入表头进行标准化处理一次
    clean_insurance_area_header = clean_header_string(insurance_area_header)
    clean_output_column_header = clean_header_string(output_column_header)

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(folder_path, filename)

            try:
                wb = openpyxl.load_workbook(filepath)

                for sheet_name in wb.sheetnames:
                    sheet = wb[sheet_name]

                    insurance_area_col_idx = -1
                    output_col_idx = -1
                    actual_header_row_for_data_start = -1

                    max_col_on_sheet = sheet.max_column

                    for h_row in header_rows:
                        if h_row > sheet.max_row:
                            continue

                        for col_idx in range(1, max_col_on_sheet + 1):
                            # 获取单元格原始值，并进行清理
                            raw_header_value = get_merged_cell_value(sheet, h_row, col_idx)
                            cleaned_header_value = clean_header_string(raw_header_value)

                            # 查找“投保面积”列
                            if cleaned_header_value == clean_insurance_area_header and insurance_area_col_idx == -1:
                                insurance_area_col_idx = col_idx
                                actual_header_row_for_data_start = h_row

                            # 查找“赔偿金额”输出列
                            if cleaned_header_value == clean_output_column_header and output_col_idx == -1:
                                output_col_idx = col_idx
                                if actual_header_row_for_data_start == -1:
                                    actual_header_row_for_data_start = h_row

                        if insurance_area_col_idx != -1 and output_col_idx != -1:
                            break
                        if insurance_area_col_idx != -1 and output_col_idx == -1: # If only insurance area is found, but compensation amount is not, it may also be necessary to exit the current header_row loop.
                            break

                    if insurance_area_col_idx == -1:
                        print(f"警告: 文件 {filename} 工作表 {sheet_name} 在指定表头行 {header_rows} 未找到 '{insurance_area_header}' 列（考虑合并单元格和字符清理），跳过此工作表。")
                        continue

                    if actual_header_row_for_data_start == -1:
                        actual_header_row_for_data_start = max(header_rows)

                    if output_col_idx == -1:
                        output_col_idx = max_col_on_sheet + 1
                        sheet.insert_cols(output_col_idx)
                        sheet.cell(row=actual_header_row_for_data_start, column=output_col_idx, value=output_column_header)
                        print(f"文件 {filename} 工作表 {sheet_name} 已创建新列 '{output_column_header}' 在 {get_column_letter(output_col_idx)} 列。")
                    else:
                        print(f"文件 {filename} 工作表 {sheet_name} 中 '{output_column_header}' 列已存在于 {get_column_letter(output_col_idx)} 列，将覆盖原有数据。")

                    data_start_row = actual_header_row_for_data_start + 1
                    max_row = sheet.max_row
                    if max_row < data_start_row:
                        print(f"警告: 文件 {filename} 工作表 {sheet_name} 在 '{data_start_row}' 行之后没有找到数据，跳过计算。")
                        continue

                    for row_idx in range(data_start_row, max_row + 1):
                        insurance_area_value = sheet.cell(row=row_idx, column=insurance_area_col_idx).value

                        if isinstance(insurance_area_value, (int, float)):
                            compensation_amount = insurance_area_value * compensation_factor
                            sheet.cell(row=row_idx, column=output_col_idx, value=round(compensation_amount, 2))
                        else:
                            sheet.cell(row=row_idx, column=output_col_idx, value="数据错误")

                    # Apply styles after data processing
                    print(f"    正在为工作表 {sheet_name} 应用样式...")
                    apply_excel_styles(sheet, header_rows, output_col_idx) # Call the new styling function


                wb.save(filepath)
                processed_files += 1
                print(f"处理成功: {filename}")

            except Exception as e:
                print(f"处理文件 {filename} 失败: {e}")

    print(f"\n处理完成！共处理 {processed_files} 个文件")