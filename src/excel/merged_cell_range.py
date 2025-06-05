import openpyxl
import re
from openpyxl.utils import get_column_letter, column_index_from_string

def unmerge_and_fill_with_original_format(input_filepath: str, output_filepath: str):
    """
    Unmerges cells in an Excel file, fills down the values,
    and attempts to preserve the original cell's number format,
    especially for percentage columns.

    Args:
        input_filepath (str): The path to the input .xlsx file with merged cells.
        output_filepath (str): The path to save the processed .xlsx file.
    """
    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(input_filepath)
        sheet = workbook.active

        # Get a list of merged cell ranges as strings
        merged_ranges_str = [str(mr) for mr in sheet.merged_cells]

        # Iterate through each merged range string
        for merged_range_str in merged_ranges_str:
            # Parse the cell range string to get min_col, min_row, max_col, max_row
            try:
                from openpyxl.utils.cell import range_boundaries
                min_col, min_row, max_col, max_row = range_boundaries(merged_range_str)
            except ImportError:
                # Fallback for older openpyxl versions
                col_start, row_start, col_end, row_end = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", merged_range_str).groups()

                min_col = column_index_from_string(col_start)
                min_row = int(row_start)
                max_col = column_index_from_string(col_end)
                max_row = int(row_end)

            # Get the top-left cell of the merged range
            top_left_cell = sheet.cell(row=min_row, column=min_col)
            top_left_cell_value = top_left_cell.value
            top_left_cell_format = top_left_cell.number_format # 获取源单元格的格式

            # Unmerge the cells
            sheet.unmerge_cells(merged_range_str)

            # Fill down the value and copy the number format to all previously merged cells
            for row_idx in range(min_row, max_row + 1):
                for col_idx in range(min_col, max_col + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    cell.value = top_left_cell_value
                    cell.number_format = top_left_cell_format # 将源单元格的格式复制过来

        # Save the modified workbook
        workbook.save(output_filepath)
        print(f"Successfully processed and saved to: {output_filepath}")

    except FileNotFoundError:
        print(f"Error: Input file not found at {input_filepath}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    """
    Main function to run the unmerge and fill script.
    """
    input_file = "/Users/a1/理赔文件/temp/大武乡散户损失程度情况表_副本.xlsx"  # 请替换为您的输入文件名
    output_file = "/Users/a1/理赔文件/temp/_大武乡散户损失程度情况表_副本.xlsx" # 请替换为您的输出文件名

    unmerge_and_fill_with_original_format(input_file, output_file)

if __name__ == "__main__":
    # 确保您安装了 openpyxl: pip install openpyxl
    # 如果您使用的是 poetry，请确保 poetry install 已经运行
    main()
