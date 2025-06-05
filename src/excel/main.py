from .user_input import get_user_input
from .file_processor import batch_process_excel_add_column
import sys

def main():
    config = get_user_input()
    if config:
        batch_process_excel_add_column(
            folder_path=config["folder_path"],
            insurance_area_header=config["insurance_area_header"],
            compensation_factor=config["compensation_factor"],
            output_column_header=config["output_column_header"],
            header_rows=config["header_rows"],
            output_path=config["output_path"]
        )
    input("\n按回车键退出...")
    sys.exit() # Ensure the script exits after user presses Enter

if __name__ == "__main__":
    main()