import openpyxl
import os
import excel_handler
import datetime


"""ЗАПУСКАЙТЕ ЭТОТ ФАЙЛ"""


def copy_data_to_single_file(source_directory, destination_file):
    dest_wb = openpyxl.Workbook()
    dest_ws = dest_wb.active
    current_row = 2
    excel_handler.add_header(dest_ws)

    for i in range(1, 51):
        filename = f'invest{i}.xlsx'
        file_path = os.path.join(source_directory, filename)

        if os.path.exists(file_path):
            source_wb = openpyxl.load_workbook(file_path)
            source_ws = source_wb.active

            for row in source_ws['A2':'AI51']:
                for cell in row:
                    dest_ws.cell(row=current_row, column=cell.col_idx, value=cell.value)
                current_row += 1
            source_wb.close()
    if os.path.exists(f"{destination_file}.xlsx"):
        new_destination_file = str(destination_file) + "_" + str(datetime.date.today())
        dest_wb.save(f"{new_destination_file}.xlsx")
        dest_wb.close()
    else:
        new_destination_file = str(destination_file)
        dest_wb.save(f"{new_destination_file}.xlsx")
        dest_wb.close()


if __name__ == "__main__":
    source_directory = 'excel_files'
    destination_file = 'merged_data'

    excel_handler.main_excel(source_directory, 13)
    copy_data_to_single_file(source_directory, destination_file)