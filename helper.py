import string
import time
from datetime import timedelta
import pandas as pd
import xlsxwriter


def generate_letter_combinations(length: int = 3):
    if length < 1 or length > 3:
        raise ValueError("Length must be between 1 and 3.")
    else:
        for c1 in string.ascii_uppercase:
            if length == 1:
                yield c1
            else:
                for c2 in string.ascii_uppercase:
                    if length == 2:
                        yield c1 + c2
                    else:
                        for c3 in string.ascii_uppercase:
                            yield c1 + c2 + c3


def save_df_as_csv(df_to_save: pd.DataFrame, csv_name_no_filetype: str, with_header_and_rows: bool = False):
    csv_base_path = r"C:\Users\fhaum\OneDrive\401 MASTER - Masterarbeit\04 Kalkulationen\pythonProject\csv_files"
    csv_full_path = csv_base_path + "\\" + csv_name_no_filetype + ".csv"
    if with_header_and_rows:
        df_to_save.to_csv(csv_full_path, index=True, header=True)
    else:
        df_to_save.to_csv(csv_full_path, index=False, header=False)
    return f"'{csv_name_no_filetype} saved successfully to {csv_base_path}"


def read_df_from_csv(csv_name_no_filetype: str, with_header_and_rows: bool = False):
    csv_base_path = r"C:\Users\fhaum\OneDrive\401 MASTER - Masterarbeit\04 Kalkulationen\pythonProject\csv_files"
    csv_full_path = csv_base_path + "\\" + csv_name_no_filetype + ".csv"
    if with_header_and_rows:
        return_df = pd.read_csv(csv_full_path, header=0, index_col=0)
    else:
        return_df = pd.read_csv(csv_full_path, header=None, index_col=None)
    return return_df


def elapsed_time(start_time):
    end_time = time.time()
    elapsed_seconds = int(end_time - start_time)
    elapsed_time = timedelta(seconds=elapsed_seconds)
    hours = elapsed_time.seconds // 3600
    minutes = (elapsed_time.seconds % 3600) // 60
    seconds = elapsed_time.seconds % 60
    time_str = '{:02d}:{:02d}:{:02d}'.format(hours, minutes, seconds)
    print(time_str)
    return time_str


def export_map_to_excel_with_formatting_deprecated(df, wb_path, ws_name):
    """
    Export a DataFrame to an Excel file with conditional formatting using openpyxl.

    Args:
        df (pandas.DataFrame): The DataFrame to be exported.
        path (str): The file path where the Excel file will be saved.

    Returns:
        None
    """
    # Define style parameter
    colo_street = "#A6A6A6"
    colo_house_no_parcel = "#95B8D1"
    colo_house_with_parcel = "#9B1D20"  # "#522B47"
    colo_depot = "#F77F00"
    colo_empty_lot = "#566E3D"
    column_width = 3
    row_height = 12

    # Create a Pandas Excel writer
    writer = pd.ExcelWriter(wb_path, engine='xlsxwriter')

    # Write the DataFrame to the Excel file
    df.to_excel(writer, sheet_name=ws_name)

    # Get the workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets[ws_name]

    format_street = workbook.add_format({"bg_color": colo_street, "font_color": colo_street})
    format_house_no_parcel = workbook.add_format({"bg_color": colo_house_no_parcel,
                                                  "font_color": colo_house_no_parcel,
                                                  "border": 1})
    format_house_with_parcel = workbook.add_format({"bg_color": colo_house_with_parcel,
                                                    "font_color": colo_house_with_parcel,
                                                    "border": 1})
    format_empty_lot = workbook.add_format({"bg_color": colo_empty_lot, "font_color": colo_empty_lot})
    format_depot = workbook.add_format({"bg_color": colo_depot, "font_color": colo_depot, "border": 1})
    format_column = workbook.add_format({"align": "center", 'valign': "vcenter"})

    (max_row, max_col) = df.shape

    # Define the conditional formatting rules
    # Street format
    worksheet.conditional_format(1, 1, max_row, max_col,
                                 {"type": "cell",
                                  "criteria": "==",
                                  "value": '" "',  # Note: Needs double quote because that's what Excel knows
                                  "format": format_street})
    # Empty lot format
    worksheet.conditional_format(1, 1, max_row, max_col,
                                 {"type": "cell",
                                  "criteria": "==",
                                  "value": '"."',
                                  "format": format_empty_lot})
    # House without parcels
    worksheet.conditional_format(1, 1, max_row, max_col,
                                 {"type": "cell",
                                  "criteria": "==",
                                  "value": '"H"',
                                  "format": format_house_no_parcel})
    # House with parcels
    worksheet.conditional_format(1, 1, max_row, max_col,
                                 {"type": "cell",
                                  "criteria": "==",
                                  "value": '"P"',
                                  "format": format_house_with_parcel})
    # Depot
    worksheet.conditional_format(1, 1, max_row, max_col,
                                 {"type": "cell",
                                  "criteria": "==",
                                  "value": '"D"',
                                  "format": format_depot})

    worksheet.set_column(0, max_col, column_width, format_column)
    worksheet.set_default_row(row_height)

    # Close the Pandas Excel writer and save the Excel file
    writer._save()



def export_map_to_excel_with_formatting(df, wb_path, ws_name):
    # Define style parameter
    colo_street = "#A6A6A6"
    colo_house_no_parcel = "#95B8D1"
    colo_house_with_parcel = "#9B1D20"  # "#522B47"
    colo_depot = "#F77F00"
    colo_empty_lot = "#566E3D"
    column_width = 3
    row_height = 12

    # Create a Pandas Excel writer
    writer = pd.ExcelWriter(wb_path, engine='xlsxwriter')

    # Write the DataFrame to the Excel file
    #df.insert(0, 'Row', df.index + 1)  # Add a new column with row numbers, starting from 1
    df.to_excel(writer, sheet_name=ws_name)

    # Get the workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets[ws_name]

    format_street = workbook.add_format({"bg_color": colo_street, "font_color": colo_street})
    format_house_no_parcel = workbook.add_format({"bg_color": colo_house_no_parcel,
                                                  "font_color": colo_house_no_parcel,
                                                  "border": 1})
    format_house_with_parcel = workbook.add_format({"bg_color": colo_house_with_parcel,
                                                    "font_color": colo_house_with_parcel,
                                                    "border": 1})
    format_empty_lot = workbook.add_format({"bg_color": colo_empty_lot, "font_color": colo_empty_lot})
    format_depot = workbook.add_format({"bg_color": colo_depot, "font_color": colo_depot, "border": 1})
    format_column = workbook.add_format({"align": "center", 'valign': "vcenter"})

    (max_row, max_col) = df.shape

    for index, row in df.iterrows():
        for col_num, value in enumerate(row):
            # Get the cell address
            cell_address = xlsxwriter.utility.xl_col_to_name(col_num + 1) + str(
                index + 2)  # +1 for the additional column, +2 to account for header row

            # Apply the cell format based on the value
            if value == " ":
                worksheet.write(cell_address, value, format_street)
            elif value == ".":
                worksheet.write(cell_address, value, format_empty_lot)
            elif value == "H":
                worksheet.write(cell_address, value, format_house_no_parcel)
            elif value == "P":
                worksheet.write(cell_address, value, format_house_with_parcel)
            elif value == "D":
                worksheet.write(cell_address, value, format_depot)
            else:
                worksheet.write(cell_address, value)
    worksheet.set_column(1, max_col, column_width, format_column)  # Start from column 1 to exclude index column
    worksheet.set_default_row(row_height)

    # Close the Pandas Excel writer and save the Excel file
    writer._save()

# data = {'Name': ['Alice', 'D', 'H'],
#         'Age': [25, 30, 35],
#         'City': [' ', '.', 'P']}
# test_df = pd.DataFrame(data)
#
# export_map_to_excel_with_formatting_2(test_df,
#                                       r"C:\Users\fhaum\OneDrive\401 MASTER - Masterarbeit\04 Kalkulationen\pythonProject\PathVisualisation_TEST.xlsx",
#                                       "VIS")
