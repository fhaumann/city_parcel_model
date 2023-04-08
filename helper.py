import string
import time
from datetime import timedelta
import pandas as pd


def generate_letter_combinations(length:int=3):
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
        df_to_save.to_csv(csv_full_path,index=True, header=True)
    else:
        df_to_save.to_csv(csv_full_path, index=False, header=False)


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



def export_pd_to_excel_with_formatting(df, wb_path, ws_name):
    import pandas as pd
    from openpyxl import load_workbook

    # Create a Pandas Excel writer object
    ws_name = 'Sheet1'  # Worksheet name
    wb_path = 'output.xlsx'  # Output file path
    writer = pd.ExcelWriter(wb_path, engine='openpyxl')
    df.to_excel(writer, sheet_name=ws_name, index=False)

    # Load the workbook
    workbook = writer.book

    # Get the worksheet object
    worksheet = writer.sheets[ws_name]

    # Save the workbook to apply conditional formatting
    workbook.save(wb_path)

    # Load the workbook again
    workbook = load_workbook(wb_path)

    # Get the worksheet object
    worksheet = workbook[ws_name]

    # Get the conditional formatting rules from the original worksheet
    for rule in worksheet.conditional_formatting._cf_rules:
        worksheet.conditional_formatting.add(rule)

    # Save the workbook with the applied conditional formatting
    workbook.save(wb_path)

    # Close the Pandas Excel writer
    writer.close()



data = {'Name': ['Alice', 'Bob', 'Charlie'],
        'Age': [25, 30, 35],
        'City': ['New York', 'Los Angeles', 'Chicago']}
test_df = pd.DataFrame(data)

export_pd_to_excel_with_formatting(test_df, r"C:\Users\fhaum\OneDrive\401 MASTER - Masterarbeit\04 Kalkulationen\pythonProject\PathVisualisation_TEST.xlsx","test")



