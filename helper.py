import string
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

