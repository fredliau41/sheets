import pandas as pd
import time, os

END_FILENAME = "Master_Copy(AIO).xlsx"
TARGET_FILENAME = "target.xlsx"
SHEET_COLUMN_NAME = "Sheet Name"


def main():
    if not os.path.isfile(TARGET_FILENAME):
        print(f"Please name excel {TARGET_FILENAME}")
        time.sleep(5)
    else:
        excel_file = pd.read_excel(TARGET_FILENAME, sheet_name=None)
        sheets = excel_file.keys()

        master_df = []
        for my_sheet_name in sheets:
            sheet = pd.read_excel(TARGET_FILENAME, sheet_name=my_sheet_name)
            total_rows = len(sheet.index)
            my_rows = [my_sheet_name] * total_rows
            sheet.insert(0, SHEET_COLUMN_NAME, my_rows, allow_duplicates=True)

            master_df.append(sheet)

        dataset_combined = pd.concat(master_df, axis=0)
        columns_to_compare = dataset_combined.columns.drop(SHEET_COLUMN_NAME)
        dataset_combined.drop_duplicates(subset=columns_to_compare, keep="first", )
        dataset_combined.to_excel(END_FILENAME, sheet_name="Master sheet", index=False)


if __name__ == '__main__':
    main()
