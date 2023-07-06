import pandas as pd
import numpy as np
import glob
import os


path = os.getcwd()
excel_files = glob.glob(os.path.join(path, "*.xlsx"))
print(excel_files)

for file in excel_files:
    df_all = pd.read_excel(file, header=None, sheet_name= None)
    print(type(df_all))
    for key in list(df_all.keys()):
        if key!='hiddenSheet':
            print(key)
            print(df_all[key].head())
            df=df_all[key]
            file_name=key+file.split("\\")[-1]
    # file_name='out'+file.split("\\")[-1]
    # print(file_name)
            label_row_index = np.nan
            for row_index, row in df.iterrows():
                for cell in row:
                    if isinstance(cell, str) and not df.isna().loc[row_index].any():
                        label_row_index = row_index
                        break
                if not np.isnan(label_row_index): 
                    break

            # Set the column labels using the identified row
            df.columns = df.iloc[label_row_index]
            # Drop the rows before the label row and reset the index
            df = df.iloc[label_row_index + 1:].reset_index(drop=True)
            df.to_excel(file_name, index=False)

