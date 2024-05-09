import pandas as pd
import openpyxl

tar_file = pd.read_excel('ServiceCodes_TAR.xlsx', skiprows=2, nrows=10, index_col=None)
# tar file: SPA is row.iloc[0], service code is 1, charge is 2, new charge is 4
# ecb file: SPA is 0, service code is 1, charge is 3, new charge is 5
for index, row in tar_file.iterrows():
    print(f"{row.iloc[0]} {row.iloc[1]} {row.iloc[2]} {row.iloc[4]}")

ecb_file = pd.read_excel('ServiceCodes_ECB.xlsx', skiprows=2, nrows=10, index_col=None)
for index, row in ecb_file.iterrows():
    print(f"{row.iloc[0]} {row.iloc[1]} {row.iloc[3]} {row.iloc[5]}")
