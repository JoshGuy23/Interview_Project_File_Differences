import pandas as pd
from pandas import ExcelWriter, ExcelFile
import openpyxl

# Notes:
# tar file: SPA is row.iloc[0], service code is 1, charge is 2, new charge is 4
# ecb file: SPA is 0, service code is 1, charge is 3, new charge is 5
# final Excel file columns SPA, Service Code, Reason For Error, TAR Charge, ECB Charge, TAR New Charge, ECB New Charge
spa_list = []
service_code_list = []
error_list = []
t_charge_list = []
e_charge_list = []
t_new_charge_list = []
e_new_charge_list = []


# Read in the tar and ecb files
tar_file = pd.read_excel('ServiceCodes_TAR.xlsx', skiprows=1, index_col=None)
ecb_file = pd.read_excel('ServiceCodes_ECB.xlsx', skiprows=1, index_col=None)
all_entries = ecb_file.merge(tar_file, on=['SPA', 'Service Code'], how='outer')

for row in all_entries.itertuples():
    if pd.isnull(row[4]) or pd.isnull(row[10]):
        spa_list.append(row[1])
        service_code_list.append(row[2])
        error_list.append("Entry not in both files")
        t_charge_list.append("N/A")
        e_charge_list.append("N/A")
        t_new_charge_list.append("N/A")
        e_new_charge_list.append("N/A")
    else:
        ecb_charge = float(row[4][1:])
        if ecb_charge != float(row[10]) and float(row[6]) != float(row[12]):
            spa_list.append(row[1])
            service_code_list.append(row[2])
            error_list.append("Differing charges and new charges")
            e_charge_list.append(f"{ecb_charge}")
            t_charge_list.append(f"{float(row[10])}")
            e_new_charge_list.append(f"{float(row[6])}")
            t_new_charge_list.append(f"{float(row[12])}")
        elif ecb_charge != float(row[10]):
            spa_list.append(row[1])
            service_code_list.append(row[2])
            error_list.append("Differing charges")
            t_charge_list.append(f"{float(row[10])}")
            e_charge_list.append(f"{ecb_charge}")
            t_new_charge_list.append("N/A")
            e_new_charge_list.append("N/A")
        elif float(row[6]) != float(row[12]):
            spa_list.append(row[1])
            service_code_list.append(row[2])
            error_list.append("Differing new charges")
            t_charge_list.append("N/A")
            e_charge_list.append("N/A")
            t_new_charge_list.append(f"{float(row[12])}")
            e_new_charge_list.append(f"{float(row[6])}")


print("Done finding differences. Now writing to file.")
# Write data to Excel file
df = pd.DataFrame({'SPA': spa_list,
                   'Service Code': service_code_list,
                   'Reason For Difference': error_list,
                   'TAR Charge': t_charge_list,
                   'ECB Charge': e_charge_list,
                   'TAR New Charge': t_new_charge_list,
                   'ECB New Charge': e_new_charge_list})

writer = ExcelWriter('ServiceCodes_DifferenceLog.xlsx')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.close()

print("Done writing to file.")

# Some testing code to determine how to access needed information
# ecb_file = pd.read_excel('ServiceCodes_ECB.xlsx', skiprows=2, nrows=10, index_col=None)
# for index, row in ecb_file.iterrows():
#     print(f"{row.iloc[0]} {row.iloc[1]} {row.iloc[3]} {row.iloc[5]}")
