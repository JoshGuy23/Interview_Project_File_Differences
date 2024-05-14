import pandas as pd
from pandas import ExcelWriter, ExcelFile
import openpyxl

# Notes:
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

# Merged files for easier comparison
all_entries = ecb_file.merge(tar_file, on=['SPA', 'Service Code'], how='outer')

# all_df = pd.DataFrame(all_entries)
# all_writer = ExcelWriter('Complete_Entries.xlsx')
# all_df.to_excel(all_writer, sheet_name='Sheet1', index=False)
# all_writer.close()

# test_file = tar_file.merge(ecb_file, on=['SPA', 'Service Code'], how='outer')
#
# test_df = pd.DataFrame(test_file)
# test_writer = ExcelWriter('ecb_to_tar_entries.xlsx')
# test_df.to_excel(test_writer, sheet_name='Sheet1', index=False)
# test_writer.close()

# Rows are iterated over as tuples.
for row in all_entries.itertuples():
    if pd.isnull(row[4]):
        # Null/NaN cells indicate that an entry doesn't exist in both files.
        spa_list.append(row[1])
        service_code_list.append(row[2])

        error_list.append("Entry only exists in the TAR file.")

        t_charge_list.append("N/A")
        e_charge_list.append("N/A")

        t_new_charge_list.append("N/A")
        e_new_charge_list.append("N/A")
    elif pd.isnull(row[10]):
        spa_list.append(row[1])
        service_code_list.append(row[2])

        error_list.append("Entry only exists in the ECB file.")

        t_charge_list.append("N/A")
        e_charge_list.append("N/A")

        t_new_charge_list.append("N/A")
        e_new_charge_list.append("N/A")
    else:
        # Remove the $
        ecb_charge = float(row[4][1:])

        # Check if the files have differing charges and/or new charges.
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
