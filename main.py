import pandas as pd
from pandas import ExcelWriter, ExcelFile
import openpyxl
from row_data import Row

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
tar_file = pd.read_excel('ServiceCodes_TAR.xlsx', skiprows=2, nrows=10, index_col=None)
ecb_file = pd.read_excel('ServiceCodes_ECB.xlsx', skiprows=2, nrows=10, index_col=None)

# Set up the classes for comparison.
tar_row = Row()
ecb_row = Row()

for index, row in tar_file.iterrows():
    match = False

    # For each row in the tar file, put in the relevant data for that row
    tar_row.set_spa(row.iloc[0])
    tar_row.set_service_code(row.iloc[1])
    tar_row.set_charges(row.iloc[2], row.iloc[4])

    # Go through each row of the ecb file to find a match
    for ecb_index, e_row in ecb_file.iterrows():
        # Add each ecb row into the relevant class
        ecb_row.set_spa(e_row.iloc[0])
        ecb_row.set_service_code(e_row.iloc[1])
        ecb_row.ecb_set_charges(e_row.iloc[3], e_row.iloc[5])

        if tar_row.spa == ecb_row.spa and tar_row.service_code == ecb_row.service_code:
            diff_charges = False
            diff_new_charges = False

            # If a row with the matching spa and code is found, compare charges
            match = True

            if tar_row.charge != ecb_row.charge:
                diff_charges = True

            if tar_row.new_charge != ecb_row.new_charge:
                diff_new_charges = True

            # Append differing data to list to go into Excel file.
            if diff_charges and diff_new_charges:
                spa_list.append(tar_row.spa)
                service_code_list.append(tar_row.service_code)

                error_list.append("Differing charges and new charges")

                t_charge_list.append(f"{tar_row.charge}")
                e_charge_list.append(f"{ecb_row.charge}")

                t_new_charge_list.append(f"{tar_row.new_charge}")
                e_new_charge_list.append(f"{ecb_row.new_charge}")
            elif diff_charges:
                spa_list.append(tar_row.spa)
                service_code_list.append(tar_row.service_code)

                error_list.append("Differing charges")

                t_charge_list.append(f"{tar_row.charge}")
                e_charge_list.append(f"{ecb_row.charge}")

                t_new_charge_list.append("N/A")
                e_new_charge_list.append("N/A")
            elif diff_new_charges:
                spa_list.append(tar_row.spa)
                service_code_list.append(tar_row.service_code)

                error_list.append("Differing new charges")

                t_charge_list.append("N/A")
                e_charge_list.append("N/A")

                t_new_charge_list.append(f"{tar_row.new_charge}")
                e_new_charge_list.append(f"{ecb_row.new_charge}")

            # Exit the inner loop
            break

    # If no match was found, report an error
    if not match:
        spa_list.append(tar_row.spa)
        service_code_list.append(tar_row.service_code)

        error_list.append("Entry not in both files")

        t_charge_list.append("N/A")
        e_charge_list.append("N/A")

        t_new_charge_list.append("N/A")
        e_new_charge_list.append("N/A")

# Check if entries in ecb file not in tar file
for index, row in ecb_file.iterrows():
    match = False

    ecb_row.set_spa(row.iloc[0])
    ecb_row.set_service_code(row.iloc[1])

    for tar_index, t_row in tar_file.iterrows():
        tar_row.set_spa(t_row.iloc[0])
        tar_row.set_service_code(t_row.iloc[1])

        if tar_row.spa == ecb_row.spa and tar_row.service_code == ecb_row.service_code:
            match = True
            # Exit inner loop if match is found
            break

    if not match:
        spa_list.append(ecb_row.spa)
        service_code_list.append(ecb_row.service_code)

        error_list.append("Entry not in both files")

        t_charge_list.append("N/A")
        e_charge_list.append("N/A")

        t_new_charge_list.append("N/A")
        e_new_charge_list.append("N/A")

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

# Some testing code to determine how to access needed information
# ecb_file = pd.read_excel('ServiceCodes_ECB.xlsx', skiprows=2, nrows=10, index_col=None)
# for index, row in ecb_file.iterrows():
#     print(f"{row.iloc[0]} {row.iloc[1]} {row.iloc[3]} {row.iloc[5]}")
