import pandas as pd
import openpyxl
from row_data import Row

# Notes:
# tar file: SPA is row.iloc[0], service code is 1, charge is 2, new charge is 4
# ecb file: SPA is 0, service code is 1, charge is 3, new charge is 5
# final Excel file columns SPA, Service Code, Reason For Error, TAR Charge, ECB Charge, TAR New Charge, ECB New Charge

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
    for e_index, e_row in ecb_file.iterrows():
        # Add each ecb row into the relevant class
        ecb_row.set_spa(e_row.iloc[0])
        ecb_row.set_service_code(e_row.iloc[1])
        ecb_row.ecb_set_charges(e_row.iloc[3], e_row.iloc[5])
        if tar_row.spa == ecb_row.spa and tar_row.service_code == ecb_row.service_code:
            # If a row with the matching spa and code is found, compare charges
            match = True
            if tar_row.charge != ecb_row.charge:
                print("Different charges")
            elif tar_row.new_charge != ecb_row.new_charge:
                print("Different new charges")
            # Exit the innter loop
            break

    # If no match was found, report an error
    if not match:
        print("No match.")

# Some testing code to determine how to access needed information
# ecb_file = pd.read_excel('ServiceCodes_ECB.xlsx', skiprows=2, nrows=10, index_col=None)
# for index, row in ecb_file.iterrows():
#     print(f"{row.iloc[0]} {row.iloc[1]} {row.iloc[3]} {row.iloc[5]}")
