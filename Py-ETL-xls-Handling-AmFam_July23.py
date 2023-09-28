
#%%

#Author: Chinmay Patane, Business Data Analyst, American Family Mutual Insurance Co., Built for Commercial Auto Line of Business in Pricing

##Details: This ETL is written in Python. At AmFam commercial Insurance Company, each month new enterprise rates need to be uploaded for every LoB 
#(Commericial Auto, Worker's Compensation, Business Owners, Commercial Property, General Liability etc.) so that competitive market rate 
#insurance premium can be offered to the customers. 

##This ETL is made for LoB commercial Auto (CA) when certain systems are still using the legacy .xls files,
#(like BaseRate,Primaryt, Secondary etc). This ETL is made with a purpose of: 

#Read the requirements from .xlsx format and extract the data from it
#Access the legacy files using this ETL, open them, write data in to them and save them
#Create a report to notify stakeholders - consisting of how many states with their effective dates are modofied. This report is used in Unit Testing and User Acceptance Testing 
#Create a report to make sure all the given requirements match data extracted and ready to be used in rating


#%%
Revision_Number = '23.3.20'
Product_Version = '23060600'
Path_to_save_files = '/Users/CXP087/Desktop/Py/CA/'

#%%
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import xlwt
import xlrd
 
#%%

# Import the data from Excel
Get_Requirements = pd.read_excel("/Users/CXP087/Desktop/Py/CA/Requirements.xlsx", sheet_name= 'Sheet1')

Get_Requirements.columns = [
    "State",
    "NB_Effective_Date",
    "RN_Effective_Date",
    "Rate_Table_Changes",
    "REVISION"
]


# Convert date columns to datetime
Get_Requirements['NB_Effective_Date'] = pd.to_datetime(Get_Requirements['NB_Effective_Date'], errors='coerce')
Get_Requirements['RN_Effective_Date'] = pd.to_datetime(Get_Requirements['RN_Effective_Date'], errors='coerce')

# Convert date columns to desired format
Get_Requirements['NB_Effective_Date'] = Get_Requirements['NB_Effective_Date'].dt.strftime('%m/%d/%Y')
#Get_Requirements['RN_Effective_Date'] = Get_Requirements['RN_Effective_Date'].dt.strftime('%m/%d/%Y')


# Initiate the report
Requirements = pd.DataFrame(columns=[
    "Rate_Table_Changes",
    "IEL_Table",
    "States",
    "States_Total",
    "Revisions",
    "Number_of_Lines",
    "Revisions_Found",
    "States_Found",
    "States_Found_Total",
    "Mismatch_Found"
])

# Extract unique Rate_Table_Changes values
unique_rates = Get_Requirements['Rate_Table_Changes'].unique()


for rate in unique_rates:
    states = Get_Requirements.loc[Get_Requirements['Rate_Table_Changes'] == rate, 'State']
    formatted_dates = pd.to_datetime(Get_Requirements.loc[Get_Requirements['Rate_Table_Changes'] == rate, 'NB_Effective_Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    states_with_dates = [f"{state} ({date})" for state, date in zip(states, formatted_dates)]
    states_string = ', '.join(states_with_dates)
    
    revisions = Get_Requirements.loc[Get_Requirements['Rate_Table_Changes'] == rate, 'REVISION'].unique()
    revisions_string = ', '.join(revisions)
    
    new_row = pd.DataFrame({
        'Rate_Table_Changes': [rate],
        'States': [states_string],
        'Revisions': [revisions_string]
    })
    
    Requirements = pd.concat([Requirements, new_row], ignore_index=True)


#%%
#Get the files copied to the location and get them prepared
import shutil
import os

# Get mapping of factors vs tables
required_tabs = Requirements["Rate_Table_Changes"].tolist()
Mapping = pd.read_excel("/Users/CXP087/Desktop/Rate Upload/CA/Mapping.xlsx")

# Filter mapping based on required factors
Mapping = Mapping[Mapping["Factor"].isin(required_tabs)]

# Re-order mapping to match the datasets, and then fill the requirements dataset IEL_Table column
Mapping = Mapping.set_index("Factor").loc[required_tabs].reset_index()
Requirements['IEL_Table'] = Mapping['Table']


Files_List = list(Requirements["IEL_Table"])

# Source and destination directories
source_dir = "/Users/CXP087/Desktop/Rate Upload/CA/CA Blank Tables"
destination_dir = Path_to_save_files

# Iterate through the files in Files_List and copy them to the destination directory
for file_name in Files_List:
    # Add the ".xls" extension to the file name
    file_name_with_extension = f"{file_name}.xls"
    
    source_path = os.path.join(source_dir, file_name_with_extension)
    destination_path = os.path.join(destination_dir, file_name_with_extension)
    shutil.copyfile(source_path, destination_path)
    print(f"File '{file_name_with_extension}' copied to {destination_dir}")
    

#%%
#Read the Factors Tabs from the Rating Matrix
Data1 = pd.read_excel("/Users/CXP087/Desktop/Py/CA/CA-Rating-Algorithm-Matrix.xlsx", sheet_name=required_tabs)


#%%

#Export the data and create the report in one 'for' loop

Factor_Name = Mapping["Factor"].tolist()
Table_Name = Mapping["Table"].tolist()

# Export Table and Report
for i in range(len(Factor_Name)):
    f = Factor_Name[i]
    t = Table_Name[i]
    
    Data3 = Data1[f].copy()
    Data3 = pd.DataFrame(Data3)
    Data3.columns = Data3.iloc[0]
    Data3 = Data3[1:]
    
    Data3 = Data3.dropna(subset=["REVISION"])  # Get only those with revision number mentioned
    Data3 = Data3[Data3["REVISION"] == Revision_Number]  # Get desired revision number for this release
    
    Data3["PRODUCT_VERSION"] = Product_Version  # Desired Product Version
    Data3["SOURCE"] = "Custom"  # Source is always Custom
    Data3 = Data3.loc[:, ~Data3.columns.isin(["RESOURCE_OWNER", "CUSTOMER_CODE", "RELEASE_NUMBER", "GID"])]
    
    #Open file named t with extention .xls from location saved in variable named 'Path_to_save_files'
    #Paste the data in Data3 in that file, wherever the column name matches
    #Save the file
    
    # Open file named t with extension .xls from the location saved in the variable named 'Path_to_save_files'
    
    
    # Open the old legacy .xls file using xlrd
    file_path = f"{Path_to_save_files}/{t}.xls"
    rb = xlrd.open_workbook(file_path, formatting_info=True)
    sheet = rb.sheet_by_index(0)

    # Create a new workbook and sheet using xlwt
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1')

    # Write the existing data from the old .xls file to the new workbook
    for row_idx in range(sheet.nrows):
        for col_idx in range(sheet.ncols):
            cell_value = sheet.cell_value(row_idx, col_idx)
            cell_format = sheet.cell_xf_index(row_idx, col_idx)
            ws.write(row_idx, col_idx, cell_value, xlwt.easyxf(rb.xf_list[cell_format].format_key))

    # Convert the dataframe to an xlwt-compatible dataframe
    df_to_excel = pd.DataFrame(Data3)
    df_to_excel = df_to_excel.rename(columns={col: str(col) for col in df_to_excel.columns})

    # Paste the dataframe into the Excel file starting from row 3
    for row_idx, row in enumerate(df_to_excel.values):
        for col_idx, cell_value in enumerate(row):
            ws.write(row_idx + 2, col_idx, cell_value)

    # Save the modified .xls file
    wb.save(file_path)
 
    
    # Make Report for whether given requirements match with the matrix data
    Requirements.loc[i, "Number_of_Lines"] = len(Data3)
    Requirements.loc[i, "Revisions_Found"] = ", ".join(Data3["REVISION"].unique())
    
    unique_states_dates = [f"{state} ({date})" for state, date in zip(Data3["STATE_CODE"], Data3["EFFECTIVE_DATE"])]
    unique_states_dates = pd.Index(unique_states_dates).unique()
    Requirements.loc[i, "States_Found"] = ", ".join(unique_states_dates)
    
    Requirements.loc[i, 'IEL_Table'] = t

    
    # Fill the last column, Mismatch_Found, in the report based on the states mentioned in the requirements vs states found in the rating matrix mentioned in States_Found
    X = [x.strip() for x in str(Requirements.loc[i, "States"]).split(",")]
    Y = [y.strip() for y in str(Requirements.loc[i, "States_Found"]).split(",")]
    
    if set(X) == set(Y):
        Requirements.loc[i, "Mismatch_Found"] = "No"
    else:
        Requirements.loc[i, "Mismatch_Found"] = "Yes"

Requirements["States_Found_Total"] = Requirements["States_Found"].str.split(",").apply(lambda x: len(set([state.strip() for state in x])))

Requirements.to_excel(f"{Path_to_save_files}Requirements_Report.xlsx", index=False)