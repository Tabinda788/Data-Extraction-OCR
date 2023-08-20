columns_to_extract = [
    "First Name", "Middle Name", "Last Name", "Social Security Number", "Social Security No",
    "Date of Birth", "DOB", "File No", "Member ID", "Address", "City", "State", "ZIP"
]
import pandas as pd
import os
data_list = []

path='C:\\Users\\nazid\\Downloads\\Programmatic Extraction Assessment\\Programmatic Extraction Assessment\\Output'
for name in os.listdir(path):
    if 'combine' in name:
        df= pd.read_csv(path + "\\"+ name)
        for index, row in df.iterrows():
            extracted_data = {
                column: row.get(column, "") for column in columns_to_extract
            }
            data_list.append(extracted_data)
            print("Adding data..",len(extracted_data))
result_df = pd.DataFrame(data_list)
result_df.to_csv("C:\\Users\\nazid\\Downloads\\Programmatic Extraction Assessment\\Output_Template.csv")