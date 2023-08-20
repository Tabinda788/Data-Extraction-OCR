import pandas as pd
import os
import csv
import xml.etree.ElementTree as ET
from csv import writer
import re
import docx
from win32com import client as wc
import re


path = "C:\\Users\\nazid\\Downloads\\Programmatic Extraction Assessment\\Programmatic Extraction Assessment"
output_path ="C:\\Users\\nazid\\Downloads\\Programmatic Extraction Assessment\\Programmatic Extraction Assessment\\Output"
file_name_list ,first_name_list, middle_name_list, last_name_list, ssn_list,dob_list,adress_lis,city_list,state_list,zip_list=[],[],[],[],[],[],[],[],[],[]
date_pattern= r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}'
zip_pattern = re.compile(r'\d{5}')
address_pattern = r"(\d+\s+[A-Za-z\s#.-]+)"
city_pattern = r'\b[A-Z]{2}\b'
ssn_pattern = r'\b\d{3}-\d{2}-\d{4}\b'

combined_df = pd.DataFrame()



for name in os.listdir(path+ "\\Bucket 01"):
    if name.endswith(".xlsx"):
        excel_df = pd.read_excel(path + "\\Bucket 01\\" + name)
        extracted_excel_df = excel_df[['Claimant_Name','Claimant_Date_of_Birth','Member_ID','Policy_Number','Claim_Number']]
        extracted_excel_df["File Name"] = name
        combined_df = pd.concat([extracted_excel_df, combined_df], ignore_index=True)
    elif name.endswith(".csv"):
        csv_df = pd.read_csv(path + "\\Bucket 01\\" + name)
        extracted_csv_df = csv_df[['Claimant_Name','Claimant_Date_of_Birth','Member_ID','Policy_Number','Claim_Number']]
        extracted_csv_df["File Name"] = name
        combined_df = pd.concat([extracted_csv_df, combined_df], ignore_index=True)
combined_df.to_csv(output_path + "\\"+ "buc01-combine.csv",index=False)
print(combined_df)


# BUCKET 2


for name in os.listdir(path + "\\Bucket 02"):
    excel_df = pd.read_excel(path + "\\Bucket 02\\" + name)
    new_df = excel_df.dropna(how= "all",axis=1)
    new_df = new_df.dropna(how= "all", axis=0)
    if new_df.columns[0].startswith("Unnamed"):
        new_df.columns = new_df.iloc[0]
        new_df = new_df[1:]
        new_df.rename(columns={new_df.columns[0]:'Employee Name',new_df.columns[1]:'Social Security Number',
                    new_df.columns[2]:'Employee ID',new_df.columns[2]:'D0B'}, inplace=True)
        # print(name)
        combined_df["File Name"] = name
        combined_df = pd.concat([new_df, combined_df], ignore_index=True)
        
    elif name.endswith("010.xlsx"):
        new_df.drop(new_df.index[:3], inplace=True)
        new_df.drop(new_df.columns[0], axis=1, inplace=True)
        new_header = new_df.iloc[0]
        new_df = new_df[1:] 
        new_df.columns = new_header
        new_df.rename(columns={new_df.columns[0]:'Employee Name',new_df.columns[1]:'Social Security Number',
                    new_df.columns[2]:'Employee ID',new_df.columns[3]:'DOB'}, inplace=True)
        combined_df["File Name"] = name
        combined_df = pd.concat([new_df, combined_df], ignore_index=True)
        
    else:
        if 'DOB' not in new_df.columns:
            new_df['DOB'] = pd.Series(dtype='datetime64[ns]')
        new_df.rename(columns={new_df.columns[0]:'Employee Name',new_df.columns[1]:'Social Security Number',
                    new_df.columns[2]:'Employee ID','File Name': name}, inplace=True)
        combined_df["File Name"] = name
        combined_df = pd.concat([new_df, combined_df], ignore_index=True)
        
    # print(name)
    combined_df.to_csv(output_path + "\\"+ "buc02-combine.csv",index=False)

print(combined_df)



# Bucket 07 - File Name, Full Name, Social Security Number


data_list = []

def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

def extract_employee_info(text):
    employee_info = {}
    employee_info["File Name"] = name
    name_match = re.search(r"Employee Name:\s+([^\n]+)", text)
    if name_match:
        employee_info["Employee Name"] = name_match.group(1).strip()
    ssn_match = re.search(r"Social Security No:\s+(\d{3}-\d{2}-\d{4})", text)
    if ssn_match:
        employee_info["Social Security No"] = ssn_match.group(1).strip()
    print(employee_info)
    return employee_info

for name in os.listdir(path + "\\Bucket 07"):
    if name.endswith('.docx'):
        extracted_text = extract_text_from_docx(path + "\\Bucket 07\\" + name)
        employee_info = extract_employee_info(extracted_text)
        data_list.append(employee_info)

# print(data_list)
df = pd.DataFrame(data_list)
df.to_csv(output_path + "\\"+ "combine_buck07.csv",index=False)
print(df)



# Bucket - 08



def write_to_csv(data,row):
	with open(data, 'a') as f_object:
		writer_object = writer(f_object)
		writer_object.writerow(row)
	return data 

for name in os.listdir(path+ "\\Bucket 08"):
    if name.endswith(".XML"):
        tree = ET.parse(path + "\\Bucket 08\\" + name)
        root = tree.getroot()
        for element in root:
            if str(element.tag).endswith('Worksheet'):
                for sub_element in element:
                    for row in sub_element:
                        lis=[]
                        for cell in row:
                            for data in cell:
                                if data.text==None:
                                    continue
                                lis.append(data.text)
                        if len(lis) > 10:   
                            add_csv =  write_to_csv( output_path + "\\" + str(name.split(".")[0]) + ".csv",lis)	
                           
for files in os.listdir(output_path):
    if files.endswith("29.csv"):
        filename= output_path + "\\" + files
        with open(filename, 'r') as csvfile:
            datareader = csv.reader(csvfile)
            for row in datareader:
                if len(row)>0:
                    ssn_list.append("")
                    file_name_list.append("ULX-IRT000029.XML")
                    first_name_list.append(row[4])
                    middle_name_list.append(row[5])
                    if row[6] not in ['Male', 'Female']:
                        last_name = row[6]
                        last_name_list.append(last_name)
                    else:
                        last_name=''
                        last_name_list.append(last_name)
                    dob_match=[re.search(date_pattern, ele) for ele in row if re.search(date_pattern, ele)]
                    try:
                        dob = dob_match[0].group()
                        dob_list.append(dob)
                    except:
                        dob=""
                        dob_list.append(dob)
                    zip_match = [re.search(zip_pattern, ele) for ele in row if re.search(zip_pattern, ele)]
                    try:
                        zip_list.append(zip_match[0].group())
                    except:
                        zip_list.append('')
                    adress_match = [re.search(address_pattern, ele) for ele in row if re.search(address_pattern, ele)]
                    try:
                        adress_lis.append(adress_match[0].group())
                    except:
                        adress_lis.append('')
                    city_match = [re.findall(city_pattern, ele) for ele in row if re.findall(city_pattern, ele)]
                    try:
                        city_list.append(city_match[2][0])
                    except:
                        city_list.append('')
                    if len(row[13]) > 2 and len(row[13]) <10 and not row[13].isdigit():
                        state_list.append(row[13])
                    elif len(row[14]) > 2 and len(row[14]) <10 and not row[14].isdigit() and len(row[14]) != 23:
                        state_list.append(row[14])
                    else:
                        state_list.append("")
    elif files.endswith("30.csv"):
        filename = output_path + "\\" + files
        with open(filename, 'r') as csvfile:
            '''We know that second file gives us ssn number, fist last name and dob of some unknown enployees'''
            datareader = csv.reader(csvfile)
            for row in datareader:
                if len(row)>1:
                    file_name_list.append("ULX-IRT000030.XML")
                    first_name_list.append(row[1])
                    middle_name_list.append("")
                    last_name_list.append(row[2])
                    ssn_list.append(row[10])
                    dob_list.append(row[3])
                    adress_lis.append("")
                    city_list.append("")
                    state_list.append("")
                    zip_list.append("")




                                    
data = {'File Name': file_name_list,'First Name': first_name_list, 'Middle Name': middle_name_list, 'Last Name': last_name_list , 
'Social Security Number' : ssn_list, 'Date of Birth': dob_list, 'Address' : adress_lis, 'City': city_list,
'State': state_list, 'ZIP': zip_list
}
df = pd.DataFrame(data)
df.to_csv(output_path + "\\" + "buc08-combine.csv", index=False)
print(df)

