import os
# import ini biar ga banyak warning muncul pas di run
import warnings
warnings.filterwarnings('ignore')

import pandas as pd
# import en_core_web_sm
# nlp = en_core_web_sm.load()
# import PyPDF2

from pyresparser import ResumeParser

# Folder path containing the CVs
cv_folder_path = "D:\\UiPath\\ExtractData_V2\\ExtractData_V2\\attachments\\"
cv_files = os.listdir(cv_folder_path)

processed_data = []
for cv_file in cv_files:
    data = ResumeParser(os.path.join(cv_folder_path, cv_file)).get_extracted_data()
    cv_data = {
        "File": data.get("name", ""),
        "Personal Info": {
            "Name": data.get("name", ""),
            "Title": data.get("title", ""),
            "Location":  data.get("location", ""),
            "Phone": data.get("mobile_number", ""),
            "Email": data.get("email", "")
            
        },
    "Summary": data.get("summary", ""),
    "Experience": data.get("experience", []),
    "Education": data.get("education",[]),
    "Skills": data.get("skills",[])
    }
    processed_data.append(cv_data)
# print(processed_data)

df = pd.DataFrame(processed_data)

output_excel_dir = "D:\\UiPath\\ExtractData_V2\\Excel"
output_excel_name = "output.xlsx"
output_excel_path = os.path.join(output_excel_dir, output_excel_name)

if os.path.exists(output_excel_path):
    #extract filenamenya dan extensionnya
    filename, extension = os.path.splitext(output_excel_name)
    counter = 1
    while True:
        #tambahkan counter ke nama filenya
        output_excel_name = f"{filename}[{counter}]{extension}"
        output_excel_path = os.path.join(output_excel_dir, output_excel_name)
        if not os.path.exists(output_excel_path):
            break
        counter += 1
df.to_excel(output_excel_path, index=False, engine='xlsxwriter')

# Open the Excel file after it's created
os.startfile(output_excel_path)