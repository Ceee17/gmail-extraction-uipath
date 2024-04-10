import os
# import ini biar ga banyak warning muncul pas di run
import warnings
warnings.filterwarnings('ignore')
import pandas as pd

# library utk parse CV
from pyresparser import ResumeParser

# path ke folder cv
cv_folder_path = "D:\\UiPath\\ExtractData_V2\\ExtractData_V2\\attachments\\"
cv_files = os.listdir(cv_folder_path)

processed_data = []
for cv_file in cv_files:
    data = ResumeParser(os.path.join(cv_folder_path, cv_file)).get_extracted_data()
    cv_data = {
        "Name": data.get("name", ""),
        "Email": data.get("email", ""),
        "Phone Number": data.get("mobile_number", ""),
        "Position": data.get("designation", ""),
        "Degree": data.get("degree", ""),
        "Company Names": data.get("company_names", ""),
        "Skills": ", ".join(data.get("skills", [])),  # Convert skills list to comma-separated string
        "Total Experience": data.get("total_experience", ""),
        "no_of_pages": data.get("no_of_pages", "")
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
        

# memasukkan dataframe ke excel
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Enable text wrapping
    for i, col in enumerate(df.columns):
        wrap_format = workbook.add_format({'text_wrap': True})
        worksheet.set_column(i, i, width=40, cell_format=wrap_format)

# buka file excel
os.startfile(output_excel_path)