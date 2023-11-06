### import os
import pandas as pd
from IPython.display import display, HTML
display(HTML("<style>.container { width:100% !important; }</style>"))	

import simple_colors
user = os.getlogin()
path = f"C:/Users/{user}/BigB/"
directory = "My bank - CE (List) - Documents"
directory = os.path.join(path, directory)



# Common folder paths
user = os.getlogin()
data_folder = "SWedAFINANCE OUTPUT"
folder_path = f"C:/Users/{user}/BigB/"
data_path = os.path.join(folder_path, data_folder)


select_path  =  { 
        "2022" : {"1": "January_2022","2": "February_2022","3": "March_2022","4": "April_2022","5": "May_2022","6": "June_2022","7": "July_2022"},
        "2021" : {"1": "January_2021","2": "February_2021","3": "March_2021","4": "April_2021","5": "May_2021","6": "June_2021","7": "July_2021"},
        "2023": {"1": "January_2023", "2": "February_2023","3": "March_2023","4": "April_2023","5": "May_2023","6": "June_2023","7": "July_2023"}
        }


def process_excel_file(filepath):
    sheets_dict = pd.read_excel(filepath, sheet_name=None)
    all_sheets = []
    for name, sheet in sheets_dict.items():
        sheet['sheet'] = name
        all_sheets.append(sheet)
        
    full_table = pd.concat(all_sheets)
    full_table.fillna("no value", inplace=True)
    full_table["Type"] = full_table["sheet"]
    full_table["txnfeed_source"] = full_table["txn_id"].str[0:3]
    return full_table

print("IF YOU DON'T NEED TO CONTINUE, JUST PUT STOP or EXIT !!!", "bold")
print("\n")

while True:
    select_year_folder = input("Which year do you want to select from (2021, 2022, 2023) ?:").lower()
    
    if select_year_folder in ["stop", "exit"]:
        print("No need to download other files....")
        break
        
    month = input("Which month's output do you want to get (1-12) ?:")
    
    if not os.path.exists(data_path):
        os.makedirs(data_path)

    if select_year_folder in select_path:
        year_dict = select_path[select_year_folder]
        dynamic_year = f"{select_year_folder} monthly lists"
        
        if month in year_dict:
            subfolder = year_dict[month]
            relative_path = os.path.join(directory, dynamic_year, subfolder)
            files = [file for file in os.listdir(relative_path) if file.endswith('.xlsx')]
            
            data_lst = [process_excel_file(os.path.join(relative_path, file)) for file in files]
            tamdata = pd.concat(data_lst)
            
            output_file_path = os.path.join(data_path, f"{subfolder}.xlsx")
            tamdata.to_excel(output_file_path, index=False)
            
            union_table = input("Do you need to append (yes, no) tables?").lower()
            
            if union_table == "yes":
                file_list = [pd.read_excel(os.path.join(data_path, file)) for file in os.listdir(data_path)]
                unioned_df = pd.concat(file_list, ignore_index=True)
                unioned_file_path = os.path.join(data_path, "unioned_df_output_2.xlsx")
                unioned_df.to_excel(unioned_file_path, index=False)
            
        else:
            break
