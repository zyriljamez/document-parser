import pandas as pds
import os
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

excel_files_list = []

def create_approved_excel():

    df_output = pds.DataFrame()

    #get list of files
    for file_name in os.listdir():
        if file_name.endswith(".xlsx"):
            excel_files_list.append(file_name)

    #create a dataframe with approved items
    for file_name in excel_files_list:

        df_input = pds.read_excel(file_name)

        for index, item in df_input.iterrows():
            if item["Status"] == "Approved":
                df_output=df_output.append(item, ignore_index=True)


    #write the dataframe with approved items to an excel file
    writer = pds.ExcelWriter('approved.xlsx')
    df_output.to_excel(writer, index=False)
    writer.save()

    print("Approved excel file created")


if __name__ == "__main__":
    create_approved_excel()
