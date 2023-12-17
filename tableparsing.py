import pandas as pd
from io import StringIO
import xlsxwriter


def saveas_newsheet_inexistingfile(concatenated_df):

    existing_filename= input("Enter Existing file name: ")
    desired_sheetname = input("Enter the desired sheetname: ")
# To add the data as a new sheet to the existing xlsx file then continue the below code else
    with pd.ExcelWriter(existing_filename, engine='openpyxl', mode='a') as writer:
        # Write the DataFrame to a new sheet in the existing Excel file
        concatenated_df.to_excel(writer, sheet_name=desired_sheetname, index=False)

        # Save the changes to the Excel file
        writer._save()




# if you want to create a new single sheet file with the provided xml of html code then do this.
def saveas_single_file(concatenated_df):

    desired_filename = input("Enter the desired file name: ")
    desired_sheetname= input("Enter the desired sheet name: ")
    excel_writer = pd.ExcelWriter(desired_filename, engine='xlsxwriter')

    # Write the cleaned concatenated DataFrame to a single sheet in the Excel file
    # Sheet_name is what you want your sheet to be named as

    concatenated_df.to_excel(excel_writer, sheet_name=desired_sheetname, index=False)

    # Save the DataFrame to the Excel file
    excel_writer._save()


def main():
    # HTML table content or xml code
    file_path = input("Enter the xml or html file path with table : ")
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content=file.read()

    # Parse the HTML content to extract tables
    tables = pd.read_html(StringIO(html_content))

    # Concatenate all DataFrames row-wise
    concatenated_df = pd.concat(tables, ignore_index=True)


    # If you want to give the 6th column value based on 5th column then you can edit the below code 
    conditions = (concatenated_df[5].isin(['As expected', 'Issue fixed', 'Done']))
    concatenated_df.loc[conditions,6] ='pass'

    # Strip leading and trailing spaces from all columns
    concatenated_df = concatenated_df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # now decide what to happen.....save as file or in an existing file as a new sheet
    saveas_newsheet_inexistingfile(concatenated_df)
    # OR
    # saveas_single_file(concatenated_df)

if  __name__ == '__main__':
    main()
