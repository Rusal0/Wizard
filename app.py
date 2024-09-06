import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import openpyxl
from openpyxl import load_workbook
import hashlib


st.title('Excel Wizard')


# Function to split an Excel file into separate files for each sheet with formatting
def split_excel(file):
    original_wb = load_workbook(file, data_only=True)

    output = BytesIO()
    with zipfile.ZipFile(output, 'w') as zf:
        for sheet_name in original_wb.sheetnames:
            new_wb = openpyxl.Workbook()
            new_sheet = new_wb.active
            new_sheet.title = sheet_name

            original_sheet = original_wb[sheet_name]

            # Copy all formatting attributes, including conditional formatting
            for row in original_sheet.iter_rows():
                for cell in row:
                    new_cell = new_sheet[cell.coordinate]
                    new_cell.value = cell.value
                    new_cell.font = cell.font.copy()
                    new_cell.border = cell.border.copy()
                    new_cell.alignment = cell.alignment.copy()
                    new_cell.fill = cell.fill.copy()
                    new_cell.number_format = cell.number_format

                    # Additionally copy conditional formatting rules
                    for rule in cell.conditional_formatting.rules:
                        new_cell.conditional_formatting.add(rule.copy())

            with BytesIO() as sheet_output:
                new_wb.save(sheet_output)
                zf.writestr(f"{sheet_name}.xlsx", sheet_output.getvalue())

    output.seek(0)
    return output


# Function to merge (union) multiple Excel files into separate sheets within one Excel file
def merge_excels(files):
    combined_output = BytesIO()

    with pd.ExcelWriter(combined_output, engine='openpyxl') as writer:
        for file in files:
            try:
                excel_data = pd.read_excel(file, engine='openpyxl')

                # Check if excel_data is a DataFrame
                if isinstance(excel_data, pd.DataFrame):
                    sheet_name = file.name.split('.')[0]  # Extract the file name without extension
                    excel_data.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    for sheet_name in excel_data.sheetnames:
                        sheet_data = excel_data[sheet_name]

                        # Generate a unique sheet name using a hash for potential duplicate sheet names
                        with BytesIO(sheet_data.read()) as f:
                            file_data = f.read()
                            unique_id = hashlib.sha1(file_data + sheet_name.encode()).hexdigest()[:10]
                            new_sheet_name = f"{unique_id}_{sheet_name}"

                        writer.book = load_workbook(sheet_data, data_only=True)  # Load sheet data only
                        writer.sheets[sheet_name] = writer.book[sheet_name]  # Add sheet to writer

                        # Loop through rows and copy formatting (including conditional formatting)
                        for row in writer.sheets[sheet_name].iter_rows():
                            for cell in row:
                                original_cell = sheet_data[cell.coordinate]
                                cell.font = original_cell.font.copy()
                                cell.border = original_cell.border.copy()
                                cell.alignment = original_cell.alignment.copy()
                                cell.fill = original_cell.fill.copy()
                                cell.number_format = original_cell.number_format

                                # Copy conditional formatting rules
                                for rule in original_cell.conditional_formatting.rules:
                                    cell.conditional_formatting.add(rule.copy())

            except Exception as e:
                st.error(f"Error processing file {file.name}: {str(e)}")

    combined_output.seek(0)
    return combined_output


# File upload options
st.sidebar.title("Excel Wizard Options")
option = st.sidebar.radio("Choose an action", ('Split Excel by Sheets', 'Merge Excel Files'))

# Split Excel File
if option == 'Split Excel by Sheets':
    uploaded_file = st.file_uploader("Upload an
