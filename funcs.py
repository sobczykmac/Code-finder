import xlrd
from datetime import datetime
import openpyxl
import csv
import glob
import os
from tkinter import StringVar, messagebox



rows_list_xls = []
rows_xlsx = []
rows_csv = []
rows_txt = []


def check_file_type(file_name):
    global ext
    ext = os.path.splitext(file_name)[-1].lower()


def copy_rows_with_text(file_path, text_to_find):
    # Load the workbook using xlrd
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)

    # Create a list to store the copied rows
    global rows_list_xls



    # Iterate over the rows in the source sheet and copy matching rows to the new sheet
    for row_index in range(sheet.nrows):
        row = sheet.row(row_index)
        for cell in row:
            if text_to_find in str(cell.value):
                # Copy the row and its formatting to the list
                new_row = []
                for cell in row:
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        # Convert the Excel date value to a Python datetime object
                        date_tuple = xlrd.xldate_as_tuple(cell.value, workbook.datemode)
                        date = str(datetime(*date_tuple))
                        # Append the datetime object to the new row
                        new_row.append(date)
                    else:
                        # Append the cell value to the new row
                        new_row.append(cell.value)
                row_as_string = " ".join([str(value) for value in new_row])
                print(row_as_string)
                rows_list_xls.append(row_as_string)

    return rows_list_xls


def copy_row_xlsx(filepath, text_to_find):
    # load the file using openpyxl
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    for row in ws.rows:
        for cell in row:
            if text_to_find in str(cell.internal_value):
                row_as_string = " ".join([str(value.internal_value) for value in row])
                rows_xlsx.append(row_as_string)
    return rows_xlsx


def copy_csv_row(file_name, text_to_find, delimiter=","):
    with open(file_name) as file:
        reader = csv.reader(file)
        for row_nums in reader:
            if text_to_find in row_nums:
                row_as_string = " ".join([str(value) for value in row_nums])
                rows_csv.append(row_as_string)

    return rows_csv


def copy_txt_row(filename, text_to_find):
    with open(filename, "r") as f:
        Lines = f.readlines()
        for line in Lines:
            if text_to_find in line:
                rows_txt.append(line)

    return rows_txt


def msg_finished():
    messagebox.showinfo(title="Code finder process", message="Done copying the data, check your output file.")