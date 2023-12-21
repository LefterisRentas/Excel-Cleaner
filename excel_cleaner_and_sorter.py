# Writing a Python script to sort the DataFrame as per the specified criteria
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np
from tkinter import Tk, filedialog, Button, Label, Entry, Frame, StringVar, IntVar
import tkinter.font as tkFont
from datetime import datetime, timedelta
import os


def sort_excel_sheet(df, sort_columns):
    return df.sort_values(by=sort_columns)


def clean_duplicates(df, subset_columns, condition_column=None):
    """
    Remove duplicates from the DataFrame based on specified columns.
    If a condition column is specified, only remove duplicates where the condition column is empty.

    :param df: DataFrame from which duplicates will be removed
    :param subset_columns: List of column names based on which duplicates are identified
    :param condition_column: Column name where the row is removed only if this column is empty
    :return: DataFrame with duplicates removed
    """
    if condition_column:
        # Mask to identify rows where the condition column is empty
        condition_mask = df[condition_column].isna() | (df[condition_column] == '')

        # Removing duplicates with the condition
        # First, sort so that rows with data in the condition column come first
        df = df.sort_values(by=condition_column, ascending=False)

        # Then, drop duplicates based on the specified columns, keeping the first occurrence
        df = df.drop_duplicates(subset=subset_columns, keep='first')
    else:
        # Remove duplicates based on the specified columns without any condition
        df = df.drop_duplicates(subset=subset_columns)

    return df


def sort_and_separate_by_column(df, column_name, order_list, rows):
    # Creating a categorical column with specified order
    df[column_name] = pd.Categorical(df[column_name], categories=order_list, ordered=True)

    # Sorting by the categorical column
    df = df.sort_values(by=[column_name])

    # Adding empty rows
    separated_df = pd.DataFrame()
    for category in order_list:
        category_df = df[df[column_name] == category]
        if not category_df.empty:
            # Create a DataFrame with 6 empty rows, matching the column types of the original DataFrame
            empty_rows_df = pd.DataFrame(np.nan, index=np.arange(rows), columns=df.columns)

            # Concatenating category DataFrame with empty rows DataFrame
            separated_df = pd.concat([separated_df, category_df, empty_rows_df], ignore_index=True)

    return separated_df


# Columns to consider for identifying duplicates
subset_columns = ['Διεύθυνση', 'Δρομολόγιο', 'Μεταφορέας']
# Conditional column
condition_column = 'Αιτιολογία'
# Columns to sort by: Δρομολόγιο -> Περιοχή -> Επωνυμία
sort_columns = ['Δρομολόγιο', 'Περιοχή', 'Επωνυμία']

# Custom order for sorting
custom_order = [
    'XVAN', 'ΠΡΑΚΤΟΡΕΙΑ', 'ΑΝΑΤΟΛΙΚΟ', 'ΒΟΡΕΙΟ', 'ΓΕΝΙΚΟ', 'ΕΞΩΤΕΡΙΚΟ', 'ΚΕΝΤΡΟΔΥΤΙΚΟ', 'ΝΟΤΙΟ', 'ΠΕΛΟΠΟΝΝΗΣΟΣ',
    'ΣΤΕΡΕΑ', 'ΜΕΤΑΦΟΡΕΑΣ', 'ΤΑΚΗΣ ΜΕΤΑΦΟΡΙΚΗ', 'ΤΣΟΥΛΟΣ', 'ΒΑΓΙΑΣ', 'ΜΕΛ.ΠΕΡΙΦΕΡΙΑΚΟ', 'ΜΕΛ.ΑΝΑΤΟΛΙΚΟ',
    'ΜΕΛ.ΑΣΠΡΟΠΥΡΓΟΣ', 'ΜΕΛΕΤΗΣ 1', 'ΑΝΔΡΕΟΥ', 'DIRECT', 'ΣΤΑΜΠΟΥΛΗ ΜΤΦ'
]


def process_file(file_path, row_number):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Apply the sorting function
    sorted_df = sort_excel_sheet(df, ['Δρομολόγιο', 'Περιοχή', 'Επωνυμία'])

    # Apply the cleaning function
    cleaned_df = clean_duplicates(sorted_df, ['Διεύθυνση', 'Δρομολόγιο', 'Μεταφορέας'], 'Αιτιολογία')

    # Apply the sorting and separation function
    sorted_separated_df = sort_and_separate_by_column(cleaned_df, 'Δρομολόγιο', custom_order, row_number)

    # Create a new workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Append DataFrame rows to Excel worksheet
    for r in dataframe_to_rows(sorted_separated_df, index=False, header=True):
        ws.append(r)

    # Apply formatting (e.g., column widths)
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length
    # Save the workbook with tomorrow's date in the filename
    tomorrow = datetime.now() + timedelta(days=1)
    save_path = f"ΔΡΟΜΟΛΟΓΙΑ {tomorrow.strftime('%d.%m.%Y')}.xlsx"
    wb.save(save_path)

    # Open the processed file
    os.startfile(save_path)


def browse_file(file_path_var, row_count_var):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
    if file_path:
        file_path_var.set(f"Selected File: {file_path}")
        process_file(file_path, row_count_var.get())


def create_ui():
    root = Tk()
    root.title("Excel File Processor")
    root.geometry("600x200")

    # Set a custom font
    custom_font = tkFont.Font(family="Helvetica", size=12)

    # Frame for better layout control
    frame = Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill='both')

    # StringVar for dynamic label text
    file_path_var = StringVar()
    file_path_var.set("No file selected")

    # Create a label to display the selected file path
    file_label = Label(frame, textvariable=file_path_var, width=80, font=custom_font, wraplength=500)
    file_label.pack(pady=10)

    # Create a button to browse for the file
    browse_button = Button(frame, text="Browse and Process File", font=custom_font)
    browse_button.pack(pady=10)

    # IntVar for row count
    row_count_var = IntVar()
    row_count_var.set(6)  # Default value

    # Create a text field for the number of rows to be added
    row_count_entry = Entry(frame, textvariable=row_count_var, font=custom_font)
    row_count_entry.pack(pady=10)

    browse_button.config(command=lambda: browse_file(file_path_var, row_count_var))

    return root


# Initialize the Tkinter window
root = create_ui()
root.mainloop()
