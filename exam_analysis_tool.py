'''
Author: Abiy Fanta
Date: 2024-11-06
File name: exam_analysis_tool.py
IDE: PyCharm - Python 3.12

Description: A Python script to analyze exam data from an Excel file and calculate pass statistics
             for each exam type, grouped by months in chronological order, including exam names
             and pass points as separate columns.

Install Required Libraries:
    Open your terminal and type the following command to install the necessary libraries:
    - type "pip install pandas openpyxl tkinter" into the terminal, admin mode might be needed

Run the Script:
    Create an executable file using "pyinstaller --onefile --windowed exam_analysis_tool.py" command in the terminal.
    Make sure the terminal is in the same directory as the .py file.
    You will be left with an exe file in the dist folder.
    Run the .exe file, and the GUI will open up. Follow the instructions.
'''

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def analyze_exam_data(file_path):
    # Load the data
    df = pd.read_excel(file_path)

    # Convert 'exam_date' to datetime format and extract the year and month
    df['exam_date'] = pd.to_datetime(df['exam_date'])
    df['year_month'] = df['exam_date'].dt.strftime('%y-%b')  # Format as "YY-MMM" (e.g., "18-Jun")

    # Separate valid exams only (ignore exams with score 0)
    valid_exams = df[df['result_percentage'] != 0]

    # Determine pass or fail based on pass point percentage
    valid_exams['passed'] = valid_exams['result_percentage'] >= valid_exams['pass_point_percentage']

    # Create a combined column for exam name and pass point (e.g., "A-66")
    valid_exams['exam_type'] = valid_exams['exam_name'] + '-' + valid_exams['pass_point_percentage'].astype(str)

    # Group by year_month and exam_type, calculate pass rates
    pass_rates = valid_exams.groupby(['year_month', 'exam_type']).agg(
        total_exams=('result_percentage', 'size'),
        passed_exams=('passed', 'sum')
    )

    # Calculate pass rate as a percentage
    pass_rates['Pass Rate (%)'] = (pass_rates['passed_exams'] / pass_rates['total_exams']) * 100

    # Reshape the table to have exam types as columns and months as rows
    pass_rate_pivot = pass_rates['Pass Rate (%)'].unstack(level='exam_type').round(2)

    # Reorder the rows by converting year_month to datetime and sorting
    pass_rate_pivot.index = pd.to_datetime(pass_rate_pivot.index, format='%y-%b')
    pass_rate_pivot = pass_rate_pivot.sort_index()

    # Convert the datetime index back to the desired "YY-MMM" format
    pass_rate_pivot.index = pass_rate_pivot.index.strftime('%y-%b')

    # Save the output to a new Excel file
    output_file_path = "monthly_pass_rates_by_exam_type.xlsx"
    pass_rate_pivot.to_excel(output_file_path, index_label="Month")

    messagebox.showinfo("Success", f"Analysis complete! Output saved to {output_file_path}")


def open_file_dialog():
    file_path = filedialog.askopenfilename(
        title="Select Exam Data File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        analyze_exam_data(file_path)


# Set up the GUI
root = tk.Tk()
root.title("Exam Pass Rate Analysis Tool")
root.geometry("300x150")

# Button to open file dialog
button = tk.Button(root, text="Select Excel File", command=open_file_dialog)
button.pack(pady=50)

# Run the application
root.mainloop()
