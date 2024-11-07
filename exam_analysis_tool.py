'''
Author: Abiy Fanta
Date: 2024-11-06
File name: exam_analysis_tool.py
IDE: PyCharm - Python 3.12

Description: A Python script to analyze exam data from an Excel file and calculate yearly pass and fail statistics.

Install Required Libraries:
    Open your terminal and type the following command to install the necessary libraries:
    - type "pip install pandas openpyxl pyinstaller" into the terminal, admin mode might be needed

Run the Script:
    Create an executable file using "pyinstaller --onefile --windowed exam_analysis_tool.py" command in the terminal
    Make sure the terminal is in the same directory as the .py file
    You will be left with an exe file in the dist folder.
    Run the .exe file and the GUI will open up and follow the instructions.
'''

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def analyze_exam_data(file_path):
    # Load the data
    df = pd.read_excel(file_path)

    # Convert 'exam_date' to datetime format and extract the year
    df['exam_date'] = pd.to_datetime(df['exam_date'])
    df['year'] = df['exam_date'].dt.year

    # Separate the data for missed and valid exams
    missed_exams = df[df['result_percentage'] == 0]
    valid_exams = df[df['result_percentage'] != 0]

    # Count missed exams per year and per exam
    missed_exams_count = missed_exams.groupby(['year', 'exam_name']).size().rename("Missed Exams Count")

    # Determine passed and failed exams for valid exams only
    valid_exams['passed'] = valid_exams['result_percentage'] >= valid_exams['pass_point_percentage']
    valid_exams['failed'] = valid_exams['result_percentage'] < valid_exams['pass_point_percentage']

    # Group by year and exam_name to calculate statistics
    yearly_exam_stats = valid_exams.groupby(['year', 'exam_name']).agg(
        total_exams=('result_percentage', 'size'),
        passed_exams=('passed', 'sum'),
        failed_exams=('failed', 'sum'),
        avg_pass_point=('pass_point_percentage', 'mean'),
        avg_score=('result_percentage', 'mean')
    )

    # Calculate pass and fail rates
    yearly_exam_stats['Pass Rate (%)'] = (yearly_exam_stats['passed_exams'] / yearly_exam_stats['total_exams']) * 100
    yearly_exam_stats['Fail Rate (%)'] = (yearly_exam_stats['failed_exams'] / yearly_exam_stats['total_exams']) * 100

    # Add missed exams count to the yearly stats by joining on year and exam_name
    yearly_exam_stats = yearly_exam_stats.join(missed_exams_count, how='left').fillna(0)

    # Limit decimal places to 2 for each numeric column
    yearly_exam_stats = yearly_exam_stats.round(2)

    # Select only the required columns
    output_df = yearly_exam_stats[['total_exams', 'Missed Exams Count', 'avg_pass_point', 'avg_score', 'Pass Rate (%)', 'Fail Rate (%)']]

    # Rename columns for clarity
    output_df.rename(columns={
        'total_exams': 'Total Exams Taken',
        'avg_pass_point': 'Average Pass Point',
        'avg_score': 'Average Score'
    }, inplace=True)

    # Save the results to an Excel file
    output_file_path = "exam_analysis_results.xlsx"
    output_df.to_excel(output_file_path, index_label=["Year", "Exam Name"])

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
root.title("Exam Data Analysis Tool")
root.geometry("300x150")

# Button to open file dialog
button = tk.Button(root, text="Select Excel File", command=open_file_dialog)
button.pack(pady=50)

# Run the application
root.mainloop()