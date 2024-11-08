# Exam-Analysis
___
 
## Description
 It takes in an excel sheet containing data related to examination results and grading.

 It outputs another excel sheet that compiles this data in fewer cells with data such as:
 - Year Taken
 - Total Exams Taken/Missed
 - Pass/Fail Rate
 - Average Score

## Features
Used TKinter to make a very simple UI. Opens a window with a single button prompting to select an excel file.

## Installation
Any Python3 veresion should work but I used Python 3.12. I used the Pycharm integraded teriminal.

Open a terminal in the same directory as the python file and type in:
```
pyinstaller --onefile --windowed exam_analysis_tool.py
```
- This will search for the .py file and generate a .exe file within the dist folder for you to directly run.

## Contributions
- pandas library ~ https://pandas.pydata.org/