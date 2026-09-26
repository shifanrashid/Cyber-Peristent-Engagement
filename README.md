Cyber Persistent Engagement Python Based Dictionary Scoring Algorithm

Included Programs
FormulaScore: Takes a folder of templated workbooks and codes them according to the categories of deterrence, norms, and persistent engagement, returning a folder of coded workbooks. Supports excel formulas for calculation blocks, making it easier to edit. Libraries: openpyxl, pandas, os, shutil

StrategyScore: Old iteration of words&phrases coder. It does the same thing as FormulaSupport, but without the excel formula support, giving raw numbers instead. It retains the order and format of the original templated workbook, as opposed to the color/code sorting of FormulaSupport. Libraries: openpyxl, os, shutil

MastersheetGenerator: Takes a folder of coded workbooks as an input and creates a master-workbook with all of the calculation block information. Libraries: openpyxl, os, xlwings

Installation
Please check the required libraries for the individual programs you want to use, and pip install them.

Python

Python is required to run all programs. Check if python is already installed on your device by running 'py --version', 'python --version', or 'python3 --version' on your command line, depending on your device.

If not installed, go to Python's website and follow the installation guide.

On Windows:

To use the program, go to the main repository page, click the 'code' dropdown, and download and extract the zip file.
Move the python file to your desired main folder.
Within the main folder, create an input folder to place formatted xlsx workbooks into.
Within the main folder, create an output folder for the program to place coded xlsx workbooks.
Move a formatted dictionary xlsx workbook into the main folder.
Open the python file for editing (right click and select 'Edit in Notepad' or use an IDE of your choice) and change the variable names at the top of the program to correctly match your input folder, output folder, and dictionary workbook names.
You are now ready to move workbooks into the input folder and run the program.
Execution / Usage
FormulaScore/StrategyScore:

Double click the python file or select 'open with Python' to run the program.

Each run of the program iterates through the input folder and generates a new sheet in the output folder for every file it finds there. Every coded file in the output folder will have the 'new_' starting file name to differentiate them from the original workbook. Be aware that when rerunning the program, the old files in the output folder MUST be deleted to prevent an error due to existing file names.

MastersheetGenerator:

Double click the python file or select 'open with Python' to run the program.

xlwings will open and close all of the excel workbooks included in the input directory to get the correct values from the formulas. The output mastersheet will be in the main directory and named 'Mastersheet.xlsx'

Features


Coding column editing for each word

'0' means no match was found with any dictionary category, with no color highlight.
'1' means a match was found with the "Deterrence" category, with a green cell highlight.
'2' means a match was found with the "Norms" category, with a blue cell highlight
'3' means a match was found with the "Persistent Engagement" category, with a purple cell highlight
Comments column editing for each word

An orange highlight and warning means the associated words CONTAIN a dictionary phrase or word, but it is not an exact match.
A red highlight and warning means the associated words contain dictionary phrases or words from DIFFERENT categories. It will only provide the most recent coding in the coding column.
Calculations for partial and relative totals

Added to the corresponding tables on the excel workbooks. StrategyScore has Flat numbers only, with only complete totals sustaining excel formula format. FormulaScore allows for edits to the new workbook with excel formulae.

Comments column editing for each word

- An orange highlight and warning means the associated words CONTAIN a dictionary phrase or word, but it is not an exact match.
- A red highlight and warning means the associated words contain dictionary phrases or words from DIFFERENT categories. It will only provide the most recent coding in the coding column.

Calculations for partial and relative totals

- Added to the corresponding tables on the excel workbooks. StrategyScore has Flat numbers only, with only complete totals sustaining excel formula format. FormulaScore allows for edits to the new workbook with excel formulae.
