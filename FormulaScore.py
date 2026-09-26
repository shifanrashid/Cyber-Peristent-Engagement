import openpyxl
import os
import shutil
import pandas as pd

######################################## CHANGE ME ######################################################

#SET DIRECTORY NAMES
#must be subdirectories of the main directory the program resides in

#INPUT DIRECTORY: for the directory that holds the formatted excel sheets to code:
input_directory = 'Templated'

#OUTPUT DIRECTORY: for the directory to place the newly-coded excel sheets:
output_directory = 'Coded'

#SET DICTIONARY NAME: for the excel workbook containing the dictionary of words for each category:
dictionary_name = './Dictionaries/Final_Dictionary.xlsx'

#######################################################################################################

#global variables
deter_list = []
norm_list = []
pe_list = []

#FUNCTION DEFINITIONS
#checks a given word against a dictionary list and codes that row, returning frequency if a match is found
def code_cell(row, col, number):
    #set flag for detecting a partial match to false
    partial_word_flag = False
    #identify cells
    word_cell = ws.cell(row=row, column=col)
    code_cell = ws.cell(row=row, column=col+2)
    comment_cell = ws.cell(row=row, column=col+3)

    #determine which category list to use by input number
    if number == 1:
        list = deter_list
    elif number == 2:
        list = norm_list
    elif number == 3:
        list = pe_list
    
    #for each entry in the dictionary, check it against the given word
    for item in list:
            if (item == str(word_cell.value)):
                #for an exact match, code number & color parameter and return positive
                code_cell.value = number
                return 1
            elif (item in str(word_cell.value)):
                #checks if the dictionary word/phrase is in the cell
                #and if it's not the first character,
                #that the previous character is not a letter (word inside a word)
                if (word_cell.value.index(item) == 0) or (not(word_cell.value[(word_cell.value.index(item))-1].isalpha())):
                    #for a partial match, code number & color and set partial match flag to true
                    code_cell.value = number
                    partial_word_flag = True
                    partial_word_index = list.index(item)
    
    #for a partial match flag, mark the comment column in orange and return positive
    if partial_word_flag:
        comment_cell.value = "from '" + list[partial_word_index] + "'"
        return 1

#code and calculate for a given column of words/phrases based on format
def code_column(col):
    #set variables
    curr_row = 2

    #iterate through phrases - each for loop checks if the phrase contains a key word from the list
    #if it does, it adds the frequency to the frequency total for its category, and codes the appropriate column with a number and colour
    while (ws.cell(row=curr_row, column=col).value != None):
        #fill with default 0
        ws.cell(row=curr_row, column=col+2).value = 0

        #set count of repeat assigned categories to 0, then run through the three categories
        repeat_categories = 0

        for i in range(0,3):
            if code_cell(curr_row, col, i+1) == 1:
                repeat_categories = repeat_categories + 1

        #code appropriate number & colors, adding the frquency number to appropriate total
        code_cell(curr_row, col, 1)
        code_cell(curr_row, col, 2)
        code_cell(curr_row, col, 3)

        #if the cell contains words from multiple categories (i.e. repeat_categories > 1), give warning
        if repeat_categories > 1: 
            ws.cell(row=curr_row, column = col+3).value = "contains CONFLICTING dictionary categories"

        #increment the row
        curr_row = curr_row + 1
    
#fills in proper cell highlights and prints formulas for calculation block
def color_and_calculate(col):
    '''
        curr_row - row to iterate through
        oldest_deter, norm, pe - first row, where category starts
        newest_deter, norm, pe - last row, where category ends
        total_count - sum of all frequency cells in col section
    '''
    #set variables
    curr_row = 2
    oldest_deter = 1000
    oldest_norm = 1000
    oldest_pe = 1000
    newest_deter = 0
    newest_norm = 0
    newest_pe = 0
    total_count = 0
    total_deter = 0
    total_norm = 0
    total_pe = 0

    #determine appropriate letter for freq. column by col
    if col == 1:
        alpha = 'B'
    elif col == 10:
        alpha = 'K'
    elif col == 19:
        alpha = 'T'

    #iterate through rows and assign appropriate highlights
    #assign oldest and newest for each category to determine formula later
    while(ws.cell(row=curr_row, column=col+2).value != None):
        if (ws.cell(row=curr_row, column=col+2).value == 0):
            ws.cell(row=curr_row, column=col+2).fill = openpyxl.styles.fills.PatternFill(fill_type=None)
        elif (ws.cell(row=curr_row, column=col+2).value == 1):
            ws.cell(row=curr_row, column=col+2).fill = openpyxl.styles.fills.PatternFill(start_color="E6FFCC", end_color="E6FFCC", fill_type="solid")
            if (curr_row < oldest_deter):
                oldest_deter = curr_row
                total_deter = alpha + str(oldest_deter)
            newest_deter = curr_row
        elif (ws.cell(row=curr_row, column=col+2).value == 2):
            ws.cell(row=curr_row, column=col+2).fill = openpyxl.styles.fills.PatternFill(start_color="99CCFF", end_color="99CCFF", fill_type="solid")
            if (curr_row < oldest_norm):
                oldest_norm = curr_row
                total_norm = alpha + str(oldest_norm)
            newest_norm = curr_row
        elif (ws.cell(row=curr_row, column=col+2).value == 3):
            ws.cell(row=curr_row, column=col+2).fill = openpyxl.styles.fills.PatternFill(start_color="B3B3FF", end_color="B3B3FF", fill_type="solid")
            if (curr_row < oldest_pe):
                oldest_pe = curr_row
                total_pe = alpha + str(oldest_pe)
            newest_pe = curr_row
        else:
            print("coloring error")

        #Check if there is a special warning in the comment column, and color code
        if(isinstance(ws.cell(row=curr_row, column=col+3).value, str)):
            if (ws.cell(row=curr_row, column=col+3).value.startswith("from")):
                ws.cell(row=curr_row, column=col+3).fill = openpyxl.styles.fills.PatternFill(start_color="FFE6B3", end_color="FFE6B3", fill_type="solid")
            elif (ws.cell(row=curr_row, column=col+3).value.startswith("contains")):
                ws.cell(row=curr_row, column=col+3).fill = openpyxl.styles.fills.PatternFill(start_color="FF8080", end_color="FF8080", fill_type="solid")

        #add to total freq count, increment row
        total_count = total_count + ws.cell(row=curr_row, column=col+1).value
        curr_row = curr_row + 1
    
    if total_deter != 0:
        total_deter = total_deter + ":" + alpha + str(newest_deter)
    if total_norm != 0:
        total_norm = total_norm + ":" + alpha + str(newest_norm)
    if total_pe != 0:
        total_pe = total_pe + ":" + alpha + str(newest_pe)

    #calculate block
    ws.cell(row=2, column=col+6).value = "=SUM(" + str(total_deter) + ") / " + str(total_count)
    ws.cell(row=3, column=col+6).value = "=SUM(" + str(total_norm) + ") / " + str(total_count)
    ws.cell(row=4, column=col+6).value = "=SUM(" + str(total_pe) + ") / " + str(total_count)
    #ws.cell(row=6, column=col+6).value = ws.cell(row=2, column=col+6).value + ws.cell(row=3, column=col+6).value + ws.cell(row=4, column=col+6).value
    
    ws.cell(row=2, column=col+7).value = "=SUM(" + str(total_deter) + ") / SUM(" + str(total_deter) + ", " + str(total_norm) + ", " + str(total_pe) + ")"
    ws.cell(row=3, column=col+7).value = "=SUM(" + str(total_norm) + ") / SUM(" + str(total_deter) + ", " + str(total_norm) + ", " + str(total_pe) + ")"
    ws.cell(row=4, column=col+7).value = "=SUM(" + str(total_pe) + ") / SUM(" + str(total_deter) + ", " + str(total_norm) + ", " + str(total_pe) + ")"
    #ws.cell(row=6, column=col+7).value = ws.cell(row=2, column=col+7).value + ws.cell(row=3, column=col+7).value + ws.cell(row=4, column=col+7).value

#sorts data in column area by descending order of code number
def sort_column(col):
    #set starting row
    total_rows = 2

    #count through to find total no. of rows in col section
    while(ws.cell(row=total_rows, column=col+1).value != None):
        total_rows = total_rows + 1

    #set empty 2d array based on standard section length and total_rows.
    array = [[0 for i in range(4)] for j in range(total_rows-2)]

    #fill in the array with data from the section
    for k in range(total_rows-2):
        for l in range(4):
            array[k][l] = ws.cell(row=k+2, column=l+col).value
    
    try:
        #convert to pandas dataframe and sort by the coding column in ascending order
        df = pd.DataFrame(array)
        df_sorted = df.sort_values(2,ascending=False)

        #iterate through the excel sheet and fill it back from dataframe
        for k in range(total_rows-2):
            for l in range(4):
                ws.cell(row=k+2, column=l+col).value = df_sorted.iat[k,l]
    except:
        return
    
#gathers dictionary reference lists from dictionary excel sheet
def gather_dictionary_lists():
    #open dictionary workbook & sheet
    dictionary = openpyxl.load_workbook(filename = dictionary_name, read_only = True)
    dic_ws = dictionary.active

    #create dictionary lists (deterrence, norms, PE)

    #deterrence
    dic_row = 2
    while (dic_ws.cell(row= dic_row, column=1).value != None):
        deter_list.append(dic_ws.cell(dic_row, column=1).value.upper())
        dic_row = dic_row + 1

    #norms
    dic_row = 2
    while (dic_ws.cell(row= dic_row, column=3).value != None):
        norm_list.append(dic_ws.cell(dic_row, column=3).value.upper())
        dic_row = dic_row + 1

    #PE
    dic_row = 2
    while (dic_ws.cell(row= dic_row, column=5).value != None):
        pe_list.append(dic_ws.cell(dic_row, column=5).value.upper())
        dic_row = dic_row + 1


#MAIN
#create dictionary lists for reference
gather_dictionary_lists()

#iterate through input directory and code each workbook, then save it
for entry in os.scandir(input_directory):
    if entry.name.endswith(".xlsx"):
        wb = openpyxl.load_workbook(filename = entry)
        ws = wb.active

        for i in range(1, 20, 9):
            code_column(i)
            sort_column(i)
            color_and_calculate(i)

        wb.save('new_' + entry.name)

#move the coded workbooks into the output directory
for entry in os.scandir('.'):
    if entry.name.endswith('.xlsx') and entry.name.startswith('new_'):
        try:
            shutil.move(entry, output_directory)
        except NameError:

            print('Name Error, failed to move file.')
