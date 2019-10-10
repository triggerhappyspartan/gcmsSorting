#alignment.py
#Written by Brian Andersen 9/9/2019
#Contact @ bdander3@ncsu.edu

import os
import sys
import yaml
import pandas
import argparse
import openpyxl
import math
import copy

def main(file_):
    """
    Main function for aligning the volatile data from the GCMS.
    
    Parameter
    ----------------
    file_: string
    Yaml file that contains the excel files and the sheet numbers to open. 
    Yaml stands for Yeah, aint markup language.
    """
    with open(file_,'r') as yaml_file:
        yaml_dick = yaml.safe_load(yaml_file) #Dick means dictionary
                                              #Dictionary is a datatype within Python

    #This block of code checks that the input yaml file has everything in it that is needed
    #to correctly sort the volatiles.
    if 'Excel' in yaml_dick:                                          
        pass
    else:
        raise ValueError("No key 'Excel' within yaml input file.")

    if 'Sheet' in yaml_dick:
        pass
    else:
        raise ValueError("No key 'Sheet' within yaml input file.")

    if 'Markers' in yaml_dick:
        pass
    else:
        raise ValueError("No key 'Markers' within yaml input file.")

    if 'Group' in yaml_dick:
        pass
    else:
        raise ValueError("No key 'Group' within yaml input file")

    #This block of code reads all the excel sheets listed in the yaml input file
    #and stores the data recorded in those sheets accordingly. Each sample is recorded as an instance 
    #of the Sample class, and all classes are catalogued within the code in a dictionary.
    volatiles = {}
    for sheet in yaml_dick['Sheet']:
        volatiles[sheet] = {}
        excel = pandas.read_excel(yaml_dick['Excel'],sheet)
        columns = excel.columns
        for col in columns:
            if col in yaml_dick['Markers']:
                marker = col
                volatiles[sheet][marker] = Sample()
            for i,cell in enumerate(excel[col]):
                if not cell:
                    pass
                else:
                    if type(cell) == str:
                        if cell == "RT (min)":
                            volatiles[sheet][marker].RT.extend(excel[col][(i+1):])
                        elif cell == "Area (Ab*s)":
                            volatiles[sheet][marker].area.extend(excel[col][(i+1):])
                        elif cell == "Hit Name":
                            volatiles[sheet][marker].hit.extend(excel[col][(i+1):])
                        elif cell == "Quality":
                            volatiles[sheet][marker].quality.extend(excel[col][(i+1):])
                        elif cell == "Normalization":
                            volatiles[sheet][marker].normalization.extend(excel[col][(i+1):])
                        break

    #Determines the maximum and minimum values of the parameter RT. The code uses the RT values
    #to align the different volatiles recorded.
    minimum_rt = 1000000
    maximum_rt = 0
    for sheet in volatiles:
        for marker in volatiles[sheet]:
            if min(volatiles[sheet][marker].RT) < minimum_rt:
                minimum_rt = min(volatiles[sheet][marker].RT)
            if max(volatiles[sheet][marker].RT) > maximum_rt:
                maximum_rt = max(volatiles[sheet][marker].RT)

    #This for loop is where the alignment actually takes place.
    #The alignment works by setting the current rt value to 0.5 lower than the minimum rt value within the excel sheets.
    #It then increases the RT value by 0.01 each loop. If there is a matching RT value in any of the samples that is within
    # +/- 0.01 of the current RT value, it is considered a match, and written to the excel file.
    excel_file = openpyxl.load_workbook(yaml_dick['Excel'])
    for gru in yaml_dick['Group']:
        volatile_copy = copy.deepcopy(volatiles)
        current_rt_value = minimum_rt - 0.5
        excel_file.create_sheet(gru)
        current_sheet = excel_file[gru]
        column = ord("A")
        line=1
        for sheet in yaml_dick['Group'][gru]:
            for mark in volatiles[sheet]:
                cell = get_cell_value(column,line)
                current_sheet[cell] = mark
                cell = get_cell_value(column,line+1)
                current_sheet[cell] = "RT (min)"
                column += 1
                cell = get_cell_value(column,line+1)
                current_sheet[cell] = "Area (Ab*s)"
                column += 1
                cell = get_cell_value(column,line+1)
                current_sheet[cell] = "Hit Name"
                column += 1
                cell = get_cell_value(column,line+1)
                current_sheet[cell] = "Quality"
                column += 1
                cell = get_cell_value(column,line+1)
                current_sheet[cell] = "Normalization"
                column += 1
                
        line=3

        while current_rt_value < maximum_rt:
            match_found = False
            column = ord("A")
            for sheet in yaml_dick['Group'][gru]:
                for mark in volatiles[sheet]:
                    for i,rt in enumerate(volatile_copy[sheet][mark].RT):
                        if abs(rt - current_rt_value) < 0.01:
                            match_found = True
                            cell1 = get_cell_value(column,line)
                            cell2 = get_cell_value(column+1,line) 
                            cell3 = get_cell_value(column+2,line)
                            cell4 = get_cell_value(column+3,line)
                            cell5 = get_cell_value(column+4,line)     
                     
                            current_sheet[cell1] = volatile_copy[sheet][mark].RT.pop(i)
                            if not volatile_copy[sheet][mark].area:
                                pass
                            else:
                                current_sheet[cell2] = volatile_copy[sheet][mark].area.pop(i)
                            if not volatile_copy[sheet][mark].hit:
                                pass
                            else:
                                current_sheet[cell3] = volatile_copy[sheet][mark].hit.pop(i)
                            if not volatile_copy[sheet][mark].quality:
                                pass
                            else:
                                current_sheet[cell4] = volatile_copy[sheet][mark].quality.pop(i)
                            if not volatile_copy[sheet][mark].normalization:
                                pass
                            else:
                                current_sheet[cell5] = volatile_copy[sheet][mark].normalization.pop(i)

                    column += 5
            if match_found:
                line += 1
            current_rt_value += 0.01
            
    excel_file.save(yaml_dick['Excel'])    

def get_cell_value(column_int,line):
    """
    Function that returns the excel alphabetic column value based on the integer value.

    This function will work for excel columns between A and ZZ. Breaks down at AAA. If you have that many values
    rewrite function or make a second spreadsheet.
    """    
    if column_int / ord("[") < 1: #[ is the next value after Z in the ASCII table. 
        col = chr(column_int)     #This is a check whether 1 or two letters are needed to designate cell.
        return "{}{}".format(col,line)
    else:
        temp1 = int(column_int / ord("["))
        temp2 = column_int % ord("[")
        first_letter = ord("A") - 1 + temp1
        second_letter = ord("A") + temp2

        return "{}{}{}".format(chr(first_letter),chr(second_letter),line)


class Sample(object):
    """
    Class for organizing the data of the treatment samples.
    """
    def __init__(self):
        self.RT = []
        self.area = []
        self.hit = []
        self.quality = []
        self.normalization = []

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", help="The yaml file containing the excel file information you want to read.",required=True,type=str)
    args = parser.parse_args()
    main(args.file)







