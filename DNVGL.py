import os
import xlrd
import pandas
import numpy as np
import openpyxl

'''
Pre: this application requires the following 3rd party modules: xlrd, pandas, numpy, and openpyxl. They must be installed first

Theree are couple of ways of doing that, but I'll stick with 2, which are: pivottable, DataFrame
This application will ask user for a file name as well as file location.
Then application will create 2 Data Frames (I just wanted to show that there are multiple ways of doing this).
Finally those DataFrames will be printed out to console as well as saved to original excel file (old data will be lost)
'''
file = str(input('Enter file name with extention (Default is DNVGL_Python_Exercise_Rev1.xlsx):\n'))
dirr = str(input("Enter full path (location) where file resides (Default is /Users/artemkovtunenko/Documents/):\n[For Mac use '/', Windows '\\']\n"))

file2 = 'DNVGL_Python_Exercise_Rev1.xlsx'
dirr2 = '/Users/artemkovtunenko/Documents/'
file = file if len(file) > 1 else file2
dirr = dirr if len(dirr) > 1 else dirr2
df_setup = None
frames = []


def setup():
    global df_setup
    global file
    global dirr

    os.chdir(dirr)
    try:
        df = pandas.read_excel(file)
    except FileNotFoundError:
        print("File wasn't found")
        return -1
    exclude = ['Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']
    #exclude unnecessary information from excel file
    df.ix[:, df.columns.difference(exclude)]
    #dataframe setup with WD_Bin and WS_Ratio added
    df_setup = pandas.DataFrame(df, columns = ['Timestamp', 'WS1', 'WD1', 'Temp', 'WS2', 'WD_Bin', 'WS_Ratio'])
    #pandas knows how to deal with 0 division
    WS_Ratio_series = (df_setup.WS2 / df_setup.WS1)
    df_setup['WS_Ratio'] = WS_Ratio_series
    #there are a few ways of doing rounding. They was this rounding works is: 0 to 0.4999999 => 0; 5 to 14.99 => 10, etc
    WD_Bin_series = (df_setup['WD1'].round(-1)).apply(int)
    df_setup['WD_Bin'] = WD_Bin_series
    pivotTable()
    dataFrame()

def pivotTable():
    global df_setup
    pivot_table = pandas.pivot_table(df_setup, index=['WD_Bin'], values=['WS_Ratio'], aggfunc=[np.mean, len, np.std], margins=True)
    print(pivot_table)
    print()
    saveBack(pivot_table)
    

def dataFrame():
    global df_setup
    count_WD = df_setup.groupby(['WD_Bin']).count()['WS_Ratio']
    average_WS = df_setup.groupby(['WD_Bin']).mean()['WS_Ratio']
    sd_WS = df_setup.groupby(['WD_Bin']).std()['WS_Ratio']

    count_WD_series = pandas.Series(count_WD, name = 'Count of WS_Ratio')
    average_WS_series = pandas.Series(average_WS, name = 'Average of WS_Ratio')
    sd_WS_series = pandas.Series(sd_WS, name = 'StdDev of WS_Ratio')

    data_frame = pandas.concat([average_WS_series, count_WD_series, sd_WS_series], axis=1)
    print(data_frame)
    saveBack(data_frame)

def saveBack(item):
    global frames
    global file
    frames.append(item)
    if len(frames) < 2:
        #it will wait untill both DataFrames inside 'frames' list
        pass
    else:
        sheet_list = pandas.ExcelFile(file).sheet_names
        sheet1 = 'Sheet3'

        writer = pandas.ExcelWriter(file,engine='openpyxl')
        frames[0].to_excel(writer, sheet_name=sheet1)
        frames[1].to_excel(writer,sheet_name=sheet1, startcol=6, startrow=2)
        writer.save()
    

if __name__ == '__main__':
    setup()

