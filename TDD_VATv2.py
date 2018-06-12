import sys
import time
import os
from tabula import wrapper
import pandas as pd
import string
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from PyQt5 import uic


#note_to_myself - insert os.path.join, so it can be used on windows
#directories and file locations - currently hardcoded
directory_link_2016 = r"C:\Users\lvpool\Documents\TDD_VAT_Return_Scanner\2016"
directory_link_2017 = r"C:\Users\lvpool\Documents\TDD_VAT_Return_Scanner\2017"
directory_link_2018 = r"C:\Users\lvpool\Documents\TDD_VAT_Return_Scanner\2018"
results_directory= r"C:\Users\lvpool\Documents\TDD_VAT_Return_Scanner\Results\Results_v8.xlsx"
file_names_list_2016 = os.listdir(directory_link_2016)
file_names_list_2017 = os.listdir(directory_link_2017)
file_names_list_2018 = os.listdir(directory_link_2018)
isDone = False


def vat_return_reader(file_name):
    """vat_return_reader(file_name)
    It creates a dataframe where key is file name and values are last column(2) 
    of the VAT return table. In dataframe index name is file_name which function used"""
    

    df = wrapper.read_pdf(file_name,pandas_options={'header':None},encoding="cp775")
    print("tabula_read executed")
    column_data = df.loc[0:45,2]
    column_name = df.loc[0:45,1]
    #replaces '.' with ',' in order to be treated as nummeric
    column_data = column_data.astype(str).str.replace(',','.')
    #converts str to float
    column_data = pd.to_numeric(column_data, errors='coerce')
    # renames row index as file names
    column_data = column_data.rename(index=f"{file_name}")
    return column_data


def whole_data(arg_file_names_list):
    """Function gathers all data that is given by function PDF_1st_Page_Scanner(file_name))
    in a specific dataframe that is given as argument (given_files_list).
    Function also adds totalling row at the bottom of dataframe"""

    whole_data_list = []
    # loop which iterates through the directory and executes function vat_return_reader for each file
    for file_name in arg_file_names_list:
        whole_data_list.append(vat_return_reader(file_name))
    # creates data frame from list
    df = pd.DataFrame(whole_data_list)
    # adds total row 
    df = pd.concat([df,pd.DataFrame(df.sum(axis=0),columns=['Total']).T])
    return df

def main():
    """ Function used to write a data frame to excel from file names list by using whole_data"""

    # runs function to create df in order for it to be writen to excel
    os.chdir(directory_link_2016)
    return_data_2016 = whole_data(file_names_list_2016)
    os.chdir(directory_link_2017)
    return_data_2017 = whole_data(file_names_list_2017)
    os.chdir(directory_link_2018)
    return_data_2018 = whole_data(file_names_list_2018)
    # creates excel to be writen
    writer = pd.ExcelWriter(results_directory)
    # writes excel file with df 
    return_data_2016.to_excel(writer,'2016')
    return_data_2017.to_excel(writer,'2017')
    return_data_2018.to_excel(writer,'2018')
    writer.save()

# Assigns locations for GUI QT Designer elements QT Designer element locations are hard-coded
uifile = r'C:\Users\lvpool\Documents\Python\TDD VAT Return Scanner\Process Confirmation.ui'
uifile1 = r'C:\Users\lvpool\Documents\Python\TDD VAT Return Scanner\Process Init.ui'
form, base = uic.loadUiType(uifile)
form1, base1 = uic.loadUiType(uifile1)

class Process_Confirmation(base, form):
    """ Class that creates Confirmation GUI. Informs how many files were processed
    One option - Accpet and stop process"""

    def __init__(self):
        super(base,self).__init__()
        self.setupUi(self)
        self.textBrowser.setText(f'{file_count} VAT Returns were processed \nYou can find the file here: {results_directory}')
        self.buttonBox.clicked.connect(self.buttonOK_slot)
        
    def buttonOK_slot(self):
        print('Going to Quit Seriously!')
        sys.exit(app.exec_())

class Process_Initiation(base1, form1):
    """QT Designer GUI to intiate the process and check whether all files are located in the correct directory
    two options 1st - Accept - start process, 2nd Cancel - abort"""    

    def __init__(self):
        super(base1,self).__init__()
        self.setupUi(self)
        self.textBrowserInit.setText(f'Press Initiate to start with VAT Return processing \nFiles should be located according to the Financial Year. We are working for programm to recognize the year ')
        self.InitateButton.clicked.connect(self.accept)
        self.CancelButton.clicked.connect(self.reject)   
        
    def accept(self):
        print('Scanning Started')
        main()
        self.hide()
        mpage1.show()


    def reject(self):
        sys.exit('process aborted') 


if __name__ == '__main__':
    print('Starting process')
    app = QApplication(sys.argv)
    form1, base1 = uic.loadUiType(uifile1)
    file_count = len(file_names_list_2016+file_names_list_2017+file_names_list_2018)
    mpage1 = Process_Confirmation()
    mpage = Process_Initiation()
    mpage.show()
    sys.exit(app.exec_())

