from os import system
from datetime import datetime
from tkinter import filedialog
from plc_history import PlcHistory

if __name__ == '__main__':
    """Version Control

    Version Control through GitHub, see link below to repository.
    https://github.com/Site-Automation/plc_tag_history_retrieval.git
    """

    file_path = filedialog.askopenfilename()

    """PlcHistory Class

    If you want to see all worksheets in the excel spreadsheet to find the correct name of the worksheet, 
    uncomment ```print(plc_history.work_sheet)```. The property class member ```PlcHistory.work_sheet``` will be set to the 
    first worksheet from list of worksheets by default.
    """

    plc_history = PlcHistory(file_path)
    # print(plc_history.work_sheet)
    plc_history.work_sheet = plc_history.work_sheet[0]
    start_time = datetime.now()
    plc_tag_names_list = plc_history.plc_tag_names()

    if plc_tag_names_list:
        plc_tag_values_list = plc_history.plc_tag_values(plc_tag_names_list)
        excel_write = plc_history.write_tag_values_wb(plc_tag_values_list)
        end_time = datetime.now()
        plc_history.write_exec_time_wb(start_time, end_time)
        system(f'start EXCEL.EXE "{file_path}"')
    else:
        print("There are no plc tag names defined, please check the configuration file?")
