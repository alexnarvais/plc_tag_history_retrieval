from datetime import datetime
from plc_history import PlcHistory

if __name__ == '__main__':
    """ Version Control

    Author: Alex Narvais
    Version: 1
    Date: 05-Oct-20222
    ECO: 2022-012
    Change: Initial code.
    """

    """ File Path Details
    Specify the location of the history excel file.
    ***BE SURE TO CHANGE THE CELL IN THE EXCEL FILE THAT CONTAINS THE IPV4 ADDRESS BEFORE RUNNING MAIN***
    """
    file_path = r"C:\Users\alexnarvais\Desktop\oakgrove_300_history.xlsx"

    """Tag Name Row and Columns 
    
    These parameters are required for the plc_tag_names() function. The parameters will be 
    different for each PLC because the number of tags are different across each plc. 
    The min_row and max_row values should be the same since there will be only one row of tag names.
    The min_col value (start tag) shouldn't change and max_col value (end tag) will always be changed based on the PLC.
    """

    min_row = 12
    max_row = 12
    min_col = 2
    max_col = 61

    """PlcHistory Class
    
    If you want to see all worksheets in the excel spreadsheet to find the correct name of the worksheet, 
    uncomment the line 39. 
    Line 41 utilizes a class property call work_sheet that'll be set to the first worksheet in the list of worksheets.
    """

    ogw_history = PlcHistory(file_path)
    # print(ogw_history.work_sheet)
    ogw_history.work_sheet = ogw_history.work_sheet[0]
    start_time = datetime.now()
    ogw_tag_names_list = ogw_history.plc_tag_names(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)
    ogw_tag_values_list = ogw_history.plc_tag_values(ogw_tag_names_list)
    excel_write = ogw_history.write_tag_values_wb(ogw_tag_values_list)
    end_time = datetime.now()
    ogw_history.write_exec_time_wb(start_time, end_time)
