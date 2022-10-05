# Python PLC History Test

___

##  Objective
    This python module is intented to do the following:
        - Open a excel file and read a row of plc tag names.
        - For each tag name reads its entire data array values (2000).
        - Parse the time thats in epoch to a datetime format or round the value for each element in that tag array depending on tag.
        - Use a nested list to build a dataset ,based on the index number per plc tag data array.
        - Write the nest listed (dataset) to an excel work book.
        - Write the time it took to excute the program, to an excel work book.

___

## Version Control
| File Name      | Version Number | Author       | File Date   | Change Control Description |
|----------------|----------------|--------------|-------------|----------------------------|
| plc_history.py | Version 1      | Alex Narvais | 05-Oct-2022 | Initial code.              |
|                |                |              |             |                            |


| File Name | Version Number | Author       | File Date   | Change Control Description |
|-----------|----------------|--------------|-------------|----------------------------|
| main.py   | Version 1      | Alex Narvais | 05-Oct-2022 | Initial code.              |
|           |                |              |             |                            |
 
___

## Python Script
### **Modules**
```python
from openpyxl import load_workbook
from pycomm3 import LogixDriver, PycommError
from datetime import datetime
from plc_history import PlcHistory
```

### **Classes**
```python
PlcHistory() # One required argument.-
```

### **Functions**
```python
plc_tag_names() # Returns a list of plc tag names from the workbook.
plc_tag_values() # Returns a nested list of plc tag values from the workbook.
write_exec_time_wb() # Write the nested list returned by plc_tag_values() function to an Excel spreadsheet.
write_exec_time_wb() # Write the time it took to execute the program to an Excel spreadsheet.
```

___

### Virtual Environment Setup
1. #### Create a new virtual environment in PyCharm.
> 1. File\New Project\Pure Python 
> 2. Specify a project name and the location.
> 3. Select Virtualenv as the new environment.

2. #### Create a new virtual environment using the command prompt.
> 1. Change to the directory where you want creation the virtual environment. `cd 'C:\<>\<>'`
> 2. execute the command `python -m venv new-venv-dir`
> 3. activate the virtual environment `new-ven-dir\Scripts\activate`
> 4. You should now be in the new virtual environment shell with this type of confirmation `(new-venv-dir)`

___


### Program Setup and Execution
> 1. Copy the latest Python Modules and README file from the Site-Automation GitHub Repository. [PLC Tag History Retrieval](https://github.com/Site-Automation/plc_tag_history_retrieval)
> 2. Change the variable names in the main.py file to fit the project and any other recommend changes are documented in the main.py file.
> 3. Run the main.py from the virtual environment shell using the following command `python -m main` or from PyCharm by pressing the run button. 
> 4. If any changes are made to the main.py or plc_history.py files update the README and commit the changes to the GitHub Repository.