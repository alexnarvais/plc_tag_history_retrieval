# PLC Tag History Retrieval

___

## Introduction
The purpose of this program is communicate with Allen Bradly Logix PLC and read PLC tag arrays that are being used to store
samples of data for history collection. A excel spreadsheet is used to setup the network configurations and communication 
to the PLC. The spreadsheet is also used to define the PLC tag array names that the python program will use to read the sampled data.
An Excel spreadsheet with the name **tag_history.xlsx** is included in the program root directory and will need to be used to show 
the proper setup that the program expects. The number and name of PLC tags is based on the PLC where the tags are created.
The cell location and the data it contains in the spreadsheet is what's important for successful program excecution. 
## Third Party Python Modules
```python
# plc_history.py
from openpyxl import load_workbook
from pycomm3 import LogixDriver, PycommError

# main.py
from plc_history import PlcHistory
```

        


