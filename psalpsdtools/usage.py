#from psalpsdtools import Edrw  # USE IF FROM LIBRARY
from Edrw import Edrw           # USE IF FROM LOCAL

# SET REGION TO PROCESS
regName = "All" # USE All for all regions or use specific region name
# SET QUARTER, ALSO USED IN READING THE S-D FILE
qtr = "Q2"
# SET FOLDER OF SOURCES/FINAL FILES INCLUDING S-D FILE
baseFolder = "C:/EDRW/Q2"
# SET S-D FILE NAME
sdFile = 'Chicken Egg_SD Q2 2023.xlsm'
# COMMODITY 08=chicken, 09=duck, etc.
commcode = '08a'
# Year
yr = '23'

myedrw = Edrw()
myedrw.update_sources(regName,qtr,baseFolder,sdFile,commcode,yr)