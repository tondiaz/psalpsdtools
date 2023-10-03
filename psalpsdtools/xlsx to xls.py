from pathlib import Path
import openpyxl
import os
import pandas
 
# get files
os.chdir(os.path.abspath(os.path.dirname(__file__)))
pdir = Path('C:/EDRW/Q2/Sources/07 Central Visayas/07 Cebu')
filelist = [filename for filename in pdir.iterdir() if filename.suffix == '.xlsx']
 
for filename in filelist:
    print(filename.name)
 
for infile in filelist:
    #workbook = openpyxl.load_workbook(infile)
    df = pandas.read_excel(infile, engine='openpyxl')
    outfile = f"{infile.name.split('.')[0]}.xls"
    df.to_excel(outfile)