# psalpsdtools
![LPSDLogo_sm](https://github.com/tondiaz/psalpsdtools/assets/3798545/643ce509-132b-47ad-b803-d75a1ffb421a)

Introducing **_psalpsdtools_**, a file maintenance package Python tool designed specifically for the Livestock and Poultry Statistics Division to streamline file management and updating processes. This comprehensive package provides a user-friendly interface and a robust set of functionalities to efficiently organize, manipulate, and validate data files. From data cleaning and merging to filtering and report generation, this package offers a reliable solution to enhance productivity and ensure only accurate information are produced. Furthermore, this package is continuously evolving with ongoing development, promising future enhancements and additional functionalities to cater to the evolving needs of the division.

# Features

Some of the features include:

####  Electronic Data Review Worksheet (EDRW)
- Looking-up and copying of values from the Supply-Disposition file
- Pasting copied values to the EDRW output file
- Creating output files by province, based on user specified inputs including:

  	- region
	- commodity
	- year

#### Built-in Functions
- _get_regions_ - returns a list of the regions.

- _get_provinces_ - returns a list of provinces of a given region.

# Requirements

Python 3.8 or later with all [requirements.txt](https://github.com/tondiaz/psalpsdtools/blob/main/docs/requirements.txt) dependencies installed. To install run:

```
pip install psalpsdtools
```
# Usage

#### EDRW Updating

#### - Pre-requisites (for Chicken)
- _baseFolder_ should contain the S-D file.
- inside _baseFolder_, a folder named _Sources_ must exist, containing the regional folders and the provincial files inside.
  
	![Screenshot 2023-06-28 171534](https://github.com/tondiaz/psalpsdtools/assets/3798545/711bc2dc-e45a-413d-9551-d064e1e73d46)

- Provincial EDRW source files should have .xlsm extensions
- Currently, source filename is expected to be "_cc_ " + _province name_ + "__year_", e.g. _08 Agusan del Norte_23.xlsm_

#### - Example code:
 
```
from psalpsdtools import Edrw

# Specify Region
regName = 'Caraga'

# Specify quarter
# Used in identifying which worksheet to paste the copied values from the S-D file.
qtr = 'Q1'

# Specify folder location of Sources and Final files.
# The S-D file should also be found here.
baseFolder = 'D:/EDRW/Q1'

# Specify S-D filename
# IMPORTANT! Only .xlsm or .xlsx extensions are accepted
sdFile = 'SD Q1 2023.xlsm'

# Commodity code i.e. 08=chicken, 09=duck, etc.
commcode = '08'

# Year
yr = '23'

# Call an instance of the Edrw package
myedrw = Edrw()

# Run update_sources with the parameters
myedrw.update_sources(regName,qtr,baseFolder,sdFile,commcode,yr)
```

#### _get_regions_
#### - Example code:
 
```
from psalpsdtools import Edrw

```

# Contribute

Issue Tracker: [github.com/psalpsdtools/psalpsdtools/issues](github.com/psalpsdtools/psalpsdtools/issues)

Source Code: [github.com/psalpsdtools/psalpsdtools](github.com/psalpsdtools/psalpsdtools)

# License

The project is licensed under the MIT license.
