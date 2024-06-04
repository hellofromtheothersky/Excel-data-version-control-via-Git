# Excel-data-version-control-via-Git
A tool for spreadsheet version control in Excel, created by Python and using with Git

## General introduction: 
In short, the solution is to making every part of Excel which is a binary file become text that is readable, changeable via text editor so that Git can check diff

We keep track of changes on Excel via text data. 

All the rows, for each sheet, in each Excel file, will be parsed into a single pair of file, including
- *.csv: store value
- *.json: store style

![image](https://github.com/hellofromtheothersky/Excel-data-version-control-via-Git/assets/84280247/3d2aad3b-9967-43c1-bc2e-c66afa230f04)

Row-oriented management make managing every single row indepedently. That mean, whenever changes have made, only files represent changed row are marked by Git to show. To check diff, instead of viewing diff on a large file, we are now focusing some files. 

And defitely, not only Excel is parsed into text, but also text can be parsed to generate Excel too. This helps solving conflict of data by modifying from the text to make change on Excel (no need to mannully finding conflict part in the Excel)

![image](https://github.com/hellofromtheothersky/Excel-data-version-control-via-Git/assets/84280247/9d22e693-4351-46b2-aaf1-d44059f80f4a)

## Setup process: 




