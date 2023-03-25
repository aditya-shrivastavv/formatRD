# formatRD
Code written in python to format reccuring deposit excel file to printable spreadsheets, using modules "openpyxl" and "xls2xlsx"

## How it works:
  1. Make a new folder of any name.
  2. Inside that folder paste the two files 'formatrd.py' and 'logo.png'.
  3. now paste the xls files to be formated in the same folder.
    the folder will look like
    
    ```
    formatrd.py
    logo.png
    filetobeformated1.xls
    filetobeformated2.xls
    filetobeformated3.xls
    .
    .
    .
    ```
    
  4. the spreadsheet files which are to be formated must be of extension `.xls`. By default downloaded files from post office have the `.xls`
  5. the file should not be tampered in any way. If you input a half edited file it will be fully disturbed and will not remain of any use.

## How to run

1. Open the folder in any terminal
2. run command `python formatrd.py` in *windows*
3. or run command `python3 formatrd.py` in *mac* or *linux*
