# macro4export
This is an easy-to-use Excel macro 4 extractor for Windows.
The tool will output the finding as well as save them into a txt file on your desktop.

###### Pay Attention:
This tool uses openpyxl and supports only **xlsx/xlsm/xltx/xltm** excel formats.
If your file is in xls format, please use xls to xlsx converter(there are plenty of converters online).


## Installation

This tool requires openpyxl, and pandas tools are installed.
You can install openpyxl by running the following command:
```
pip install openpyxl
```
You can install pandas by running the following command:
```
pip install pandas
```

## User Guide
Run the tool with python command like that:
```
python macro4export.py
```
![ScreenShot 1](https://github.com/genzsecurity/macro4export/blob/main/ScreenShot%201.png)

**Don’t forget that you have to be in the tool folder path.**

Then enter files full path and hit enter. 

![ScreenShot 2](https://github.com/genzsecurity/macro4export/blob/main/ScreenShot%202.png)

The output gonna look like that:

![Output](https://github.com/genzsecurity/macro4export/blob/main/ScreenShot%203.png)

You can see the column name and a row number of the macro 4.

The output also saved on desktop:

![File_Created](https://github.com/genzsecurity/macro4export/blob/main/ScreenShot%204.png)
