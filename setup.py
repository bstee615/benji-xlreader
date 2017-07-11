from cx_Freeze import setup, Executable

setup(name = "xlreader" ,
      includes = ['openpyxl', 'PyQt5'],
      version = "0.1" ,
      description = "XLReader" ,
      executables = [Executable("xlreader.py")])