from cx_Freeze import setup, Executable

setup(name = "xlreader" ,
      includes = ['openpyxl'],
      version = "0.1" ,
      description = "Custom Excel parser. Converts a datasheet to a fractionally more competent datasheet." ,
      executables = [Executable("xlreader.py")])