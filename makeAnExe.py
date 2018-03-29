from cx_Freeze import *
import os

build_exe_options = {"packages": ["os", "sys"],"include_files": ["tcl86t.dll", "tk86t.dll"], "includes": ["pandas","tkinter", "easygui", "numpy", "xlrd", "xlwt", "xlrd", "math", "smtplib", "string", "xlsxwriter", "email.mime.multipart", "email.mime.text"]}

os.environ['TCL_LIBRARY'] = "C:\\Users\\patri\\AppData\\Local\\Programs\\Python\\Python36-32\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = "C:\\Users\\patri\\AppData\\Local\\Programs\\Python\\Python36-32\\tcl\\tk8.6"

#os.environ['TK_LIBRARY'] = "C:\\LOCAL_TO_PYTHON\\Python35-32\\tcl\\tk8.6"

setup(name = "inputy" ,
      version = "0.1" ,
      description = "" ,
      options = {"build_exe": build_exe_options},
      executables = [Executable("inputy.py")])
