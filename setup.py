import sys
from cx_Freeze import setup, Executable


base = "Win32GUI"

buildOptions = dict(packages= ["sys","sip","os","time","selenium","openpyxl", "io"])
exe = [Executable("mainwindow.py", base = None)]

setup(
    name='Test Application',
    version = '0.1',
    description = "cho",
options= dict(build_exe = buildOptions),
    executables = exe
)
