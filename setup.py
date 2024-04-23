import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["tkinter", "openpyxl", "os"], "include_files": []}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

exe = Executable(
    script="main.py",
    base=base,
    icon="favicon.ico"  # Dodaj ścieżkę do pliku ikony
)

setup(
    name="MetalMaterial",
    version="1.0",
    description="Metal Material",
    options={"build_exe": build_exe_options},
    executables=[exe]
)
