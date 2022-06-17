import sys
from cx_Freeze import setup, Executable
# Dependências
build_exe_options = {"packages":["os"], "includes":["tkinter","re","win32com"]}

# Se aplicação windows, insere base diferente
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name = "PapelCarta",
    version = "2.11",
    description = "Emissor de Papel de Carta",
    options = {"build_exe": build_exe_options},
    executables = [Executable("GeradorPapelCarta.py",base=base)]
)