import sys
from cx_Freeze import setup, Executable
# Dependências
build_exe_options = {"packages":["os"],
                    "includes":["tkinter","re","win32com"],
                    "include_files":["Modelo.html","GeradorPapelCarta.png"]
                     }

# Se aplicação windows, insere base diferente
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable('GeradorPapelCarta.py',
               base=base,
               icon='.\GeradorPapelCarta.png',
               target_name='PapelCarta_QD')
]

setup(
    name = "PapelCarta",
    version = "2.16",
    description = "Emissor de Papel de Carta",
    options = {"build_exe": build_exe_options},
    executables = executables
)