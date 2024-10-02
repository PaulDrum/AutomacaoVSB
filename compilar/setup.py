from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('central.py', base=base, target_name = 'Gestão de Estoque Automatico', icon='vallourec.ico')
]

setup(name='Gestão de Estoque Automatico',
      version = '0.1',
      description = 'Aplicativo para Geração e Tratamento de Relatórios de Forma Automatizada',
      options = {'build_exe': build_options},
      executables = executables)