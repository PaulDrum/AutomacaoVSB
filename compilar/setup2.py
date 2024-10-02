from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': ['matplotlib', 'PIL', 'numpy']}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('6telas.py', base=base, target_name = 'Telas BI e DDMX', icon='vallourec.ico')
]

setup(name='Telas BI e DDMX',
      version = '0.1',
      description = 'Automatização Relatórios',
      options = {'build_exe': build_options},
      executables = executables)