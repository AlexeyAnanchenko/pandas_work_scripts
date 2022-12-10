"""Последлвательный запуск отчётов проекта"""

import subprocess

from service import BASE_DIR

scripts = [
    'assortment',
    'price',
    'directory',
    'holdings',
    'purchases',
    'remains',
    'reserve',
    'sales',
]


for script in scripts:
    subprocess.Popen([
        'python.exe', f'{BASE_DIR}/scripts/base_scripts/{script}.py'
    ], shell=True).wait()
