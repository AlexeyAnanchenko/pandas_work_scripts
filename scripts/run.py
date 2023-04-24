"""Последовательный запуск отчётов проекта"""

import subprocess
import time
import os

from service import BASE_DIR

start_time = time.time()
os.environ.setdefault('SRS_DOWNLOAD', 'FALSE')

scripts_base = [
    'assortment',
    'sales_by_date',
    'sales',
    'price',
    'directory',
    'holdings',
    'purchases',
    'remains',
    'reserve',
    'factors'
]

for script in scripts_base:
    subprocess.Popen([
        'python.exe', f'{BASE_DIR}/scripts/base_scripts/{script}.py'
    ], shell=True).wait()

reports = [
    'potential_sales',
    'registry_potential_sales',
    'future_potential_sales',
    'check_factors',
    'not_sold',
    'order_form'
]

for script in reports:
    subprocess.Popen([
        'python.exe', f'{BASE_DIR}/scripts/reports/{script}.py'
    ], shell=True).wait()

os.environ.pop('SRS_DOWNLOAD')

end_time = time.time()
elapsed_time = end_time - start_time
min, sec = divmod(elapsed_time, 60)
print('Время выполнения: {:.0f} минут {:.2f} секунд'.format(min, sec))
