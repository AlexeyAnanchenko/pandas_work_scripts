"""Обновление отчёта по не проданному товару за все месяца"""

import subprocess
import time
import os

from service import BASE_DIR

start_time = time.time()
os.environ.setdefault('SRS_DOWNLOAD', 'FALSE')

subprocess.Popen(
    ['python.exe', f'{BASE_DIR}/scripts/base_scripts/full_sales.py'],
    shell=True
).wait()

subprocess.Popen([
    'python.exe', f'{BASE_DIR}/scripts/reports/tracking_not_sold.py'],
    shell=True
).wait()

os.environ.pop('SRS_DOWNLOAD')

end_time = time.time()
elapsed_time = end_time - start_time
min, sec = divmod(elapsed_time, 60)
print('Время выполнения: {:.0f} минут {:.2f} секунд'.format(min, sec))
