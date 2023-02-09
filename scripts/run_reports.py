"""Последовательный запуск отчётов проекта"""

import subprocess
import time

from service import BASE_DIR

start_time = time.time()

scripts = [
    'potential_sales',
    'registry_potential_sales',
    'future_potential_sales'
]

for script in scripts:
    subprocess.Popen([
        'python.exe', f'{BASE_DIR}/scripts/reports/{script}.py'
    ], shell=True).wait()

end_time = time.time()
elapsed_time = end_time - start_time
min, sec = divmod(elapsed_time, 60)
print('Время выполнения: {:.0f} минут {:.2f} секунд'.format(min, sec))
