"""Последовательный запуск отчётов проекта"""

import subprocess
import datetime
import time

from service import BASE_DIR

start_time = time.time()

scripts = [
    'archive',
    'remains',
    'price',
    'factors',
    'registry_factors'
]

for script in scripts:
    subprocess.Popen([
        'python.exe', f'{BASE_DIR}/scripts/base_scripts/{script}.py'
    ], shell=True).wait()

flag = True
while flag:
    now = datetime.datetime.now()
    if now.hour == 9 and now.minute >= 10 or now.hour > 9:
        subprocess.Popen(
            ['python.exe', f'{BASE_DIR}/scripts/base_scripts/reserve.py'],
            shell=True
        ).wait()
        flag = False
    else:
        time.sleep(20)

end_time = time.time()
elapsed_time = end_time - start_time
min, sec = divmod(elapsed_time, 60)
print('Время выполнения: {:.0f} минут {:.2f} секунд'.format(min, sec))
