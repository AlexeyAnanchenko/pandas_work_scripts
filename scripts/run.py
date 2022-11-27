import subprocess


scripts = [
    'assortment',
    'directory',
    'holdings',
    'purchases',
    'remains',
    'reserve',
    'sales',
]

for script in scripts:
    subprocess.Popen([
        'python.exe',
        'C:/Users/ananchenko.as/Desktop/pandas_work_scripts'
        '/scripts/{}.py'.format(script)
    ], shell=True).wait()
