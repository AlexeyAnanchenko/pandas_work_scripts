import utils
utils.path_append()

import os
import shutil
from datetime import datetime

from settings import ARCHIVE_DIR, RESULT_DIR, TABLE_REGISTRY_FACTORS
from settings import TABLE_REMAINS, TABLE_RESERVE, TABLE_FIXING_FACTORS


table_copy = [
    TABLE_REGISTRY_FACTORS,
    TABLE_RESERVE,
    TABLE_REMAINS,
    TABLE_FIXING_FACTORS
]

current_date = datetime.now().strftime("%d.%m.%Y")
list_dir = os.listdir(ARCHIVE_DIR)

if current_date not in list_dir:
    current_fold = ARCHIVE_DIR + '\\' + current_date
    os.mkdir(current_fold)

    for table in table_copy:
        src_file = RESULT_DIR + table
        mtime = os.path.getmtime(src_file)
        mtime_dt = datetime.fromtimestamp(mtime).strftime("%d.%m.%Y")
        shutil.copy(src_file, current_fold)
        current_name = current_fold + '\\' + table
        new_name = current_fold + '\\' + mtime_dt + ' ' + table
        os.rename(current_name, new_name)

    print(f'Архив на дату {current_date} создан!')
