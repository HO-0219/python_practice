import pandas as pd
import shutil
import os
from datetime import datetime

def backup_excel_files():
    src_dir = os.path.dirname(__file__)
    dst_dir = os.path.join(src_dir, 'backup', datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))

    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)

    files_to_backup = ['Sales.xlsx', 'Inventory.xlsx']
    
    for file in files_to_backup:
        shutil.copy(os.path.join(src_dir, file), dst_dir)
