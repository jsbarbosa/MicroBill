import os
from shutil import copy, copytree

FROM_PATH = 'microbill'
TO_PATH = os.path.join('dist', 'Microbill')

FILES = ['icon.ico', 'logo.png', 'login.py', 'config.py']
FOLDERS = ['Old', 'PDF', 'Registers']

for folder in FOLDERS:
    from_ = os.path.join(FROM_PATH, folder)
    to_ = os.path.join(TO_PATH, folder)

    try: copytree(from_, to_)
    except PermissionError as e: print('PermissionError: ', e)

    print("%s has been copied: %s" % (from_, to_))


for file in FILES:
    from_ = os.path.join(FROM_PATH, file)
    to_ = os.path.join(TO_PATH, file)

    try: copy(from_, TO_PATH)
    except PermissionError as e: print('PermissionError: ', e)

    print("%s has been copied: %s" % (from_, to_))
