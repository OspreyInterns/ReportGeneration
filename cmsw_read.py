
import os

TEST_FILE = r'C:\Users\zpedersen\Desktop\Case Data\GarbageTestName_CMSWdatabase.sqlite'

# Reads files to find the correct name for the cmsw


def cmsw_id_read(file_name):

    cmsw = file_name
    p, f = os.path.split(cmsw)
    serial = f.lower().split('_')
    # print(serial)
    return serial[0]

# Testing function
# cmsw_id_read(TEST_FILE)
