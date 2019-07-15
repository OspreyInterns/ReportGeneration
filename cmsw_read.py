
# reads files to find the correct name for the cmsw


def cmsw_id_read(file_name):

    cmsw = file_name
    cmsw = cmsw[:-20]  # slices off the _CMSWdatabase.sqlite from the end of the file
    cursor = cmsw[-1]
    iterate = -2
    serial = []

    while cursor != '/':
        serial.insert(0, cursor)
        cursor = cmsw[iterate]
        iterate = iterate -1

    serial = ''.join(serial)
    return serial
