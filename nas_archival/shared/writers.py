import os

from shared.constants import FILE_DIR


def write_error(directory=FILE_DIR, error='', exception=None, filename='error.txt'):
    errormsg = '{}{}\r\n'.format(error, '' if exception is None else ', Exception: ' + str(exception))

    with open(os.path.join(directory, filename), 'a') as f:
        # print(errormsg)
        f.write(errormsg)


def write_debug(directory=FILE_DIR, msg='', exception=None, filename='debug.txt'):
    debugmsg = '{}{}\r\n'.format(msg, '' if exception is None else ', Exception: ' + str(exception))

    with open(os.path.join(directory, filename), 'a') as f:
        f.write(debugmsg)


def write_details(directory=FILE_DIR, title='', url='', filename='details.txt'):
    with open(os.path.join(directory, filename), 'a') as f:
        f.write(title + '\r\n')
        f.write(url + '\r\n')
