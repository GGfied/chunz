import os

from shared.constants import FILE_DIR


def write_error(directory=FILE_DIR, error='', exception=None):
    errormsg = '{}_{}\r\n'.format(error, '' if exception is None else str(exception))

    with open(os.path.join(directory, 'error.txt'), 'a') as f:
        f.write(errormsg)


def write_debug(directory=FILE_DIR, msg='', exception=None):
    debugmsg = '{}_{}\r\n'.format(msg, '' if exception is None else str(exception))

    with open(os.path.join(directory, 'debug.txt'), 'a') as f:
        f.write(debugmsg)


def write_details(directory=FILE_DIR, title='', url=''):
    with open(os.path.join(directory, 'details.txt'), 'a') as f:
        f.write(title + '\r\n')
        f.write(url + '\r\n')
