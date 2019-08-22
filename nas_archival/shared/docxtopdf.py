"""
Source: https://gist.github.com/MichalZalecki/92fd007699004ae7d806274d3a0d5476
"""
import os
import re
import subprocess
import sys

from shared.constants import FILE_DIR


def convert_to(folder, source, timeout=None):
    args = [libreoffice_exec(), '-env:UserInstallation=file://{}'.format(FILE_DIR), '--headless', '--convert-to',
            'pdf:writer_pdf_Export', '--outdir', folder, source]

    process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)


def libreoffice_exec():
    # TODO: Provide support for more platforms
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'


def setup():
    print('Setting up LibreOffice pdf1a export at {}'.format(FILE_DIR))
    args = [libreoffice_exec(), '-env:UserInstallation=file://{}'.format(FILE_DIR), '--headless', '--invisible',
            '--terminate_after_init']
    try:
        subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
    except subprocess.TimeoutExpired as ex:
        print('ERROR: LibreOffice missing or profile not working', ex)
    registry_file = os.path.join(FILE_DIR, 'user', 'registrymodifications.xcu')
    with open(registry_file, 'r') as f:
        lines = f.readlines()
    with open(registry_file, 'w') as f:
        select_pdf_version_1a_re = re.compile(r'"SelectPdfVersion"')
        if len(list(filter(lambda v: re.search(select_pdf_version_1a_re, v), lines))) == 0:
            print('Add SelectPdfVersion to registry')
            select_pdf_version_1a = '<item oor:path="/org.openoffice.Office.Common/Filter/PDF/Export">\
            <prop oor:name="SelectPdfVersion" oor:op="fuse"><value>1</value></prop></item>'
            lines.insert(2, select_pdf_version_1a)

        use_tagged_pdf_re = re.compile(r'"UseTaggedPDF"')
        if len(list(filter(lambda v: re.search(use_tagged_pdf_re, v), lines))) == 0:
            print('Add UseTaggedPDF to registry')
            use_tagged_pdf = '<item oor:path="/org.openoffice.Office.Common/Filter/PDF/Export">\
            <prop oor:name="UseTaggedPDF" oor:op="fuse"><value>true</value></prop></item>'
            lines.insert(2, use_tagged_pdf)

        f.writelines(lines)


class LibreOfficeError(Exception):
    def __init__(self, output):
        self.output = output


if __name__ == '__main__':
    print('Converted to ' + convert_to(sys.argv[1], sys.argv[2]))
