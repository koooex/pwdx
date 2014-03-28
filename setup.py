#!/usr/bin/env python
# encoding: utf-8
#
from __future__ import with_statement

import pkg_resources

from make import main, config

# config.py2exe.packages += ['sqlalchemy']
# config.py2exe.zipfile = "stdlib.dll"
config.py2exe.dll_excludes += ['oci.dll']
config.py2exe.includes += ['xlwt', 'encodings.latin_1', 'encodings.utf_16_le', 'sqlalchemy', 'sqlalchemy.ext.sessioncontext', 'sqlalchemy.databases.mysql', 'MySQLdb']
config.py2exe.includes += ['uuid']
config.py2exe.excludes.remove('optparse')
config.py2exe.excludes += ['bx', 'boot']
config.py2exe.ascii = True
config.py2exe.optimize = 2
config.py2exe.compressed = 1
config.py2exe.bundle_files = 1

config.py2exe.setup.console = ['pwdx.py']
# config.py2exe.setup.service = [
#     {'modules' : ['gateway'], 'cmdline_style' : 'pywin32' }, 
#     {'modules' : ['servant'], 'cmdline_style' : 'pywin32' }
# ]

config.bzlib.bize.name = 'pwdx.dll'
#config.bzlib.bize.dirs.pop(1)
#config.bzlib.bize.skip_files += ['pwdx.py']

config.cython.reset(['pwdx'], 0)

if __name__ == '__main__':
    main()



