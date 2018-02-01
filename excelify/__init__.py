import sys

__version__ = '0.1'

if sys.version_info >= (3, 0):
    from excelify.excelify import *
else:
    from excelify import *

__all__ = ['excelify']