from setuptools import setup, find_packages
import os
import excelify

VERSION = excelify.__version__

setup(
    name='excelify',
    version=VERSION,
    license='newBSD',
    description=('IPython magic function to export pandas objects to excel'),
    author='Peter Baumgartner',
    author_email='',
    url='https://github.com/pbaumgartner/excelify',
    packages=find_packages(exclude=[]),
    install_requires=['ipython', 'pandas', 'XlsxWriter'],
    long_description=""""""
)