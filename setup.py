from setuptools import setup, find_packages
import os
import excelify

VERSION = 0.4

setup(
    name='excelify',
    version=VERSION,
    license='MIT',
    description=('IPython magic function to export pandas objects to excel'),
    author='Peter Baumgartner',
    author_email='petermbaumgartner@gmail.com',
    url='https://github.com/pbaumgartner/excelify',
    packages=find_packages(exclude=[]),
    install_requires=['ipython', 'pandas', 'XlsxWriter', 'xlrd'],
    include_package_data=True
)