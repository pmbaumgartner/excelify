"""
IPython magic function to save pandas objects to excel
"""

from time import strftime
from time import time
import datetime

from pandas.core.frame import DataFrame
from pandas.core.series import Series
from pandas import ExcelWriter

from IPython.core.magic import needs_local_scope
from IPython.core.magic_arguments import (argument, magic_arguments, parse_argstring)


@magic_arguments()
@argument('dataframe', help='Dataframe to Save')
@argument('-f', '--filepath', help=u'Filepath to Excel spreadsheet. Default: {object}_{timestamp}.xlsx')
@argument(
    '-s', '--sheetname', type=str, help=u'Sheet name to output data. Default: {object}_{timestamp}')
@needs_local_scope
def excel(string, local_ns=None):
    '''Saves a DataFrame or Series to Excel'''
    args = parse_argstring(excel, string)

    try:
        dataframe = local_ns[args.dataframe]
    except KeyError:
        raise NameError("name '{}' is not defined".format(args.dataframe))

    if not (isinstance(dataframe, DataFrame) or isinstance(dataframe, Series)):
        raise TypeError("Object must be pandas Series or DataFrame. Object passed is: {}".format(
            type(dataframe)))

    if not args.filepath:
        filepath = args.dataframe + "_" + datetime.datetime.now().strftime(
            "%Y%m%d-%H%M%S") + '.xlsx'
    else:
        filepath = args.filepath
        if filepath.find('.xlsx') == -1:
            filepath += '.xlsx'

    if not args.sheetname:
        sheetname = args.dataframe + "_" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    else:
        sheetname = args.sheetname

    writer = ExcelWriter(filepath, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name=sheetname)
    writer.save()

    print("{dataframe} saved to {filepath} on sheet {sheetname}".format(
        dataframe=args.dataframe, filepath=filepath, sheetname=sheetname))


@magic_arguments()
@argument('-f', '--filepath', help=u'Filepath to excel spreadsheet. Default: all_data_{timestamp}.xlsx')
@argument('-n', '--nosort', help=u'Turns off alphabetical sorting of objects for export to sheets')
@needs_local_scope
def excel_all(string, local_ns=None):
    '''
    Saves all Series or DataFrame objects in the namespace to Excel.
    Use at your own peril. Will not allow more than 100 objects.
    '''

    args = parse_argstring(excel_all, string)

    pandas_object = lambda d: isinstance(d, DataFrame) or isinstance(d, Series)

    pandas_objects = [(name, obj) for (name, obj) in local_ns.items() if pandas_object(obj)]

    if len(pandas_objects) == 0:
        raise RuntimeError("No pandas objects in local namespace.")
    if len(pandas_objects) > 100:
        raise RuntimeError("Over 100 pandas objects in local namespace.")

    pandas_objects = sorted(pandas_objects, key=lambda x: x[0])

    if not args.filepath:
        filepath = "all_data_" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + '.xlsx'
    else:
        filepath = args.filepath
        if filepath[-4:] != '.xlsx':
            filepath += '.xlsx'

    writer = ExcelWriter(filepath, engine='xlsxwriter')
    for (name, obj) in pandas_objects:
        obj.to_excel(writer, sheet_name=name)
    writer.save()

    n_objects = len(pandas_objects)

    print("{n_objects} saved to {filepath}".format(n_objects=n_objects, filepath=filepath))


def load_ipython_extension(ipython):
    ipython.register_magic_function(excel)
    ipython.register_magic_function(excel_all)