import unittest
import tempfile
import pathlib
import datetime

from IPython.testing.globalipapp import start_ipython, get_ipython

import pandas.util.testing as tm
from pandas.core.frame import DataFrame
from pandas.core.series import Series
from pandas import read_excel

import pytest

ip = get_ipython()
ip.magic('load_ext excelify')

class TestMagicExportImport(unittest.TestCase):

    def setUp(self):
        self.tempexcel = tempfile.NamedTemporaryFile(suffix='.xlsx')

    def test_series(self):
        series = Series()
        excel_name = self.tempexcel.name
        ip.run_line_magic('excel', 'series -f {filepath}'.format(filepath=excel_name))
        loaded_series = read_excel(excel_name, squeeze=True, dtype=series.dtype)
        tm.assert_series_equal(series, loaded_series, check_names=False)

    def test_dataframe(self):
        df = DataFrame()
        excel_name = self.tempexcel.name
        ip.run_line_magic('excel', 'df -f {filepath}'.format(filepath=excel_name))
        loaded_df = read_excel(excel_name, dtype=df.dtypes)
        tm.assert_frame_equal(df, loaded_df, check_names=False)

    def test_all_pandas_objects(self):
        df1 = DataFrame()
        df2 = DataFrame()
        series1 = Series()
        series2 = Series()
        pandas_objects = [(name, obj) for (name, obj) in locals().items()
                          if isinstance(obj, (DataFrame, Series))]
        excel_name = self.tempexcel.name
        ip.run_line_magic('excel_all', '-f {filepath}'.format(filepath=excel_name))
        for (name, obj) in pandas_objects:
            if isinstance(obj, Series):
                loaded_data = read_excel(excel_name, sheet_name=name, squeeze=True, dtype=obj.dtype)
                tm.assert_series_equal(obj, loaded_data, check_names=False)
            elif isinstance(obj, DataFrame):
                loaded_data = read_excel(excel_name, sheet_name=name, dtype=obj.dtypes)
                tm.assert_frame_equal(obj, loaded_data, check_names=False)
    
    def test_sheet_timestamp(self):
        series = Series()
        excel_name = self.tempexcel.name
        ip.run_line_magic('excel', 'series -f {filepath}'.format(filepath=excel_name))
        loaded_excel = read_excel(excel_name, sheet_name=None)
        sheet_names = list(loaded_excel.keys())
        for sheet in sheet_names:
            _, date_string = sheet.split('_')
            saved_date = datetime.datetime.strptime(date_string, "%Y%m%d-%H%M%S")
            load_to_read = datetime.datetime.now() - saved_date
            # there is probably a better way to test this
            assert load_to_read.seconds < 10

    def tearDown(self):
        self.tempexcel.close()

@pytest.fixture
def no_extension_file():
    file = tempfile.NamedTemporaryFile()
    yield file
    file.close()

def test_filepath_append(no_extension_file):
    series = Series()
    excel_name = no_extension_file.name
    ip.run_line_magic('excel', 'series -f {filepath}'.format(filepath=excel_name))
    exported_filepath = pathlib.PurePath(excel_name + '.xlsx')
    assert exported_filepath.suffix == '.xlsx'
    
def test_all_filepath_append(no_extension_file):
    series = Series()
    df = DataFrame()
    excel_name = no_extension_file.name
    ip.run_line_magic('excel_all', '-f {filepath}'.format(filepath=excel_name))
    exported_filepath = pathlib.Path(excel_name + '.xlsx')
    exported_filepath = pathlib.PurePath(excel_name + '.xlsx')
    assert exported_filepath.suffix == '.xlsx'

def test_no_object():
    with pytest.raises(NameError):
        ip.run_line_magic('excel', 'nonexistantobject')

def test_non_pandas_object():
    integer = 3
    with pytest.raises(TypeError):
        ip.run_line_magic('excel', 'integer')

    string = 'string'
    with pytest.raises(TypeError):
        ip.run_line_magic('excel', 'string')

def test_all_no_objects():
    with pytest.raises(RuntimeError):
        ip.run_line_magic('excel_all', '')

def test_all_too_many_objects():
    objects = [Series() for _ in range(100)]
    with pytest.raises(RuntimeError):
        ip.run_line_magic('excel_all', '')


if __name__ == '__main__':
    unittest.main()