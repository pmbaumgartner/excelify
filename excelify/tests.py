import unittest
import tempfile
import pathlib
import datetime
import warnings

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

    def test_sheet_name(self):
        series = Series()
        excel_name = self.tempexcel.name
        sheetname = 'test_sheet_name'
        ip.run_line_magic('excel', 'series -f {filepath} -s {sheetname}'.format(filepath=excel_name, sheetname=sheetname))
        loaded_excel = read_excel(excel_name, sheet_name=None)
        assert 'test_sheet_name' in loaded_excel

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

    def test_all_long_name(self):
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            locals().update({'a' * 33 : Series()})
            excel_name = self.tempexcel.name
            ip.run_line_magic('excel_all', '-f {filepath}'.format(filepath=excel_name))
            assert len(w) == 1
            assert issubclass(w[-1].category, UserWarning)
            assert "truncated" in str(w[-1].message)

    def test_long_name_provided(self):
        with warnings.catch_warnings(record=True) as w:
            series = Series()
            excel_name = self.tempexcel.name
            longsheet = 'a' * 33
            ip.run_line_magic('excel', 'series -f {filepath} -s {longsheet}'.format(filepath=excel_name, longsheet=longsheet))
            assert len(w) == 1
            assert issubclass(w[-1].category, UserWarning)
            assert "truncated" in str(w[-1].message)

    def test_long_name_default(self):
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            longsheet = 'a' * 33
            locals().update({longsheet : Series()})
            excel_name = self.tempexcel.name
            ip.run_line_magic('excel', '{longsheet} -f {filepath}'.format(longsheet=longsheet, filepath=excel_name))
            assert len(w) == 1
            assert issubclass(w[-1].category, UserWarning)
            assert "truncated" in str(w[-1].message)

    def tearDown(self):
        self.tempexcel.close()


def test_filename():
    series = Series()
    ip.run_line_magic('excel', 'series')
    excel_name = list(pathlib.Path().glob('series_*.xlsx'))[0]
    assert excel_name.exists()
    excel_name.unlink()

def test_all_filename():
    series = Series()
    df = DataFrame()
    ip.run_line_magic('excel_all', '')
    excel_name = list(pathlib.Path().glob('all_data_*.xlsx'))[0]
    assert excel_name.exists()
    excel_name.unlink()


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
    # this seems like a bad idea...
    for i in range(102):
        locals().update({'series' + str(i) : Series()})
    with pytest.raises(RuntimeError):
        ip.run_line_magic('excel_all', '')

