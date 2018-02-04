import unittest
import tempfile

from IPython.testing.globalipapp import start_ipython, get_ipython

import pandas.util.testing as tm
from pandas.core.frame import DataFrame
from pandas.core.series import Series
from pandas import read_excel

import pytest


class TestMagicExportImport(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.ip = get_ipython()
        cls.ip.magic('load_ext excelify')

    def setUp(self):
        self.tempexcel = tempfile.NamedTemporaryFile(suffix='.xlsx')

    def test_series(self):
        series = Series()
        self.ip.push({'series': Series()})
        excel_name = self.tempexcel.name
        self.ip.run_line_magic('excel', 'series -f {filepath}'.format(filepath=excel_name))
        series = self.ip.user_ns['series']
        loaded_series = read_excel(excel_name, squeeze=True, dtype=series.dtype)
        tm.assert_series_equal(series, loaded_series, check_names=False)

    def test_dataframe(self):
        df = DataFrame()
        excel_name = self.tempexcel.name
        print(excel_name)
        self.ip.run_line_magic('excel', 'df -f {filepath}'.format(filepath=excel_name))
        loaded_df = read_excel(excel_name, dtype=df.dtypes)
        tm.assert_frame_equal(df, loaded_df, check_names=False)

    def test_all_pandas_objects(self):
        df1 = DataFrame()
        df2 = DataFrame()
        series1 = Series()
        series2 = Series()
        pandas_object = lambda d: isinstance(d, (DataFrame, Series))
        pandas_objects = [(name, obj) for (name, obj) in locals().items() if pandas_object(obj)]
        excel_name = self.tempexcel.name
        self.ip.run_line_magic('excel_all', '-f {filepath}'.format(filepath=excel_name))
        for (name, obj) in pandas_objects:
            if isinstance(obj, Series):
                loaded_data = read_excel(excel_name, sheet_name=name, squeeze=True, dtype=obj.dtype)
                tm.assert_series_equal(obj, loaded_data, check_names=False)
            elif isinstance(obj, DataFrame):
                loaded_data = read_excel(excel_name, sheet_name=name, dtype=obj.dtypes)
                tm.assert_frame_equal(obj, loaded_data, check_names=False)

    def test_no_object(self):
        with pytest.raises(NameError):
            self.ip.run_line_magic('excel', 'df')

    def test_no_objects(self):
        with pytest.raises(RuntimeError):
            self.ip.run_line_magic('excel_all', '')

    def test_too_many_objects(self):
        objects = [Series() for _ in range(100)]
        with pytest.raises(RuntimeError):
            self.ip.run_line_magic('excel_all', '')

    def test_non_pandas_object(self):
        integer = 3
        with pytest.raises(TypeError):
            self.ip.run_line_magic('excel', 'integer')

    def tearDown(self):
        self.tempexcel.close()


if __name__ == '__main__':
    unittest.main()