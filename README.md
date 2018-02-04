# Excelify

Easily export `pandas` objects to Excel spreadsheets with IPython magic.

[![Build Status](https://travis-ci.org/pmbaumgartner/excelify.svg?branch=master)](https://travis-ci.org/pmbaumgartner/excelify) [![codecov](https://codecov.io/gh/pmbaumgartner/excelify/branch/master/graph/badge.svg)](https://codecov.io/gh/pmbaumgartner/excelify)

## Example

### `%excel`
```python
%load_ext excelify

data = [
    {'name' : 'Greg', 'age' : 30},
    {'name' : 'Alice', 'age' : 36}
    ]
df = pd.DataFrame(data)

%excel df -f spreadsheet.xlsx -s sample_data
```

## Magics

### `%excel`

```
%excel [-f FILEPATH] [-s SHEETNAME] dataframe

Saves a DataFrame or Series to Excel

positional arguments:
  dataframe             DataFrame or Series to Save

optional arguments:
  -f FILEPATH, --filepath FILEPATH
                        Filepath to Excel spreadsheet.Default:
                        './{object}_{timestamp}.xlsx'
  -s SHEETNAME, --sheetname SHEETNAME
                        Sheet name to output data.Default:
                        {object}_{timestamp}

```

### `%excel_all`

```
%excel_all [-f FILEPATH] [-n NOSORT]

Saves all Series or DataFrame objects in the namespace to Excel.
Use at your own peril. Will not allow more than 100 objects.

optional arguments:
  -f FILEPATH, --filepath FILEPATH
                        Filepath to excel spreadsheet.Default:
                        './all_data_{timestamp}.xlsx'
  -n NOSORT, --nosort NOSORT
                        Turns off alphabetical sorting of objects for export
                        to sheets
```

## Dependencies

- IPython
- Pandas
- XlsxWriter
## Why?

I had several Jupyter notebooks that were outputting crosstabs or summary statistics that would eventually end up in a Word doc. Depending on the size and complexity of the table, I would either copy/paste or export to Excel. Due to the inconsistency, this made managing all these tables a pain. I figured a tool like this would make it easier to collect everything in a notebook as part of an analysis into one excel file, deal with formatting in excel, and review and insert into a doc from there.