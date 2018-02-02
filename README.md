# Excelify

Easily export `pandas` objects to Excel spreadsheets with IPython magic.

## Example

```python
%load_ext excelify
data = [
    {'name' : 'Greg', 'age' : 30},
    {'name' : 'Alice', 'age' : 36}
    ]
df = pd.DataFrame(data)
%excel df -f spreadsheet.xlsx -s sample_data
```

## Why?

I had several Jupyter notebooks that were outputting crosstabs or summary statistics that would eventually end up in a Word doc. Depending on the size and complexity of the table, I would either copy/paste or export to Excel. Due to the inconsistency, this made managing all these tables a pain. I figured a tool like this would make it easier to collect everything in a notebook as part of an analysis into one excel file, deal with formatting in excel, and review and insert into a doc from there.