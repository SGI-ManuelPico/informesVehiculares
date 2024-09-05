import pandas as pd
from texttable import Texttable

# Create a sample DataFrame
df = pd.DataFrame({
    'Name': ['John', 'Jane', 'Bob'],
    'Age': [28, 34, 22],
    'City': ['New York', 'Los Angeles', 'Chicago']
})

# Create a Texttable object
table = Texttable()
table.set_deco(Texttable.HEADER)
table.set_cols_dtype(['t', 'i', 't'])  # t is text, i is integer
table.set_cols_align(["l", "r", "l"])  # l is left, r is right

# Add header
table.add_row(df.columns.tolist())

# Add rows
for row in df.itertuples(index=False):
    table.add_row(row)

# Generate the table
print(table.draw())