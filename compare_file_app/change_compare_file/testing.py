import pandas as pd

# Create a sample DataFrame
df = pd.DataFrame({
    'col1': ['AH123', 'XYZ', 'AH456', 'PQR', 'AH789'],
    'col2': [1, 2, 3, 4, 5]
})

# Count the number of rows containing "AH"
num_rows_containing_ah = df['col1'].str.contains('AH').sum()

print(type(df))