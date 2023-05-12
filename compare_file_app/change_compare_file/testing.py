import pandas as pd
import re

# Create a sample DataFrame
df = pd.DataFrame({
    'col1': ['AH123', 'XYZ', 'AH456', 'PQR', 'AH789'],
    'col2': [1, 2, 3, 4, 5]
})

# Count the number of rows containing "AH"
num_rows_containing_ah = df['col1'].str.contains('AH').sum()


pattern1 = r"\d{3}-\d{8}"
file_name = "dadsa dfs f 123- 98789098"

file_name = file_name.replace(" ", "")
matches = re.findall(pattern1, file_name)
print(matches[0] + '_T86')