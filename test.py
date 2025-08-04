
import pandas as pd

# Test completo
df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
df.to_excel('test.xlsx', engine='openpyxl')
print("Test openpyxl riuscito!")