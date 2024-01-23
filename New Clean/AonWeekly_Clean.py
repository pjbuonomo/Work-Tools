import pandas as pd
import pyodbc
from datetime import datetime

# Replace with your actual connection details
conn = pyodbc.connect('DSN=ILS;Trusted_Connection=yes;')

# FILE NEEDS TO BE CLEANED BEFORE USE
file_path = "S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/Aon_Weekly/Aon20231229Clean.xlsx"
table = pd.read_excel(file_path, sheet_name='RLS', skiprows=1)

# DATE NEEDS TO BE MANUALLY CHANGED
table['QDate'] = pd.Timestamp('2023-12-29')

# Dropping specific columns
table = table.drop(table.columns[[3, 12, 13, 14]], axis=1)

# Getting column names from the SQL table
cursor = conn.cursor()
cursor.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Aon'")
col_names = [row[0] for row in cursor.fetchall()]
table.columns = col_names

# Handling NaN and 'n/a' values
table['LongTermAsk'].fillna(0, inplace=True)
table['LongTermEL'].fillna(0, inplace=True)
table['NearTermAsk'].fillna(0, inplace=True)
table['NearTermEL'].fillna(0, inplace=True)
table['BidSpread'] = table['BidSpread'].replace('n/a', 0)
table['OfferSpread'] = table['OfferSpread'].replace('n/a', 0)

# Parsing and formatting numbers
table['Coupon'] = pd.to_numeric(table['Coupon'], errors='coerce').fillna(0).astype(int)
table['BidPrice'] = pd.to_numeric(table['BidPrice'], errors='coerce').fillna(0)
table['OfferPrice'] = pd.to_numeric(table['OfferPrice'], errors='coerce').fillna(0)

for column in ['Size', 'LongTermAsk', 'LongTermEL', 'NearTermAsk', 'NearTermEL', 'Coupon', 'BidPrice', 'OfferPrice']:
    table[column] = table[column].apply(lambda x: format(x, '.2f'))

# Saving to SQL database
for index, row in table.iterrows():
    cursor.execute("INSERT INTO Aon (columns) VALUES (values)", tuple(row))
    conn.commit()

# Closing the database connection
conn.close()
PS C:\Users\buonomo> & C:/ProgramData/Anaconda3/python.exe //ad-its.credit-agricole.fr/Amundi_Boston/Homedirs/buonomo/@Config/Desktop/BetterDatabaseTool/AonWeekly_Clean.py
Traceback (most recent call last):
  File "\\ad-its.credit-agricole.fr\Amundi_Boston\Homedirs\buonomo\@Config\Desktop\BetterDatabaseTool\AonWeekly_Clean.py", line 38, in <module>
    table[column] = table[column].apply(lambda x: format(x, '.2f'))
  File "C:\ProgramData\Anaconda3\lib\site-packages\pandas\core\series.py", line 4357, in apply
    return SeriesApply(self, func, convert_dtype, args, kwargs).apply()
  File "C:\ProgramData\Anaconda3\lib\site-packages\pandas\core\apply.py", line 1043, in apply
    return self.apply_standard()
  File "C:\ProgramData\Anaconda3\lib\site-packages\pandas\core\apply.py", line 1098, in apply_standard
    mapped = lib.map_infer(
  File "pandas\_libs\lib.pyx", line 2859, in pandas._libs.lib.map_infer
  File "\\ad-its.credit-agricole.fr\Amundi_Boston\Homedirs\buonomo\@Config\Desktop\BetterDatabaseTool\AonWeekly_Clean.py", line 38, in <lambda>
    table[column] = table[column].apply(lambda x: format(x, '.2f'))
ValueError: Unknown format code 'f' for object of type 'str'
