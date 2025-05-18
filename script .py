import pandas as pd
import pyodbc
import math

# conn = pyodbc.connect(
#     "DRIVER={ODBC Driver 17 for SQL Server};"
#     "SERVER=SEVERNAME;"
#     "DATABASE=DATABASENAMEr;"
#     "UID=USER;"
#     "PWD=1928"
# )

# CONNECTION WITH WINDOWS AUTHENTICATION
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SERVERNAME;"
    "DATABASE=master;"
    "Trusted_Connection=yes;"
)


query = "SELECT * FROM MOCK_DATA;"  

df = pd.read_sql(query, conn)

conn.close()

MAX_ROWS = 200
total_rows = len(df)
num_sheets = math.ceil(total_rows / MAX_ROWS)

with pd.ExcelWriter("salida_excel.xlsx", engine="openpyxl") as writer:
    for i in range(num_sheets):
        start_row = i * MAX_ROWS
        end_row = (i + 1) * MAX_ROWS
        sheet_df = df[start_row:end_row]
        sheet_name = f'Parte_{i+1}'
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Excel file generated successfully!")
