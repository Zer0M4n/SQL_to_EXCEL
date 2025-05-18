# SQL_to_EXCEL
This project is a Python script that connects to a SQL Server database using pyodbc, executes a SQL query that can return millions of rows, and exports the results to an Excel (.xlsx) file. Since Excel has a row limit of about 1,048,576 rows per sheet, the script automatically splits the data across multiple sheets to prevent data loss

## Technologies Used:
- Python 3
- pandas (for data processing)
- pyodbc (for database connection)
- openpyxl (for writing .xlsx Excel files)
