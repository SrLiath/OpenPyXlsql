# Excel File Processing Script

This Python script allows you to process Excel files, query the SQL Server database and move the corresponding lines to the "Processed" or "Rejected" folders, according to certain conditions.

## Requirements

- Python 3.x
- Python libraries: pyodbc, pandas, openpyxl

Make sure you have the necessary libraries installed before running the script.

## Settings

1. Configure the SQL Server database settings in code, providing the server name, database name, and credentials (if applicable).

2. Define folder paths where Excel files will be read, processed, rejected and where logs will be stored.

3. Create the necessary tables in the database using the SQL script provided in the `banco.sql` file or edit your own tables. Make sure you have SQL Server configured correctly and have privileges to create tables and insert data.

## Usage

1. Place the Excel files in the specified folder (`folder_path`).

2. Run the `main.py` Python script.

3. The script will read each Excel file, query the database, update the corresponding lines and move the files to the "Processed" or "Rejected" folders.

4. The processing logs will be registered in the `log_basico.txt` and `log_<excel_file>.txt` files, providing information about processed, updated and rejected lines.

Make sure you have proper permissions on the required directories and files.

## Comments

- The script assumes that the Excel files have "Sequential" and "Cnpj" columns that will be used for querying and updating the database.

- Make sure the database tables and columns correctly match the queries and updates performed in the script.

- Check and adjust the script according to the specific needs of your environment and data.
