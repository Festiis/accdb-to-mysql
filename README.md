# accdbToMySQL • Access Database to SQL Export

This PowerShell script is designed to facilitate the conversion of an Access database to SQL format.
It generates SQL queries for table creation, data insertion, and relationships, allowing for easy migration of the database structure and content to other SQL-based systems.

## Features

- Generates SQL CREATE TABLE statements for each table in the Access database.
- Provides optional data export by generating SQL INSERT INTO statements for table data.
- Handles primary keys and various field types.
- Generates SQL for defining relationships (foreign keys) between tables.
- Supports exclusion of system tables.

## Usage

1. **Setup**:

   - Ensure you have PowerShell installed on your system.
   - Place the script in the same directory as your Access database file (.accdb).
   - Modify the script if necessary to match your Access database file name or structure.

2. **Run the Script**:

   - Open PowerShell in the script's directory.
   - Run the script using the following command:
     ```
     .\script.ps1
     ```

3. **Follow the Prompts**:

   - The script will prompt you to:
     - Choose whether to export data to SQL files.
     - Enter the Access database password.

4. **Output**:
   - The generated SQL queries (table creation, data insertion, and relationships) will be saved to the "mysql" subfolder in a file named "database.sql".

## Note

- System tables (e.g., those starting with "msys" or "~") are ignored during processing.
