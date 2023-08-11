# VBA Module: Excel to SQL Command Generator

This VBA module simplifies the conversion of Excel data into SQL commands, aiding in generating SQL statements from Excel files and storing them in text format. Its purpose is to expedite the process of transforming Excel data into SQL queries for database operations.

## Key Features

- Generate SQL commands for table creation and data insertion.
- Define Excel input file and output text file paths.
- Filter rows using specific keywords.
- Exclude duplicate rows based on chosen columns.
- Omit rows containing empty cells within a designated range.
- Integrate unique identifiers for rows.
- Specify primary keys for the table.
- Set up foreign keys within the table.
- Define NOT NULL constraints for columns.
- Implement indexes and constraints within the table.
- Assign default values to columns.
- Seamless handling of european languages special characters for SQL compatibility.
- Directly connect to SQL Server and execute SQL statements from a `.txt` file.
- Safeguard data by removing existing tables before creating new ones.

## Latest Release: Enhanced Usability

The updated version offers improved user-friendliness. Here's how to utilize it:

1. **Installation**: Download the `.xlsm` file.

2. **Execute the Main Procedure**:
   - Click the "Run" button.
   - A dialog box will prompt for inputs:
     - Select the Excel file.
     - Designate the output text file to store generated SQL commands.
     - Optionally, specify keyword-based row filtering (e.g., "yes" or "no").
     - If filtering, provide the keyword for retrieving specific rows.
     - Optionally, skip duplicate entries.
     - Optionally, skip rows with empty cells.
     - Optionally, enable unique ID generation.

3. **VBA Processing**:
   - The VBA code will open the Excel file.
   - Data will be extracted based on defined ranges.
   - SQL statements will be generated for creating tables and inserting data.

4. **Output**:
   - Generated SQL statements will be saved in the designated output text file.

5. **Polish Special Characters Handling**:
   - The code effectively manages Polish special characters for SQL compatibility.

6. **Automatic SQL Server Connection**.

## Previous Installation and Usage

1. **Installation**: Download the provided `.xlsm` file.

2. **Usage**:
   - Open the downloaded `.xlsm` file in Excel.
   - Access the VBA editor by pressing "Alt + F11."
   - Create a new module through "Insert" in the top menu, then choose "Module."
   - Copy and paste the supplied code into the new module.
   - Add a button and assign the macro.
   - Execute the main procedure "GenerateSQL" by clicking "Run" or pressing "F5."
   - Follow the prompts to input essential details, including Excel file path, output text file path, filtering preferences, unique IDs, keys, constraints, indexes, and more.
   - The VBA code will handle Excel data and produce SQL statements as specified.
   - The resultant SQL statements will be stored in the output text file.

## Important Considerations

- **Backup**: Prior to code execution, ensure a backup of the Excel file exists. The code may manipulate Excel data during operations.
- **Sheet Selection**: Incorrect active sheet may cause issues. Utilize "RESTRICTION" mode for accuracy. (Note: DELETED NOW)
- **Mixed Data Type Detection**: Version v2 addressed mixed data type detection using NVARCHAR.
- **Clarity and Prompting**: Versions v3 and v4 improved user clarity and prompts during execution.
- **Review Generated SQL**: The code might open the output text file for your review before uploading.

## Changelog

### v3
- Resolved an issue causing merged values between runs when workbook wasn't closed.
- Enhanced selection range for improved accuracy and reliability.

### v3.11
- Ensured active sheet activation prior to user range selection.

### v3.35
- Refined range selection process for enhanced usability.

### v3.36
- Automatically opens the generated output text file for review before upload.
- Rectified output issues and related concerns.
- Expanded options for specifying keys.

### v3.4
- Introduced the option to skip rows containing specific words.
- Enhanced prompts for better user comprehension.
- Enabled entering multiple filter words.
- Data range is now selected for filter range by default.
- Now handle all europeans languages with a latin alphabet, not just polish

### v3.43
- Now handle Russian, Ukrainian, Greek, etc. Converted alphabet to their latin counterpart to make it work in sql.

### Additional Enhancements
- Included additional prompts for clear guidance and options.

## Further Information

The subsequent section details additional intricacies and explains each feature through examples:

---

---

## Additional Details

Before generating an output, the following queries may arise:

- **Omitting Rows without a Specific Keyword**: This function enables row exclusion based on a designated keyword. For instance, with the following dataset:
  ```
  1 / Michael / 123
  2 / Julius / 456
  1 / Michael / 123
  1 / Olga / 789
  3 / Kevin / 789
  ```
  Selecting data range A1:C4 and filtering range A1:A5 with keyword "1" results in:
  ```
  1 / Michael / 123
  1 / Michael / 123
  1 / Olga / 789
  ```
  Choosing data range A1:C4 and filter range A1:A4 with keyword "1" results in:
  ```
  1 / Michael / 123
  1 / Michael / 123
  1 / Olga / 789
  3 / Kevin / 789
  ```
  Similarly, selecting range C1:C5 with keyword "123" yields:
  ```
  1 / Michael / 123
  1 / Michael / 123
  ```

- **Excluding Duplicate Rows Based on Specific Columns**: This feature facilitates the removal of duplicate rows based on designated columns. Using column A:
  ```
  1 / Michael / 123
  2 / Julius / 456
  ```
  And with column C:
  ```
  1 / Michael / 123
  2 / Julius / 456
  1 / Olga / 789
  ```

- **Omitting Rows with Empty Cells in a Specific Range**: This functionality discards rows containing empty cells within a specified range.

- **Inclusion of a Unique ID for Each Row**: Activation of this feature appends an ID column "[Id] [int] IDENTITY(1,1) NOT NULL" for each row in the table.

- **Setting Primary Key for the Table**: This choice empowers the establishment of a primary key for the table. Multiple columns can be selected and set as NOT NULL via the "ALTER TABLE ... ADD PRIMARY KEY ..." statement.

- **Adding a Foreign Key**: Selection of a column, its NOT NULL assignment, reference table, and column creation results in an "ALTER TABLE ... ADD FOREIGN KEY ... REFERENCES ..." statement.

- **Imposing NOT NULL Constraint on Other Columns**: This attribute streamlines the process of setting specific columns as NOT NULL through the "ALTER TABLE ... ALTER COLUMN ... NOT NULL" statement.

- **Inclusion of an Index**: Options include "Non-clustered," "Clustered," or "Unique Non-clustered with Sort Order." Column selection leads to the creation of 'CREATE INDEX' / 'CREATE

 CLUSTERED INDEX' / 'CREATE UNIQUE INDEX … ON … DESC' (or ASC).

- **Addition of a Constraint**: Constraints such as "unique" or "check" can be set using the "ALTER TABLE … ADD CONSTRAINT …" statement.

- **Assignment of Default Value for Any Column**: Utilize this functionality to assign default values to columns via the "ALTER TABLE … ADD CONSTRAINT … DEFAULT … FOR …" statement.

---
