This VBA module helps you generate SQL statements from data in an Excel file and save them to a text file. It can be handy when you want to convert your Excel data into SQL commands for database operations.

NEW RELEASE > NOW EASIER TO USE
Here's how to use it:
0. INSTALLATION // Just download xlsm file. 
1. RUN the main procedure "GenerateSQL" by clicking "Run" button.
2. A pop-up window will ask you for some inputs: - Choose your Excel file 
3. Choose the output text file where you want to save the generated. Then choose the name pf the txt file. 
- Optionally, specify if you want to filter rows based on a specific keyword (e.g., "yes" or "no"). If you choose "yes," it will ask you to enter the keyword to filter rows (e.g., "zm"). This will only retrieve rows containing the keyword. 
- Optionnally, you can avoid the rows containing duplicates entry. 
- Optionnally, you can avoid the rows containing empty cells. 
- Optionnally, you can generate an ID. Note that the ID will be generated in Microsoft SQL Server and not in VBA (no loop). The ID consists of a simple increment starting from 1. You can change the function in the output or in VBA to use a random number using NEWID or NEWSEQUENTIALID. 
4. After providing the inputs, the VBA code will open your Excel file, extract data based on your specified ranges, and generate SQL statements for creating a table and inserting data. 
5. The generated SQL statements will be saved in the output text file you specified. 
6. The code will also handle Polish special characters in your data to ensure compatibility with SQL. 

7. COPY PASTE THE CONTENT IN TXT FILE TO MICROSOFT SQL QUERY

**WIP : EXCEPTION HANDLING, allow the users to leave the program at any time.


////// AUTOMATIC TXT TO SQL CONVERTER
VBA Code Overview for Microsoft Access and SQL Server Integration

1. Introduction:
This VBA module provides functionalities to connect to a SQL Server database, fetch and display user tables, and execute SQL statements from a selected `.txt` file. Users can also specify the server and database names for a custom connection.

2. Functions and Procedures:

- ConnectToSQL(serverName, dbName): 
  Connects to the specified SQL Server and database. Returns a Boolean indicating success or failure.

- GetTableNames(): 
  Fetches the names of user tables in the connected SQL Server database and returns them as a concatenated string.

- DeleteTableIfExists(tableName): 
  Deletes the specified table from the SQL Server database if it exists.

- UpdateSQLWithTxtContent(): 
  This is the primary procedure. It prompts the user for server and database names, shows existing user tables, asks for a `.txt` file containing SQL commands, and executes the SQL commands.

- GetQueryFromTxt(filePath): 
  Reads and returns the content of the specified `.txt` file.

3. Usage:
To use the module, integrate it into an Access VBA project. Users can run the `UpdateSQLWithTxtContent` procedure, either directly or by attaching it to a form button, to start the process.

4. Notes:
- Ensure all functions and procedures are in the same VBA module.
- Before executing a new table creation, the script will check if a table with the same name already exists and will delete it. This can lead to data loss, so use with caution.
- System tables and tables from special schemas are excluded from the table list display.


______
Older version

Here's how to use it:

1. Open your Excel file and press "Alt + F11" to access the VBA editor.
2. In the editor, click on "Insert" from the top menu and choose "Module." This will create a new module called "ModuleX"
3. Copy the entire code provided in the "ModuleX" and paste it into your newly created module. You might need to delete the declaration before option explicit.
4. Now, you can run the main procedure "GenerateSQL" by clicking "Run" or pressing "F5." A pop-up window will ask you for some inputs:
   - Enter the path of your Excel file (e.g., C:\Users\yourname\Documents\YourExcelFile.xls). > NOW AUTOMATIC
   - Enter the path of the output text file where you want to save the generated SQL (e.g., C:\Users\yourname\Documents\Output.txt). > NOW AUTOMATIC
   - Optionally, specify if you want to filter rows based on a specific keyword (e.g., "yes" or "no"). If you choose "yes," it will ask you to enter the keyword to filter rows (e.g., "zm"). 
     This will only retrieve rows containing the keyword.
   - Optionnally, you can avoid the rows containing duplicates entry.
   - Optionnally, you can avoid the rows containing empty cells.
   - Optionnally, you can generate an ID. Note that the ID will be generated in Microsoft SQL Server and not in VBA (no loop). The ID consists of a simple increment starting from 1. 
     You can change the function in the output or in VBA to use a random number using NEWID or NEWSEQUENTIALID.
5. After providing the inputs, the VBA code will open your Excel file, extract data based on your specified ranges, and generate SQL statements for creating a table and inserting data.
6. The generated SQL statements will be saved in the output text file you specified.
7. The code will also handle Polish special characters in your data to ensure compatibility with SQL.

Make sure to have a backup of your Excel file before running the code, just in case. This tool can save you time when dealing with large datasets and needing SQL commands for database interactions.


NOTE: it was made to handle polish characters, adapt the code if needed.

To pozwala na generowanie instrukcji SQL z danych zawartych w pliku Excel i zapisywanie ich do pliku tekstowego. Przyda się, gdy chcesz przekształcić dane z arkusza Excela na instrukcje SQL do operacji na bazie danych.

Instrukcje, jak to użyć:

1. Otwórz plik Excela i wciśnij "Alt + F11", aby uzyskać dostęp do edytora VBA.
2. W edytorze, kliknij "Insert" w górnym menu, a następnie wybierz "Module". Spowoduje to utworzenie nowego modułu o nazwie "Module4".
3. Skopiuj cały kod dostarczony w "Module4" i wklej go do nowo utworzonego modułu.
4. Teraz możesz uruchomić główną procedurę "GenerateSQL", klikając "Run" lub naciskając "F5". Pojawi się okno z prośbą o podanie kilku informacji:
   - Wprowadź ścieżkę do Twojego pliku Excela (np. C:\Users\nazwaużytkownika\Dokumenty\TwójPlikExcel.xls).
   - Wprowadź ścieżkę do pliku tekstowego, w którym chcesz zapisać wygenerowane instrukcje SQL (np. C:\Users\nazwaużytkownika\Dokumenty\Output.txt).
   - Opcjonalnie, określ, czy chcesz filtrować wiersze na podstawie określonego słowa kluczowego (np. "yes" lub "no"). Jeśli wybierzesz "yes", zostaniesz poproszony o wprowadzenie słowa kluczowego do filtrowania wierszy>
5. Po podaniu informacji, kod VBA otworzy Twój plik Excela, wydobędzie dane na podstawie określonych zakresów i wygeneruje instrukcje SQL do tworzenia tabeli i wstawiania danych.
6. Wygenerowane instrukcje SQL zostaną zapisane w pliku tekstowym, który wcześniej podałeś jako wynik.
7. Kod również poradzi sobie z polskimi znakami specjalnymi w Twoich danych, aby zapewnić kompatybilność z SQL.

Upewnij się, że masz kopię zapasową swojego pliku Excela przed uruchomieniem kodu, na wszelki wypadek. To narzędzie może zaoszczędzić Ci czas, gdy masz do czynienia z dużymi zestawami danych i potrzebujesz instrukcji SQL do interakcji z bazą danych.
