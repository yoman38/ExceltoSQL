Attribute VB_Name = "Txt2SQL"


Public outputFilePath As String ' Declare a public variable to store the output file path.

Dim conn As Object ' Declare an object variable to store the connection to the SQL Server.
Dim cmd As Object ' Declare an object variable to store the SQL command.
Dim rs As Object ' Declare an object variable to store the result set.

Function ConnectToSQL(serverName As String, dbName As String) As Boolean ' Function to connect to the SQL Server.
    On Error GoTo ErrorHandler ' Error handling starts.

    Set conn = CreateObject("ADODB.Connection") ' Create a new ADODB connection object.
    Set cmd = CreateObject("ADODB.Command") ' Create a new ADODB command object.

    ' Connection string to connect to SQL Server using user input.
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & dbName & ";Integrated Security=SSPI;"
    conn.Open ' Open the connection to the SQL Server.

    Set cmd.ActiveConnection = conn ' Set the command object's active connection to the opened connection.
    ConnectToSQL = True ' Return True to indicate successful connection.
    Exit Function ' Exit the function.

ErrorHandler: ' Error handling starts if any error occurs during the function.
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical ' Display an error message with the error number and description.
    ConnectToSQL = False ' Return False to indicate connection failure.
End Function

Function GetTableNames() As String ' Function to get the list of table names from the connected SQL Server.
    Dim tableList As String ' Variable to store the list of table names as a string.
    Set rs = conn.OpenSchema(20) '20 = adSchemaTables, open a schema recordset containing information about the tables in the database.

    Do While Not rs.EOF ' Loop through the records in the schema recordset.
        ' Exclude system tables and special schemas by checking the table name and schema name.
        If Not (Left(rs("TABLE_NAME"), 3) = "sys" Or _
                Left(rs("TABLE_NAME"), 2) = "dt" Or _
                Left(rs("TABLE_NAME"), 4) = "MSys" Or _
                rs("TABLE_SCHEMA") = "sys" Or _
                rs("TABLE_SCHEMA") = "INFORMATION_SCHEMA") Then
            tableList = tableList & rs("TABLE_NAME") & vbCrLf ' Add the table name to the list along with a new line.
        End If
        rs.MoveNext ' Move to the next record in the recordset.
    Loop

    GetTableNames = tableList ' Return the list of table names.
End Function

Sub DeleteTableIfExists(tableName As String) ' Subroutine to delete a table if it exists in the database.
    On Error Resume Next ' Continue execution even if an error occurs (i.e., table not found).
    cmd.CommandText = "DROP TABLE " & tableName ' Set the command text to delete the specified table.
    cmd.Execute ' Execute the command to delete the table.
    On Error GoTo 0 ' Disable the error handling.
End Sub

Sub UpdateSQLWithTxtContent() ' Subroutine to update the SQL Server with the content from a text file.

    Dim sqlText As String ' Variable to store the SQL query read from a text file.
    Dim fd As Object ' File dialog object (not used in the provided code).
    Dim serverName As String ' Variable to store the SQL Server name.
    Dim dbName As String ' Variable to store the SQL Server database name.
    Dim tableName As String ' Variable to store the table name extracted from the output file path.
    Dim fso As Object ' File system object for handling file operations.
    Set fso = CreateObject("Scripting.FileSystemObject") ' Create the file system object.

    ' Get server and database names from the user using input boxes.
    serverName = InputBox("Enter the server name:", "Server Name", ".\SQLEXPRESS")
    If serverName = "" Then
        MsgBox "Operation cancelled by user."
        Exit Sub
    End If
    
    dbName = InputBox("Enter the database name:", "Database Name", "test_db")
    If dbName = "" Then
        MsgBox "Operation cancelled by user."
        Exit Sub
    End If
    
    ' Connect to the SQL Server using the ConnectToSQL function.
    If Not ConnectToSQL(serverName, dbName) Then Exit Sub ' If the connection fails, exit the subroutine.

    ' Show tables in the database using GetTableNames function and display them in a message box.
    MsgBox "Tables in the database:" & vbCrLf & GetTableNames(), vbInformation, "Table List"

    ' Get the SQL query from the text file specified by the outputFilePath variable using the GetQueryFromTxt function.
    sqlText = GetQueryFromTxt(outputFilePath) ' Using the output file from Modified_ExcelSQL0.bas

    ' Extract the table name from the outputFilePath and check if the table already exists. If yes, delete it using DeleteTableIfExists subroutine.
    tableName = fso.GetBaseName(outputFilePath)
    If tableName <> "" Then
        DeleteTableIfExists tableName
    End If

    ' Execute the new SQL query by setting the command text and using the cmd.Execute method.
    
' Check and drop primary key constraint if it exists
cmd.CommandText = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE CONSTRAINT_TYPE = 'PRIMARY KEY' AND TABLE_NAME = '" & tableName & "') " & _
                  "BEGIN " & _
                  "   ALTER TABLE " & tableName & " DROP CONSTRAINT PK_" & tableName & " " & _
                  "END"
cmd.Execute

' Check and drop foreign key constraint if it exists
' Assuming a naming convention of FK_<TableName>_<ColumnName>
cmd.CommandText = "IF EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'FK_" & tableName & "_" & columnName & "') " & _
                  "BEGIN " & _
                  "   ALTER TABLE " & tableName & " DROP CONSTRAINT FK_" & tableName & "_" & columnName & " " & _
                  "END"
cmd.Execute

cmd.CommandText = sqlText
    cmd.Execute
    MsgBox "Query executed successfully!", vbInformation ' Display a success message.

    ' Clean up by setting the cmd and conn objects to Nothing to release memory resources.
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Function GetQueryFromTxt(filePath As String) As String ' Function to read the content of a text file and return it as a string.
    Dim fileContent As String ' Variable to store the file content as a string.
    Dim fileNumber As Integer ' File number to use when opening the text file.

    ' Get an available file number to be used for file operations.
    fileNumber = FreeFile

    ' Open the file specified by the filePath for reading.
    Open filePath For Input As fileNumber

    ' Read the entire content of the file into the fileContent variable.
    fileContent = Input$(LOF(fileNumber), fileNumber)

    ' Close the file.
    Close fileNumber

    ' Return the file content as the result of the function.
    GetQueryFromTxt = fileContent
End Function








