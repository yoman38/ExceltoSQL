Attribute VB_Name = "Txt2SQL"
Dim conn As Object
Dim cmd As Object
Dim rs As Object

Function ConnectToSQL(serverName As String, dbName As String) As Boolean
    On Error GoTo ErrorHandler

    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    ' Connection string to connect to SQL Server using user input
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & dbName & ";Integrated Security=SSPI;"
    conn.Open

    Set cmd.ActiveConnection = conn
    ConnectToSQL = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    ConnectToSQL = False
End Function

Function GetTableNames() As String
    Dim tableList As String
    Set rs = conn.OpenSchema(20) '20 = adSchemaTables
    
    Do While Not rs.EOF
        ' Exclude system tables and special schemas
        If Not (Left(rs("TABLE_NAME"), 3) = "sys" Or _
                Left(rs("TABLE_NAME"), 2) = "dt" Or _
                Left(rs("TABLE_NAME"), 4) = "MSys" Or _
                rs("TABLE_SCHEMA") = "sys" Or _
                rs("TABLE_SCHEMA") = "INFORMATION_SCHEMA") Then
            tableList = tableList & rs("TABLE_NAME") & vbCrLf
        End If
        rs.MoveNext
    Loop
    
    GetTableNames = tableList
End Function


Sub DeleteTableIfExists(tableName As String)
    On Error Resume Next
    cmd.CommandText = "DROP TABLE " & tableName
    cmd.Execute
    On Error GoTo 0
End Sub

Sub UpdateSQLWithTxtContent()
    Dim sqlText As String
    Dim fd As Object
    Dim serverName As String
    Dim dbName As String
    Dim tableName As String

    ' Get server and database names from user
    serverName = InputBox("Enter the server name:", "Server Name", ".\SQLEXPRESS")
    dbName = InputBox("Enter the database name:", "Database Name", "test_db")
    
    ' Connect to SQL
    If Not ConnectToSQL(serverName, dbName) Then Exit Sub

    ' Show tables in the database
    MsgBox "Tables in the database:" & vbCrLf & GetTableNames(), vbInformation, "Table List"

    ' Prompt user for file selection
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    fd.AllowMultiSelect = False
    fd.Title = "Select a SQL text file"
    If fd.Show = -1 Then
        sqlText = GetQueryFromTxt(fd.SelectedItems(1))
    Else
        MsgBox "File not selected", vbExclamation
        Exit Sub
    End If

    ' Check if table already exists and delete
    tableName = InputBox("Enter the name of the table you're updating/creating:", "Table Name")
    If tableName <> "" Then
        DeleteTableIfExists tableName
    End If

    ' Execute the new query
    cmd.CommandText = sqlText
    cmd.Execute
    MsgBox "Query executed successfully!", vbInformation

    ' Clean up
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Function GetQueryFromTxt(filePath As String) As String
    Dim fileContent As String
    Dim fileNumber As Integer
    
    ' Get an available file number
    fileNumber = FreeFile
    
    ' Open the file for reading
    Open filePath For Input As fileNumber
    
    ' Read the file content
    fileContent = Input$(LOF(fileNumber), fileNumber)
    
    ' Close the file
    Close fileNumber
    
    ' Return the file content
    GetQueryFromTxt = fileContent
End Function

