Attribute VB_Name = "MSql"
Option Explicit

' Main Subroutine that gets user input and generates SQL statements
Sub GenerateSQL()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim headers As Range
    Dim filterRange As Range
    Dim duplicateCheckRange As Range
    Dim sqlCreate As String
    Dim sqlInsert As String
    Dim filename As String
    Dim outputFile As String
    Dim filterKeyword As String
    Dim useFilter As String
    Dim checkDuplicates As String
    Dim duplicateCheckColumns As String
    Dim skipEmpty As String
    Dim emptyCheckRange As Range

    ' Prompt user for necessary inputs
    filename = GetUserInput("Enter the path of the Excel file:", "C:\Users\name\Desktop\data\PROJECT3_work_schedule\TP1.xls")
    outputFile = GetUserInput("Enter the path of the output text file:", "C:\Users\name\Desktop\output.txt")
    useFilter = GetUserInput("Do you want to skip rows without a specific keyword? (yes/no)", "no")

    If LCase(useFilter) = "yes" Then
        filterKeyword = GetUserInput("Enter the keyword to filter rows:", "zm")
    Else
        filterKeyword = "" ' No filtering
    End If

    ' Error handling
    On Error GoTo ErrorHandler

    ' Open the workbook
    Set wb = Workbooks.Open(filename)
    Set ws = wb.Sheets("TP1 grafik brygad 2022-2023")

    ' Ask the user about skipping duplicate rows
    checkDuplicates = GetUserInput("Do you want to skip duplicate rows based on specific columns? (yes/no)", "no")

    If LCase(checkDuplicates) = "yes" Then
        duplicateCheckColumns = GetUserInput("Enter the range of columns to check for duplicates (e.g., 'F3:H100'):", "F3:H100")
        Set duplicateCheckRange = ws.Range(duplicateCheckColumns)
    End If
    
    ' Ask the user about skipping rows with empty cells
    skipEmpty = GetUserInput("Do you want to skip rows with empty cells in a specific range? (yes/no)", "no")
    
    If LCase(skipEmpty) = "yes" Then
        ' Ask the user to specify the range to check for empty cells
        Set emptyCheckRange = ws.Range(GetUserInput("Enter the range to check for empty cells (e.g., 'F3:H100'):", "F3:H100"))
    End If

    ' Ask the user to specify the range of data and headers
    Set rng = ws.Range(GetUserInput("Enter the range of data (e.g., 'F3:BK100'):", "F3:BK100"))
    Set headers = ws.Range(GetUserInput("Enter the range of headers (e.g., 'F2:BK2'):", "F2:BK2"))

    ' Ask the user to specify the range of column for keyword filtering if filter is being used
    If filterKeyword <> "" Then
        Set filterRange = ws.Range(GetUserInput("Enter the range of column for keyword filtering (e.g., 'G3:G100'):", "G3:G100"))
    End If
    
    ' Extract the base name from the output file path
    Dim tableName As String
    tableName = GetBaseName(outputFile)
    
    ' Generate SQL
    sqlCreate = GenerateCreateTable(headers, tableName)
    sqlInsert = GenerateInsertStatements(rng, filterKeyword, filterRange, duplicateCheckRange, emptyCheckRange)
     
    ' Write output to file
    WriteToFile outputFile, sqlCreate & vbCrLf & sqlInsert

    ' Close the workbook without saving changes
    wb.Close False

    MsgBox "SQL generation complete! The output was written to " & outputFile, vbInformation, "Success"

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Resume ExitSub
End Sub

' Function to show an input box and return the user's input
Function GetUserInput(prompt As String, Optional defaultText As String) As String
    GetUserInput = InputBox(prompt, "Input Required", defaultText)
End Function

Function GenerateCreateTable(headers As Range, tableName As String) As String
    Dim i As Long
    Dim sqlCreate As String
    Dim columnName As String
    Dim dict As Object
    Dim suffix As Long

    Set dict = CreateObject("Scripting.Dictionary")

    sqlCreate = "CREATE TABLE [" & tableName & "] ("
    For i = 1 To headers.Columns.Count
        columnName = ReplaceSpecialCharacters(headers.Cells(1, i).Value) ' Get column name from headers
        If columnName = "" Then
            columnName = "UnnamedColumn" & i
        Else
            If dict.Exists(columnName) Then
                suffix = dict(columnName) + 1
                dict(columnName) = suffix
                columnName = columnName & "_" & suffix
            Else
                dict.Add columnName, 0
            End If
        End If
        sqlCreate = sqlCreate & "[" & columnName & "] NVARCHAR(50), "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon

    GenerateCreateTable = sqlCreate
End Function

Function GenerateInsertStatements(rng As Range, filterKeyword As String, Optional filterRange As Range, Optional duplicateCheckRange As Range, Optional emptyCheckRange As Range) As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sqlInsert As String
    Dim filterCondition As Boolean
    Dim duplicateCheck As String
    Dim duplicateDict As Object
    Dim emptyCheck As Boolean

    Set duplicateDict = CreateObject("Scripting.Dictionary")

    For i = 1 To rng.Rows.Count
        If filterKeyword = "" Then
            filterCondition = True
        Else
            filterCondition = InStr(1, filterRange.Cells(i, 1).Value, filterKeyword, vbTextCompare) > 0
        End If
        
        ' Check for empty cells in the specified range
        If Not emptyCheckRange Is Nothing Then
            emptyCheck = Application.WorksheetFunction.CountA(emptyCheckRange.Rows(i)) = emptyCheckRange.Columns.Count
        Else
            emptyCheck = True
        End If

        ' Combine values of selected columns to create a unique key
        If Not duplicateCheckRange Is Nothing Then
            duplicateCheck = ""
            For k = 1 To duplicateCheckRange.Columns.Count
                duplicateCheck = duplicateCheck & duplicateCheckRange.Cells(i, k).Value
            Next k
        Else
            duplicateCheck = "none"
        End If

        If filterCondition And Not duplicateDict.Exists(duplicateCheck) And emptyCheck Then
            duplicateDict.Add duplicateCheck, 0
            sqlInsert = sqlInsert & "INSERT INTO [" & tableName & "] VALUES ("
            For j = 1 To rng.Columns.Count
                ' Replace single quotes with two single quotes to escape them and Polish special characters
                sqlInsert = sqlInsert & "N'" & Replace(ReplaceSpecialCharacters(rng.Cells(i, j).Value), "'", "''") & "', "
            Next j
            sqlInsert = Left(sqlInsert, Len(sqlInsert) - 2) & ");" & vbCrLf ' Remove trailing comma and space, add closing bracket, semicolon and newline
        End If
    Next i

    GenerateInsertStatements = sqlInsert
End Function

Sub WriteToFile(filename As String, text As String)
    Dim fso As Object
    Dim outputFile As Object

    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Open the file for output
    Set outputFile = fso.CreateTextFile(filename, True)

    outputFile.Write text

    ' Close the file
    outputFile.Close
End Sub

Function ReplaceSpecialCharacters(str As String) As String
    ' Lower case
    str = Replace(str, "ê", "e")
    str = Replace(str, "¹", "a")
    str = Replace(str, "ñ", "n")
    str = Replace(str, "¿", "z")
    str = Replace(str, "Ÿ", "z")
    str = Replace(str, "œ", "s")
    str = Replace(str, "æ", "c")
    str = Replace(str, "ó", "o")
    str = Replace(str, "³", "l")
    
    ' Upper case
    str = Replace(str, "Ê", "E")
    str = Replace(str, "¥", "A")
    str = Replace(str, "Ñ", "N")
    str = Replace(str, "¯", "Z")
    str = Replace(str, "", "Z")
    str = Replace(str, "Œ", "S")
    str = Replace(str, "Æ", "C")
    str = Replace(str, "Ó", "O")
    str = Replace(str, "£", "L")

    ReplaceSpecialCharacters = str
End Function

' Function to extract the base name from a file path
Function GetBaseName(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function
