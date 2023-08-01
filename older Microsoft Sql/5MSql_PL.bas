Attribute VB_Name = "MSql"
Option Explicit

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

    ' Get the user's input
    filename = InputBox("Enter the path of the Excel file:", "Input Required", "C:\Users\name\Desktop\data\PROJECT3_work_schedule\ex.xls")
    outputFile = InputBox("Enter the path of the output text file:", "Input Required", "C:\Users\name\Desktop\output.txt")
    useFilter = InputBox("Do you want to skip rows without a specific keyword? (yes/no)", "Input Required", "no")

    If LCase(useFilter) = "yes" Then
        filterKeyword = InputBox("Enter the keyword to filter rows:", "Input Required", "zm")
    Else
        filterKeyword = "" ' No filtering
    End If

    On Error GoTo ErrorHandler

    ' Open the workbook
    Set wb = Workbooks.Open(filename)
    Set ws = wb.Sheets("TP1 grafik brygad 2022-2023")

    checkDuplicates = InputBox("Do you want to skip duplicate rows based on specific columns? (yes/no)", "Input Required", "no")

    If LCase(checkDuplicates) = "yes" Then
        duplicateCheckColumns = InputBox("Enter the range of columns to check for duplicates (e.g., 'F3:H100'):", "Input Required", "F3:H100")
        Set duplicateCheckRange = ws.Range(duplicateCheckColumns)
    End If

    ' Define the range of data manually
    Set rng = ws.Range(InputBox("Enter the range of data (e.g., 'F3:BK100'):", "Input Required", "F3:BK100"))

    ' Define the row of headers manually
    Set headers = ws.Range(InputBox("Enter the range of headers (e.g., 'F2:BK2'):", "Input Required", "F2:BK2"))

    ' Define the range of column for keyword filtering manually
    If filterKeyword <> "" Then
        Set filterRange = ws.Range(InputBox("Enter the range of column for keyword filtering (e.g., 'G3:G100'):", "Input Required", "G3:G100"))
    End If

    ' Generate SQL
    sqlCreate = GenerateCreateTable(headers)
    sqlInsert = GenerateInsertStatements(rng, filterKeyword, filterRange, duplicateCheckRange)

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

Function GenerateCreateTable(headers As Range) As String
    Dim i As Long
    Dim sqlCreate As String
    Dim columnName As String
    Dim dict As Object
    Dim suffix As Long

    Set dict = CreateObject("Scripting.Dictionary")

    sqlCreate = "CREATE TABLE [test_vba] ("
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
        sqlCreate = sqlCreate & "[" & columnName & "] NVARCHAR(100), "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon

    GenerateCreateTable = sqlCreate
End Function

Function GenerateInsertStatements(rng As Range, filterKeyword As String, Optional filterRange As Range, Optional duplicateCheckRange As Range) As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sqlInsert As String
    Dim filterCondition As Boolean
    Dim duplicateCheck As String
    Dim duplicateDict As Object

    Set duplicateDict = CreateObject("Scripting.Dictionary")

    For i = 1 To rng.Rows.Count
        If filterKeyword = "" Then
            filterCondition = True
        Else
            filterCondition = InStr(1, filterRange.Cells(i, 1).Value, filterKeyword, vbTextCompare) > 0
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

        If filterCondition And Not duplicateDict.Exists(duplicateCheck) Then
            duplicateDict.Add duplicateCheck, 0
            sqlInsert = sqlInsert & "INSERT INTO [test_vba] VALUES ("
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

