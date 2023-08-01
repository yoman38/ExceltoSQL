Attribute VB_Name = "Module2"
Option Explicit

Sub GenerateSQL()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim headers As Range
    Dim sqlCreate As String
    Dim sqlInsert As String
    Dim filename As String
    Dim outputFile As String

    ' Get the user's input
    filename = InputBox("Enter the path of the Excel file:", "Input Required", "C:\Users\name\Desktop\data\PROJECT3_work_schedule\ex.xls")
    outputFile = InputBox("Enter the path of the output text file:", "Input Required", "C:\Users\name\Desktop\output.txt")

    On Error GoTo ErrorHandler

    ' Open the workbook
    Set wb = Workbooks.Open(filename)
    Set ws = wb.Sheets("TP1 grafik brygad 2022-2023")

    ' Define the range of data manually
    Set rng = ws.Range(InputBox("Enter the range of data (e.g., 'F3:BK100'):", "Input Required", "F3:BK100"))

    ' Define the row of headers manually
    Set headers = ws.Range(InputBox("Enter the range of headers (e.g., 'F2:BK2'):", "Input Required", "F2:BK2"))

    ' Generate SQL
    sqlCreate = GenerateCreateTable(headers)
    sqlInsert = GenerateInsertStatements(rng)

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

    sqlCreate = "CREATE TABLE `test_vba` ("
    For i = 1 To headers.Columns.Count
        sqlCreate = sqlCreate & "[" & headers.Cells(1, i).Value & "] NVARCHAR(100), "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon

    GenerateCreateTable = sqlCreate
End Function

Function GenerateInsertStatements(rng As Range) As String
    Dim i As Long
    Dim j As Long
    Dim sqlInsert As String

    For i = 1 To rng.Rows.Count
        If InStr(1, rng.Cells(i, 2).Value, "zm", vbTextCompare) > 0 Then ' column G corresponds to the 2nd column in the range
            sqlInsert = sqlInsert & "INSERT INTO `test_vba` VALUES ("
            For j = 1 To rng.Columns.Count
                sqlInsert = sqlInsert & "'" & rng.Cells(i, j).Value & "', "
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

