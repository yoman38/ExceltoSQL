Attribute VB_Name = "Module5"
Option Explicit

' Constants for User Prompts
Const PROMPT_YES As String = "yes"
Const PROMPT_NO As String = "no"

' Main Subroutine
Sub GenerateSQL()
    ' Workbook and Worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet

    ' Ranges for data and headers
    Dim rngData As Range
    Dim rngHeaders As Range

    ' Filter variables
    Dim filterKeyword As String
    Dim rngFilter As Range
    Dim useFilter As Boolean

    ' Duplicate check variables
    Dim rngDuplicateCheck As Range
    Dim useDuplicatesCheck As Boolean

    ' Empty cells check variables
    Dim rngEmptyCheck As Range
    Dim skipEmpty As Boolean

    ' SQL statements
    Dim sqlCreate As String
    Dim sqlInsert As String

    ' File paths
    Dim inputFilePath As String
    Dim outputFilePath As String

    ' Initialize FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get inputs
    inputFilePath = SelectFile
    If inputFilePath = "" Then Exit Sub
    outputFilePath = SelectFolder
    If outputFilePath = "" Then Exit Sub
    Set wb = Workbooks.Open(inputFilePath)
    Set ws = SelectWorksheet(wb)
    If ws Is Nothing Then Exit Sub

    ' Get the range for data and headers
    Set rngData = SelectRange("Select the data range.")
    Set rngHeaders = SelectRange("Select the headers range.")

    ' Ask if the user wants to use filters
    useFilter = GetUserResponse("Do you want to skip rows without a specific keyword? (yes/no)", PROMPT_NO) = PROMPT_YES
    If useFilter Then
        filterKeyword = GetUserInput("Enter the keyword to filter rows.")
        Set rngFilter = SelectRange("Select the range for filtering.")
    End If

    ' Ask if the user wants to check for duplicates
    useDuplicatesCheck = GetUserResponse("Do you want to skip duplicate rows based on specific columns? (yes/no)", PROMPT_NO) = PROMPT_YES
    If useDuplicatesCheck Then
        Set rngDuplicateCheck = SelectRange("Select the range for checking duplicates.")
    End If

    ' Ask if the user wants to skip rows with empty cells in a specific range
    skipEmpty = GetUserResponse("Do you want to skip rows with empty cells in a specific range? (yes/no)", PROMPT_NO) = PROMPT_YES
    If skipEmpty Then
        Set rngEmptyCheck = SelectRange("Select the range for checking empty cells.")
    End If

    ' Generate SQL
    sqlCreate = GenerateCreateTable(rngHeaders, fso.GetBaseName(inputFilePath))
    sqlInsert = GenerateInsertStatements(rngData, filterKeyword, fso.GetBaseName(inputFilePath), rngFilter, rngDuplicateCheck, rngEmptyCheck)

    ' Write output to file
    WriteToFile outputFilePath, sqlCreate & vbCrLf & sqlInsert

    ' Close the workbook without saving changes
    wb.Close SaveChanges:=False

    ' Show success message
    MsgBox "SQL generation complete! Don't forget to change column types (default=NVARCHAR(50). The output was written to " & outputFilePath, vbInformation, "Success"
End Sub

' Function to show an input box and return the user's input
Function GetUserInput(prompt As String, Optional defaultText As String = "") As String
    GetUserInput = InputBox(prompt, "Input Required", defaultText)
End Function

' Function to show a yes/no input box and return the user's response
Function GetUserResponse(prompt As String, Optional defaultText As String = "") As String
    GetUserResponse = Application.InputBox(prompt, "Input Required", defaultText, Type:=2) ' Type 2 = Text
End Function

' Function to show a file selection dialog and return the selected file path
Function SelectFile() As String
    MsgBox "Please select the file to extract datas from.", vbInformation, "Select file"
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Excel file."
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            SelectFile = .SelectedItems(1)
        Else
            SelectFile = "" ' Return empty string if user clicked Cancel
        End If
    End With
End Function

' Function to show a folder selection dialog and return the selected folder path
Function SelectFolder() As String
    Dim outputFileName As String
    MsgBox "Please select the folder where the output SQL file should be created.", vbInformation, "Select Folder"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the output folder."
        If .Show = -1 Then
            outputFileName = GetUserInput("Enter the output filename.")
            If outputFileName = "" Then Exit Function ' Exit function if user clicked Cancel
            SelectFolder = .SelectedItems(1) & "\" & outputFileName & ".txt" 'Here you can change the extension
        Else
            SelectFolder = "" ' Return empty string if user clicked Cancel
        End If
    End With
End Function

' Function to open a workbook and return the Workbook object
Function OpenWorkbook(filePath As String) As Workbook
    Set OpenWorkbook = Workbooks.Open(filePath)
End Function

' Function to show a range selection dialog and return the selected Range object
Function SelectRange(prompt As String) As Range
    Dim ws As Worksheet
    Dim wsName As String

    On Error GoTo ErrorHandler

    wsName = GetUserInput("Enter the name of the worksheet where you want to " & prompt)
    If wsName = "" Then Exit Function ' Exit function if user clicked Cancel

    Set ws = ThisWorkbook.Worksheets(wsName)
    Set SelectRange = Application.InputBox(prompt, "Select Range", Type:=8) ' Type 8 = Range

    Exit Function
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Function


Function GenerateCreateTable(headers As Range, tableName As String) As String
    Dim i As Long
    Dim sqlCreate As String
    Dim columnName As String

    sqlCreate = "CREATE TABLE [" & tableName & "] ("
    For i = 1 To headers.Columns.Count
        columnName = headers.Cells(1, i).Value ' Get column name from headers
        If columnName = "" Then
            columnName = "UnnamedColumn" & i
        End If
        sqlCreate = sqlCreate & "[" & columnName & "] NVARCHAR(50), "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon

    GenerateCreateTable = sqlCreate
End Function

Function GenerateInsertStatements(rng As Range, filterKeyword As String, tableName As String, Optional filterRange As Range, Optional duplicateCheckRange As Range, Optional emptyCheckRange As Range) As String
    Dim i As Long
    Dim j As Long
    Dim sqlInsert As String
    Dim filterCondition As Boolean
    Dim duplicateCondition As Boolean
    Dim emptyCondition As Boolean

    For i = 1 To rng.Rows.Count
        If filterKeyword = "" Then
            filterCondition = True
        Else
            filterCondition = InStr(1, filterRange.Cells(i, 1).Value, filterKeyword, vbTextCompare) > 0
        End If

        ' Add duplicate checking condition
        If duplicateCheckRange Is Nothing Then
            duplicateCondition = True
        Else
            ' Check if the cell value in the duplicateCheckRange is unique so far
            duplicateCondition = Application.WorksheetFunction.CountIf(duplicateCheckRange.Resize(i, 1), duplicateCheckRange.Cells(i, 1).Value) = 1
        End If

        ' Add empty checking condition
        If emptyCheckRange Is Nothing Then
            emptyCondition = True
        Else
            ' Check if the cell value in the emptyCheckRange is not empty
            emptyCondition = Len(Trim(emptyCheckRange.Cells(i, 1).Value)) <> 0
        End If

        ' Add the row to the SQL insert statement only if it satisfies all conditions
        If filterCondition And duplicateCondition And emptyCondition Then
            sqlInsert = sqlInsert & "INSERT INTO [" & tableName & "] VALUES ("
            For j = 1 To rng.Columns.Count
                ' Replace single quotes with two single quotes to escape them
                ' Use ReplaceSpecialCharacters function to replace polish special characters with latin ones
                sqlInsert = sqlInsert & "N'" & ReplaceSpecialCharacters(Replace(rng.Cells(i, j).Value, "'", "''")) & "', "
            Next j
            ' Remove trailing comma and space, add closing bracket and semicolon
            sqlInsert = Left(sqlInsert, Len(sqlInsert) - 2) & ");" & vbCrLf
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
    str = Replace(str, "Í", "e")
    str = Replace(str, "π", "a")
    str = Replace(str, "Ò", "n")
    str = Replace(str, "ø", "z")
    str = Replace(str, "ü", "z")
    str = Replace(str, "ú", "s")
    str = Replace(str, "Ê", "c")
    str = Replace(str, "Û", "o")
    str = Replace(str, "≥", "l")
    
    ' Upper case
    str = Replace(str, " ", "E")
    str = Replace(str, "•", "A")
    str = Replace(str, "—", "N")
    str = Replace(str, "Ø", "Z")
    str = Replace(str, "è", "Z")
    str = Replace(str, "å", "S")
    str = Replace(str, "∆", "C")
    str = Replace(str, "”", "O")
    str = Replace(str, "£", "L")

    ReplaceSpecialCharacters = str
End Function
' Function to extract the base name from a file path
Function GetBaseName(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function
Function SelectWorksheet(wb As Workbook) As Worksheet
    Dim index As Variant
    Dim wsCount As Integer

    DisplayWorksheets wb

    index = GetUserInput("Enter the number of the worksheet you want to extract data from.")
    If index = "" Then Exit Function ' Exit function if user clicked Cancel

    ' check if index is a number
    If IsNumeric(index) Then
        index = CInt(index)
    Else
        MsgBox "You must enter a numeric value.", vbCritical, "Error"
        Exit Function
    End If

    wsCount = wb.Sheets.Count
    If index < 1 Or index > wsCount Then
        MsgBox "Invalid worksheet number. Please enter a number between 1 and " & wsCount & ".", vbCritical, "Error"
        Set SelectWorksheet = Nothing
    Else
        Set SelectWorksheet = wb.Sheets(index)
    End If
End Function



' Function to display all worksheet names
Sub DisplayWorksheets(wb As Workbook)
    Dim ws As Worksheet
    Dim msg As String
    Dim index As Integer

    msg = "Please select a worksheet by entering its corresponding number:" & vbCrLf & vbCrLf
    index = 1

    For Each ws In wb.Sheets
        msg = msg & index & ". " & ws.name & vbCrLf
        index = index + 1
    Next ws

    MsgBox msg, vbInformation, "Select Worksheet"
End Sub

