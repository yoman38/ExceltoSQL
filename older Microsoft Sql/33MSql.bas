Attribute VB_Name = "Module1"
Option Explicit

' Constants for User Prompts
Const PROMPT_YES As String = "yes"
Const PROMPT_NO As String = "no"

    'Unique ID
    Dim addUniqueID As Boolean

    ' Workbook and Worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    
    'Main Initialize FileSystemObject
    Dim fso As Object

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
    
    'createtable
    Dim i As Long
    Dim j As Long
    Dim columnName As String
    Dim guessedType As String
    Dim maxStrLength As Long
    Dim numRows As Long
    
    'insert statement
    Dim filterCondition As Boolean
    Dim duplicateCondition As Boolean
    Dim emptyCondition As Boolean
    Dim cellValue As Variant
    Dim startColumn As Long
    
    'getusernumber
    Dim userNumber As Variant
     
     'Selectfolder
    Dim outputFileName As String
    
    'Write to file
    Dim outputFile As Object
    
    'selectworksheet
    Dim index As Variant
    Dim wsCount As Integer
    
    'guessdata
    Dim absValue As Double
    Dim strLength As Integer
    
    'display worksheet
    Dim msg As String
     
' Main Subroutine
Sub GenerateSQL()

    ' Initialize FileSystemObject
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
    Set rngData = SelectRange(ws, "Select the data range.")
    Set rngHeaders = SelectRange(ws, "Select the headers range.")


    ' Ask if the user wants to use filters
    useFilter = GetUserResponse("Do you want to skip rows without a specific keyword? (yes/no)", PROMPT_NO) = PROMPT_YES
    If useFilter Then
        filterKeyword = GetUserInput("Enter the keyword to filter rows.")
        Set rngFilter = SelectRange(ws, "Select the range for filtering.")
    End If

    ' Ask if the user wants to check for duplicates
    useDuplicatesCheck = GetUserResponse("Do you want to skip duplicate rows based on specific columns? (yes/no)", PROMPT_NO) = PROMPT_YES
    If useDuplicatesCheck Then
        Set rngDuplicateCheck = SelectRange(ws, "Select the range for checking duplicates.")
    End If

    ' Ask if the user wants to skip rows with empty cells in a specific range
    skipEmpty = GetUserResponse("Do you want to skip rows with empty cells in a specific range? (yes/no)", PROMPT_NO) = PROMPT_YES
    If skipEmpty Then
        Set rngEmptyCheck = SelectRange(ws, "Select the range for checking empty cells.")
    End If

      ' Ask if the user wants to include a unique ID column
    addUniqueID = GetUserResponse("Do you want to add a unique ID for each row in the table? (yes/no)", PROMPT_NO) = PROMPT_YES

    ' Generate SQL
    sqlCreate = GenerateCreateTable(rngHeaders, rngData, outputFileName)
    sqlInsert = GenerateInsertStatements(rngData, filterKeyword, outputFileName, rngFilter, rngDuplicateCheck, rngEmptyCheck)

    ' Write output to file
    WriteToFile outputFilePath, sqlCreate & vbCrLf & sqlInsert

    ' Close the workbook without saving changes
    wb.Close SaveChanges:=False

    ' Show success message
    MsgBox "SQL generation complete! Don't forget to change column types. The output was written to " & outputFilePath, vbInformation, "Success"
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
Function SelectRange(ws As Worksheet, prompt As String) As Range
    On Error GoTo ErrorHandler

    Set SelectRange = Application.InputBox(prompt, "Select Range", Type:=8) ' Type 8 = Range

    Exit Function
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Function
Function GenerateCreateTable(headers As Range, data As Range, outputFileName As String) As String

    ' Prompt the user for the number of rows to consider
    numRows = GetUserNumber("Enter the number of rows to consider for determining data type:", 1)

    ' If the user clicked Cancel, exit the function
    If numRows = -1 Then Exit Function
    
    ' Limit the number of rows to the actual number of rows in the data range
    If numRows > data.Rows.Count Then numRows = data.Rows.Count

    sqlCreate = "CREATE TABLE [" & outputFileName & "] ("
    
    ' Add a unique ID column if requested
    If addUniqueID Then sqlCreate = sqlCreate & "[Id] [int] IDENTITY(1,1) NOT NULL, "

    For i = 1 To headers.Columns.Count
        columnName = headers.Cells(1, i).value ' Get column name from headers
        If columnName = "" Then
            columnName = "UnnamedColumn" & i
        Else
            ' Replace special characters in the column name
            columnName = ReplaceSpecialCharacters(columnName)
        End If

        ' Get the datatype for all the rows in the column
        guessedType = "NVARCHAR(10)" ' Set default value
        maxStrLength = 0 ' Reset the maximum string length for each column
        For j = 1 To numRows
            Dim tempGuessedType As String
            tempGuessedType = GuessDataType(data.Cells(j, i).value)
            ' If it's a string, we update maxStrLength if current string is longer
            If TypeName(data.Cells(j, i).value) = "String" Then
                maxStrLength = Application.WorksheetFunction.Max(maxStrLength, Len(data.Cells(j, i).value))
                guessedType = "NVARCHAR(" & maxStrLength * 2 & ")"
            ElseIf tempGuessedType <> "NVARCHAR(10)" And tempGuessedType <> guessedType Then
                ' We assume that GuessDataType will always return a data type higher in hierarchy
                guessedType = tempGuessedType
            End If
        Next j

        sqlCreate = sqlCreate & "[" & columnName & "] " & guessedType & ", "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon

    GenerateCreateTable = sqlCreate
End Function


Function GenerateInsertStatements(rng As Range, filterKeyword As String, outputFileName As String, Optional filterRange As Range, Optional duplicateCheckRange As Range, Optional emptyCheckRange As Range) As String


    ' If includeId is False, start from the second column
    If addUniqueID Then
        startColumn = 1
    Else
        startColumn = 2
    End If
    
    For i = 1 To rng.Rows.Count
        If filterKeyword = "" Then
            filterCondition = True
        Else
            filterCondition = InStr(1, filterRange.Cells(i, 1).value, filterKeyword, vbTextCompare) > 0
        End If

        ' Add duplicate checking condition
        If duplicateCheckRange Is Nothing Then
            duplicateCondition = True
        Else
            ' Check if the cell value in the duplicateCheckRange is unique so far
            duplicateCondition = Application.WorksheetFunction.CountIf(duplicateCheckRange.Resize(i, 1), duplicateCheckRange.Cells(i, 1).value) = 1
        End If

        ' Add empty checking condition
        If emptyCheckRange Is Nothing Then
            emptyCondition = True
        Else
            ' Check if the cell value in the emptyCheckRange is not empty
            emptyCondition = Len(Trim(emptyCheckRange.Cells(i, 1).value)) <> 0
        End If

        ' Add the row to the SQL insert statement only if it satisfies all conditions
        If filterCondition And duplicateCondition And emptyCondition Then
            sqlInsert = sqlInsert & "INSERT INTO [" & outputFileName & "] VALUES ("
            For j = 1 To rng.Columns.Count
                ' Replace single quotes with two single quotes to escape them
                ' Use ReplaceSpecialCharacters function to replace polish special characters with latin ones
                cellValue = rng.Cells(i, j).value
                If IsDate(cellValue) Then
                    If TimeValue(cellValue) = 0 Then
                        cellValue = Format$(cellValue, "yyyy-mm-dd")
                    ElseIf DateValue(cellValue) = 0 Then
                        cellValue = Format$(cellValue, "hh:nn:ss")
                    Else
                        cellValue = Format$(cellValue, "yyyy-mm-ddThh:nn:ss")
                    End If
                End If
                sqlInsert = sqlInsert & "N'" & ReplaceSpecialCharacters(Replace(cellValue, "'", "''")) & "', "
            Next j
            ' Remove trailing comma and space, add closing bracket and semicolon
            sqlInsert = Left(sqlInsert, Len(sqlInsert) - 2) & ");" & vbCrLf
        End If
    Next i

    GenerateInsertStatements = sqlInsert
End Function



Sub WriteToFile(filename As String, text As String)

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
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function
Function SelectWorksheet(wb As Workbook) As Worksheet

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
    Dim index As Integer
    msg = "Please select a worksheet by entering its corresponding number:" & vbCrLf & vbCrLf
    index = 1

    For Each ws In wb.Sheets
        msg = msg & index & ". " & ws.Name & vbCrLf
        index = index + 1
    Next ws

    MsgBox msg, vbInformation, "Select Worksheet"
End Sub

Function GuessDataType(value As Variant) As String


    If IsNumeric(value) Then
        absValue = Abs(CDbl(value))
        
        If Int(absValue) = absValue Then ' value is integer
            Select Case absValue
                Case 0, 1
                    GuessDataType = "BIT"
                Case Is <= 255
                    GuessDataType = "TINYINT"
                Case Is <= 32767
                    GuessDataType = "SMALLINT"
                Case Is <= 2147483647
                    GuessDataType = "INT"
                Case Is <= 9.22337203685478E+18
                    GuessDataType = "BIGINT"
                Case Else
                    GuessDataType = "UNKNOWN"
            End Select
        Else ' value is decimal
            ' Guess decimal type
            Select Case absValue
                Case Is <= 3.4028235E+38
                    GuessDataType = "REAL"
                Case Is <= 1.7976931348623E+308
                    GuessDataType = "FLOAT"
                Case Else
                    GuessDataType = "UNKNOWN"
            End Select
        End If
    ElseIf IsDate(value) Then
        If TimeValue(value) = 0 Then
            GuessDataType = "DATE"
        ElseIf DateValue(value) = 0 Then
            GuessDataType = "TIME"
        Else
            GuessDataType = "DATETIME2"
        End If
    ElseIf TypeName(value) = "String" Then
        strLength = Len(value) * 2
        GuessDataType = "NVARCHAR(" & strLength & ")" ' Reflects the actual string length
    ElseIf IsEmpty(value) Then
        GuessDataType = "NULL"
    Else
        GuessDataType = "UNKNOWN"
    End If
End Function



Function GetUserNumber(prompt As String, defaultNumber As Long) As Long
    
    ' Keep asking for input until the user either provides valid input or cancels
    Do
        ' Prompt the user for a number
        userNumber = Application.InputBox(prompt, Type:=1, Default:=defaultNumber)

        ' If the user clicked Cancel, return -1
        If userNumber = False Then
            GetUserNumber = -1
            Exit Function
        End If
    Loop Until IsNumeric(userNumber) And userNumber >= 0

    GetUserNumber = CLng(userNumber)
End Function

