Attribute VB_Name = "ExcelSQL"


Option Explicit

' Constants for User Prompts
Const PROMPT_YES As String = "yes"
Const PROMPT_NO As String = "no"

    'Unique ID
    Dim addUniqueID As Boolean

    ' Workbook and Worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim userResponse As VbMsgBoxResult

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
    
    'columntype dictionary
    Dim columnDataTypes As Object
     
' Main Subroutine
Sub GenerateSQL()

    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Initialize Workbook and Worksheet variables
    Set wb = Nothing
    Set ws = Nothing
    
    ' Initialize Ranges for data and headers
    Set rngData = Nothing
    Set rngHeaders = Nothing
    
    ' Initialize Filter variables
    filterKeyword = ""
    Set rngFilter = Nothing
    useFilter = False

    ' Initialize Duplicate check variables
    Set rngDuplicateCheck = Nothing
    useDuplicatesCheck = False

    ' Initialize Empty cells check variables
    Set rngEmptyCheck = Nothing
    skipEmpty = False

    ' Initialize SQL statements
    sqlCreate = ""
    sqlInsert = ""

    ' Initialize File paths
    inputFilePath = ""
    outputFilePath = ""
    
    ' Initialize createtable variables
    i = 0
    j = 0
    columnName = ""
    guessedType = ""
    maxStrLength = 0
    numRows = 0
    
    ' Initialize insert statement variables
    filterCondition = False
    duplicateCondition = False
    emptyCondition = False
    cellValue = Empty
    startColumn = 0
    
    ' Initialize getusernumber variable
    userNumber = Empty
     
    ' Initialize Selectfolder variable
    outputFileName = ""
    ' Extract filename from outputFilePath and assign to outputFileName
    

    
    ' Initialize Write to file variable
    Set outputFile = Nothing
    
    ' Initialize selectworksheet variables
    index = Empty
    wsCount = 0
    
    ' Initialize guessdata variables
    absValue = 0
    strLength = 0
    
    ' Initialize display worksheet variable
    msg = ""
    ' Get inputs
    inputFilePath = SelectFile
    If inputFilePath = "" Then Exit Sub
    outputFilePath = SelectFolder
    outputFileName = fso.GetBaseName(outputFilePath)
    If outputFilePath = "" Then Exit Sub

    ' Open workbook and set local worksheet variable
    Set wb = Workbooks.Open(inputFilePath)

    ' Choose the worksheet
    Set ws = SelectWorksheet(wb)
    If ws Is Nothing Then Exit Sub

    ' Get the range for data and headers
    Set rngData = SelectRange(ws, "Select the data range.")
    Set rngHeaders = SelectRange(ws, "Select the headers range. (Try to give it the same number of columns as the data range)")

    ' Ask if the user wants to use filters
    useFilter = MsgBox("Do you want to skip rows without a specific keyword?", vbYesNo) = vbYes
    If useFilter Then
        filterKeyword = GetUserInput("Enter the keyword to filter rows.")
        Set rngFilter = SelectRange(ws, "Select the range for filtering. (Try to give it the same number of rows as the data range)")
    End If
    
    ' Ask if the user wants to check for duplicates
    useDuplicatesCheck = MsgBox("Do you want to skip duplicate rows based on specific columns?", vbYesNo) = vbYes
    If useDuplicatesCheck Then
        Set rngDuplicateCheck = SelectRange(ws, "Select the range for checking duplicates. (Try to give it the same number of rows as the data range)")
    End If
    
    ' Ask if the user wants to skip rows with empty cells in a specific range
    skipEmpty = MsgBox("Do you want to skip rows with empty cells in a specific range?", vbYesNo) = vbYes
    If skipEmpty Then
        Set rngEmptyCheck = SelectRange(ws, "Select the range for checking empty cells. (Try to give it the same number of rows as the data range)")
    End If


    ' Ask if the user wants to include a unique ID column
    addUniqueID = MsgBox("Do you want to add a unique ID for each row in the table?", vbYesNo) = vbYes


    ' Generate SQL
    sqlCreate = GenerateCreateTable(rngHeaders, rngData, outputFileName)
    sqlInsert = GenerateInsertStatements(rngData, filterKeyword, outputFileName, rngFilter, rngDuplicateCheck, rngEmptyCheck)
    
    ' Check if outputFileName is empty
    If outputFileName = "" Then
        outputFileName = InputBox("Please enter the output filename:", "Filename Required")
        If outputFileName = "" Then
            MsgBox "No filename provided. Exiting subroutine.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Now call the GetAdditionalTableDetails function
    sqlCreate = sqlCreate & vbCrLf & GetAdditionalTableDetails(rngHeaders, outputFileName)

    ' Write output to file
    WriteToFile outputFilePath, sqlCreate & vbCrLf & sqlInsert

    ' Close the workbook without saving changes
    wb.Close SaveChanges:=False

    ' Show success message
    MsgBox "SQL query generation complete! Don't forget to change column types. The output was written to " & outputFilePath, vbInformation, "Success"
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
    SelectFile = gSelectedExcelFile
End Function


' Function to show a folder selection dialog and return the selected folder path

Function SelectFolder() As String
    Dim outputFileName As String
    Dim folderPath As String
    
    folderPath = Left(gSelectedExcelFile, InStrRev(gSelectedExcelFile, "\") - 1) ' Extract directory from gSelectedExcelFile
    
    outputFileName = GetUserInput("Enter the output filename.")
    If outputFileName = "" Then Exit Function ' Exit function if user clicked Cancel
    
    SelectFolder = folderPath & "\" & outputFileName & ".txt" ' Here you can change the extension
End Function

' Function to open a workbook and return the Workbook object
Function OpenWorkbook(filePath As String) As Workbook
    Set OpenWorkbook = Workbooks.Open(filePath)
End Function

Function SelectRange(ws As Worksheet, prompt As String) As Range

    ' Activate the worksheet
    ws.Activate
    
    On Error Resume Next ' Ignore error if user clicks Cancel or selects an invalid range
    Set SelectRange = Application.InputBox(prompt, "Select Range", Type:=8) ' Type 8 = Range
    On Error GoTo 0 ' Revert to normal error handling

    ' If the selected range is on another sheet, ask the user to select a range on the correct sheet
    If Not SelectRange Is Nothing Then
        If Not SelectRange.Parent Is ws Then
            MsgBox "Please select a range on the correct worksheet (" & ws.Name & ").", vbInformation, "Wrong Worksheet"
            Set SelectRange = SelectRange(ws, prompt) ' Recursively call the function until a valid range is selected
        End If
    End If
End Function

Function GenerateCreateTable(headers As Range, data As Range, outputFileName As String) As String

    Dim sqlCreate As String
    Dim addUniqueID As Boolean
    Dim dict As Object
    Dim suffix As Long
    Dim i As Integer
    Dim columnName As String
    Dim guessedType As String
    Dim previousType As String
    Dim hasChanged As Boolean
    Dim maxStrLength As Integer
    Dim j As Integer
    Dim tempGuessedType As String
    Dim numRows As Integer
    
    Set columnDataTypes = CreateObject("Scripting.Dictionary")



    ' Prompt the user for the number of rows to consider
    numRows = GetUserNumber("Enter the number of rows to consider for determining data type:", 1)

    ' If the user clicked Cancel, exit the function
    If numRows = -1 Then Exit Function
    
    ' Limit the number of rows to the actual number of rows in the data range
    If numRows > data.Rows.Count Then numRows = data.Rows.Count

    sqlCreate = "CREATE TABLE [" & outputFileName & "] ("
    
    ' Add a unique ID column if requested
    If addUniqueID Then sqlCreate = sqlCreate & "[Id] [int] IDENTITY(1,1) NOT NULL, "
    
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To headers.Columns.Count
        columnName = headers.Cells(1, i).value ' Get column name from headers
        If columnName = "" Then
            columnName = "UnnamedColumn" & i
        Else
            ' Replace special characters in the column name
            columnName = ReplaceSpecialCharacters(columnName)
            If dict.Exists(columnName) Then
                suffix = dict(columnName) + 1
                dict(columnName) = suffix
                columnName = columnName & "_" & suffix
            Else
                dict.Add columnName, 0
            End If
        End If
        columnDataTypes(columnName) = guessedType


        ' Get the datatype for all the rows in the column
        guessedType = "NVARCHAR(10)" ' Set default value
        previousType = "" ' To keep track of the previous guessed type
        hasChanged = False ' To keep track of whether the type has changed
        maxStrLength = 0 ' Reset the maximum string length for each column
For j = 1 To numRows
    If Not IsEmpty(data.Cells(j, i).value) Then ' Skip empty cells
        tempGuessedType = GuessDataType(data.Cells(j, i).value)
        ' If it's a string, we update maxStrLength if current string is longer
        If TypeName(data.Cells(j, i).value) = "String" Then
            maxStrLength = Application.WorksheetFunction.Max(maxStrLength, Len(data.Cells(j, i).value))
            guessedType = "NVARCHAR(" & maxStrLength * 2 & ")"
            hasChanged = True ' Setting hasChanged to True ensures that NVARCHAR is selected
        ElseIf tempGuessedType <> "NVARCHAR(10)" And tempGuessedType <> guessedType Then

                    ' Check if the type has changed, except if it is within numeric or string types
                    If previousType <> "" And previousType <> tempGuessedType And _
                       Not ((previousType = "TINYINT" And tempGuessedType = "SMALLINT") Or _
                            (previousType = "SMALLINT" And tempGuessedType = "INT")) And _
                       Not (Left(previousType, 8) = "NVARCHAR" And Left(tempGuessedType, 8) = "NVARCHAR") Then
                        hasChanged = True
                    End If
                    ' We assume that GuessDataType will always return a data type higher in hierarchy
                    guessedType = tempGuessedType
                End If
                previousType = tempGuessedType
            End If
        Next j

        ' If the type has changed, set it to NVARCHAR(maxStrLength*2)
        If hasChanged Then guessedType = "NVARCHAR(" & maxStrLength * 2 & ")"
        
        
        columnDataTypes(columnName) = guessedType
sqlCreate = sqlCreate & "[" & columnName & "]" & guessedType & ", "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");"

 ' Remove trailing comma and space, add closing bracket and semicolon

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


Sub WriteToFile(path As String, text As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outputFile As Object
    ' Make sure the second parameter is True to overwrite existing files
    Set outputFile = fso.CreateTextFile(path, True)
    outputFile.Write text
    outputFile.Close
    Set outputFile = Nothing
    Set fso = Nothing
End Sub

' Function to replace special characters in a string
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
    Dim index As Integer
    Dim msg As String
    Dim wsCount As Integer
    
    msg = "Please select a worksheet by entering its corresponding number:" & vbCrLf & vbCrLf
    index = 1

    For Each ws In wb.Sheets
        msg = msg & index & ". " & ws.Name & vbCrLf
        index = index + 1
    Next ws

    index = Application.InputBox(msg, "Select Worksheet", Type:=1) ' Type 1 = Number
    
    ' Exit function if user clicked Cancel
    If index = False Then Exit Function

    wsCount = wb.Sheets.Count
    If index < 1 Or index > wsCount Then
        MsgBox "Invalid worksheet number. Please enter a number between 1 and " & wsCount & ".", vbCritical, "Error"
        Set SelectWorksheet = Nothing
    Else
        Set SelectWorksheet = wb.Sheets(index)
    End If
End Function



Function GuessDataType(value As Variant) As String


    If IsNumeric(value) Then
        absValue = Abs(CDbl(value))
        
        If Int(absValue) = absValue Then ' value is integer
            Select Case absValue
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
Function GetAdditionalTableDetails(rngHeaders As Range, outputFileName As String) As String
    Dim result As String
    Dim userResponse As VbMsgBoxResult
    Dim columnName As String
    Dim referenceTable As String
    Dim referenceColumn As String
    Dim constraintType As String
    Dim defaultValue As String
    Dim selectedCells As Range
    Dim columnNames() As String
    Dim i As Integer

    
        ' Ask the user about Primary Key
        Dim response As VbMsgBoxResult
        response = MsgBox("Do you want to set a Primary Key for the table?", vbYesNo, "Primary Key Confirmation")
        
        If response = vbYes Then
            ' Prompt the user to select the headers cells for the Primary Key columns
            Set selectedCells = Application.InputBox("Select the header cells for the Primary Key columns:", "Primary Key", Type:=8)
            
        If Not selectedCells Is Nothing Then
            ReDim columnNames(1 To selectedCells.Cells.Count)
            For i = 1 To selectedCells.Cells.Count
                columnName = selectedCells.Cells(i).value
            columnNames(i) = columnName
                ' Generate the ALTER COLUMN statement first for each column
                result = result & "ALTER TABLE [" & outputFileName & "] ALTER COLUMN " & "[" & columnName & "] " & columnDataTypes(columnName) & " NOT NULL;" & vbCrLf
                MsgBox columnName & " has been set to NOT NULL constraint.", vbInformation, "Information"
            Next i
                
                ' Generate the ALTER TABLE ADD PRIMARY KEY statement using all the selected columns
                result = result & "ALTER TABLE [" & outputFileName & "] ADD PRIMARY KEY (" & Join(columnNames, ",") & ");" & vbCrLf
            End If
        End If


    ' Ask the user about the Foreign Key
    userResponse = MsgBox("Do you want to add a Foreign Key?", vbYesNo, "Foreign Key")
    
    If userResponse = vbYes Then
    ' Prompt the user to select the header cell for the Foreign Key column
    Set selectedCells = Application.InputBox("Select the header cell for the Foreign Key column:", "Foreign Key", Type:=8)
    
    If Not selectedCells Is Nothing Then
        ' Get the column name, reference table name, and reference column name from the user
        columnName = selectedCells.Cells(1, 1).value
        referenceTable = InputBox("Enter the reference table name for the Foreign Key:", "Foreign Key")
        referenceColumn = InputBox("Enter the reference column name from the other table:", "Foreign Key")
    
        ' Generate the ALTER COLUMN statement first
        result = result & "ALTER TABLE [" & outputFileName & "] ALTER COLUMN " & "[" & columnName & "] " & columnDataTypes(columnName) & " NOT NULL;" & vbCrLf
    
        ' Generate the ALTER TABLE ADD FOREIGN KEY statement
        result = result & "ALTER TABLE [" & outputFileName & "] ADD FOREIGN KEY (" & columnName & ") REFERENCES " & referenceTable & "(" & referenceColumn & ");" & vbCrLf
    End If
    End If


    ' Ask user if they want to set other columns to NOT NULL
    Dim anotherResponse As VbMsgBoxResult
    Dim continueSetting As Boolean
    
    continueSetting = True
    
    Do While continueSetting
        anotherResponse = MsgBox("Do you want to set other columns to NOT NULL?", vbYesNo, "NOT NULL Setting")
        
        If anotherResponse = vbYes Then
            ' Prompt the user to select the header cells for columns they want to set as NOT NULL
            Set selectedCells = Application.InputBox("Select the header cells for columns to set as NOT NULL:", "NOT NULL Setting", Type:=8)
            
            If Not selectedCells Is Nothing Then
                For i = 1 To selectedCells.Cells.Count
                    columnName = selectedCells.Cells(i).value
                    ' Generate the ALTER COLUMN statement for each selected column
                    result = result & "ALTER TABLE [" & outputFileName & "] ALTER COLUMN " & "[" & columnName & "] " & columnDataTypes(columnName) & " NOT NULL;" & vbCrLf
                    MsgBox columnName & " has been set to NOT NULL constraint.", vbInformation, "Information"
                Next i
            End If
        Else
            continueSetting = False
        End If
    Loop

    ' Ask user about Indexes
    userResponse = MsgBox("Do you want to add an Index?", vbYesNo, "Indexes")
        
    If userResponse = vbYes Then
        ' Prompt the user to select the header cell for the Index column
        Set selectedCells = Application.InputBox("Select the header cell for the Index column:", "Indexes", Type:=8)
        
        If Not selectedCells Is Nothing Then
            columnName = selectedCells.Cells(1, 1).value
            result = result & "CREATE INDEX idx_" & columnName & " ON [" & outputFileName & "] (" & columnName & ");" & vbCrLf
        End If
    End If
    
    ' Ask user about constraints
    Do
        userResponse = MsgBox("Do you want to add a Constraint?", vbYesNo, "Constraints")
        If userResponse = vbNo Then Exit Do
        Set selectedCells = Application.InputBox("Select the header cell for the Constraint column:", "Constraints", Type:=8)
        If Not selectedCells Is Nothing Then
            columnName = selectedCells.Cells(1, 1).value
            constraintType = InputBox("Enter the type of constraint (e.g. UNIQUE, CHECK):", "Constraints")
            result = result & "ALTER TABLE [" & outputFileName & "] ADD CONSTRAINT UC_" & columnName & " " & constraintType & " ([" & columnName & "]);" & vbCrLf
        End If
    Loop While userResponse = vbYes

    ' Ask user about default
    Do
        userResponse = MsgBox("Do you want to set a Default Value for any column?", vbYesNo, "Default Values")
        If userResponse = vbNo Then Exit Do
        Set selectedCells = Application.InputBox("Select the header cell for which you want to set the Default Value:", "Default Values", Type:=8)
        If Not selectedCells Is Nothing Then
            columnName = selectedCells.Cells(1, 1).value
            defaultValue = InputBox("Enter the default value:", "Default Values")
            result = result & "ALTER TABLE [" & outputFileName & "] ADD CONSTRAINT DF_" & outputFileName & "_" & columnName & " DEFAULT '" & defaultValue & "' FOR " & columnName & ";" & vbCrLf
        End If
    Loop While userResponse = vbYes

    GetAdditionalTableDetails = result
End Function

