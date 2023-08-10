Attribute VB_Name = "TPtoPIVOT"


' Declare a public variable to store the path of the selected Excel file
Public gSelectedExcelFile As String

' Subroutine to get user input and run other subroutines accordingly
Function GetUserInputAndRunSubroutines() As Boolean


    ' Ask the user to select the source file
    Dim srcFile As Variant
    srcFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", _
                                          Title:="Please select the source Excel file")
    
    ' Exit if the user didn't choose a file
    If srcFile = False Then
        MsgBox "No file selected. Exiting..."
        GetUserInputAndRunSubroutines = False
        Exit Function
    End If

    ' Open the selected workbook
    Dim srcWorkbook As Workbook
    Set srcWorkbook = Workbooks.Open(srcFile)
    gSelectedExcelFile = srcFile


    ' Get the name of the source worksheet from the user
    Dim ws As Worksheet
    Dim wsNames() As String
    Dim wsNum As Long
    Dim i As Long
    ReDim wsNames(1 To srcWorkbook.Sheets.Count)
    For Each ws In srcWorkbook.Sheets
        i = i + 1
        wsNames(i) = i & ". " & ws.Name
    Next ws
    
    Dim inputResult As Variant

    inputResult = InputBox("Please enter the number of the worksheet you want to use as the source:" & vbNewLine & _
                          Join(wsNames, vbNewLine), "Input needed", "1")
    
    If inputResult = "" Then
        MsgBox "Operation cancelled by user."
        GetUserInputAndRunSubroutines = False
        Exit Function
    End If
    
    If Not IsNumeric(inputResult) Then
        MsgBox "Please enter a valid worksheet number."
        Exit Function
    Else
        wsNum = CLng(inputResult)
    End If


    ' Activate the selected source worksheet
    srcWorkbook.Sheets(wsNum).Activate


    ' Get the last row from the user
    Dim lastRow As Long
    
    Dim lastRowRange As Range
    On Error Resume Next
    
    ' Ensure the user is still on the selected source worksheet and select the last row
    If Not ActiveSheet Is srcWorkbook.Sheets(wsNum) Then
        MsgBox "Please stay on the selected source worksheet!", vbExclamation, "Warning"
        Exit Function
    End If
    
    ' Find the last row with content in the first column
    Dim actualLastRow As Long

    actualLastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
    Set lastRowRange = Application.InputBox("Please click on the last row of the table:", Type:=8, Default:=ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Address)
    
    If lastRowRange Is Nothing Then
        MsgBox "Operation cancelled by user."
        GetUserInputAndRunSubroutines = False
        Exit Function
    End If
    
    lastRow = lastRowRange.Row

    ' Get the column of "squad" from the user
    Dim squadColRange As Range
    Dim squadCol As String
    
    On Error Resume Next
    Set squadColRange = Application.InputBox("Please click on the 'Brygada' header cell or column:", Type:=8, Default:=ActiveSheet.Cells(1, 1).Address)
    On Error GoTo 0
    
    If squadColRange Is Nothing Then
        MsgBox "Operation cancelled by user."
        GetUserInputAndRunSubroutines = False
        Exit Function
    Else
        squadCol = Split(squadColRange.Address, "$")(1)
    End If
    
    ' If all inputs were successful, return True
    GetUserInputAndRunSubroutines = True
    
    ' Call both subroutines with the user input
    PivotData wsNum, lastRow, squadCol, srcWorkbook
    PivotDataUnique wsNum, lastRow, squadCol, srcWorkbook

    ' Check for empty cells in the workbook
    CheckForEmptyCells srcWorkbook
    
End Function




Sub PivotData(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String, srcWorkbook As Workbook)

    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet, destSheet2 As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim Name As String
    Dim monthYear As String
    Dim day As Variant, data As Variant
    Dim shift As String
    Dim month As Integer, year As String
    Dim monthDict As Object


    ' Dictionary to map Polish month names to month numbers
    Set monthDict = CreateObject("Scripting.Dictionary")
    monthDict("styczeñ") = 1
    monthDict("luty") = 2
    monthDict("marzec") = 3
    monthDict("kwiecieñ") = 4
    monthDict("maj") = 5
    monthDict("czerwiec") = 6
    monthDict("lipiec") = 7
    monthDict("sierpieñ") = 8
    monthDict("wrzesieñ") = 9
    monthDict("paŸdziernik") = 10
    monthDict("listopad") = 11
    monthDict("grudzieñ") = 12

    ' Convert column letter to column number for 'Brygada'
    Dim squadColNum As Long
    squadColNum = Range(squadCol & "1").Column
    
    ' Set source worksheet
    Set srcSheet = srcWorkbook.Sheets(wsNum)
        
    ' Check if "WorkersShifts" and "WorkersMonthData" sheets exist. If not, create them.
    Dim ws As Worksheet
    Dim sheetExists As Boolean, sheetExists2 As Boolean
    sheetExists = False
    sheetExists2 = False
    For Each ws In srcWorkbook.Sheets
        If ws.Name = "WorkersShifts" Then
            sheetExists = True
        ElseIf ws.Name = "WorkersMonthData" Then
            sheetExists2 = True
        End If
    Next ws
    
    If Not sheetExists Then
        Set destSheet = srcWorkbook.Sheets.Add(After:=srcWorkbook.Sheets(srcWorkbook.Sheets.Count))
        destSheet.Name = "WorkersShifts"
    Else
        Set destSheet = srcWorkbook.Sheets("WorkersShifts")
    End If
    
    If Not sheetExists2 Then
        Set destSheet2 = srcWorkbook.Sheets.Add(After:=srcWorkbook.Sheets(srcWorkbook.Sheets.Count))
        destSheet2.Name = "WorkersMonthData"
    Else
        Set destSheet2 = srcWorkbook.Sheets("WorkersMonthData")
    End If
    
    ' Calculate relative column positions
    Dim nameCol As Long, monthYearCol As Long, dataStartCol As Long, dataEndCol As Long, shiftStartCol As Long, shiftEndCol As Long
    nameCol = squadColNum + 2
    monthYearCol = squadColNum + 3
    dataStartCol = squadColNum + 42
    dataEndCol = squadColNum + 60
    shiftStartCol = squadColNum + 4
    shiftEndCol = squadColNum + 41

    ' Before you start filling the PivotTable sheets with data, clear the existing contents.
    Application.EnableEvents = False ' disable events
    destSheet.Cells.ClearContents
    destSheet2.Cells.ClearContents
    Application.EnableEvents = True ' enable events again

    destSheet.Cells(1, 1).value = "WorkerName"
    destSheet.Cells(1, 2).value = "DateShifts"
    destSheet.Cells(1, 3).value = "NumberShifts"
    destSheet.Cells(1, 1).EntireRow.Font.Bold = True

    ' Setup headers for PivotTable2
    destSheet2.Cells(1, 1).value = "WorkerName"
    destSheet2.Cells(1, 2).value = "DateMonth"
    destSheet2.Cells(1, 3).value = "DataHeader"
    destSheet2.Cells(1, 4).value = "DataValue"
    destSheet2.Cells(1, 1).EntireRow.Font.Bold = True

    Dim destSheetRow As Long, destSheet2Row As Long
    destSheetRow = 2
    destSheet2Row = 2

    For i = 3 To lastRow
        Name = CStr(srcSheet.Cells(i, nameCol).value) ' Convert cell content to a string
        monthYear = srcSheet.Cells(i, monthYearCol).value
        
        ' Skip if name is empty, equals 'Nazwisko i imiê', equals '-', or equals '0'
        If Name = "" Or Name = "Nazwisko i imiê" Or Name = "-" Or Name = "0" Then
            GoTo NextRow
        End If

        ' Split monthYear string into month and year
        Dim splitMonthYear() As String
        splitMonthYear = Split(monthYear, " ")
        
        If UBound(splitMonthYear) >= 1 Then
            If InStr(monthYear, "zm.") > 0 Then
                GoTo NextRow
            Else
                month = monthDict(splitMonthYear(0))
                year = splitMonthYear(1)
            End If
        Else
            ' If monthYear does not contain a space, handle accordingly (e.g., skip row)
            GoTo NextRow
        End If

        ' Check if the month and year for the shift data is the same as the month and year for the day data
        Dim nextMonthYear As String
        nextMonthYear = srcSheet.Cells(i + 1, monthYearCol).value
        Dim nextMonth As Integer, nextYear As String
        If InStr(nextMonthYear, "zm.") > 0 Then
            nextMonth = monthDict(Split(nextMonthYear, " zm. ")(0))
            nextYear = Split(nextMonthYear, " zm. ")(1)
        Else
            GoTo NextRow
        End If
        
        If nextMonth <> month Or nextYear <> year Then
            MsgBox "Error: Month and year for rows " & i & " and " & i + 1 & " do not match."
            GoTo NextRow
        End If

        ' Add data to PivotTable2 for columns AT to BF
        For k = dataStartCol To dataEndCol
            data = srcSheet.Cells(i + 1, k).value
            If Not IsEmpty(data) And Not IsError(data) Then
                destSheet2.Cells(destSheet2Row, 1).value = Name
                destSheet2.Cells(destSheet2Row, 2).value = DateSerial(year, month, 1)
                destSheet2.Cells(destSheet2Row, 2).NumberFormat = "yyyy-mm-dd" ' Change date format here
                destSheet2.Cells(destSheet2Row, 3).value = srcSheet.Cells(2, k).value
                destSheet2.Cells(destSheet2Row, 4).value = data
                destSheet2Row = destSheet2Row + 1
            End If
        Next k

        ' Add data to PivotTable for columns I to AS
        For j = shiftStartCol To shiftEndCol
            day = srcSheet.Cells(i, j).value
            shift = srcSheet.Cells(i + 1, j).value
            
            ' Only process cells that contain a numeric day value
            If IsNumeric(day) Then
                
                If Not IsEmpty(shift) And Not IsError(shift) Then
                    destSheet.Cells(destSheetRow, 1).value = Name
                    destSheet.Cells(destSheetRow, 2).value = DateSerial(year, month, day)
                    destSheet.Cells(destSheetRow, 2).NumberFormat = "yyyy-mm-dd" ' Change date format here
                    destSheet.Cells(destSheetRow, 3).value = shift
                    destSheetRow = destSheetRow + 1
                End If
            End If
        Next j

NextRow:
    Next i

    ' Autofit columns in the PivotTable sheet
    destSheet.Columns("A:C").EntireColumn.AutoFit
    
    ' Autofit columns in the PivotTable2 sheet
    destSheet2.Columns("A:D").EntireColumn.AutoFit
    
    ' Cleanup
    Set srcSheet = Nothing
    Set destSheet = Nothing
    Set destSheet2 = Nothing
    Set monthDict = Nothing
    
End Sub



Sub PivotDataUnique(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String, srcWorkbook As Workbook)
    
    Dim srcSheet As Worksheet
    Dim destSheet3 As Worksheet
    Dim i As Long
    Dim Group As String
    Dim Squad As String
    Dim Abbreviation As String
    Dim Name As String
    Dim dictUnique As Object
    Dim GroupRelevant As Integer
    
    ' Set source worksheet
    Set srcSheet = srcWorkbook.Sheets(wsNum)
        
    ' Check if "WorkersStatus" sheet exists. If not, create it.
    Dim ws As Worksheet
    Dim sheetExists3 As Boolean
    sheetExists3 = False
    For Each ws In srcWorkbook.Sheets
        If ws.Name = "WorkersStatus" Then
            sheetExists3 = True
        End If
    Next ws
    
    If Not sheetExists3 Then
        Set destSheet3 = srcWorkbook.Sheets.Add(After:=srcWorkbook.Sheets(srcWorkbook.Sheets.Count))
        destSheet3.Name = "WorkersStatus"
    Else
        Set destSheet3 = srcWorkbook.Sheets("WorkersStatus")
    End If


    ' Clear PivotTable3 sheet before adding new data
    destSheet3.Cells.ClearContents

    GroupRelevant = MsgBox("Is the column before Brygada relevant?", vbYesNo)
    
    If GroupRelevant = vbYes Then
        destSheet3.Cells(1, 1).value = "WorkerGroup"
        destSheet3.Cells(1, 2).value = "WorkerSquad"
        destSheet3.Cells(1, 3).value = "SquadSymbol"
        destSheet3.Cells(1, 4).value = "WorkerName"
    Else
        destSheet3.Cells(1, 1).value = "WorkerSquad"
        destSheet3.Cells(1, 2).value = "SquadSymbol"
        destSheet3.Cells(1, 3).value = "WorkerName"
    End If

    destSheet3.Cells(1, 1).EntireRow.Font.Bold = True
    
    Dim destSheet3Row As Long
    destSheet3Row = 2

    Set dictUnique = CreateObject("Scripting.Dictionary")

    For i = 3 To lastRow
        Group = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) - 1)).value) ' Convert cell content to a string
        Squad = CStr(srcSheet.Cells(i, squadCol).value) ' Convert cell content to a string
        Abbreviation = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) + 1)).value) ' Convert cell content to a string
        Name = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) + 2)).value) ' Convert cell content to a string


        ' Skip if name is empty, equals 'Nazwisko i imiê', equals '-', or equals '0'
        If Name = "" Or Name = "Nazwisko i imiê" Or Name = "-" Or Name = "0" Then
            GoTo NextRowUnique
        End If

        ' Only add unique values to PivotTable3
        If Not dictUnique.Exists(Name) Then
            dictUnique(Name) = ""

            If GroupRelevant = vbYes Then
                destSheet3.Cells(destSheet3Row, 1).value = Group
                destSheet3.Cells(destSheet3Row, 2).value = Squad
                destSheet3.Cells(destSheet3Row, 3).value = Abbreviation
                destSheet3.Cells(destSheet3Row, 4).value = Name
            Else
                destSheet3.Cells(destSheet3Row, 1).value = Squad
                destSheet3.Cells(destSheet3Row, 2).value = Abbreviation
                destSheet3.Cells(destSheet3Row, 3).value = Name
            End If
            
            destSheet3Row = destSheet3Row + 1
        End If

NextRowUnique:
    Next i

    ' Autofit columns in the PivotTable3 sheet
    destSheet3.Columns("A:D").EntireColumn.AutoFit
    
    ' Cleanup
    Set srcSheet = Nothing
    Set destSheet3 = Nothing
    Set dictUnique = Nothing
End Sub

Sub CheckForEmptyCells(srcWorkbook As Workbook)

    ' Check sheets "WorkersShifts", "WorkersMonthData", and "WorkersStatus" for empty cells
    Dim wsNames As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long, lastCol As Long
    Dim rowEmpty As Boolean
    Dim msg As String
    
    wsNames = Array("WorkersShifts", "WorkersMonthData", "WorkersStatus")
    
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = srcWorkbook.Sheets(wsNames(i))
        ' Find the last header column
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        msg = ""
        For j = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            rowEmpty = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(j, 1), ws.Cells(j, lastCol))) = 0
            If rowEmpty Then
                Exit For
            Else
                If Application.WorksheetFunction.CountBlank(ws.Range(ws.Cells(j, 1), ws.Cells(j, lastCol))) > 0 Then
                    msg = msg & vbNewLine & "Row " & j & " in sheet " & wsNames(i)
                End If
            End If
        Next j
        If msg <> "" Then
            MsgBox "The following rows in " & wsNames(i) & " have empty cells:" & msg
        End If
        Set ws = Nothing
    Next i

End Sub











