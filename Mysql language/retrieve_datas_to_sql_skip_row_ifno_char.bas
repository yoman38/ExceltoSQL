Attribute VB_Name = "Module1"
Sub GenerateSQL()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim headers As Range
    Dim i As Long
    Dim j As Long
    Dim sqlCreate As String
    Dim sqlInsert As String
    Dim filename As String
    Dim fso As Object
    Dim outputFile As Object
    
    filename = "C:\Users\gbray\Desktop\data\PROJECT3_work_schedule\TP1 KAZIMIERZ i JULIUSZ GRAFIK 2022-2023-wraz z remontami agata.xls"
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open the file for output
    Set outputFile = fso.CreateTextFile("C:\Users\gbray\Desktop\output3.txt", True)
    
    ' Open the workbook
    Set wb = Workbooks.Open(filename)
    Set ws = wb.Sheets("TP1 grafik brygad 2022-2023")
    
    ' Define the range of data manually
    Set rng = ws.Range("F3:BK576") ' Update this to the correct range
    
    ' Define the row of headers manually
    Set headers = ws.Range("F2:BK2") ' Update this to the correct range

    ' Generate the CREATE TABLE statement
    sqlCreate = "CREATE TABLE `test_vba` ("
    For i = 1 To headers.Columns.Count
        sqlCreate = sqlCreate & "[" & headers.Cells(1, i).Value & "] NVARCHAR(100), "
    Next i
    sqlCreate = Left(sqlCreate, Len(sqlCreate) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon
    
    outputFile.WriteLine sqlCreate ' This writes the create table statement to the file
    
    ' Generate the INSERT INTO statements
    For i = 1 To rng.Rows.Count
        If InStr(1, rng.Cells(i, 2).Value, "zm", vbTextCompare) > 0 Then ' column G corresponds to the 2nd column in the range
            sqlInsert = "INSERT INTO `test_vba` VALUES ("
            For j = 1 To rng.Columns.Count
                sqlInsert = sqlInsert & "'" & rng.Cells(i, j).Value & "', "
            Next j
            sqlInsert = Left(sqlInsert, Len(sqlInsert) - 2) & ");" ' Remove trailing comma and space, add closing bracket and semicolon
            
            outputFile.WriteLine sqlInsert ' This writes the insert into statement to the file
        End If
    Next i
    
    ' Close the file
    outputFile.Close
    
    ' Close the workbook without saving changes
    wb.Close False

End Sub

