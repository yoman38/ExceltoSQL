Attribute VB_Name = "Excel2SQLconverter"

Sub Excel2SQLconverter()

    ' Call the main procedure from TPtoPivot and check its result
    
    ' Ask the user to select the source file
    Dim srcFile As Variant
    srcFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", _
                                          Title:="Please select the source Excel file")
    
    ' Exit if the user didn't choose a file
    If srcFile = False Then
        MsgBox "No file selected. Exiting..."
        Exit Sub
    End If
    
    ' Open the selected workbook
    Dim srcWorkbook As Workbook
    Set srcWorkbook = Workbooks.Open(srcFile)
    gSelectedExcelFile = srcFile


    ' Continue with the rest of the processes
        
        ' Call the main procedure from ExcelSQL
        GenerateSQL
        
        ' Prompt the user to check if they want to upload the table to Microsoft SQL Server
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("Do you want to upload your table to Microsoft SQL Server?", vbYesNo)
        
        If userResponse = vbYes Then
            ' Call the main procedure from Txt2SQL
            UpdateSQLWithTxtContent
        End If
        

        ' Do not run other modules
        MsgBox "Processes were cancelled or not completed successfully. Other modules will not run.", vbExclamation

End Sub


