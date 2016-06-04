Attribute VB_Name = "CalcsM"
Private Const nonIndexedCalcSheetName = "Non Indexed Calcs"

Public Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

Sub checkCalcsM()

    ' This sub checks that all of the excel files in CalcsM:
    ' have an info tab - if they don't, it's added
    ' are indexed in the Calculations index workbook - if they're not, they're added to the unindexed sheet (if the unindexed sheet doesn't exist, it's added)
    
    ' All issues found are assigned to the Mech Engineer (if present) rather than the Project Runner
    
    ' Cross our fingers...
    On Error Resume Next
        
    Application.ScreenUpdating = False
        
    Dim aylaDashboard, aylaJ As Worksheet
    Set aylaDashboard = Sheets("Dashboard")
    Set aylaJ = Sheets("J")
    
    ' Check for a calculations index
    ' Where the workbook should be saved
    Dim SourceFolderName As String
    SourceFolderName = "J:\" & projectNumber & "\Calculations"
    
    Dim targetWorkbook As String
    targetWorkbook = "CalculationsM"
    
    Dim fullWorkBookname As String
    fullWorkBookname = targetWorkbook & ".xls"
    
    ' Cycle through each file in the folder
    Dim FSO As New FileSystemObject, SourceFolder As Folder, Subfolder As Folder, FileItem As File
    
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    indexExists = False
    For Each FileItem In SourceFolder.Files
            
        ' Cycle through each file
            
        ' separate file name & extension
        For g = Len(FileItem.Name) To 1 Step -1
        
            If Mid(FileItem.Name, g, 1) = "." Then
                trimmedName = Mid(FileItem.Name, 1, g - 1)
                fileExtension = Mid(FileItem.Name, g + 1, Len(FileItem.Name) - g)
                Exit For
            End If
            
        Next g
        
        ' If it's the correct file don't scan any more
        If trimmedName = targetWorkbook And InStr(1, LCase(fileExtension), "xl") Then
            indexExists = True
            Exit For
        End If
                    
    Next FileItem
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    
    ' Define the workbooks we're going to be using
    Dim aylaWb, indexWb, currentCalculation As Workbook
    ' Must also have a reference to Ayla otherwise the opened workbooks becomes the default one referred to
    Set aylaWb = ThisWorkbook
    
    
    ' If the CalculationsM index exists, make sure it has a non indexed calcs sheet
    If indexExists = True Then
    
        ' Count how many unindexed calcs there are
        Dim unindexedCalcCount As Integer
        unindexedCalcCount = 0
    
        ' Index exists (if it doesn't we'll already produce an error from our rules)
    
        ' At this point the index workbook is definately available at the usual path
        Dim actualPath As String
        actualPath = SourceFolderName & "\" & fullWorkBookname
        
        ' There's a known bug that makes the macro stop when you open a workbook in excel
        ' Seemingly, you need to check if the shift button is pressed
        ' See link: https://support.microsoft.com/en-us/kb/555263
        
        ' Open up with CalculationsM workbook
        Do While ShiftPressed()
            DoEvents
        Loop
        Set indexWb = Workbooks.Open(actualPath, False) ' false prevents links updating (and the annoying pop up that comes with it)
        
        ' Check if non-indexed calcs sheet exists in indexWb
        Dim sh As Worksheet
        Dim nonIndexedSheetFound As Boolean
        
        For Each sh In indexWb.Worksheets
            If sh.Name Like nonIndexedCalcSheetName Then nonIndexedSheetFound = True: Exit For
        Next
        
        ' If there's no Non Indexed Calcs Sheet
        If nonIndexedSheetFound = False Then
            
            ' Add it
            Dim nonIndexedSheet As Worksheet
            Set nonIndexedSheet = indexWb.Sheets.Add(After:=indexWb.Sheets(Sheets.Count))
            
            nonIndexedSheet.Select
            nonIndexedSheet.Name = nonIndexedCalcSheetName
            
            ActiveWindow.DisplayGridlines = False
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "Non Indexed Mechanical Calculations"
            Range("B4").Select
            ActiveCell.FormulaR1C1 = "Project Name:"
            Range("B5").Select
            ActiveCell.FormulaR1C1 = "Project Number:"
            Range("B6").Select
            ActiveCell.FormulaR1C1 = "Date:"
            Range("B7").Select
            Columns("B:B").ColumnWidth = 13
            Columns("A:A").ColumnWidth = 0.92
            Range("B2").Select
            Selection.Font.Bold = True
            Range("C4").Select
            ActiveCell.FormulaR1C1 = projectName
            Range("C5").Select
            ActiveCell.FormulaR1C1 = projectNumber
            Range("C6").Select
            ActiveCell.FormulaR1C1 = "='Job info'!R[1]C[2]"
            Range("C9").Select
            ActiveCell.FormulaR1C1 = "Name"
            Range("D9").Select
            ActiveCell.FormulaR1C1 = "Path"
            Columns("C:C").Select
            Selection.ColumnWidth = 40
            Range("D9").Select
            Columns("D:D").ColumnWidth = 80
            Range("C9:D9").Select
            Selection.Font.Bold = True
            Range("C10").Select
            Range("B4:C6").Select
            
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
                    
        End If
        
        ' Empty the non indexed register as we're going to recreate it below
        ' This ensures calcs that have been deleted are removed and that any incorrectly identified by Ayla are also removed (once the corresponding bug is fixed)
        
        indexWb.Sheets(nonIndexedCalcSheetName).Select
        Range("C10:D1000").Select
        Selection.ClearContents
        Range("A1").Select
        indexWb.Sheets("Index").Select
        Range("A1").Select
        
        ' Find the next available row for non indexed calcs
        ' (this is now redundant because we clear the register but it still works so I'll leave it in)
    
        nonIndexedNameColumn = 3
        For nonIndexedRow = 10 To 1000 ' 10 is first empty row
            If indexWb.Sheets(nonIndexedCalcSheetName).Cells(nonIndexedRow, nonIndexedNameColumn) = "" Then Exit For
        Next nonIndexedRow
        
        ' If we find a calculation that isn't indexed
        ' we can now record it on indexWb.Sheets(nonIndexedCalcSheetName).Cells(nonIndexedRow, nonIndexedNameColumn)
        ' but must increment nonIndexedRow after doing so
        
    
    End If
    
    
    ' At this point the aylaWb and indexWb are loaded into memory
    ' (if indexExists = true, otherwise its nothing. If it does exist, there's definitely a non indexed calcs tab)
    ' They remain loaded while we cycle through all of the calculation workbooks within CalcsM
        
    
    ' Now cycle through each of the calculations in the calcM folder
       
    ' Count how many unindexed calcs there are
    Dim unCheckedCalcCount As Integer
    unCheckedCalcCount = 0
                    
    ' Cycle through each of the j drive files
    For J = 3 To 10000
    
        ' For Testing (stop at a particular J drive file)
        If J = 11 Then
            test = "put a breakpoint here"
        End If
                                         
        ' Grab the current file properties
        currentFileName = aylaJ.Cells(J, 1)
        currentFileType = aylaJ.Cells(J, 5)
        currentFilePath = aylaJ.Cells(J, 3)
        currentFileFullpath = currentFilePath & currentFileName & "." & currentFileType
        
        ' Don't bother continuing if we're at the end
        If currentFileName = "" Then Exit For
        
        ' Only check documents in the CalcsM folder
        If InStr(1, LCase(currentFilePath), LCase("\CalcsM\")) Then
        
            ' Only check excel documents
            If InStr(1, LCase(currentFileType), LCase("xl")) Then
            
                ' Check whether the calc is in an excluded path
                excludedFilePath = False
                For excludedStringRow = 11 To 1000
                
                    excludedString = aylaWb.Sheets("Misc").Cells(excludedStringRow, 5)
                    
                    ' If it's empty don't continue
                    If excludedString = "" Then Exit For
                    
                    ' Check the path and name for the excluded string
                    If InStr(1, LCase(currentFileFullpath), LCase(excludedString)) Then
                        
                        excludedFilePath = True
                        Exit For
                    
                    End If
                
                Next excludedStringRow
                
                
                ' If not excluded, continue
                If Not excludedFilePath Then
                
                    ' Open the workbook
                    Do While ShiftPressed()
                        DoEvents
                    Loop
                    Set currentCalculation = Workbooks.Open(currentFileFullpath, False) ' false prevents links updating (and the annoying pop up that comes with it)
                            
                    ' Check if there's an INFO tab
                    Dim sht As Worksheet
                    Dim infoSheetFound As Boolean
                    
                    infoSheetFound = False
                    For Each sht In currentCalculation.Worksheets
                        If sht.Name = "INFO" Then infoSheetFound = True: Exit For
                    Next
                
                    ' If there's no INFO tab, add one in
                    If infoSheetFound = False Then
                    
                        ' Check if no INFO tab and do this if there isn't
                        Dim infoWs As Worksheet
                        Set infoWs = currentCalculation.Sheets.Add(After:=currentCalculation.Sheets(currentCalculation.Sheets.Count))
                        
                        ' Configure the INFO Tab
                        infoWs.Select
                        infoWs.Name = "INFO"
                        ActiveWindow.DisplayGridlines = False
                        Range("B2").Select
                        ActiveCell.FormulaR1C1 = "Calculations Information Sheet"
                        Range("A4").Select
                        ActiveCell.FormulaR1C1 = "Note: Do not change the name of this sheet's tab"
                        Range("A6").Select
                        ActiveCell.FormulaR1C1 = _
                            "The following project information is filled out when the calculation is first created:"
                        Range("A7").Select
                        ActiveCell.FormulaR1C1 = "Project Name"
                        Range("B7").Select
                        ActiveCell.FormulaR1C1 = projectName
                        Range("A8").Select
                        ActiveCell.FormulaR1C1 = "Project Number"
                        Range("B8").Select
                        ActiveCell.FormulaR1C1 = projectNumber
                        Range("A9").Select
                        ActiveCell.FormulaR1C1 = "Date of Calc Creation"
                        Range("B9").Select
                        'ActiveCell.FormulaR1C1 = Date
                        Dim stringDate As String
                        stringDate = Date
                        ActiveCell.Value = stringDate
                        Range("A10").Select
                        ActiveCell.FormulaR1C1 = "Engineer for First Use"
                        Range("B10").Select
                        ActiveCell.FormulaR1C1 = "Ayla"
                        Range("A11").Select
                        ActiveCell.FormulaR1C1 = "Index File Location"
                        Range("B11").Select
                        ActiveCell.FormulaR1C1 = currentFileFullpath
                        Range("I5").Select
                        ActiveCell.FormulaR1C1 = "Checked"
                        Range("B2").Select
                        Selection.Font.Bold = True
                        Range("A16").Select
                        Selection.Font.Bold = True
                        Range("I5:J5").Select
                        Selection.Font.Bold = True
                        Range("I5:J5").Select
                        Range("J5").Activate
                        With Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlBottom
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = False
                            .ReadingOrder = xlContext
                            .MergeCells = False
                        End With
                        Selection.Merge
                        Range("I6").Select
                        ActiveCell.FormulaR1C1 = "Date"
                        Range("J6").Select
                        ActiveCell.FormulaR1C1 = "Engineer"
                        Range("I5:J11").Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlInsideVertical)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlInsideHorizontal)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Range("I6:J6").Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0.349986266670736
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0.499984740745262
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorDark1
                            .TintAndShade = -0.349986266670736
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorDark1
                            .TintAndShade = -0.249977111117893
                            .PatternTintAndShade = 0
                        End With
                        Range("A7:G11").Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                        Range("A7:G7").Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                        Range("A7").Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorDark1
                            .TintAndShade = -0.249977111117893
                            .PatternTintAndShade = 0
                        End With
                        Range("B7:G7").Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.799981688894314
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.599993896298105
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                        Range("A7:G7").Select
                        Selection.Copy
                        Range("A8:G11").Select
                        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                            SkipBlanks:=False, Transpose:=False
                        Application.CutCopyMode = False
                        Columns("A:A").Select
                        Selection.ColumnWidth = 19.14
                        Range("B7:G11").Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.599993896298105
                            .PatternTintAndShade = 0
                        End With
                        Range("A7:A11").Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        With Selection.Borders(xlInsideHorizontal)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Range("B7:G11").Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.799981688894314
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                        Range("B7:G11").Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent4
                            .TintAndShade = 0.599993896298105
                            .PatternTintAndShade = 0
                        End With
                        Range("A16").Select
                        ActiveCell.FormulaR1C1 = "Record Of Edits"
                        Range("A18").Select
                        ActiveCell.FormulaR1C1 = "Date"
                        Range("B18").Select
                        ActiveCell.FormulaR1C1 = "Engineer / Comments"
                        Range("A19:E19").Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                        Range("A19").Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                        Range("A19:E19").Select
                        Selection.Copy
                        Range("A20:E36").Select
                        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                            SkipBlanks:=False, Transpose:=False
                        Application.CutCopyMode = False
                        Range("A1").Select
                        
                        ' Only save the calculation workbook if we added the info tab to it
                        currentCalculation.Sheets(1).Select
                        ActiveWorkbook.CheckCompatibility = False   ' turn off compatibility check
                        currentCalculation.Save
                    
                    End If
                       
                
                    ' At this point the current calculation definitely has an INFO tab
                    
                    ' Only check if the calculation is indexed if we know we have an index workbook
                    If indexExists Then
                    
                        ' Check if the current calculation is indexed
                        calculationIsIndexed = False
                        For indexRow = 30 To 1000       ' first indexed path is in (30,11) on the Job info sheet
                            
                            indexedFilePath = indexWb.Sheets("Job info").Cells(indexRow, 11)
                            If indexedFilePath = "" Then Exit For
                            
                            If LCase(indexedFilePath) = LCase(currentFileFullpath) Then  ' Chris added the Lcase
                                calculationIsIndexed = True
                                Exit For
                            End If
                        
                        Next indexRow
                        
                        ' If the calculation wasn't found on the index,
                        If Not calculationIsIndexed Then
                        
                            ' Increment the counter
                            unindexedCalcCount = unindexedCalcCount + 1
                        
                            ' Check if it's already been recorded on the non index sheet
                            alreadyRecorded = False
                            For nonIndexedRowToCheck = 10 To 1000
                            
                                ' if name is blank there's no more to check
                                If indexWb.Sheets(nonIndexedCalcSheetName).Cells(nonIndexedRowToCheck, nonIndexedNameColumn) = "" Then Exit For
                                
                                ' check if the path is the same as the current file's path
                                If LCase(indexWb.Sheets(nonIndexedCalcSheetName).Cells(nonIndexedRowToCheck, nonIndexedNameColumn + 1)) = LCase(currentFileFullpath) Then
                                    alreadyRecorded = True
                                    Exit For
                                End If
                                
                            Next nonIndexedRowToCheck
                        
                            ' add it to the non indexed sheet (if not already there)
                            If Not alreadyRecorded Then
                                indexWb.Sheets(nonIndexedCalcSheetName).Cells(nonIndexedRow, nonIndexedNameColumn) = currentFileName            ' record name
                                indexWb.Sheets(nonIndexedCalcSheetName).Cells(nonIndexedRow, nonIndexedNameColumn + 1) = currentFileFullpath    ' record path
                                nonIndexedRow = nonIndexedRow + 1                                                                               ' NB: increment row
                            End If
                        
                        ' if the calculation was found on the index, make sure it was checked
                        Else
                        
                            ' Check if the calculation has been checked
                            ' Cells(7,9) is the first date's cell, if it's empty
                            If currentCalculation.Sheets("INFO").Cells(7, 9) = "" Then
                                ' increment the counter
                                unCheckedCalcCount = unCheckedCalcCount + 1
                            End If
                        
                        End If
                    
                    End If
                    
                    ' We're finished checking the current calculation so we can close it now
                    currentCalculation.Close (False)    ' false = don't save (already saved if INFO tab added)
                    ' Don't set as nothing yet, do at end
                
                End If
            
            End If
        
        End If
        
                                                
    Next J
    
    ' Check if we found any unindexed calcs
    If unindexedCalcCount > 0 Then
    
        ' Output the error
        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        
        ' Check for project mech engineer
        If projectMech <> "" Then
            ' Assign to mech if present
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Else
            ' otherwise assign to runner
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        End If
        
        
        ' Grammar Nazi!!
        If unindexedCalcCount = 1 Then
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 4) = "Mech Calcs: There is " & unindexedCalcCount & " unindexed calculation in the CalcsM folder. Refer to the CalculationsM index for more details."
        Else
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 4) = "Mech Calcs: There are " & unindexedCalcCount & " unindexed calculations in the CalcsM folder. Refer to the CalculationsM index for more details."
        End If
        
        nextBlankRow = nextBlankRow + 1
        
        ' Add to the total number of unindexed calcs found (over the entire audit)
        totalNumberOfUnindexedMechCalcs = totalNumberOfUnindexedMechCalcs + unindexedCalcCount
    
    End If
    
    ' Check if we found any unindexed calcs
    If unCheckedCalcCount > 0 Then
    
        ' Output the error
        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        
        ' Check for project mech engineer
        If projectMech <> "" Then
            ' Assign to mech if present
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Else
            ' otherwise assign to runner
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        End If
        
        ' Grammar Nazi!!
        If unCheckedCalcCount = 1 Then
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 4) = "Mech Calcs: There is " & unCheckedCalcCount & " indexed calculation that hasn't been checked. Refer to the CalculationsM index for more details."
        Else
            aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 4) = "Mech Calcs: There are " & unCheckedCalcCount & " indexed calculations that haven't been checked. Refer to the CalculationsM index for more details."
        End If
        
        nextBlankRow = nextBlankRow + 1
        
        ' Add to the total number of unchecked calcs found (over the entire audit)
        totalNumberOfUnCheckedMechCalcs = totalNumberOfUnCheckedMechCalcs + unCheckedCalcCount
    
    End If
    
    ' Ouput the current totals even if they didn't change
    ' (this ensures they're overwritten even if none found - so the previous audit's values don't remain)
    aylaWb.Sheets("Misc").Cells(14, 6) = totalNumberOfUnCheckedMechCalcs
    aylaWb.Sheets("Misc").Cells(11, 6) = totalNumberOfUnindexedMechCalcs
    
    ' Save the index Wb (if it exists) - And also check if the currently indexed calcs exist
    If indexExists Then
        
        ' Cycle through the indexed calcs and make sure they exist
        For indexRow = 30 To 1000
        
            ' grab the calc name and path
            calcName = indexWb.Sheets("Job info").Cells(indexRow, 10)
            calcPath = indexWb.Sheets("Job info").Cells(indexRow, 11)
            
            If calcName = "" Then Exit For
            
            ' MOD: 26/05/16
            ' Only check if the calculation exists if it has a path as sometimes a refernece to hand calcs / SBEm / IES models are put in the name but there's no associated path
            
            ' If the calc doesn't exist, output an error
            If calcPath <> "" And Not IsFile(calcPath) Then
                aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                
                ' Check for project mech engineer
                If projectMech <> "" Then
                    ' Assign to mech if present
                    aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                Else
                    ' otherwise assign to runner
                    aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                End If
                
                aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 4) = "Mech Calcs: The calculation '" & calcName & "' is indexed but could not be found. Refer to the CalculationsM index for more details."
                nextBlankRow = nextBlankRow + 1
            End If
        
        Next indexRow
    
        indexWb.Sheets("Index").Select
        indexWb.CheckCompatibility = False   ' turn off compatibility check
        indexWb.Save
        indexWb.Close (False)    ' false = don't save
        Set indexWb = Nothing
    
    End If
    
    Set currentCalculation = Nothing
    Set aylaWb = Nothing
    
    
End Sub



