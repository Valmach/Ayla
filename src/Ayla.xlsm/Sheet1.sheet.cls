Private Sub CommandButton1_Click()

End Sub

Private Sub cmdAutoCAD_Click()
Excel.Workbooks.Open ("K:\M&E\Calculations\APPLICATIONS\Autocad 2014 Drawing cleaner and noprint layer fix.xlsm")
End Sub

Private Sub cmdcompressimages_Click()
Shell ("K:\M&E\Calculations\APPLICATIONS\compressor.exe")
End Sub

Private Sub cmdfoldersize_Click()
Excel.Workbooks.Open ("K:\M&E\Calculations\APPLICATIONS\folder size extraction run macro.xls")
End Sub

Private Sub cmdsorttype_Click()

' Sort the Results table by Issue

   Range("A15:G2814").Select
       ActiveWorkbook.Worksheets("Dashboard").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dashboard").sort.SortFields.Add Key:=Range( _
        "D17:D969"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Dashboard").sort
        .SetRange Range("A15:G969")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-3
    Range("A16").Select
End Sub