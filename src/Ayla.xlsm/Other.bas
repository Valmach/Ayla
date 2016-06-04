Attribute VB_Name = "Other"
Sub sortbyname()

' Sort the Results table by Runner

   Range("A15:G2814").Select
       ActiveWorkbook.Worksheets("Dashboard").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dashboard").sort.SortFields.Add Key:=Range( _
        "C17:C969"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
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


' clear the existing list
For i = 8 To 50
If Sheets("Sortbyname").Cells(i, 1) = "" Then Exit For
Sheets("Sortbyname").Cells(i, 1) = ""
Sheets("Sortbyname").Cells(i, 2) = ""
Next i

For J = 16 To 1000 ' cycle issues
myname = Sheets("Dashboard").Cells(J, 3)
If myname = "" Then Exit For

' check for name exists

For K = 8 To 50
If Sheets("Sortbyname").Cells(K, 1) = "" Then ' the name isn't on the list
Sheets("Sortbyname").Cells(K, 1) = myname
Sheets("Sortbyname").Cells(K, 2) = 1
Exit For
End If

If Sheets("Sortbyname").Cells(K, 1) = myname Then

Sheets("Sortbyname").Cells(K, 2) = Sheets("Sortbyname").Cells(K, 2) + 1


Exit For
End If


Next K


Next J


' Sort the Runners by issues found

       Sheets("Sortbyname").sort.SortFields.Clear
    Sheets("Sortbyname").sort.SortFields.Add Key:=Range( _
        "B8:B37"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Sheets("Sortbyname").sort
        .SetRange Range("A8:B37")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



End Sub


Sub changeExceptionPaths()

    ' This macro is used to change pre-existing exceptions for Copied Files and Duplicate Files
    ' We need to get rid of the filename from the path as Ayla now gives a link to the folder name rather than the file itself
    ' In order for the exceptions still to work, all columns must be the same (otherwise Ayla will think it's a new issue)
    ' Therefore we have to go back and change them (so people won't have to re-ignore them... and get mad)

    Dim exceptions As Worksheet
    Set exceptions = Sheets("Exceptions")
    
    issueColumn = 6
    pathColumn = 7
    infoColumn1 = 8
    infoColumn2 = 9
    
    For exceptionRow = 16 To 10000
    
       issue = exceptions.Cells(exceptionRow, issueColumn)
       Path = exceptions.Cells(exceptionRow, pathColumn)
       info1 = exceptions.Cells(exceptionRow, infoColumn1)
       info2 = exceptions.Cells(exceptionRow, infoColumn2)
       
       ' Check if it's the right issue
       If InStr(1, LCase(issue), LCase("Copied Files:")) Or InStr(1, LCase(issue), LCase("Duplicate Files:")) Then
                   
           ' CHANGE THE PATH
                   
           ' Cycle through the file path characters
           ' Make sure there's a file in the path by checking for a file extension (or a ".")
           ' If we find an extension, scan through until the first backslash
           ' Trim the path to this backslash
           
           ' trackers for logic
           foundExtension = False
           trimPath = False
           
           ' Scan through the path
           For g = Len(Path) To 1 Step -1
           
               ' Check for an extension (once found don't bother checking again)
               If Mid(Path, g, 1) = "." And foundExtension = False Then
                   foundExtension = True
               End If
               
               ' Check for a backslash
               If Mid(Path, g, 1) = "\" Then
                   
                   ' only trim the name if we found a file extension
                   If foundExtension Then
                       ' We do want to trim the path (to the current g value)
                       trimPath = True
                       
                   End If
                   
                   ' if we find a backslash but haven't found an extension, don't check anymore (and don't change the g value)
                   Exit For
                   
               End If
               
           Next g
           
           If trimPath = True Then
                   
               trimmedPath = Mid(Path, 1, g)
               exceptions.Cells(exceptionRow, pathColumn) = trimmedPath
           
           
           End If
           
           
           ' REPEAT FOR INFO1
           ' trackers for logic
           foundExtension = False
           trimPath = False
           
           ' Scan through the path
           For g = Len(info1) To 1 Step -1
           
               ' Check for an extension (once found don't bother checking again)
               If Mid(info1, g, 1) = "." And foundExtension = False Then
                   foundExtension = True
               End If
               
               ' Check for a backslash
               If Mid(info1, g, 1) = "\" Then
                   
                   ' only trim the name if we found a file extension
                   If foundExtension Then
                       ' We do want to trim the path (to the current g value)
                       trimPath = True
                       
                   End If
                   
                   ' if we find a backslash but haven't found an extension, don't check anymore (and don't change the g value)
                   Exit For
                   
               End If
               
           Next g
           
           If trimPath = True Then
                   
               trimmedPath = Mid(info1, 1, g)
               exceptions.Cells(exceptionRow, infoColumn1) = trimmedPath
           
           
           End If
       
       End If
       
       
           
    
    Next exceptionRow




End Sub
