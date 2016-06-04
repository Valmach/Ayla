Attribute VB_Name = "PathLength"
Private Const pathLengthError = "Path Error: path must be less than 255 characters."
Private Const additionalInfoForError = "Reduce the length of the file name or its directory. Or move the file to its parent directory."

Sub checkForMaxPathLengths()

    ' This sub searches through the J drive files for paths that exceed 255 characters.
    ' Each time a file is found with a path longer than this, an error is output
    
    ' Cross our fingers...
    On Error Resume Next
            
    ' Cycle through each of the j drive files
    For J = 3 To 10000
    
        ' For Testing (stop at a particular J drive file)
        If J = 11 Then
            test = "put a breakpoint here"
        End If
                                         
        ' Grab the current file properties and construct the full path
        currentFileName = Sheets("J").Cells(J, 1)
        currentFileType = Sheets("J").Cells(J, 5)
        currentFilePath = Sheets("J").Cells(J, 3)
        currentFileFullpath = currentFilePath & currentFileName & "." & currentFileType
        
        ' Don't bother continuing if we're at the end
        If currentFileName = "" Then Exit For
        
        ' Check if the path length exceeds 255 characters
        If Len(currentFileFullpath) >= 255 Then
        
            ' path length is too long, output an error
            
            Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
            Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
            Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
            Sheets("Dashboard").Cells(nextBlankRow, 4) = pathLengthError & " (Currently " & Len(currentFileFullpath) & ")"
            Sheets("Dashboard").Cells(nextBlankRow, 5).Formula = "=HYPERLINK(""" & currentFilePath & """,""" & currentFilePath & """)"  ' Use Path instead of full path so it goes to the folder instead of the file
            Sheets("Dashboard").Cells(nextBlankRow, 6) = additionalInfoForError
        
            ' Increment the outout row
            nextBlankRow = nextBlankRow + 1
        
        End If
                                           
    Next J
    
End Sub




