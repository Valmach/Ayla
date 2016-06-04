Attribute VB_Name = "Export"


Sub export()

    ' cross our fingers
    On Error Resume Next
    
    ' Before we overwrite by exporting, make sure we import the exceptions from the current Runner workbooks
    Call importExceptions
    
    ' There's no point in trying to remove any newly imported exceptions because they won't have been approved yet
    ' Call removeExceptions
        
    Application.ScreenUpdating = False
        
    Dim aylaDashboard, output, runnerDashboard As Worksheet
    Set aylaDashboard = Sheets("Dashboard")
    Set output = Sheets("Output")
    
    ' Find the last row
    For lastRow = 16 To 10000
        
        If aylaDashboard.Cells(lastRow, 1) = "" Then Exit For
    
    Next lastRow
    
    ' Sort the dashboard by job runner
    aylaDashboard.sort.SortFields.Clear
    aylaDashboard.sort.SortFields.Add Key:=Range( _
        "C16:C" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With aylaDashboard.sort
        .SetRange Range("A15:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    ' Clear the job runners found list
    output.Range("A16:A100").Clear
    runnerRecordRow = 16
    
    ' Where the individual workbooks should be saved
    Dim outputFolderPath, workBookname As String
    outputFolderPath = output.Cells(8, 1)
    workBookname = output.Cells(11, 1)  ' to be preceeded by job runner name
    
    ' Before we export the workbooks, we need to delete the current ones
    ' Otherwise, those who have no errors will still have a workbook (as it won't be overwritten)
    ' If something goes drastically wrong here, or if somebody complains that they filled it out but they weren't imported and now it's gone, just check the previous versions
    Kill outputFolderPath + "*.xlsx"    ' deletes all xlsx files in folder
    
    
    ' Where the template workbook is saved
    Dim templateFullPath As String
    templateFullPath = "K:\M&E\Project Management\Ayla\Ayla Job Runner Template.xlsx"
    
    
    ' Create the workbooks
    Dim objFSO
    Set objFSO = CreateObject("scripting.filesystemobject")
    
    Dim aylaWb, runnerWb As Workbook
    
    
    ' Must also have a reference to Ayla otherwise the opened workbook (runnerWb) becomes the default one referred to
    Set aylaWb = ThisWorkbook
    
    ' The first available row in the template (and initally set this to the output row)
    firstAvailableRowInTemplate = 16
    currentOutputRow = firstAvailableRowInTemplate
    
    lastRunner = ""
    
    For dashboardRow = 16 To 10000
    
        currentRunner = aylaDashboard.Cells(dashboardRow, 3)
        
        ' If we've reached the end of the list, save the currently opened runners workbook
        If currentRunner = "" Then
            If runnerWb <> Null Then
                runnerWb.Save
                runnerWb.Close
            End If
            Exit For
        End If
        
        
        ' If there's a change in job runner between one row and the next
        ' - save and close the previous runner's workbook
        ' - create and open the next runner's workbook
        ' - reset the output row of the new workbook
        ' - input the job runner into the job runner cell
        ' - they should be in order so duplicates shouldn't happen...)
        If currentRunner <> lastRunner Then
        
            ' Save the previous workbook
            If runnerWb Is Nothing Then
                ' Don't know how to not do this so I'll use an else...
            Else
                runnerWb.Save
                runnerWb.Close
            End If
        
            ' Record the job runner (for info only)
            output.Cells(runnerRecordRow, 1) = currentRunner
            runnerRecordRow = runnerRecordRow + 1
                        
            ' Set the destination (and rename accordingly)
            destinationFullPath = outputFolderPath & currentRunner & workBookname & ".xlsx"
            
            ' Create a new workbook (overwrite if existing)
            objFSO.CopyFile Source:=templateFullPath, Destination:=destinationFullPath
            
            ' There's a known bug that makes the macro stop when you open a workbook in excel
            ' Seemingly, you need to check if the shift button is pressed
            ' See link: https://support.microsoft.com/en-us/kb/555263
            Do While ShiftPressed()
                DoEvents
            Loop
            
            ' Set the runner objects
            Set runnerWb = Workbooks.Open(destinationFullPath)
            Set runnerDashboard = runnerWb.Sheets("Dashboard")
            
            ' Update the job runner cell
            runnerDashboard.Cells(13, 1) = currentRunner
            
            ' Reset the ouput row
            currentOutputRow = firstAvailableRowInTemplate
            
            
        End If
        
        
        
        ' Ouput the issue from Ayla's dashboard to the Runner's dashboard
        firstColumn = 1
        lastColumn = 7
        
        For currentColumn = firstColumn To lastColumn
        
            ' Copy over the issue except replace the job runner with a blank (for the exceptions column in the template)
            If currentColumn = 3 Then
            
                runnerDashboard.Cells(currentOutputRow, currentColumn) = ""
                
            Else
                
                aylaColumnContents = aylaDashboard.Cells(dashboardRow, currentColumn)
                
                ' If it's a path, use a hyperlink
                If InStr(1, LCase(aylaColumnContents), LCase("J:\")) Then
                    
                    ' runnerDashboard.Cells(currentOutputRow, currentColumn).Hyperlink.Add runnerDashboard.Cells(currentOutputRow, currentColumn), aylaColumnContents
                    runnerDashboard.Cells(currentOutputRow, currentColumn).Formula = "=HYPERLINK(""" & aylaColumnContents & """,""" & aylaColumnContents & """)"
                    
                ElseIf aylaColumnContents = "Open Equipment Schedule" Then
                
                    ' If there's a link to the equipment schedule, then copy over the formula (the path isn't displayed)
                    runnerDashboard.Cells(currentOutputRow, currentColumn).Formula = aylaDashboard.Cells(dashboardRow, currentColumn).Formula
                
                Else
                
                    ' If it's not a path, just do a straight copy over
                    runnerDashboard.Cells(currentOutputRow, currentColumn) = aylaDashboard.Cells(dashboardRow, currentColumn)
                    
                End If
                
            End If
            
        
        Next currentColumn
        
        ' increment the output row
        currentOutputRow = currentOutputRow + 1
        
        ' Update the last job runner
        lastRunner = currentRunner
    
    Next dashboardRow
    
    Set objFSO = Nothing
    Set runnerWb = Nothing
    Set aylaWb = Nothing
        

End Sub
