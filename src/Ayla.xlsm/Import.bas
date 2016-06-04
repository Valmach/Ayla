Attribute VB_Name = "Import"

Sub importExceptions()
    
    
    Application.ScreenUpdating = False
    
    ' read all the files in Ayla's PM folder
    On Error Resume Next
    
    Dim aylaSourceFoldername As String
    aylaSourceFoldername = Sheets("Output").Cells(8, 1)
    
    Dim runnerWb As Workbook
    Dim runnerDashboard, aylaExceptions As Worksheet
    
    Set aylaExceptions = ThisWorkbook.Sheets("Exceptions")
    
    Dim FSO As New FileSystemObject, SourceFolder As Folder, FileItem As File
    
    Set SourceFolder = FSO.GetFolder(aylaSourceFoldername)
    
    Dim g, firstColumn, lastColumn, currentColumn, runnerStartRow, runnerRow, aylaStartRow, aylaExceptionRow, exceptionRowForAyla As Integer
    Dim trimmedName, fileExtension, aylaOutputFilename As String
    Dim includeFile, alreadyRecorded, matchFound As Boolean
    
    
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
        
        ' Check wheter the file is to be included
        includeFile = False
        
        ' Exclude files with ~ in their name
        If InStr(1, LCase(trimmedName), LCase("~")) Then
            ' Don't include
        Else
            ' No ~ , so continue checking
            ' Check if it's an excel file
            If LCase(fileExtension) = LCase("xls") Or LCase(fileExtension) = LCase("xlsx") Or LCase(fileExtension) = LCase("xlsm") Then
                
                ' Make sure it's one of Ayla's files
                aylaOutputFilename = Sheets("Output").Cells(11, 1)
                If InStr(1, LCase(trimmedName), LCase(aylaOutputFilename)) Then includeFile = True
                
            End If
        
        End If
        
        ' Only check the file if it is to be included
        If includeFile = True Then
        
            ' There's a known bug that makes the macro stop when you open a workbook in excel
            ' Seemingly, you need to check if the shift button is pressed
            ' See link: https://support.microsoft.com/en-us/kb/555263
            Do While ShiftPressed()
                DoEvents
            Loop
        
            ' Open the workbook
            Set runnerWb = Workbooks.Open(FileItem.Path, , True)
            Set runnerDashboard = runnerWb.Sheets("Dashboard")
            
            firstColumn = 1
            lastColumn = 7
                   
            ' Cycle through each issue in the runner workbook
            runnerStartRow = 16
            
            For runnerRow = runnerStartRow To 10000
            
                ' Don't continue if empty
                If runnerDashboard.Cells(runnerRow, 1) = "" Then Exit For
                
                ' Check if the runner has asked for it to be ignored
                If runnerDashboard.Cells(runnerRow, 3) <> "" Then
                
                    ' Cycle through Ayla's exception list to see if it's already been recorded
                    alreadyRecorded = False
                    aylaStartRow = 16
                    For aylaExceptionRow = aylaStartRow To 10000
                    
                        ' Exit if no more records to check
                        If aylaExceptions.Cells(aylaExceptionRow, 1) = "" Then Exit For
                        
                        matchFound = True
                        For currentColumn = firstColumn To lastColumn
                            
                            If currentColumn = 3 Then
                                ' Do nothing, this is job runner in Ayla but "ignore" in runner's wb
                            ElseIf aylaExceptions.Cells(aylaExceptionRow, currentColumn + 2) <> runnerDashboard.Cells(runnerRow, currentColumn) Then
                                matchFound = False
                                Exit For
                            End If
                    
                        Next currentColumn
                        
                        ' If we found a match, stop checking Ayla's exception list
                        If matchFound Then
                            alreadyRecorded = True
                            Exit For
                        End If
                    
                    Next aylaExceptionRow
                    
                    ' If it hasn't already been recorded, add it to Ayla's list
                    If alreadyRecorded = False Then
                    
                        ' Get the first available row in Ayla's records
                        For exceptionRowForAyla = aylaStartRow To 10000
                            
                            If aylaExceptions.Cells(exceptionRowForAyla, 3) = "" Then
                                ' If it's blank, this row is available, so exit the for loop
                                Exit For
                            End If
                        
                        Next exceptionRowForAyla
                        
                        ' Copy over the exception to ayla's records
                        aylaExceptions.Cells(aylaExceptionRow, 1) = 0   ' Approval column
                        aylaExceptions.Cells(aylaExceptionRow, 2) = runnerDashboard.Cells(runnerRow, 3)   ' Reason to ignore column
                        For currentColumn = firstColumn To lastColumn
                            
                            If currentColumn = 3 Then
                                ' this is job runner in Ayla but "ignore" in runner's wb
                                aylaExceptions.Cells(aylaExceptionRow, currentColumn + 2) = runnerDashboard.Cells(13, 1)  ' job runner is listed in a single cell
                            Else
                                aylaExceptions.Cells(aylaExceptionRow, currentColumn + 2) = runnerDashboard.Cells(runnerRow, currentColumn)
                            End If
                    
                        Next currentColumn
                    
                    End If
                    
                
                End If
            
            Next runnerRow
            
            ' Close the current runner workbook
            runnerWb.Close False
            
            
        End If
    
        
    Next FileItem
    
        
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    
    Set runnerWb = Nothing
    Set runnerDashboard = Nothing
    Set aylaExceptions = Nothing


End Sub
