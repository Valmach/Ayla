Attribute VB_Name = "Rule3"
Sub rulethree()

    ' These rules require specific code for each rule (they're not generic) - so there's no overall loop
        
    ' These rules are combined
    ' Note: You can't check for copies if you don't check for duplicates first etc.
    checkForDuplicates = Sheets("Rules 3").Cells(12, 3)
    
    ' Only do this if the rule has been activated
    If checkForDuplicates = 1 Then
    
        checkCompleteString = "Complete"
    
        ' Clear the column to record files have been checked already
        Sheets("J").Select
        Range("G3:G10000").Select
        Selection.ClearContents
             Sheets("j").Range("A3").Select
        Sheets("Dashboard").Select
            
        ' Grab the rule parameters
        ruleStage = Sheets("Rules 3").Cells(12, 1)
        errorIfDuplicate = Sheets("Rules 3").Cells(12, 4)
        errorIfCopy = Sheets("Rules 3").Cells(13, 4)
        errorIfDiffFileTypes = Sheets("Rules 3").Cells(14, 4)
    
        If ruleStage <> "" Then
        
            ' If the stage is relevant
        
            ' assign the rule stage a number
            For K = 2 To 30
            
                If Sheets("Stages").Cells(K, 1) = "" Then
                    MsgBox ("The rule stage doesn't exist in the stages tab")
                    Exit Sub
                End If
                
                If LCase(ruleStage) = LCase(Sheets("Stages").Cells(K, 1)) Then
                    Exit For ' k will be the rule stage number
                End If
                
            Next K
                    
            If K <= projectStageNumber Then
                            
                ' Cycle through each of the j drive files
                For J = 3 To 10000
                
                    ' Check if file has already been registered as duplicate/copy
                    If Sheets("J").Cells(J, 7) <> checkCompleteString Then
                                        
                        ' Grab the current file properties (against which we'll check for duplicates/copies)
                        currentFileName = Sheets("J").Cells(J, 1)
                        currentFileSize = Sheets("J").Cells(J, 6)
                        currentFileType = Sheets("J").Cells(J, 5)
                        currentFilePath = Sheets("J").Cells(J, 3)
                        
                        ' Don't bother continuing if we're at the end
                        If currentFileName = "" Then Exit For
                        
                        ' Record that this file has already been checked for duplicates/copies
                        ' This has to be done now so that it won't check against itself (later in this code)
                        Sheets("J").Cells(J, 7) = checkCompleteString
                        
                        ' Check if it's in an excluded path
                        excludedFilePath = False
                        For excludedStringRow = 3 To 100
                        
                            excludedString = Sheets("Rules 3").Cells(excludedStringRow, 5)
                            
                            ' If it's empty don't continue
                            If excludedString = "" Then Exit For
                            
                            If InStr(1, LCase(currentFilePath), LCase(excludedString)) Then
                                
                                excludedFilePath = True
                                Exit For
                            
                            End If
                        
                        Next excludedStringRow
                        
                        ' Check if the file needs to be checked
                        nameContainsIncludedString = False
                        For includedStringRow = 3 To 100
                        
                            includedString = Sheets("Rules 3").Cells(includedStringRow, 6)
                            
                            ' If it's empty don't continue
                            If includedString = "" Then Exit For
                            
                            If InStr(1, LCase(currentFileName), LCase(includedString)) Then
                                
                                nameContainsIncludedString = True
                                Exit For
                            
                            End If
                        
                        Next includedStringRow
                        
                        
                        ' Only continue with this file if it's not in an excluded path and it's name has a string that we're looking for
                        If excludedFilePath = False And nameContainsIncludedString = True Then
                        
                            ' cycle through each of the remaining files on the j drive
                            For jFileRow = 3 To 10000
                            
                                ' NOTE: The same intial process is used to check the comparison jFiles as the currentFile
                                    
                                ' Grab the file properties to check
                                jFileName = Sheets("J").Cells(jFileRow, 1)
                                jFileSize = Sheets("J").Cells(jFileRow, 6)
                                jFileType = Sheets("J").Cells(jFileRow, 5)
                                jFilePath = Sheets("J").Cells(jFileRow, 3)
                                
                                ' If it's empty don't continue
                                If jFileName = "" Then Exit For
                                
                                ' Check if it's in an excluded path
                                excludedFilePath = False
                                For excludedStringRow = 3 To 100
                                
                                    excludedString = Sheets("Rules 3").Cells(excludedStringRow, 5)
                                    
                                    ' If it's empty don't continue
                                    If excludedString = "" Then Exit For
                                    
                                    If InStr(1, LCase(jFilePath), LCase(excludedString)) Then
                                        
                                        excludedFilePath = True
                                        Exit For
                                    
                                    End If
                                
                                Next excludedStringRow
                                
                                ' Check if the file needs to be checked
                                nameContainsIncludedString = False
                                For includedStringRow = 3 To 100
                                
                                    includedString = Sheets("Rules 3").Cells(includedStringRow, 6)
                                    
                                    ' If it's empty don't continue
                                    If includedString = "" Then Exit For
                                    
                                    If InStr(1, LCase(jFileName), LCase(includedString)) Then
                                        
                                        nameContainsIncludedString = True
                                        Exit For
                                    
                                    End If
                                
                                Next includedStringRow
                                
                                ' Only continue with this j-file if it's not in an excluded path and it's name has a string that we're looking for
                                If excludedFilePath = False And nameContainsIncludedString = True Then
                                
                                    ' Check if they're the same (and haven't already been checked)
                                    If currentFileName = jFileName And Sheets("J").Cells(jFileRow, 7) <> checkCompleteString Then
                                                                                                   
                                        ' Check if it's a copy? (are the file sizes the same?)
                                    
                                        If currentFileSize = jFileSize Then
                                            
                                            ' The file is a copy
                                            
                                            ' Output the error
                                            Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                                            Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                                            Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                            Sheets("Dashboard").Cells(nextBlankRow, 4) = errorIfCopy & currentFileName
                                            path1 = Sheets("J").Cells(J, 3)
                                            path2 = Sheets("J").Cells(jFileRow, 3)
                                            Sheets("Dashboard").Cells(nextBlankRow, 5).Formula = "=HYPERLINK(""" & path1 & """,""" & path1 & """)"
                                            Sheets("Dashboard").Cells(nextBlankRow, 6).Formula = "=HYPERLINK(""" & path2 & """,""" & path2 & """)"
                                            
                                            nextBlankRow = nextBlankRow + 1
                                                
                                        Else
                                        
                                            ' The file just has the same name, it's not a copy
                                            
                                            ' Check if it's the same file type
                                            If currentFileType = jFileType Then
                                            
                                                ' Both files are the same type and name (so must be in different folders), but have different sizes
                                                ' Ouput the error
                                                Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                                                Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                                                Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                                Sheets("Dashboard").Cells(nextBlankRow, 4) = errorIfDuplicate & currentFileName
                                                path1 = Sheets("J").Cells(J, 3)
                                                path2 = Sheets("J").Cells(jFileRow, 3)
                                                Sheets("Dashboard").Cells(nextBlankRow, 5).Formula = "=HYPERLINK(""" & path1 & """,""" & path1 & """)"
                                                Sheets("Dashboard").Cells(nextBlankRow, 6).Formula = "=HYPERLINK(""" & path2 & """,""" & path2 & """)"
                                                nextBlankRow = nextBlankRow + 1
                                            
                                            Else
                                            
                                                ' Files are the same name but not the same type or size (possibly a doc and a pdf)
                                                ' NOTE: This error has been disabled for the time being
                                                '       Enabling will throw error if doc and pdf of same name exist
                                                
                                                ' Ouput the error
'                                                Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
'                                                Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
'                                                Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
'                                                Sheets("Dashboard").Cells(nextBlankRow, 4) = errorIfDiffFileTypes & currentFileName
'                                                Sheets("Dashboard").Cells(nextBlankRow, 5) = Sheets("J").Cells(j, 3) & Sheets("J").Cells(j, 1) & "." & Sheets("J").Cells(j, 5)
'                                                Sheets("Dashboard").Cells(nextBlankRow, 6) = Sheets("J").Cells(jFileRow, 3) & Sheets("J").Cells(jFileRow, 1) & "." & Sheets("J").Cells(jFileRow, 5)
'
'                                                nextBlankRow = nextBlankRow + 1
                                            
                                            End If
                                                                                 
                                              
                                        End If
                                        
                                        
                                    ' Record that this file has already been checked for duplicates/copies
                                    Sheets("J").Cells(jFileRow, 7) = checkCompleteString
                                                                                
                                    End If
                                
                                End If
                                                    
                            Next jFileRow
                        
                        End If
                                                
                    End If
                                        
                Next J
        
            End If
        
        End If
    
    
    
    End If
    
    
                 
    
End Sub


