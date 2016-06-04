Attribute VB_Name = "Rule2"
Sub ruletwo()

    For i = 12 To 100000 ' read through rules
    
        ruleStage = Sheets("Rules 2").Cells(i, 1)
        checkFileType = Sheets("Rules 2").Cells(i, 2)
        checkFileTypeLocation = Sheets("Rules 2").Cells(i, 3)
        requiredFile = Sheets("Rules 2").Cells(i, 4)
        requiredFileLocation = Sheets("Rules 2").Cells(i, 5)
        errorIfNotFound = Sheets("Rules 2").Cells(i, 6)
        errorIfFoundInWrongLocation = Sheets("Rules 2").Cells(i, 7)
        
        If ruleStage = "" Then Exit For
                
        ' Is the stage relevant for the project, if not then go to next rule
        
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
        
        If K <= projectStageNumber Then  ' the stage referred to in the rule is valid
            
            fileTypeFoundInCorrectLocation = False
            requiredFileFound = False
            requiredFileWrongLocation = False
            
            For J = 3 To 10000
            
                foundFileType = Sheets("J").Cells(J, 5)
                
                If foundFileType = "" Then Exit For
                
                If InStr(1, LCase(foundFileType), LCase(checkFileType)) Then
                                        
                    ' current file is the correct type
                                        
                    ' but is it in the required location
                    currentFileTypeLocation = LCase(Sheets("J").Cells(J, 3))
                    correctFileTypeLocation = LCase(Sheets("Stages").Cells(2, 2) & "\" & projectNumber & checkFileTypeLocation)
                                     
                                        
                    If currentFileTypeLocation = correctFileTypeLocation Then
                
                        ' in correct place
                        fileTypeFoundInCorrectLocation = True
                        
                        ' now check that the required file exists
                                                
                        For jFileRow = 3 To 10000
                        
                            foundName = Sheets("J").Cells(jFileRow, 1)
                    
                            If foundName = "" Then Exit For
                            
                            ' Add the extension to the file name
                            currentFile = foundName & "." & Sheets("J").Cells(jFileRow, 5)
                            
                            If InStr(1, LCase(currentFile), LCase(requiredFile)) Then
                                
                                ' required file name
                                
                                ' but is it in the required location
                                currentFileLocation = LCase(Sheets("J").Cells(jFileRow, 3))
                                correctFilelocation = LCase(Sheets("Stages").Cells(2, 2) & "\" & projectNumber & requiredFileLocation)
                    
                                If currentFileLocation = correctFilelocation Then
                                
                                    requiredFileFound = True
                                    
                                Else
                                
                                    requiredFileWrongLocation = True
                                    
                                End If
                                
                                Exit For
                
                            End If
                        
                        Next jFileRow
                        
                        ' Don't bother checking for any more file types (once one exists, we check for the required file anyway)
                        Exit For
                                            
                    End If
                    
                
                End If
                
                
            Next J
            
            
            ' Ouput any errors (NB: order of If statements is important!)
            
            If requiredFileWrongLocation = True Then
            
                ' If the required file was found but it's in the wrong location
                        
                ' Output the error
                Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                Sheets("Dashboard").Cells(nextBlankRow, 4) = errorIfFoundInWrongLocation
                
                nextBlankRow = nextBlankRow + 1
                
            ElseIf requiredFileFound = False And fileTypeFoundInCorrectLocation = True Then
                       
                ' If there are file types present but the required file has not been found
                                  
                ' Output the error
                Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                Sheets("Dashboard").Cells(nextBlankRow, 4) = errorIfNotFound
                
                nextBlankRow = nextBlankRow + 1
                        
            End If
                     
            
        End If
         
    Next i ' next rule
    
End Sub

