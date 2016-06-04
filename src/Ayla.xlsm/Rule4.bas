Attribute VB_Name = "Rule4"
Sub rulefour()

    ' This rule searches through the J drive files for doc or docx type files.
    ' Each time a correct file is found, it's opened for inspection
    ' Each of the activated rules are checked in the active document
    ' Once finished, the document is closed and the next J file is inspected
    
    ' Cross our fingers...
    On Error Resume Next
    
    ' Only create a word object once (change it within the loops)
    Set objWord = CreateObject("word.application")
    
    objWord.Visible = False
    
    ' For testing purposes
    ' objword.Visible = True
        
    ' Cycle through each of the j drive files
    For J = 3 To 10000
    
        ' For Testing (stop at a particular J drive file)
        If J = 11 Then
            test = "put a breakpoint here"
        End If
                                         
        ' Grab the current file properties (against which we'll check for duplicates/copies)
        currentFileName = Sheets("J").Cells(J, 1)
        currentFileType = Sheets("J").Cells(J, 5)
        currentFilePath = Sheets("J").Cells(J, 3)
        currentFileFullpath = currentFilePath & currentFileName & "." & currentFileType
        
        ' Don't bother continuing if we're at the end
        If currentFileName = "" Then Exit For
        
        ' Don't bother continuing if it's not the right file type
        If currentFileType = "doc" Or currentFileType = "docx" Then
        
            ' Check if it's in an excluded path
            excludedFilePath = False
            For excludedStringRow = 3 To 100
            
                excludedString = Sheets("Rules 4").Cells(excludedStringRow, 6)
                
                ' If it's empty don't continue
                If excludedString = "" Then Exit For
                
                If InStr(1, LCase(currentFileFullpath), LCase(excludedString)) Then
                    
                    excludedFilePath = True
                    Exit For
                
                End If
            
            Next excludedStringRow
            
            ' Check if the file needs to be checked
            nameContainsIncludedString = False
            For includedStringRow = 3 To 100
            
                includedString = Sheets("Rules 4").Cells(includedStringRow, 7)
                
                ' If it's empty don't continue
                If includedString = "" Then Exit For
                
                If InStr(1, LCase(currentFileName), LCase(includedString)) Then
                    
                    nameContainsIncludedString = True
                    Exit For
                
                End If
            
            Next includedStringRow
            
            
            ' Only continue with this file if it's not in an excluded path and it's name has a string that we're looking for
            ' Or if it contains the specific string provided
            If excludedFilePath = False And nameContainsIncludedString = True Then
                                
                ' CODE TO OPEN THE CURRENT FILE FOR INSPECTION GOES HERE
                
                ' Open the current document
                Set wordDoc = objWord.documents.Open(Filename:=currentFileFullpath, ReadOnly:=True)
                
                ' Make sure it's not empty (it will be if there's an error opening it)
                If Not wordDoc Is Nothing Then
                
                    For i = 12 To 100000 ' read through rules
                    
                        ' For Testing (stop at a particular rule)
                        If i = 37 Then
                            test = "put a breakpoint here"
                        End If
        
                        ruleStage = Sheets("Rules 4").Cells(i, 1)
                        ruleSpecificFileName = Sheets("Rules 4").Cells(i, 2)
                        ruleActivated = Sheets("Rules 4").Cells(i, 3)
                        rulePhrase = Sheets("Rules 4").Cells(i, 4)
                        errorIfPhraseFound = Sheets("Rules 4").Cells(i, 5)
                        phraseFound = False
                        
                        If ruleStage = "" Then Exit For
                        
                        ' Check if a specific file name was provided for the check (if not, continue anyway)
                        continueCheckingFile = True
                        
                        If ruleSpecificFileName <> "" Then
                            
                            ' A specific file name to check has been provided
                            
                            If InStr(1, LCase(currentFileName), LCase(ruleSpecificFileName)) Then
                                ' The current file does match the specific file
                                continueCheckingFile = True
                            Else
                                continueCheckingFile = False
                            End If
                        
                        End If
                        
                        If continueCheckingFile Then
                        
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
                            
                            ' For testing purposes
                            ' projectStageNumber = 100
                            
                            If K <= projectStageNumber Then  ' the stage referred to in the rule is valid
                            
                                ' CODE FOR CHECKING FOR PHRASES HERE
    
                               With wordDoc.Content.Find
                                    Do While .Execute(findText:=rulePhrase, Forward:=True, Format:=True, MatchWholeWord:=False, MatchCase:=False, Wrap:=wdFindStop) = True
                                    
                                        ' Known Error: For some reason some documents find phrases that actually aren't there
                                        '   This is easily spoted on the dashboard as each phrase is found (which is unlikely to be genuine)
                                        '   Haven't had time to look into why this is happening
                                        
                                        ' There is a macro in the Audit module called "removePhraseErrors" that removes these false errors
                                        
                                        ' The current phrase has been found
                                        ' output the error and exit the loop (we don't want multiple errors for the same phrase)
        
                                        ' Output the error
                                        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                                        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                                        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                        Sheets("Dashboard").Cells(nextBlankRow, 4) = errorIfPhraseFound
                                        Sheets("Dashboard").Cells(nextBlankRow, 5).Formula = "=HYPERLINK(""" & currentFileFullpath & """,""" & currentFileFullpath & """)"
        
                                        nextBlankRow = nextBlankRow + 1
        
                                        Exit Do
    
                                    Loop
                                End With
                                                        
                                       
                            End If
                        
                        End If
                            
                    Next i ' next rule
                
                End If
                
                ' Close the current document
                wordDoc.Close
            
            End If
        
        End If
                        
        
                                                
    Next J
    
    ' Quit word
    objWord.Quit
    Set objWord = Nothing
    
End Sub



