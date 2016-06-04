Attribute VB_Name = "Rule1"
Sub ruleone()

    notFoundPreString = Sheets("Rules 1").Cells(7, 9)
    wrongLocationPreString = Sheets("Rules 1").Cells(7, 10)

    For i = 12 To 100000 ' read through rules
    
        ruleStage = Sheets("Rules 1").Cells(i, 1)
        string1 = Sheets("Rules 1").Cells(i, 2)
        string2 = Sheets("Rules 1").Cells(i, 3)
        string3 = Sheets("Rules 1").Cells(i, 4)
        string4 = Sheets("Rules 1").Cells(i, 5)
        string5 = Sheets("Rules 1").Cells(i, 6)
        string6 = Sheets("Rules 1").Cells(i, 7)
        ruleFileReqLocation = Sheets("Rules 1").Cells(i, 8)
        errorIfNotFound = Sheets("Rules 1").Cells(i, 9)
        errorIfFoundInWrongPlace = Sheets("Rules 1").Cells(i, 10)
        If ruleStage = "" Then Exit For
        
        ' For Testing (stop at a particular rule)
        If i = 37 Then
            test = "put a breakpoint here"
        End If
    
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
        
            foundCorrectFile = False
            
            For J = 3 To 10000
            
                foundName = Sheets("J").Cells(J, 1)
                foundPath = Sheets("J").Cells(J, 3)
                
                ' For Testing (stop at a particular J drive file)
                If J = 44 Then
                    test = "put a breakpoint here"
                End If
                
                If foundName = "" Then Exit For
                
                ' Check if it's in an excluded path (checks name & path)
                excludedFilePath = False
                For excludedStringRow = 3 To 100
                
                    excludedString = Sheets("Rules 1").Cells(excludedStringRow, 11)
                    
                    ' If it's empty don't continue
                    If excludedString = "" Then Exit For
                    
                    ' Check the path and name for the excluded string
                    ' But don't check the name for SS as it pops up in a number of file names such as 'risk assessment'
                    If InStr(1, LCase(foundPath), LCase(excludedString)) Or (InStr(1, LCase(foundName), LCase(excludedString)) And excludedString <> "SS") Then
                        
                        excludedFilePath = True
                        Exit For
                    
                    End If
                
                Next excludedStringRow
                
                If InStr(1, LCase(foundName), LCase("template")) Or excludedFilePath = True Then
                
                    ' Don't include if it has template in the name or if it's in an excluded path
                
                Else
                    
                    ' Check whether the file's name has the required strings
                    checkFile = False
                    
                    ' Check combination 1 & 2 & 3
                    If string1 <> "" Then
                        If InStr(1, LCase(foundName), LCase(string1)) Then
                            If InStr(1, LCase(foundName), LCase(string2)) Then
                                If InStr(1, LCase(foundName), LCase(string3)) Then
                                    checkFile = True
                                End If
                            End If
                        End If
                    End If
                    
                    
                    ' Check combination 4 & 5 & 6 (if wasn't found already in 1,2,3)
                    If string4 <> "" And checkFile = False Then
                        If InStr(1, LCase(foundName), LCase(string4)) Then
                            If InStr(1, LCase(foundName), LCase(string5)) Then
                                If InStr(1, LCase(foundName), LCase(string6)) Then
                                    checkFile = True
                                End If
                            End If
                        End If
                    End If
                                        
                    
                    ' Check if the filename contains (1 & 2 & 3) Or (4 & 5 & 6) - only if 1 and 4 are not empty respectively
                    If checkFile Then
                        
                        ' the file we are looking for has been found
                        
                        foundCorrectFile = True ' Could potentially still be wrong, we only know that the searched for term is in the name and that "template" is not
                        
                        ' but is it in the right location
                        foundlocation = LCase(Sheets("J").Cells(J, 3))
                        rulelocation = LCase(Sheets("Stages").Cells(2, 2) & "\" & projectNumber & ruleFileReqLocation)
                                           
                        ' By using InStr() we can account for subfolders
                        ' If foundlocation = rulelocation Then
                        If InStr(1, LCase(foundlocation), LCase(rulelocation)) Then
                        
                            ' in correct place
                            Exit For
                        
                        Else
                            
                            ' in the wrong location
                            ' output the error
                            
                            Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                            Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                            Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                            Sheets("Dashboard").Cells(nextBlankRow, 4) = wrongLocationPreString & errorIfFoundInWrongPlace
                            Sheets("Dashboard").Cells(nextBlankRow, 5) = foundlocation
                            Sheets("Dashboard").Cells(nextBlankRow, 6) = foundName
                            
                            nextBlankRow = nextBlankRow + 1
                            
                            Exit For
                            
                        End If
                    
                    End If
                
                
                End If
                                
                
            Next J
            
            ' What Ayla was searching for
            Dim searchString As String
            
            searchString = "Filename must contain: "
            initialLength = Len(searchString)
            
            If string1 <> "" Then
                searchString = searchString & string1 & ", "
            End If
            If string2 <> "" Then
                searchString = searchString & string2 & ", "
            End If
            If string3 <> "" Then
                searchString = searchString & string3
            End If
            
            ' if nothing was added from the first 3 strings, use the next 3
            If Len(searchString) <= initialLength Then
                If string4 <> "" Then
                    searchString = searchString & string4 & ", "
                End If
                If string5 <> "" Then
                    searchString = searchString & string5 & ", "
                End If
                If string6 <> "" Then
                    searchString = searchString & string6
                End If
            End If
            
            ' If there still isn't anything added
            If Len(searchString) <= initialLength Then
                searchString = "Ayla couldn't find what she was looking for"
            End If
            
            ' Output if not found
            If foundCorrectFile = False Then
                            
                Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                Sheets("Dashboard").Cells(nextBlankRow, 4) = notFoundPreString & errorIfNotFound
                Sheets("Dashboard").Cells(nextBlankRow, 6) = searchString
                
                nextBlankRow = nextBlankRow + 1
                        
            End If
                     
            
        End If
         
    Next i ' next rule
    
End Sub
