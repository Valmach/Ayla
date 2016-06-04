Attribute VB_Name = "RulesPHB"

' These rule parameters have to be available from all subs
Private ruleErrorMessage As String
Private ruleOtherParameter As String

'Prevents repetitive error messages
Private alreadyShownKeyDateError As Boolean



Sub rulePHB()

    ' Each rule does something different and must be hard coded
    ' If an additional rule is added, you must also add a macro for it called phbRule#() where # is the rule number
    
    ' Note that the global variables are set by the Audit function for each project so they are available for checking here
    
    
    ' Do anything that's required before the checks commence here:
    
    ' Set the initial next blank row
    For nextBlankRow = 16 To 10000
        If Sheets("Dashboard").Cells(nextBlankRow, 1) = "" Then
        Exit For
        End If
    Next nextBlankRow
    
    alreadyShownKeyDateError = False
    
    preErrorString = Sheets("Rules PHB").Cells(3, 5)
    
    
    ' Now read through each of the rules

    For i = 12 To 100000
    
        ruleStage = Sheets("Rules PHB").Cells(i, 1)
        ruleActivated = Sheets("Rules PHB").Cells(i, 2)
        ruleErrorMessage = preErrorString & Sheets("Rules PHB").Cells(i, 4)
        ruleOtherParameter = Sheets("Rules PHB").Cells(i, 5)
        
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
        
        ' For testing purposes
        ' projectStageNumber = 100
        
        ' Check if rule is relevant for this stage
        If K <= projectStageNumber Then
        
            'Check if rule is activated
            If ruleActivated = 1 Then
                
                Dim ruleNumber As Integer
                
                ruleNumber = i - 11 ' Rule 1 in row 12 (i is row number)
                
                ' Decide which rule number to call
                If ruleNumber >= 11 And ruleNumber <= 16 Then
                    ' Rules 11 to 16 do the same thing (but rely on parameters being set above)
                    callPhbRule 11
                Else
                    ' most rules have their own sub
                    callPhbRule ruleNumber
                End If
                
            End If
        
        End If
        
    Next i
    

End Sub

Private Sub callPhbRule(ruleNumber As Integer)
    ' This is used to call the rule associated with the rule number
    On Error GoTo DynamicCallError
    Application.Run "phbRule" & ruleNumber
    Exit Sub
DynamicCallError:
    Debug.Print "Failed dynamic call: " & Err.Description
End Sub

Private Sub phbRule1()
    
    ' Checks for a project name
    If projectName = "" Then
        
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = "Error"
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
  
End Sub

Private Sub phbRule2()

    ' Checks for a project number
    If projectNumber = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = "Error"
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule3()
    
    ' Checks for a project area
    If projectArea = 0 Then ' Int variable so even if it's blank in the database it will be 0 in memory
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule4()
    
    ' Checks for a project occupancy
    If projectOccupancy = 0 Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule5()
    
    ' Checks for a project type
    If projectType = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule6()
    
    ' Checks for a project DES Roll (if projecttype is a school)
    If InStr(1, LCase(projectType), "school") Then
    
        If projectDesRoll = "" Then
    
            ' Output the error
            Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
            Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
            Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
            Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
            ' increment the output row number
            nextBlankRow = nextBlankRow + 1
        
        End If
    
    End If
    
    
End Sub
    
Private Sub phbRule7()
    
    ' Checks for a project director
    If projectDirector = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule8()
    
    ' Checks for a job runner
    If projectJobRunner = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = "Error"
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule9()
    
    ' Checks for a lead mech
    If projectMech = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule10()
    
    ' Checks for a lead elec
    If projectElec = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule11()
    
    ' Also called for rules 12-16
    
    ' Checks for a sufficient number of key dates
    If alreadyShownKeyDateError = False Then
    
        ' Grab the number of dates required (only if a number is entered)
        If IsNumeric(ruleOtherParameter) Then
        
            minDates = Int(ruleOtherParameter) ' Cast it to an integer
            datesEntered = 0    ' track how many dates have been entered
            
            ' Cycle through the key dates array and check how many contain text
            For arrayIndex = 0 To UBound(projectKeyDates) - 1
            
                If (projectKeyDates(arrayIndex) <> "") Then
                
                    datesEntered = datesEntered + 1
                
                End If
            
            Next arrayIndex
            
            ' If there hasn't been enough key dates entered, output an error
            If datesEntered < minDates Then
            
                ' Output the error
                Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
                Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
                Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
                ' increment the output row number
                nextBlankRow = nextBlankRow + 1
                
                ' Prevent this error from being shown again from future rules
                alreadyShownKeyDateError = True
            
            End If
            
        Else
        
            ' If the ruleOtherparameter is not a number
            MsgBox ("PHB Rule for checking minimum number of dates could not be carried out as the x value was not defined" & vbNewLine & vbNewLine & "See column E row 22-27 on Rules PHB sheet")
            
        End If
        
    End If
    
End Sub

Private Sub phbRule17()
    
    ' Checks for professions
    If projectProfessions = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule18()
    
    ' Checks for a project description
    
    ' This is the string if there's rtf but no text (but some of the characters can change, so examine the lenght instead)
    baseString = ruleOtherParameter
    
    ' Add a tolerance incase minor character changes in baseString (also, the description shjould be longer than 5 characters anyway)
    baseLength = Len(baseString) + 5
        
    If projectDesc = "" Or Len(projectDesc) < baseLength Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule19()
    
    ' Checks for risks
    
    ' This is the string if there's rtf but no text (but some of the characters can change, so examine the lenght instead)
    baseString = ruleOtherParameter
    
    ' Add a tolerance incase minor character changes in baseString (also, the description should be longer than 5 characters anyway)
    baseLength = Len(baseString) + 5
        
    If projectRisks = "" Or Len(projectDesc) < baseLength Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub

Private Sub phbRule20()
    
    ' Checks for a project location address
    If projectAddress = "" Then
    
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = projectNumber
        Sheets("Dashboard").Cells(nextBlankRow, 2) = projectName
        Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
        Sheets("Dashboard").Cells(nextBlankRow, 4) = ruleErrorMessage
        ' increment the output row number
        nextBlankRow = nextBlankRow + 1
    
    End If
    
End Sub
