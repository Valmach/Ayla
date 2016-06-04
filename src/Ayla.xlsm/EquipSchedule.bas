Attribute VB_Name = "EquipSchedule"
Private Const errorBeginning = "Equip Schedule: "   ' all errors to be preceeded by
Private Const errorIfReqSheetNotFound = " schedule was found but none of the following schedules were found: " ' to be preceeded by reference sheet name and followed by required sheet name(s)
Private Const errorIfTypeSheetNotFound = " requires the following schedule: " ' to be preceeded by project type and followed by required schedule

' The strings used to find the equipment schedule
Private Const fileNameContains = "Equipment Schedule"
Private Const filePathContains = "\Specs\Mechanical\"

Sub checkEquipSchedule()

    ' This sub searches checks the project's mechanical equipment schedule (if it exists).
    ' Refer to the "Equip Schedule" sheet for rules applied
    
    ' All issues found are assigned to the Mech Engineer (if present) rather than the Project Runner
    
    ' Cross our fingers...
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    ' A single Stage is defined for all checks to be carried out from
    ruleStage = Sheets("Equip Schedule").Cells(1, 4)
                
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
        
        ' First things first, make sure we have an equipment schedule...
        ' Cycle through the j Drive (don't use the assumed path in case the schedule's name has been altered slightly)
        For J = 3 To 10000
                
            ' Grab the current file properties and construct the full path
            currentFileName = Sheets("J").Cells(J, 1)
            currentFileType = Sheets("J").Cells(J, 5)
            currentFilePath = Sheets("J").Cells(J, 3)
            currentFileFullpath = currentFilePath & currentFileName & "." & currentFileType
            
            If currentFileName = "" Then Exit For     ' at the end, don't continue
            
            If J = 199 Then
                test = "put a breakpoint here"
            End If
                    
            ' A couple of checks to make sure it's the right file
            If InStr(1, LCase(currentFileFullpath), LCase(filePathContains)) And InStr(1, LCase(currentFileFullpath), LCase(fileNameContains)) And (currentFileType = "xls" Or currentFileType = "xlsx" Or currentFileType = "xlsm") Then
                                    
                ' There is an equipment schedule and we now have its path
                ' Crack open the workbooks...
                
                ' Define the workbooks we're going to be using
                Dim aylaWb, equipWb As Workbook
                ' Must also have a reference to Ayla otherwise the opened workbooks becomes the default one referred to
                Set aylaWb = ThisWorkbook
                
                ' Grab the worksheets we'll mainly be using
                Dim dashboard, schedule As Worksheet
                Set dashboard = aylaWb.Sheets("Dashboard")
                Set schedule = aylaWb.Sheets("Equip Schedule")
                
                ' There's a known bug that makes the macro stop when you open a workbook in excel
                ' Seemingly, you need to check if the shift button is pressed
                ' See link: https://support.microsoft.com/en-us/kb/555263
                
                ' Open up with CalculationsM workbook
                Do While ShiftPressed()
                    DoEvents
                Loop
                Set equipWb = Workbooks.Open(currentFileFullpath, False) ' false prevents links updating (and the annoying pop up that comes with it)
                  
                ' At this point both workbooks are open and available
                
                ' Check what date it was created (needed so we only apply the relevant checks below)
                created = aylaWb.Sheets("J").Cells(J, 8)
                
                
                ' START CHECKS HERE
                
                ' See function at bottom of module for SheetExists()
                
                ' Cycle through the sheets that should exist if another sheet exists
                For ruleRow = 11 To 1000
                
                    If ruleRow = 13 Then
                        test = 5
                    End If
                
                    Dim refSheetName  As String
                    Dim reqSheet1 As String
                    Dim reqSheet2 As String
                    Dim reqSheet3 As String
                    refSheetName = schedule.Cells(ruleRow, 3)
                
                    If refSheetName = "" Then Exit For    ' if no more rules, quit
                
                    ' Make sure rule is active
                    If schedule.Cells(ruleRow, 1) = 1 Then
                    
                        ' Check if rule is valid for this schedule (wasn't created too long ago)
                        If schedule.Cells(ruleRow, 2) < created Then
                        
                            ' Rule is valid for this project, check if the reference sheet exists
                            If SheetExists(refSheetName, equipWb) Then
                                                    
                                ' The ref sheet exists, check if any of the three required sheets exist
                                reqSheet1 = schedule.Cells(ruleRow, 4)
                                reqSheet2 = schedule.Cells(ruleRow, 5)
                                reqSheet3 = schedule.Cells(ruleRow, 6)
                                
                                requiredScheduleFound = False
                                
                                If SheetExists(reqSheet1, equipWb) Then
                                    requiredScheduleFound = True
                                ElseIf SheetExists(reqSheet2, equipWb) Then
                                    requiredScheduleFound = True
                                ElseIf SheetExists(reqSheet3, equipWb) Then
                                    requiredScheduleFound = True
                                End If
                                
                                If Not requiredScheduleFound Then
                                
                                    ' Didn't find the required sheet (any of them, only 1 needed)
                                    ' Construct the error
                                    errorToOutput = errorBeginning & refSheetName & errorIfReqSheetNotFound & reqSheet1
                                    
                                    If reqSheet2 <> "" Then
                                        errorToOutput = errorToOutput & ", " & reqSheet2
                                    End If
                                    If reqSheet3 <> "" Then
                                        errorToOutput = errorToOutput & ", " & reqSheet3
                                    End If
                                    
                                    linkDesc = "Open Equipment Schedule"
                                    
                                    ' Output the error
                                    dashboard.Cells(nextBlankRow, 1) = projectNumber
                                    dashboard.Cells(nextBlankRow, 2) = projectName
                                    ' Check for project mech engineer
                                    If projectMech <> "" Then
                                        ' Assign to mech if present
                                        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                    Else
                                        ' otherwise assign to runner
                                        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                    End If
                                    dashboard.Cells(nextBlankRow, 4) = errorToOutput
                                    dashboard.Cells(nextBlankRow, 5).Formula = "=HYPERLINK(""" & currentFileFullpath & """,""" & linkDesc & """)"
                                    
                                    nextBlankRow = nextBlankRow + 1
                                
                                
                                End If
                                
                            
                            End If
                            
                        
                        End If
                    
                    
                    End If
                
                Next ruleRow
                
                
                ' Cycle through the sheets that should exist if the project is a specific type
                For ruleRow = 11 To 1000
                    
                    Dim reqSheet As String
                    
                    projectTypeForRule = schedule.Cells(ruleRow, 13)
                
                    If projectTypeForRule = "" Then Exit For    ' if no more rules, quit
                    
                    If ruleRow = 17 Then
                        test = 5
                    End If
                
                    ' Make sure rule is active
                    If schedule.Cells(ruleRow, 11) = 1 Then
                    
                        ' Check if rule is valid for this schedule (wasn't created too long ago)
                        If schedule.Cells(ruleRow, 12) < created Then
                        
                            ' Rule is valid for this project, check if it's the correct type
                            If LCase(projectTypeForRule) = LCase(projectType) Or LCase(projectTypeForRule) = LCase("Any") Then
                                                    
                                ' The project is the correct type, check if the required sheet exists
                                reqSheet = schedule.Cells(ruleRow, 14)
                                                            
                                If Not SheetExists(reqSheet, equipWb) Then
                                   
                                    ' The required sheet doesn't exist
                                    ' Construct the error
                                    errorToOutput = errorBeginning & projectTypeForRule & errorIfTypeSheetNotFound & reqSheet
                                                                    
                                    linkDesc = "Open Equipment Schedule"
                                    
                                    ' Output the error
                                    dashboard.Cells(nextBlankRow, 1) = projectNumber
                                    dashboard.Cells(nextBlankRow, 2) = projectName
                                    ' Check for project mech engineer
                                    If projectMech <> "" Then
                                        ' Assign to mech if present
                                        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                    Else
                                        ' otherwise assign to runner
                                        aylaWb.Sheets("Dashboard").Cells(nextBlankRow, 3) = projectJobRunner
                                    End If
                                    dashboard.Cells(nextBlankRow, 4) = errorToOutput
                                    dashboard.Cells(nextBlankRow, 5).Formula = "=HYPERLINK(""" & currentFileFullpath & """,""" & linkDesc & """)"
                                    
                                    nextBlankRow = nextBlankRow + 1
                                    
                                End If
                                
                            
                            End If
                            
                        
                        End If
                    
                    
                    End If
                
                Next ruleRow
                
                
                ' CHECKS ARE COMPLETE HERE
                
                
                ' Get rid of the workbook references
                equipWb.Close (False)    ' false = don't save
                Set equipWb = Nothing
                Set aylaWb = Nothing
                
                ' Don't check any more J files, we've already found it and done our checks
                Exit For
            
            End If
            
            
        Next J
    
    
    End If
    
    Application.ScreenUpdating = True

End Sub

Public Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean

    ' http://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
    
    Dim sht As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
     
 End Function
