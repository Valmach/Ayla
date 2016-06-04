Attribute VB_Name = "Audit"
' Database variables are reassigned each time a new project is audited in the below function
' Database variables are public so any module can use them

Public projectRow As Integer
Public projectNumber As String
Public projectName As String
Public projectStage As String
Public projectDirector As String
Public projectJobRunner As String
Public projectStageNumber As Integer
Public projectIsValid As Boolean
Public projectArea As Double
Public projectElec As String
Public projectMech As String
Public projectStartDate As Date
Public projectCompletionYear As String
Public projectCompletionDate As Date
Public projectKeyDates(7) As String ' there are 7 key date entries (arrays are 0 based)
Public projectDesc As String
Public projectType As String
Public projectNew As String 'yes or no (or Yes or No etc. so use LCase())
Public projectOccupancy As Integer
Public projectRisks As String
Public projectProfessions As String
Public projectDesRoll As String
Public projectAddress As String

' This saves us having to cycle through to find the next available row each time there's an error to output
Public nextBlankRow As Integer

' keep track of total number of unindexed mech calcs found and total number of uncheckedcalcs
Public totalNumberOfUnindexedMechCalcs As Integer
Public totalNumberOfUnCheckedMechCalcs As Integer

' The check all projects and check specific project buttons set a start row and number of projects to check on the dashboard
' The cells used to record these are hidden off to the right
' The text in these cells are white so the user will not see them
' This work around is required due to how Ayla has developed over time (I don't want to rewrite a load of code)

Sub checkAllProjects()

    ' turn off error reporting
    Application.DisplayAlerts = False

    ' Set the start project row and number of projects to check
    Sheets("Dashboard").Cells(3, 14) = 2
    Sheets("Dashboard").Cells(4, 14) = 1000
    
    ' Start the check
    Call Audit

End Sub

Sub checkSpecificProject()
    
    ' turn off error reporting
    Application.DisplayAlerts = False

    ' Grab the specific project number entered by the user
    specificProjectNumber = Sheets("Dashboard").Cells(4, 4)
    
    ' Find the project
    For databaseRow = 2 To 10000
        
        ' Grab the current project number from the database
        currentProjectNumber = Sheets("Database").Cells(databaseRow, 2)
        
        ' if it's at the end, without finding the project, throw an error and stop running
        If currentProjectNumber = "" Then
            MsgBox ("The specific project number was not found.")
            Exit Sub
        End If
        
        ' Check if it's the correct project number, if so, the required row will be the databaseRow
        If currentProjectNumber = specificProjectNumber Then Exit For
    
    Next databaseRow


    ' Set the start project row and number of projects to check
    Sheets("Dashboard").Cells(3, 14) = databaseRow
    Sheets("Dashboard").Cells(4, 14) = 1
    
    ' Start the check
    Call Audit

End Sub

Private Sub Audit()

    On Error Resume Next

    ' Ouput the start time
    Sheets("Dashboard").Cells(5, 5) = "Start Time:"
    Sheets("Dashboard").Cells(6, 5) = Now
    
    ' Clear the end time
    Sheets("Dashboard").Cells(9, 5) = "Still Running..."
        
    ' Refresh the database
    ' Note that it is necessary to turn off "Enable background refresh" on the data connection for this to work in VBA
    ' Go to the data tab in the ribbon -> Connections section -> Connections -> select "projectnames" and hit properties in the window that appears -> uncheck the option under refresh control
    ' See link: http://www.mrexcel.com/forum/excel-questions/388633-refreshing-data-connections-through-visual-basic-applications-only-working-if-macro-stepped-through-debugger.html
    ActiveWorkbook.RefreshAll

    Dim SourceFolderName As String
            
    ' Clear the dashboard
    Call cleardash
    
    'Starting & ending point for the audit (make a UI for these eventually)
    Dim startProjectRow, numberOfProjectsToCheck As Integer
    startProjectRow = Sheets("Dashboard").Cells(3, 14)
    numberOfProjectsToCheck = Sheets("Dashboard").Cells(4, 14)
        
    ' Check valid inputs
    If Not IsNumeric(startProjectRow) Or Not IsNumeric(numberOfProjectsToCheck) Then
        MsgBox ("invalid start row or number of projects entered")
        Exit Sub
    End If
    
    ' Use Int in case the user entered a decimal
    endProjectRow = Int(startProjectRow) + Int(numberOfProjectsToCheck) - 1 ' minus one because it's a for loop based on row numbers
    
    Excel.Application.ScreenUpdating = True
    
    ' Set up dashboard progress
    Sheets("Dashboard").Cells(7, 3) = "Project Number:"
    Sheets("Dashboard").Cells(8, 3) = "Project Name:"
    Sheets("Dashboard").Cells(9, 3) = "Job Runner:"
    Sheets("Dashboard").Cells(11, 3) = "Status:"
    Sheets("Dashboard").Cells(13, 3) = "Error Count:"
    
    ' Clear the error count
    Sheets("Dashboard").Cells(13, 4) = ""
    
    ' Set the initial next blank row
    For nextBlankRow = 16 To 10000
        If Sheets("Dashboard").Cells(nextBlankRow, 1) = "" Then
        Exit For
        End If
    Next nextBlankRow
    
    ' keep track of # of projects audited and skipped
    audited = 0
    skipped = 0
    
    ' keep track of total number of unidexed mech calcs
    totalNumberOfUnindexedMechCalcs = 0
    totalNumberOfUnCheckedMechCalcs = 0
    
    Dim i As Integer
    For i = startProjectRow To endProjectRow ' work through valid project
    
        ' Update dashboard progress
        Sheets("Dashboard").Cells(7, 4) = Sheets("Database").Cells(i, 2)    'Number
        Sheets("Dashboard").Cells(8, 4) = Sheets("Database").Cells(i, 3)    'Name
        Sheets("Dashboard").Cells(9, 4) = Sheets("Database").Cells(i, 9)    'Runner
        Sheets("Dashboard").Cells(11, 4) = "Checking Project..."            'Status
        
        If Sheets("Database").Cells(i, 2) = "" Then Exit For
        
        Application.Wait (Now + TimeValue("00:00:01"))
        Excel.Application.ScreenUpdating = False
        
        ' Check if this project is valid
        projectIsValid = False
        Call checkIfprojectIsValid(i)   ' also assigns global variables such as project number and runner (also makes sure type isn't PSDP)
                
        If projectIsValid = True Then
            
            ' Increment the number of projects audited
            audited = audited + 1
        
            ' Clear the previous jobs files
            Call clearj
            
            ' Which folder on the j drive to check
            drivepath = Sheets("Stages").Cells(2, 2)
            SourceFolderName = drivepath & "\" & projectNumber & "\"
            
            ' read all files
            Excel.Application.ScreenUpdating = True
            Sheets("Dashboard").Cells(11, 4) = "Reading J Drive..."
            Application.Wait (Now + TimeValue("00:00:01"))
            Excel.Application.ScreenUpdating = False
            
            Call readjdrive(SourceFolderName)
            
            ' Check what's activated
            If Sheets("Rules PHB").Cells(1, 1) = 1 Then
                checkRulePHB = True
            Else
                checkRulePHB = False
            End If
            If Sheets("Rules 1").Cells(1, 1) = 1 Then
                checkRule1 = True
            Else
                checkRule1 = False
            End If
            If Sheets("Rules 2").Cells(1, 1) = 1 Then
                checkRule2 = True
            Else
                checkRule2 = False
            End If
            If Sheets("Rules 3").Cells(1, 1) = 1 Then
                checkRule3 = True
            Else
                checkRule3 = False
            End If
            If Sheets("Rules 4").Cells(1, 1) = 1 Then
                checkRule4 = True
            Else
                checkRule4 = False
            End If
            If Sheets("Equip Schedule").Cells(1, 1) = 1 Then
                checkEquipScheduleEnabled = True
            Else
                checkEquipScheduleEnabled = False
            End If
            If Sheets("Misc").Cells(11, 1) = 1 Then
                recordPhbEnabled = True
            Else
                recordPhbEnabled = False
            End If
            If Sheets("Misc").Cells(12, 1) = 1 Then
                checkPathLengths = True
            Else
                checkPathLengths = False
            End If
            If Sheets("Misc").Cells(13, 1) = 1 Then
                checkMechanicalCalcs = True
            Else
                checkMechanicalCalcs = False
            End If
            If Sheets("Misc").Cells(14, 1) = 1 Then
                checkWordDocsEnabled = True
            Else
                checkWordDocsEnabled = False
            End If
            
            
            ' check if any files were found
            If Sheets("J").Cells(3, 1) <> "" Then
            
                ' Run Rules for J drive files
                
                ' Rule 1
                If checkRule1 Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Rules 1..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call ruleone
                End If


                ' Rule 2
                If checkRule2 Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Rules 2..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call ruletwo
                End If

                ' Rule 3
                If checkRule3 Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Rules 3..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call rulethree
                End If

                ' Rule 4
                If checkRule4 Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Rules 4..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call rulefour
                End If

                ' Equipment Schedule
                If checkEquipScheduleEnabled Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Equipment Schedule..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call checkEquipSchedule
                End If

                ' Check Path Lengths
                If checkPathLengths Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Path Lengths..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call checkForMaxPathLengths
                End If


                ' Check checkCalcsM
                If checkMechanicalCalcs Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking Mechanical Calcs..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call checkCalcsM
                End If

                ' Check checkWordDocs
                If checkWordDocsEnabled Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Turning on spell check..."
                    Application.Wait (Now + TimeValue("00:00:01"))
                    Excel.Application.ScreenUpdating = False

                    Call checkWordDocs
                End If


                ' Run rules for the PHB

                ' PHB Rules
                If checkRulePHB Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Checking PHB Rules..."
                    Application.Wait (Now + TimeValue("00:00:02"))
                    Excel.Application.ScreenUpdating = False

                    Call rulePHB
                End If

                ' PHB Record
                If recordPhbEnabled Then
                    Excel.Application.ScreenUpdating = True
                    Sheets("Dashboard").Cells(11, 4) = "Recording PHB settings..."
                    Application.Wait (Now + TimeValue("00:00:02"))
                    Excel.Application.ScreenUpdating = False

                    Call recordPhbData(projectNumber, i)
                End If
                                
                
            Else
                
                ' Empty J drive found
                
                ' Output the error
                Sheets("Dashboard").Cells(nextBlankRow, 1) = Sheets("Database").Cells(i, 2)
                Sheets("Dashboard").Cells(nextBlankRow, 2) = Sheets("Database").Cells(i, 3)
                Sheets("Dashboard").Cells(nextBlankRow, 4) = "Project is live but there are no files on the J drive"
                Sheets("Dashboard").Cells(nextBlankRow, 3) = Sheets("Database").Cells(i, 9)
                                    
                nextBlankRow = nextBlankRow + 1
            
            End If
            
            
        Else
        
            ' invalid project, increment the skipped variable
            skipped = skipped + 1
        
        End If
                
        ' Only update UI if didn't skip project
        Excel.Application.ScreenUpdating = True
        
        ' Update dashboard progress
        Sheets("Dashboard").Cells(11, 4) = "Finished Checking Project"   'Status
        Application.Wait (Now + TimeValue("00:00:01"))
        Excel.Application.ScreenUpdating = False
        
        ' Include a delay so that the user can see what's happening and BREAK if needed
'        For delay = 3 To 1 Step -1
'
'            Sheets("Dashboard").Cells(11, 4) = "Next Project in " & delay & "..."   'Status
'            Application.Wait (Now + TimeValue("00:00:01"))
'
'        Next delay
        
                
    Next i
        
    
    Excel.Application.ScreenUpdating = True
    
    ' Clear project info
    Sheets("Dashboard").Cells(7, 4) = ""    'Number
    Sheets("Dashboard").Cells(8, 4) = ""    'Name
    Sheets("Dashboard").Cells(9, 4) = ""    'Runner
            
    ' There's no need to import any new exceptions because they won't be approved anyway
'    ' Import Exceptions
'    Sheets("Dashboard").Cells(11, 4) = "Importing exceptions..."
'    Application.Wait (Now + TimeValue("00:00:02"))
'    Excel.Application.ScreenUpdating = False
'
'    Call importExceptions

    ' Remove Phrase Errors (instances when all phrases were found in a doc)
    Excel.Application.ScreenUpdating = True
    Sheets("Dashboard").Cells(11, 4) = "Removing invalid phrase errors..."
    Application.Wait (Now + TimeValue("00:00:02"))
    Excel.Application.ScreenUpdating = False
    
    Call removePhraseErrors
    
    ' Exclude Exceptions
    Excel.Application.ScreenUpdating = True
    Sheets("Dashboard").Cells(11, 4) = "Removing approved exceptions..."
    Application.Wait (Now + TimeValue("00:00:02"))
    Excel.Application.ScreenUpdating = False
    
    Call removeExceptions
    
    Excel.Application.ScreenUpdating = True
    
    ' Update dashboard progress
    Sheets("Dashboard").Cells(7, 4) = ""    'Number
    Sheets("Dashboard").Cells(8, 4) = ""    'Name
    Sheets("Dashboard").Cells(9, 4) = ""    'Runner
    Sheets("Dashboard").Cells(11, 4) = "Audit Complete. Successfully auditted " & audited & " projects. (Skipped " & skipped & " projects.)"   'Status
        
    ' Update the end time
    Sheets("Dashboard").Cells(8, 5) = "End Time:"
    Sheets("Dashboard").Cells(9, 5) = Now
    
    ' Count the errors
    itemCount = 0
    For dashboardRow = 16 To 100000
    
        If Sheets("Dashboard").Cells(dashboardRow, 1) = "" Then
            Exit For
        Else
            itemCount = itemCount + 1
        End If
        
    Next dashboardRow
    
    ' Output the count
    Sheets("Dashboard").Cells(13, 3) = "Error Count:"
    Sheets("Dashboard").Cells(13, 4) = itemCount
    
    
End Sub

Sub checkIfprojectIsValid(projectRow)
    
    projectRow = Int(projectRow)

    ' Check if project number is empty
    If Sheets("Database").Cells(projectRow, 2) = "" Then
        Exit Sub    ' just exit, projectIsValid is already set to false
    End If

    ' Checks if the current project is valid
    skipit = 0
    
    projectStage = LCase(Sheets("Database").Cells(projectRow, 19))
    
    ' if the stage is blank
    If projectStage = "" Then
        
        ' Output the error
        Sheets("Dashboard").Cells(nextBlankRow, 1) = Sheets("Database").Cells(projectRow, 2)
        Sheets("Dashboard").Cells(nextBlankRow, 2) = Sheets("Database").Cells(projectRow, 3)
        Sheets("Dashboard").Cells(nextBlankRow, 4) = "CRITICAL: Stage entered in the PHB is blank. Project could not be audited."
        Sheets("Dashboard").Cells(nextBlankRow, 3) = Sheets("Database").Cells(projectRow, 9)
        
        nextBlankRow = nextBlankRow + 1
        
        projectNumber = Sheets("Database").Cells(projectRow, 2)
        
        If projectNumber = "" Then End ' must be at the end of the database
    
        skipit = 1 ' get us to next i
                        
    End If
    
    If skipit = 0 Then
            
        ' check if it is a valid stage
        For J = 2 To 30 ' traverse valid stage options
        
            validStage = LCase(Sheets("Stages").Cells(J, 1))
                        
            If validStage = "" Then
            
                ' end of valid stages reached so check if we are to ignore it ' but we first check if it is an unknown stage
                skipit = 1
                
                For K = 33 To 40    ' Rows with stage to ignore
                
                    validStage = LCase(Sheets("Stages").Cells(K, 1))
                    
                    If InStr(1, LCase(validStage), LCase(projectStage)) Then ' it is a stage to ingore
                    
                        projectNumber = Sheets("Database").Cells(projectRow, 2)
                        
                        If projectNumber = "" Then End ' must be at the end of the database
                        
                        If validStage = "" Then
                        
                            ' the stage is unknown so record that and skip to next stage
                           
                            Sheets("Dashboard").Cells(nextBlankRow, 1) = Sheets("Database").Cells(projectRow, 2)
                            Sheets("Dashboard").Cells(nextBlankRow, 2) = Sheets("Database").Cells(projectRow, 3)
                            Sheets("Dashboard").Cells(nextBlankRow, 4) = "CRITICAL: Stage entered in the PHB is not recognised - " & projectStage
                            Sheets("Dashboard").Cells(nextBlankRow, 3) = Sheets("Database").Cells(projectRow, 9)
                            
                            nextBlankRow = nextBlankRow + 1
                            
                        End If
                        
                        Exit For    ' don't bother checking the rest of the stages to ignore (we're already ignoring it)
                        
                    End If
                
                Next K
                
                Exit For
            
            End If
            
            
            If projectStage = "" Then Exit For
            
            ' if we are here then the project stage is valid
            
            If InStr(1, validStage, projectStage) Then
                
                ' Assign project's database values to project variables
                
                projectNumber = Sheets("Database").Cells(projectRow, 2)
                If projectNumber = "" Then End ' must be at the end of the database
                
                projectName = Sheets("Database").Cells(projectRow, 3)
                projectStage = Sheets("Database").Cells(projectRow, 19)
                
                projectDirector = Sheets("Database").Cells(projectRow, 8)
                If projectDirector = "" Then projectDirector = "None Assigned"
                projectJobRunner = Sheets("Database").Cells(projectRow, 9)
                If projectJobRunner = "" Then projectJobRunner = "None Assigned"    ' needed for outputting errors
                projectElec = Sheets("Database").Cells(projectRow, 10)
                projectMech = Sheets("Database").Cells(projectRow, 12)
                
                projectStageNumber = J
                
                projectIsValid = True
                
                projectArea = Sheets("Database").Cells(projectRow, 5)
                projectStartDate = Sheets("Database").Cells(projectRow, 14)
                projectCompletionYear = Sheets("Database").Cells(projectRow, 29)
                projectCompletionDate = Sheets("Database").Cells(projectRow, 30)
                For i = 0 To 6
                    projectKeyDates(i) = Sheets("Database").Cells(projectRow, 22 + i)
                Next i
                projectAddress = LCase(Sheets("Database").Cells(projectRow, 31))
                projectDesc = LCase(Sheets("Database").Cells(projectRow, 33))
                projectType = LCase(Sheets("Database").Cells(projectRow, 34))
                projectNew = LCase(Sheets("Database").Cells(projectRow, 35))
                projectOccupancy = Sheets("Database").Cells(projectRow, 37)
                projectRisks = Sheets("Database").Cells(projectRow, 40)
                projectProfessions = Sheets("Database").Cells(projectRow, 43)
                projectDesRoll = Sheets("Database").Cells(projectRow, 54)
                
                ' Make sure the project type isn't PSDP
                If projectType = "PSDP" Then
                    projectIsValid = False
                End If
                
                Exit Sub ' get out as we have a valid project
            
            End If
            
        Next J
    
    End If ' end skipit
        
End Sub

Sub removePhraseErrors()

    ' There is a known error with Rule 4 that results in every phrase being found in a document even though they're not present.
    ' I can't seem to find why this is (something to do with the word object?)
    ' so for the time being, I'll just search for documents that have found every phrase searched for, and remove the errors associated with them.
    
    Dim dashboard, rule As Worksheet
    Set dashboard = Sheets("Dashboard")
    Set rule = Sheets("Rules 4")
    
    ' The first phrase error in the list of phrases searched for
    ' This will tell us when to start comparing
    initialPhraseError = rule.Cells(12, 5)
    
    For dashboardRow = 16 To 100000
        
        ' Exit if it's the last row
        If dashboard.Cells(dashboardRow, 1) = "" Then Exit For
        
        ' Grab the current error on the dashboard
        currentError = dashboard.Cells(dashboardRow, 4)
        
        ' Only search consecutive rows if it's the inital error
        If currentError = initialPhraseError Then
            
            ' Tracker boolean to determine whether all of the phrases have been found (and hence, if it's a result of the bug)
            allPhraseErrorsFound = False
            rowOffset = 0
            
            ' First phrase error has been found, check if ALL of the others follow it.
            For addedRows = 1 To 10000
                        
                ' The error doesn't occur when the rule is only applied to files that contain "x" (i.e. column 2 on Rules 4)
                For checkIndex = 0 To 1000
                
                    If rule.Cells(12 + addedRows + rowOffset, 2) = "" Then Exit For
                    rowOffset = rowOffset + 1
                    
                Next checkIndex
                
                nextPhraseError = rule.Cells(12 + addedRows + rowOffset, 5)
                
                ' If there's no more errors to check, they have all been found
                If nextPhraseError = "" Then
                
                    allPhraseErrorsFound = True
                    Exit For
                
                End If
                
                nextDashboardError = dashboard.Cells(dashboardRow + addedRows, 4)
                
                ' If the next dashboard error and next phrase error do not match, then these are genuine errors and not a result of the known bug
                ' So we should just exit the loop, as allPhrasesFound is already set to false
                If nextDashboardError <> nextPhraseError Then Exit For
                            
            Next addedRows
            
            If allPhraseErrorsFound = True Then
            
                ' Correct addedRows (it currently points to the next genuine error)
                addedRows = addedRows - 1
                
                ' Delete the correct number of rows.
                For deleteCount = 0 To addedRows
                
                    ' Delete the dashboard entry (we will always be deleting dashboardRow as all the rows below will shift up when one is deleted)
                    dashboard.Rows(dashboardRow).EntireRow.Delete
                
                Next deleteCount
                
                ' dashboardRow will be incremented at the end of this loop but we there's a new row in the current dashboard row so we need to reset it
                dashboardRow = dashboardRow - 1
            
            End If
        
        End If
        
    Next dashboardRow



End Sub


Sub forceRemoveExceptions()

    ' This is usually done from the audit macro
    ' But if we want to exclude exceptions we can use this function
        
    Call removeExceptions
    
    ' Recount the errors
    itemCount = 0
    For dashboardRow = 16 To 100000
    
        If Sheets("Dashboard").Cells(dashboardRow, 1) = "" Then
            Exit For
        Else
            itemCount = itemCount + 1
        End If
        
    Next dashboardRow
    
    ' Output the count
    Sheets("Dashboard").Cells(13, 3) = "Error Count:"
    Sheets("Dashboard").Cells(13, 4) = itemCount

End Sub


Private Sub removeExceptions()
    
    
    Dim dashboard, exceptions As Worksheet
    Set dashboard = Sheets("Dashboard")
    Set exceptions = Sheets("Exceptions")
    
    firstColumn = 1
    lastColumn = 7
    
    ' Cycle through each of the errors found on the dashboard
    For dashboardRow = 16 To 10000
    
        ' for troubleshooting (stop at a particular row)
        If dashboardRow = 87 Then
            test = "put a breakpoint here"
        End If
        
        ' Exit if no more
        If dashboard.Cells(dashboardRow, 1) = "" Then Exit For
    
        ' Check against each of the exceptions
        exceptionFound = False
        For exceptionRow = 16 To 10000
            
            ' Exit if no more
            If exceptions.Cells(exceptionRow, 2) = "" Then Exit For
        
            ' Only check rules that have been approved
            If exceptions.Cells(exceptionRow, 1) = 1 Then
            
                matchFound = True
                For currentColumn = firstColumn To lastColumn
                    
                    ' Make sure every column is the same, otherwise it's not a match
                    ' The offset by 2 is required for the Approved & Reason to ignore columns in Ayla's exceptions list (reason to ignore replaces Job runner column in Job runner's workbooks)
                    If currentColumn = 3 Then
                        ' Dont check the project runner column as the runner may have changed but the exception is still valid
                    ElseIf exceptions.Cells(exceptionRow, currentColumn + 2) <> dashboard.Cells(dashboardRow, currentColumn) Then
                        matchFound = False
                        Exit For
                    End If
            
                Next currentColumn
                
                ' If we found an exception, stop checking the exception list
                If matchFound Then
                    exceptionFound = True
                    Exit For
                End If
            
            End If
        
        Next exceptionRow
        
        If exceptionFound Then
        
            ' Delete the dashboard entry
            dashboard.Rows(dashboardRow).EntireRow.Delete
            
            ' Decrement the current row so we don't skip one
            dashboardRow = dashboardRow - 1
        
        End If
    
    Next dashboardRow
    
    Set dashboard = Nothing
    Set exceptions = Nothing

End Sub

Private Sub clearj()
    '
    ' clearj Macro
    '
    Sheets("J").Select
    
      Range("A3:G10000").Select
       Selection.ClearContents
            Sheets("j").Range("A3").Select
    Sheets("Dashboard").Select
    
End Sub
Private Sub cleardash()

    Sheets("Dashboard").Cells(13, 3) = ""
    Sheets("Dashboard").Cells(13, 4) = ""
    
    Sheets("Dashboard").Select
    
    Range("A16:G10000").Select
        Selection.ClearContents
        
    Sheets("Dashboard").Range("A16").Select

End Sub

