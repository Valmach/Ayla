Attribute VB_Name = "RecordPHB"
' There's a known bug that makes the macro stop when you open a workbook in excel
' Seemingly, you need to check if the shift button is pressed
' See link: https://support.microsoft.com/en-us/kb/555263
' That's what the below two blocks of code are for

'Declare API
Public Declare Function GetKeyState Lib "User32" _
(ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Public Function ShiftPressed() As Boolean
'Returns True if shift key is pressed
    ShiftPressed = GetKeyState(16) < 0
End Function


Sub testPhbRecord()

    ' Use this macro to set a project row and number for testing below
    ' Refer to Ayla's Database Sheet
    
    ' For testing purposes only
    Dim projectNumber As String
    Dim projectRowInDatabase As Integer
    
    ' Add project number to test and row where it occurs in database below
    projectNumber = "P5000652"
    projectRowInDatabase = 434
        
    Call recordPhbData(projectNumber, projectRowInDatabase)

End Sub

' Created as a separate sub so that it can be called from other Subs or exported to other versions of Ayla etc.

Sub recordPhbData(projectNumber As String, projectRowInDatabase As Integer)

    ' Where the workbook should be saved
    Dim SourceFolderName As String
    SourceFolderName = "J:\" & projectNumber & "\QA\Project handbook records"
    
    Dim targetWorkbook As String
    targetWorkbook = "PHB historic records"
    
    Dim fullWorkBookname As String
    fullWorkBookname = targetWorkbook & ".xlsx"
    
    ' cross our fingers
    On Error Resume Next
    
    Dim trimmedName, fileExtension As String
    trimmedName = ""
    fileExtension = ""
    
    ' Cycle through each file in the folder
    Dim FSO As New FileSystemObject, SourceFolder As Folder, Subfolder As Folder, FileItem As File
    
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    correctFileFound = False
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
        
        ' If it's the correct file don't scan any more
        If trimmedName = targetWorkbook And InStr(1, LCase(fileExtension), "xl") Then
            correctFileFound = True
            Exit For
        End If
                    
    Next FileItem
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    
    ' Create the file if it doesn't exist
    If correctFileFound = False Then
    
        ' Source file
        templateSourceFolder = "K:\M&E\QA\"
        ' Destination file
        destinationFolder = SourceFolderName & "\"
                
        Dim objFSO
        Set objFSO = CreateObject("scripting.filesystemobject")
        
        ' Check if destination folder exists.
        If objFSO.FolderExists(destinationFolder) = False Then
            
            ' If there are any errors they'll be ignored because we have On Error Resume Next above
            MkDir destinationFolder
            
        End If
        
        objFSO.CopyFile Source:=templateSourceFolder & fullWorkBookname, Destination:=destinationFolder
        
        Set objFSO = Nothing
    
    End If
    
    
    ' At this point the excel workbook is definately available at the usual path
    Dim actualPath As String
    actualPath = SourceFolderName & "\" & fullWorkBookname

    ' Open the workbook
    Dim aylaWb, historicWb As Workbook
        
    ' Must also have a reference to Ayla otherwise the opened workbook (historic) becomes the default one referred to
    Set aylaWb = ThisWorkbook
    
    ' There's a known bug that makes the macro stop when you open a workbook in excel
    ' Seemingly, you need to check if the shift button is pressed
    ' See link: https://support.microsoft.com/en-us/kb/555263
    Do While ShiftPressed()
        DoEvents
    Loop
    Set historicWb = Workbooks.Open(actualPath)
    
    aylaRecordSheet = "Database"
    historicRecordSheet = "Records"

    ' Find most recent entry in the records
    mostRecentRecordRow = 12    ' must initally be set to the first row
    For recordRow = mostRecentRecordRow To 10000
        
        ' Cycle through the records, when you reach an empty one then record the previous row as the most recent
        If historicWb.Sheets(historicRecordSheet).Cells(recordRow, 1) = "" Then
        
            ' make an execption for the first row
            If recordRow = mostRecentRecordRow Then Exit For
            
            ' otherwise
            mostRecentRecordRow = recordRow - 1
            Exit For
            
        End If
    
    Next recordRow
    
    ' The columns of the values to be compared (same for both Ayla and the PHB historic record)
    firstEntryColumn = 2
    lastEntryColumn = 54
    
    changeDetected = False
    For currentColumn = firstEntryColumn To lastEntryColumn
    
        aylaRecord = aylaWb.Sheets(aylaRecordSheet).Cells(projectRowInDatabase, currentColumn)
        historicRecord = historicWb.Sheets(historicRecordSheet).Cells(mostRecentRecordRow, currentColumn)
        
        ' Check for differences in any column
        If aylaRecord <> historicRecord Then
            changeDetected = True
            Exit For
        End If
    
    Next currentColumn
    
    ' If there's been a change, make a new record
    If changeDetected Then
        
        ' Add the date first
        historicWb.Sheets(historicRecordSheet).Cells(mostRecentRecordRow + 1, 1) = Now
        ' Set fill to green to associate with M&E (other colour used in Architect's version of Ayla)
        With historicWb.Sheets(historicRecordSheet).Cells(mostRecentRecordRow + 1, 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        ' Populate the records
        For currentColumn = firstEntryColumn To lastEntryColumn
    
            aylaRecord = aylaWb.Sheets(aylaRecordSheet).Cells(projectRowInDatabase, currentColumn)
            historicWb.Sheets(historicRecordSheet).Cells(mostRecentRecordRow + 1, currentColumn) = aylaRecord
            ' Set fill to green to associate with M&E (other colour used in Architect's version of Ayla)
            With historicWb.Sheets(historicRecordSheet).Cells(mostRecentRecordRow + 1, currentColumn).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        
        Next currentColumn
        
        ' Turn wrap text off (otherwise the rows expand becuase of the rtf values)
        historicWb.Sheets(historicRecordSheet).Rows(mostRecentRecordRow + 1).WrapText = False
        
        ' Only save it if you added a record
        historicWb.Save
    
    End If
    
    ' Close the records and remove from memory
    historicWb.Close
    
    Set historicWb = Nothing
    Set aylaWb = Nothing

End Sub



