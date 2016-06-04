Attribute VB_Name = "JDrive"
Sub readjdrive(SourceFolderName As String)

    ' read all the files on the j drive for the current project
    On Error Resume Next
    
    Dim FSO As New FileSystemObject, SourceFolder As Folder, Subfolder As Folder, FileItem As File
    
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
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
            
            ' Check if it's the correct type
            For fileTypeRow = 2 To 7
        
                includedFileType = Sheets("Stages").Cells(fileTypeRow, 3)
                
                If LCase(fileExtension) = LCase(includedFileType) Then
                    
                    includeFile = True
                    Exit For
                    
                End If
            
            Next fileTypeRow
        
        End If
        
        ' Only record the file if it is to be included
        If includeFile = True Then
        
            ' Find a blank row
            For i = 3 To 10000
                If Sheets("J").Cells(i, 1) = "" Then
                    Exit For
                End If
            Next i
                    
            ' Output the file info
            Sheets("J").Cells(i, 2) = FileItem.Type
            Sheets("J").Cells(i, 3) = FileItem.Path
            Sheets("J").Cells(i, 4) = FileItem.DateLastModified
            Sheets("J").Cells(i, 1) = trimmedName
            Sheets("J").Cells(i, 5) = fileExtension
            Sheets("J").Cells(i, 6) = FileItem.Size
            Sheets("J").Cells(i, 8) = FileItem.DateCreated  ' Column 8 because this was added at a later date and columns 7 is populated by the duplicate check macro
            
            ' trim the filepath down to just the path
            Sheets("J").Cells(i, 3) = Mid(Sheets("J").Cells(i, 3), 1, Len(Sheets("J").Cells(i, 3)) - Len(FileItem.Name))
            
        
        End If
    
        
    Next FileItem
    
    '  IncludeSubfolders Then
    For Each Subfolder In SourceFolder.SubFolders
        readjdrive Subfolder.Path
    Next Subfolder
    
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing

End Sub
