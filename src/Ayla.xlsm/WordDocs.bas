Attribute VB_Name = "WordDocs"
Sub checkWordDocs()

    ' This sub searches through the J drive files for word documents.
    ' Each time a file is found, it's opened and if spell check isn't on, it's turned on and then saved
    
    Dim wordDoc As Word.Document
    
    ' Cross our fingers...
    On Error Resume Next
    
    ' Only create a word object once (change it within the loops)
    Set objWord = CreateObject("word.application")
    
    ' TODO: change back to false
    objWord.Visible = False  ' leave this on in case some dialog appears that needs input (even though we have set it not to show, sometimes it can still appear - if it's read only for eg)
    objWord.DisplayAlerts = False
            
    ' Cycle through each of the j drive files
    For J = 3 To 10000
    
        ' For Testing (stop at a particular J drive file)
        If J = 4855 Then
            test = "put a breakpoint here"
        End If
                                         
        ' Grab the current file properties and construct the full path
        currentFileName = Sheets("J").Cells(J, 1)
        currentFileType = Sheets("J").Cells(J, 5)
        currentFilePath = Sheets("J").Cells(J, 3)
        currentFileFullpath = currentFilePath & currentFileName & "." & currentFileType
        
        ' Don't bother continuing if we're at the end
        If currentFileName = "" Then Exit For
        
        ' Check if it's a word document (doc, docx)
        If InStr(1, LCase(currentFileType), LCase("doc")) Then
        
            ' Check if it's in an excluded path
            excludedFilePath = False
            For excludedStringRow = 11 To 1000
            
                excludedString = Sheets("Misc").Cells(excludedStringRow, 8)
                
                ' If it's empty don't continue
                If excludedString = "" Then Exit For
                
                If InStr(1, LCase(currentFileFullpath), LCase(excludedString)) Then
                    
                    excludedFilePath = True
                    Exit For
                
                End If
            
            Next excludedStringRow
            
            ' Check if the file needs to be checked
            nameContainsIncludedString = False
            For includedStringRow = 11 To 1000
            
                includedString = Sheets("Misc").Cells(includedStringRow, 9)
                
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
            
                ' Open the current document
                Set wordDoc = objWord.documents.Open(Filename:=currentFileFullpath, ReadOnly:=False)
                
                ' Only do the checks if it's not read only
                If wordDoc.ReadOnly = False Then
                                    
                    ' Make sure it's not empty (it will be if there's an error opening it)
                    If Not wordDoc Is Nothing Then
                    
                       changeMade = False
                    
                       ' Check if the language is correct
                       If wordDoc.Range.LanguageID <> wdEnglishUK Then
                       
                           wordDoc.Range.LanguageID = wdEnglishUK
                           changeMade = True
                       
                       End If
                       
                       ' Check if no proofing is true
                       If wordDoc.Range.NoProofing = True Then
                       
                           wordDoc.Range.NoProofing = False
                           changeMade = True
                       
                       End If
                       
                       ' Save it if we made a change
                       If changeMade = True Then
                       
                           wordDoc.DisplayAlerts = False
                           wordDoc.Save
                       
                       End If
                    
                    End If
                
                End If
                
                ' Close the current document
                wordDoc.Close savechanges:=False
            
            
            End If
        
        End If
                                           
    Next J

    ' Quit word
    objWord.Quit
    Set objWord = Nothing

End Sub
