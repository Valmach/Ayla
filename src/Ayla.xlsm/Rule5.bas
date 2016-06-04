Attribute VB_Name = "Rule5"
Sub rulefive()

    ' THIS RULE IS NOT COMPLETE
    
    ' This rule will find and replace phrases in word documents
    ' At the moment, i'm having trouble replacing with superscript in part of the word (i.e. changing m2 to m^2).

    ' Cross our fingers...
    On Error Resume Next
    
    ' Only create a word object once (change it within the loops)
    Set objWord = CreateObject("word.application")
    
    objWord.Visible = False
    
    
    
    testFilePath = "K:\M&E\Calculations\APPLICATIONS\In Development\test find and replace.docx"
    
    ' Open the current document
    Set wordDoc = objWord.documents.Open(Filename:=currentFileFullpath, ReadOnly:=False)
    
    ' Make sure it's not empty (it will be if there's an error opening it)
    If Not wordDoc Is Nothing Then
    
        With wordDoc.Content.Find
            Do While .Execute(findText:=rulePhrase, Forward:=True, Format:=True, MatchWholeWord:=False, MatchCase:=False, Wrap:=wdFindStop) = True
            
                
                
                

            Loop
        End With
    
    
    End If
                
                

End Sub
