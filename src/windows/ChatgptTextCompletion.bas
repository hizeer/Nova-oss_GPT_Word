Option Explicit

Function ReplaceIllegalCharacters(strIn As String, strChar As String) As String
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>{}[]" & Chr(10) & Chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next

    ReplaceIllegalCharacters = strIn
End Function

Function ToJson(ByVal dict As Object) As String
    Dim key As Variant, result As String, value As String
    
    result = "{"
    For Each key In dict.Keys
        result = result & IIf(Len(result) > 1, ",", "")
        If TypeName(dict(key)) = "Dictionary" Then
            value = SubToJson(dict(key))
            ToJson = value
            result = result & """" & key & """" & ": " & value
        Else
            value = dict(key)
            If IsNumeric(value) Then
                result = result & """" & key & """" & ": " & CInt(value)
            Else
                result = result & """" & key & """" & ": " & """" & value & """"
            End If
        End If
        
    Next key
    
    result = result & "}"
    
    ToJson = result
End Function

Function SubToJson(ByVal dict As Object) As String
    Dim key As Variant, result As String, value As String
    For Each key In dict.Keys
        If IsNumeric(key) Then
            value = SubToJson(dict(key))
            value = Left(value, Len(value) - 2)
            result = "[{" & value & "}]"
        Else
            value = dict(key)
            result = result & """" & key & """: " & """" & value & """" & ", "
        End If
        
    Next key
    
    SubToJson = result
    
End Function

Sub TextCompletion()
'
' Text Completion Macro - Version 0.1.4
'
'
  If Selection.Type = wdSelectionIP Then
    Exit Sub
  End If
  
  If Selection.Text = ChrW$(13) Then
    Exit Sub
  End If

  Dim strAPIKey As String
  Dim strURL As String
  Dim strPrompt As String
  Dim strModel As String
  Dim intMaxTokens As Integer
  Dim strResponse As String
  Dim objCurlHttp As Object
  Dim strJSONdata As String
  Dim dictData As New Scripting.Dictionary
  Dim dictMessages As New Scripting.Dictionary
  Dim intMaxToken As Integer
  Dim strSafePrompt As String
  
  strAPIKey = Environ("NOVA-OSS_API_KEY")
  strURL = "https://api.nova-oss.com/v1/chat/completions"
  strPrompt = Replace(Selection, ChrW$(13), "")
  strSafePrompt = ReplaceIllegalCharacters(strPrompt, "")

  strModel = "gpt-3.5-turbo"
  intMaxToken = 2048
  
  Set dictMessages(0) = New Scripting.Dictionary
  dictMessages(0).Add "role", "user"
  dictMessages(0).Add "content", strSafePrompt
  
  dictData.Add "model", strModel
  dictData.Add "messages", dictMessages
  dictData.Add "max_tokens", intMaxToken
  
  strJSONdata = ToJson(dictData)
  
  Set objCurlHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

  With objCurlHttp
    .Open "POST", strURL, False
    .SetRequestHeader "Content-type", "application/json"
    .SetRequestHeader "Authorization", "Bearer " + strAPIKey
    .Send strJSONdata
    
    Dim strStatus As Integer
    strStatus = .Status
    Dim strStatusText As String
    strStatusText = .StatusText
    
    If strStatus <> 200 Then
      MsgBox Prompt:="The Nova-oss servers have experienced an error while processing your request! Please try again shortly."
      Exit Sub
    End If

    strResponse = .ResponseText
    
    Dim strResponseLenght As Integer
    strResponseLenght = Len(strResponse)
    
    Dim Output As Object
    Set Output = ParseJSON(strResponse)
    
    Dim strOutput As String
    strOutput = Output("obj.choices(0).message.content")
    
    Dim strOutputFormatted As String, strOutputFormatted1 As String, strOutputFormatted2 As String
    strOutputFormatted1 = Replace(strOutput, "\n\n", vbCrLf)
    strOutputFormatted2 = Replace(strOutputFormatted1, "\n", vbCrLf)
    strOutputFormatted = strOutputFormatted2
    
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertAfter vbCr & strOutputFormatted
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 12
    Selection.Font.ColorIndex = wdViolet
    Selection.Paragraphs.Alignment = wdAlignParagraphJustify
    Selection.InsertAfter vbCr
    Selection.Collapse Direction:=wdCollapseEnd
    
  End With
  
  Set objCurlHttp = Nothing

End Sub
