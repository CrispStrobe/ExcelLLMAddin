' modLLMFunctions:
Option Explicit

' Detect platform
Private Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function

' Main PROMPT function
Public Function prompt(promptText As String, Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo ErrorHandler
    
    Application.Volatile
    
    ' Load config if needed
    If CurrentProvider = "" Or OLLAMA_BASE_URL = "" Then
        Call LoadConfig
    End If
    
    If provider = "" Then provider = CurrentProvider
    If model = "" Then model = CurrentModel
    
    provider = LCase(Trim(provider))
    
    Dim baseURL As String
    Dim apiKey As String
    Dim endpoint As String
    Dim jsonBody As String
    Dim response As String
    Dim systemMsg As String
    
    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)
    
    If baseURL = "" Then
        prompt = "Error: Invalid provider"
        Exit Function
    End If
    
    If apiKey = "" And provider <> "ollama" Then
        prompt = "Error: No API key for " & provider
        Exit Function
    End If
    
    ' System message to avoid emojis
    systemMsg = "You are a helpful assistant. Use only standard text characters. Avoid emojis and special unicode symbols."
    
    ' Build request
    Select Case provider
        Case "ollama"
            endpoint = baseURL & "/api/chat"
            jsonBody = "{""model"":""" & model & """,""messages"":[" & _
                      "{""role"":""system"",""content"":""" & EscapeJSON(systemMsg) & """}," & _
                      "{""role"":""user"",""content"":""" & EscapeJSON(promptText) & """}],""stream"":false}"
        Case Else
            endpoint = baseURL & "/chat/completions"
            jsonBody = "{""model"":""" & model & """,""messages"":[" & _
                      "{""role"":""system"",""content"":""" & EscapeJSON(systemMsg) & """}," & _
                      "{""role"":""user"",""content"":""" & EscapeJSON(promptText) & """}]}"
    End Select
    
    response = HTTPPost(endpoint, jsonBody, apiKey, provider)
    
    If Left(response, 6) = "Error:" Then
        prompt = response
        Exit Function
    End If
    
    prompt = ParseJSONContent(response, provider)
    prompt = FixEncoding(prompt)
    
    Exit Function
    
ErrorHandler:
    prompt = "Error: " & Err.Number & " - " & Err.Description
End Function

' List models
Public Function LIST_MODELS(Optional provider As String = "") As String
    On Error GoTo ErrorHandler
    
    If CurrentProvider = "" Then Call LoadConfig
    If provider = "" Then provider = CurrentProvider
    provider = LCase(Trim(provider))
    
    Dim baseURL As String
    Dim apiKey As String
    Dim endpoint As String
    Dim response As String
    Dim models As String
    
    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)
    
    If baseURL = "" Then
        LIST_MODELS = "Error: Invalid provider"
        Exit Function
    End If
    
    If apiKey = "" And provider <> "ollama" Then
        LIST_MODELS = "Error: No API key for " & provider
        Exit Function
    End If
    
    Select Case provider
        Case "ollama": endpoint = baseURL & "/api/tags"
        Case Else: endpoint = baseURL & "/models"
    End Select
    
    response = HTTPGet(endpoint, apiKey, provider)
    
    If Left(response, 6) = "Error:" Then
        LIST_MODELS = response
        Exit Function
    End If
    
    models = ParseModelList(response, provider)
    
    If Left(models, 6) <> "Error:" Then
        LIST_MODELS = Replace(models, ", ", vbCrLf)
    Else
        LIST_MODELS = models
    End If
    
    Exit Function
    
ErrorHandler:
    LIST_MODELS = "Error: " & Err.Description
End Function

' Config display
Public Function LLM_CONFIG() As String
    Application.Volatile
    If CurrentProvider = "" Then Call LoadConfig
    LLM_CONFIG = "Provider: " & CurrentProvider & " | Model: " & CurrentModel
End Function

' Test curl
Public Function TestCurl() As String
    Dim outputFile As String
    Dim curlCmd As String
    Dim result As String
    Dim waitCount As Integer
    
    If IsMac() Then
        outputFile = Environ("TMPDIR") & "test_curl.txt"
        curlCmd = "curl -s http://localhost:11434/api/version > '" & outputFile & "' 2>&1 &"
        MacScript "do shell script """ & curlCmd & """"
    Else
        outputFile = Environ("TEMP") & "\test_curl.txt"
        curlCmd = "curl -s http://localhost:11434/api/version > """ & outputFile & """ 2>&1"
        Shell "cmd /c " & curlCmd, vbHide
    End If
    
    waitCount = 0
    Do While Dir(outputFile) = "" And waitCount < 10
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    
    If Dir(outputFile) = "" Then
        TestCurl = "Error: No response"
        Exit Function
    End If
    
    Application.Wait Now + TimeValue("00:00:01")
    result = ReadFile(outputFile)
    
    On Error Resume Next
    Kill outputFile
    
    TestCurl = Trim(result)
End Function

' HTTP POST
Private Function HTTPPost(url As String, jsonBody As String, Optional apiKey As String = "", Optional provider As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim bodyFile As String
    Dim outputFile As String
    Dim curlCmd As String
    Dim result As String
    Dim waitCount As Integer
    Dim fileNum As Integer
    
    If IsMac() Then
        bodyFile = Environ("TMPDIR") & "llm_body_" & Format(Now, "hhnnss") & ".json"
        outputFile = Environ("TMPDIR") & "llm_out_" & Format(Now, "hhnnss") & ".txt"
    Else
        bodyFile = Environ("TEMP") & "\llm_body_" & Format(Now, "hhnnss") & ".json"
        outputFile = Environ("TEMP") & "\llm_out_" & Format(Now, "hhnnss") & ".txt"
    End If
    
    ' Write body
    fileNum = FreeFile
    Open bodyFile For Output As #fileNum
    Print #fileNum, jsonBody
    Close #fileNum
    
    ' Build command
    If IsMac() Then
        curlCmd = "curl -s -X POST '" & url & "' -H 'Content-Type: application/json'"
        If apiKey <> "" Then curlCmd = curlCmd & " -H 'Authorization: Bearer " & apiKey & "'"
        If LCase(provider) = "openrouter" Then curlCmd = curlCmd & " -H 'HTTP-Referer: https://excel-addin'"
        curlCmd = curlCmd & " --data-binary '@" & bodyFile & "' --max-time 60 -o '" & outputFile & "' 2>&1 &"
        MacScript "do shell script """ & curlCmd & """"
    Else
        curlCmd = "curl -s -X POST """ & url & """ -H ""Content-Type: application/json"""
        If apiKey <> "" Then curlCmd = curlCmd & " -H ""Authorization: Bearer " & apiKey & """"
        If LCase(provider) = "openrouter" Then curlCmd = curlCmd & " -H ""HTTP-Referer: https://excel-addin"""
        curlCmd = curlCmd & " --data-binary ""@" & bodyFile & """ --max-time 60 -o """ & outputFile & """"
        Shell "cmd /c " & curlCmd, vbHide
    End If
    
    ' Wait for response
    waitCount = 0
    Do While Dir(outputFile) = "" And waitCount < 60
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    
    If Dir(outputFile) = "" Then
        HTTPPost = "Error: Timeout"
        GoTo Cleanup
    End If
    
    Application.Wait Now + TimeValue("00:00:02")
    result = ReadFile(outputFile)
    
    If result = "" Then
        HTTPPost = "Error: Empty response"
    Else
        HTTPPost = result
    End If
    
Cleanup:
    On Error Resume Next
    If Dir(bodyFile) <> "" Then Kill bodyFile
    If Dir(outputFile) <> "" Then Kill outputFile
    Exit Function
    
ErrorHandler:
    HTTPPost = "Error: " & Err.Description
    GoTo Cleanup
End Function

' HTTP GET
Private Function HTTPGet(url As String, Optional apiKey As String = "", Optional provider As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim outputFile As String
    Dim curlCmd As String
    Dim result As String
    Dim waitCount As Integer
    
    If IsMac() Then
        outputFile = Environ("TMPDIR") & "llm_get_" & Format(Now, "hhnnss") & ".txt"
        curlCmd = "curl -s '" & url & "'"
        If apiKey <> "" Then curlCmd = curlCmd & " -H 'Authorization: Bearer " & apiKey & "'"
        If LCase(provider) = "openrouter" Then curlCmd = curlCmd & " -H 'HTTP-Referer: https://excel-addin'"
        curlCmd = curlCmd & " --max-time 30 -o '" & outputFile & "' 2>&1 &"
        MacScript "do shell script """ & curlCmd & """"
    Else
        outputFile = Environ("TEMP") & "\llm_get_" & Format(Now, "hhnnss") & ".txt"
        curlCmd = "curl -s """ & url & """"
        If apiKey <> "" Then curlCmd = curlCmd & " -H ""Authorization: Bearer " & apiKey & """"
        If LCase(provider) = "openrouter" Then curlCmd = curlCmd & " -H ""HTTP-Referer: https://excel-addin"""
        curlCmd = curlCmd & " --max-time 30 -o """ & outputFile & """"
        Shell "cmd /c " & curlCmd, vbHide
    End If
    
    waitCount = 0
    Do While Dir(outputFile) = "" And waitCount < 30
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    
    If Dir(outputFile) = "" Then
        HTTPGet = "Error: Timeout"
        Exit Function
    End If
    
    Application.Wait Now + TimeValue("00:00:01")
    result = ReadFile(outputFile)
    HTTPGet = result
    
    On Error Resume Next
    Kill outputFile
    Exit Function
    
ErrorHandler:
    HTTPGet = "Error: " & Err.Description
End Function

' Read file with encoding conversion
Private Function ReadFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim fileNum As Integer
    Dim line As String
    
    If IsMac() Then
        ' Try iconv conversion
        Dim convertedFile As String
        Dim convCmd As String
        
        convertedFile = filePath & ".conv"
        convCmd = "iconv -f UTF-8 -t MACROMAN '" & filePath & "' > '" & convertedFile & "' 2>/dev/null"
        
        On Error Resume Next
        MacScript "do shell script """ & convCmd & """"
        Application.Wait Now + TimeValue("00:00:01")
        On Error GoTo ErrorHandler
        
        If Dir(convertedFile) <> "" Then
            fileNum = FreeFile
            Open convertedFile For Input As #fileNum
            result = ""
            Do While Not EOF(fileNum)
                Line Input #fileNum, line
                result = result & line & vbCrLf
            Loop
            Close #fileNum
            
            On Error Resume Next
            Kill convertedFile
            On Error GoTo ErrorHandler
            
            ReadFile = Trim(result)
            Exit Function
        End If
    End If
    
    ' Fallback: normal read
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    result = ""
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        result = result & line & vbCrLf
    Loop
    Close #fileNum
    
    ReadFile = Trim(result)
    Exit Function
    
ErrorHandler:
    ' Ultimate fallback
    Dim fn As Integer, ln As String, rs As String
    fn = FreeFile
    Open filePath For Input As #fn
    rs = ""
    Do While Not EOF(fn)
        Line Input #fn, ln
        rs = rs & ln & vbCrLf
    Loop
    Close #fn
    ReadFile = Trim(rs)
End Function

' Fix encoding issues
Private Function FixEncoding(text As String) As String
    Dim result As String
    result = text
    
    ' German umlauts
    result = Replace(result, "√§", "ä")
    result = Replace(result, "√∂", "ö")
    result = Replace(result, "√º", "ü")
    result = Replace(result, "√Ñ", "Ä")
    result = Replace(result, "√Ö", "Ö")
    result = Replace(result, "√ú", "Ü")
    result = Replace(result, "√ü", "ß")
    
    ' French
    result = Replace(result, "√©", "é")
    result = Replace(result, "√®", "è")
    result = Replace(result, "√†", "à")
    
    ' Artifacts
    result = Replace(result, "‚Äô", "'")
    result = Replace(result, "üòä", "")
    
    FixEncoding = result
End Function

' Parse JSON content
Private Function ParseJSONContent(jsonText As String, Optional provider As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim startPos As Long, endPos As Long, content As String
    
    If LCase(Trim(provider)) = "ollama" Then
        startPos = InStr(jsonText, """message""")
        If startPos > 0 Then startPos = InStr(startPos, jsonText, """content"":""")
    End If
    
    If startPos = 0 Then
        startPos = InStr(jsonText, """content"":""")
        If startPos = 0 Then startPos = InStr(jsonText, """content"": """)
    End If
    
    If startPos > 0 Then
        startPos = InStr(startPos, jsonText, ":""") + 2
        endPos = startPos
        
        Do While endPos < Len(jsonText)
            endPos = InStr(endPos + 1, jsonText, """")
            If endPos = 0 Then Exit Do
            If Mid(jsonText, endPos - 1, 1) <> "\" Then Exit Do
        Loop
        
        If endPos > startPos Then
            content = Mid(jsonText, startPos, endPos - startPos)
            content = Replace(content, "\n", vbCrLf)
            content = Replace(content, "\r", vbCr)
            content = Replace(content, "\t", vbTab)
            content = Replace(content, "\""", """")
            content = Replace(content, "\\", "\")
            ParseJSONContent = content
        Else
            ParseJSONContent = "Error: Parse failed"
        End If
    Else
        ParseJSONContent = "Error: No content found"
    End If
    
    Exit Function
    
ErrorHandler:
    ParseJSONContent = "Error: " & Err.Description
End Function

' Parse model list
Private Function ParseModelList(jsonText As String, provider As String) As String
    On Error GoTo ErrorHandler
    
    Dim models As String, startPos As Long, endPos As Long
    Dim modelId As String, count As Integer
    
    models = ""
    startPos = 1
    count = 0
    provider = LCase(Trim(provider))
    
    Select Case provider
        Case "ollama"
            Do While True
                startPos = InStr(startPos, jsonText, """name"":""")
                If startPos = 0 Then Exit Do
                startPos = startPos + 8
                endPos = InStr(startPos, jsonText, """")
                If endPos = 0 Then Exit Do
                
                modelId = Mid(jsonText, startPos, endPos - startPos)
                If models <> "" Then models = models & ", "
                models = models & modelId
                count = count + 1
                startPos = endPos + 1
            Loop
        Case Else
            Do While True
                startPos = InStr(startPos, jsonText, """id"":""")
                If startPos = 0 Then Exit Do
                startPos = startPos + 6
                endPos = InStr(startPos, jsonText, """")
                If endPos = 0 Then Exit Do
                
                modelId = Mid(jsonText, startPos, endPos - startPos)
                If models <> "" Then models = models & ", "
                models = models & modelId
                count = count + 1
                startPos = endPos + 1
            Loop
    End Select
    
    If models = "" Then
        ParseModelList = "Error: No models found"
    Else
        ParseModelList = models
    End If
    
    Exit Function
    
ErrorHandler:
    ParseModelList = "Error: " & Err.Description
End Function

' Escape JSON
Private Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function

