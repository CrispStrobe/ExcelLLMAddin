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

' Main PROMPT function - WITH FULL DEBUG
Public Function prompt(promptText As String, Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo ErrorHandler
    
    Application.Volatile
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "[PROMPT] === CALLED ==="
    Debug.Print "========================================="
    
    ' Load config if needed
    If CurrentProvider = "" Or OLLAMA_BASE_URL = "" Then
        Debug.Print "[PROMPT] Loading config..."
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
    
    Debug.Print "[PROMPT] Provider: '" & provider & "'"
    Debug.Print "[PROMPT] Model: '" & model & "'"
    Debug.Print "[PROMPT] BaseURL: '" & baseURL & "'"
    Debug.Print "[PROMPT] Prompt: '" & Left(promptText, 100) & "'"

    If model = "" Then
        ' Try to auto-detect a valid model
        Dim availableModels As String
        availableModels = LIST_MODELS(provider)
        If InStr(availableModels, CurrentModel) > 0 Then
            model = CurrentModel
        ElseIf availableModels <> "" Then
            ' Use first available model
            model = Split(Split(availableModels, vbCrLf)(0), "|")(0)
            Debug.Print "[PROMPT] Auto-selected model: " & model
        Else
            prompt = "Error: No valid model available for " & provider
            Exit Function
        End If
    End If
    
    If baseURL = "" Then
        prompt = "Error: Invalid provider"
        Debug.Print "[PROMPT] *** " & prompt
        Exit Function
    End If
    
    If apiKey = "" And provider <> "ollama" Then
        prompt = "Error: No API key for " & provider
        Debug.Print "[PROMPT] *** " & prompt
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
    
    Debug.Print "[PROMPT] Endpoint: " & endpoint
    Debug.Print "[PROMPT] Request body length: " & Len(jsonBody)
    Debug.Print "[PROMPT] Calling HTTPPost..."
    
    response = HTTPPost(endpoint, jsonBody, apiKey, provider)
    
    Debug.Print "[PROMPT] HTTPPost returned, length: " & Len(response)
    Debug.Print "[PROMPT] Response preview: " & Left(response, 200)
    
    If Left(response, 6) = "Error:" Then
        prompt = response
        Debug.Print "[PROMPT] *** HTTP Error: " & response
        Exit Function
    End If
    
    Debug.Print "[PROMPT] Calling ParseJSONContent..."
    prompt = ParseJSONContent(response, provider)
    
    Debug.Print "[PROMPT] ParseJSONContent returned: " & Left(prompt, 100)
    
    If Left(prompt, 6) <> "Error:" Then
        Debug.Print "[PROMPT] Calling FixEncoding..."
        prompt = FixEncoding(prompt)
        Debug.Print "[PROMPT] After FixEncoding: " & Left(prompt, 100)
    End If
    
    Debug.Print "[PROMPT] *** FINAL RESULT: " & Left(prompt, 100)
    Debug.Print "========================================="
    Debug.Print ""
    
    Exit Function
    
ErrorHandler:
    prompt = "Error: " & Err.Number & " - " & Err.Description
    Debug.Print "[PROMPT] *** EXCEPTION: " & prompt
    Debug.Print "========================================="
End Function

' List models
Public Function LIST_MODELS(Optional provider As String = "") As String
    On Error GoTo ErrorHandler

    ' Ensure configuration is loaded
    If CurrentProvider = "" Then Call LoadConfig
    If provider = "" Then provider = CurrentProvider
    provider = LCase(Trim(provider))

    Dim baseURL As String, apiKey As String, endpoint As String
    Dim response As String, models As String

    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)

    ' Validation checks
    If baseURL = "" Then
        LIST_MODELS = "Error: Invalid provider"
        Exit Function
    End If

    If apiKey = "" And provider <> "ollama" Then
        LIST_MODELS = "Error: No API key for " & provider
        Exit Function
    End If

    ' Determine endpoint
    Select Case provider
        Case "ollama": endpoint = baseURL & "/api/tags"
        Case Else: endpoint = baseURL & "/models"
    End Select

    ' Fetch data
    response = HTTPGet(endpoint, apiKey, provider)

    If Left(response, 6) = "Error:" Then
        LIST_MODELS = response
        Exit Function
    End If

    ' CENTRALIZED FIX: Delegate all parsing to ParseModelList
    models = ParseModelList(response, provider)

    ' Final display formatting for Excel
    If Left(models, 6) <> "Error:" Then
        ' Convert internal pipes to newlines for the cell
        LIST_MODELS = Replace(models, "|", vbCrLf)
    Else
        LIST_MODELS = models
    End If

    Exit Function

ErrorHandler:
    LIST_MODELS = "Error: " & Err.Description
End Function

' Config display
Public Function LLM_CONFIG() As String
    On Error Resume Next
    Application.Volatile
    
    ' Force config reload if global variables are lost
    If CurrentProvider = "" Or CurrentModel = "" Then 
        Call LoadConfig 
    End If
    
    ' If still empty after loading, provide a helpful status message
    If CurrentProvider = "" Then
        LLM_CONFIG = "Status: Not Configured (Run ShowSettings)"
    Else
        LLM_CONFIG = "Provider: " & CurrentProvider & " | Model: " & CurrentModel 
    End If
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
' HTTP POST - FIXED VERSION
Private Function HTTPPost(url As String, jsonBody As String, Optional apiKey As String = "", Optional provider As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim bodyFile As String
    Dim outputFile As String
    Dim curlCmd As String
    Dim result As String
    Dim waitCount As Integer
    Dim fileNum As Integer
    Dim lastSize As Long
    Dim currentSize As Long
    Dim stableCount As Integer
    
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
        curlCmd = curlCmd & " --data-binary '@" & bodyFile & "' --max-time 60 -o '" & outputFile & "' 2>&1"
        MacScript "do shell script """ & curlCmd & """"
    Else
        curlCmd = "curl -s -X POST """ & url & """ -H ""Content-Type: application/json"""
        If apiKey <> "" Then curlCmd = curlCmd & " -H ""Authorization: Bearer " & apiKey & """"
        If LCase(provider) = "openrouter" Then curlCmd = curlCmd & " -H ""HTTP-Referer: https://excel-addin"""
        curlCmd = curlCmd & " --data-binary ""@" & bodyFile & """ --max-time 60 -o """ & outputFile & """"
        Shell "cmd /c " & curlCmd, vbHide
    End If
    
    ' Wait for file to be created
    waitCount = 0
    Do While Dir(outputFile) = "" And waitCount < 60
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    
    If Dir(outputFile) = "" Then
        HTTPPost = "Error: Timeout - no response file"
        GoTo Cleanup
    End If
    
    ' NEW: Wait for file size to stabilize (curl finished writing)
    lastSize = -1
    stableCount = 0
    waitCount = 0
    Do While stableCount < 3 And waitCount < 30
        On Error Resume Next
        currentSize = FileLen(outputFile)
        On Error GoTo ErrorHandler
        
        If currentSize = lastSize And currentSize > 0 Then
            stableCount = stableCount + 1
        Else
            stableCount = 0
            lastSize = currentSize
        End If
        
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    
    ' Additional safety wait
    Application.Wait Now + TimeValue("00:00:01")
    
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
' Read file - SIMPLIFIED VERSION
Private Function ReadFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    ' Read the entire file at once
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
    Close #fileNum
    
    ReadFile = fileContent
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ReadFile = ""
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

' Parse JSON content - WITH FULL DEBUG
' Parse JSON content - UPDATED VERSION
Private Function ParseJSONContent(jsonText As String, Optional provider As String = "") As String
    On Error GoTo ErrorHandler

    Debug.Print "[Parse] ========================================="
    Debug.Print "[Parse] Provider: '" & provider & "'"
    Debug.Print "[Parse] JSON Length: " & Len(jsonText)
    Debug.Print "[Parse] JSON Preview (first 300 chars): " & Left(jsonText, 300)

    ' First check for error response
    If InStr(jsonText, """error"":") > 0 Then
        Dim errorStart As Long, errorEnd As Long
        errorStart = InStr(jsonText, """error"":""") + 9
        If errorStart > 9 Then
            errorEnd = InStr(errorStart, jsonText, """")
            If errorEnd > errorStart Then
                ParseJSONContent = "Error: " & Mid(jsonText, errorStart, errorEnd - errorStart)
                Debug.Print "[Parse] *** Detected error response: " & ParseJSONContent
                Debug.Print "[Parse] ========================================="
                Exit Function
            End If
        End If
    End If

    Dim startPos As Long, endPos As Long, content As String

    ' Try Ollama format first
    If LCase(Trim(provider)) = "ollama" Then
        Debug.Print "[Parse] Looking for Ollama format..."
        startPos = InStr(jsonText, """message""")
        Debug.Print "[Parse] Found 'message' at position: " & startPos

        If startPos > 0 Then
            startPos = InStr(startPos, jsonText, """content"":""")
            Debug.Print "[Parse] Found 'content' at position: " & startPos
        End If
    End If

    ' Try standard format
    If startPos = 0 Then
        Debug.Print "[Parse] Looking for standard format..."
        startPos = InStr(jsonText, """content"":""")
        If startPos = 0 Then
            startPos = InStr(jsonText, """content"": """)
        End If
        Debug.Print "[Parse] Found 'content' at position: " & startPos
    End If

    If startPos > 0 Then
        startPos = InStr(startPos, jsonText, ":""") + 2
        Debug.Print "[Parse] Content starts at position: " & startPos
        endPos = startPos

        ' Find closing quote
        Do While endPos < Len(jsonText)
            endPos = InStr(endPos + 1, jsonText, """")
            If endPos = 0 Then
                Debug.Print "[Parse] *** No closing quote found!"
                Exit Do
            End If
            If Mid(jsonText, endPos - 1, 1) <> "\" Then
                Debug.Print "[Parse] Found closing quote at position: " & endPos
                Exit Do
            End If
        Loop

        If endPos > startPos Then
            content = Mid(jsonText, startPos, endPos - startPos)
            Debug.Print "[Parse] Extracted content length: " & Len(content)
            Debug.Print "[Parse] Raw content: " & Left(content, 150)

            ' Unescape JSON
            content = Replace(content, "\n", vbCrLf)
            content = Replace(content, "\r", vbCr)
            content = Replace(content, "\t", vbTab)
            content = Replace(content, "\""", """")
            content = Replace(content, "\\", "\")

            ParseJSONContent = content
            Debug.Print "[Parse] *** SUCCESS! Final content: " & Left(content, 100)
        Else
            ParseJSONContent = "Error: Parse failed - no closing quote (endPos=" & endPos & ", startPos=" & startPos & ")"
            Debug.Print "[Parse] *** " & ParseJSONContent
        End If
    Else
        ' If no content found, return the raw JSON for debugging
        ParseJSONContent = "Error: No content found in JSON. Raw response: " & jsonText
        Debug.Print "[Parse] *** " & ParseJSONContent
    End If

    Debug.Print "[Parse] ========================================="

    Exit Function

ErrorHandler:
    ParseJSONContent = "Error: " & Err.Number & " - " & Err.Description
    Debug.Print "[Parse] *** EXCEPTION: " & ParseJSONContent
End Function

' Parse model list
Private Function ParseModelList(jsonText As String, provider As String) As String
    On Error GoTo ErrorHandler

    Dim models As String, startPos As Long, endPos As Long
    Dim modelId As String, count As Integer
    Dim i As Long, jsonLen As Long
    Dim inQuotes As Boolean, prevChar As String

    models = ""
    count = 0
    provider = LCase(Trim(provider))
    jsonLen = Len(jsonText)

    Select Case provider
        Case "ollama"
            ' Standard Ollama /api/tags parsing
            startPos = 1
            Do While True
                startPos = InStr(startPos, jsonText, """name"":""")
                If startPos = 0 Then Exit Do
                startPos = startPos + 8
                endPos = InStr(startPos, jsonText, """")
                If endPos = 0 Then Exit Do

                modelId = Mid(jsonText, startPos, endPos - startPos)
                
                ' Fix: Multi-line If block for syntax compatibility
                If models <> "" Then
                    models = models & "|"
                End If
                models = models & modelId

                startPos = endPos + 1
            Loop

            ' Alternative parsing fallback
            If models = "" Then
                startPos = 1
                inQuotes = False
                modelId = ""
                For i = 1 To jsonLen
                    prevChar = Mid(jsonText, i, 1)
                    If prevChar = """" Then
                        If inQuotes Then
                            If modelId <> "" And (InStr(modelId, ":") > 0 Or InStr(modelId, "-") > 0) Then
                                ' Fix: Standardize delimiter block
                                If models <> "" Then
                                    models = models & "|"
                                End If
                                models = models & modelId
                            End If
                            modelId = ""
                        End If
                        inQuotes = Not inQuotes
                    ElseIf inQuotes Then
                        modelId = modelId & prevChar
                    End If
                Next i
            End If

        Case Else
            ' Standard API (OpenAI, Mistral, etc.)
            startPos = 1
            Do While True
                startPos = InStr(startPos, jsonText, """id"":""")
                If startPos = 0 Then Exit Do
                startPos = startPos + 6
                endPos = InStr(startPos, jsonText, """")
                If endPos = 0 Then Exit Do

                modelId = Mid(jsonText, startPos, endPos - startPos)
                
                ' Fix: Explicit block for standard API delimiters
                If models <> "" Then
                    models = models & "|"
                End If
                models = models & modelId
                
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

