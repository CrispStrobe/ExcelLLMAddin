' modConfig:

Option Explicit

' API Keys
Public OPENAI_API_KEY As String
Public MISTRAL_API_KEY As String
Public NEBIUS_API_KEY As String
Public SCALEWAY_API_KEY As String
Public OPENROUTER_API_KEY As String

' Provider URLs
Public OPENAI_URL As String
Public MISTRAL_URL As String
Public NEBIUS_URL As String
Public SCALEWAY_URL As String
Public OPENROUTER_URL As String
Public OLLAMA_BASE_URL As String

' Current settings
Public CurrentProvider As String
Public CurrentModel As String

' Get config file location
Public Function GetConfigLocation() As String
    Dim paths() As String
    Dim i As Integer
    Dim testPath As String
    
    ReDim paths(0 To 4)
    
    If Environ("HOME") <> "" Then
        paths(0) = Environ("HOME") & "/LLMExcelAddin_config.txt"
    End If
    
    If Environ("USERPROFILE") <> "" Then
        paths(1) = Environ("USERPROFILE") & "\LLMExcelAddin_config.txt"
    End If
    
    On Error Resume Next
    paths(2) = Application.DefaultFilePath & Application.PathSeparator & "LLMExcelAddin_config.txt"
    paths(3) = ThisWorkbook.Path & Application.PathSeparator & "LLMExcelAddin_config.txt"
    On Error GoTo 0
    
    paths(4) = Environ("TMPDIR") & "LLMExcelAddin_config.txt"
    If paths(4) = "LLMExcelAddin_config.txt" Then
        paths(4) = Environ("TEMP") & "\LLMExcelAddin_config.txt"
    End If
    
    ' Return first existing file
    For i = LBound(paths) To UBound(paths)
        testPath = paths(i)
        If testPath <> "" And testPath <> "\LLMExcelAddin_config.txt" And testPath <> "/LLMExcelAddin_config.txt" Then
            If Dir(testPath) <> "" Then
                GetConfigLocation = testPath
                Exit Function
            End If
        End If
    Next i
    
    ' Return first valid path for creation
    For i = LBound(paths) To UBound(paths)
        testPath = paths(i)
        If testPath <> "" And testPath <> "\LLMExcelAddin_config.txt" And testPath <> "/LLMExcelAddin_config.txt" Then
            GetConfigLocation = testPath
            Exit Function
        End If
    Next i
    
    GetConfigLocation = "LLMExcelAddin_config.txt"
End Function

' Load configuration from file
Public Sub LoadConfig()
    Dim filePath As String
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    Dim lineNum As Integer
    
    Debug.Print "[Config] === LoadConfig START ==="
    
    ' Initialize defaults FIRST
    Call InitializeDefaults
    Debug.Print "[Config] Defaults set: OLLAMA_BASE_URL='" & OLLAMA_BASE_URL & "'"
    
    filePath = GetConfigLocation()
    Debug.Print "[Config] Config path: " & filePath
    
    ' Check if file exists
    If Dir(filePath) = "" Then
        Debug.Print "[Config] No config file, creating with defaults"
        Call SaveConfig
        Exit Sub
    End If
    
    Debug.Print "[Config] File exists, reading..."
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    lineNum = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lineNum = lineNum + 1
        line = Trim(line)
        
        ' Skip comments and empty lines
        If line <> "" And Left(line, 1) <> "#" Then
            If InStr(line, "=") > 0 Then
                parts = Split(line, "=", 2)
                If UBound(parts) >= 1 Then
                    Dim key As String
                    Dim value As String
                    key = Trim(parts(0))
                    value = Trim(parts(1))
                    
                    ' CRITICAL FIX: Only set if value is NOT empty
                    If value <> "" Then
                        Debug.Print "[Config] Line " & lineNum & ": " & key & " = '" & Left(value, 30) & "'"
                        
                        Select Case key
                            Case "OPENAI_API_KEY": OPENAI_API_KEY = value
                            Case "MISTRAL_API_KEY": MISTRAL_API_KEY = value
                            Case "NEBIUS_API_KEY": NEBIUS_API_KEY = value
                            Case "SCALEWAY_API_KEY": SCALEWAY_API_KEY = value
                            Case "OPENROUTER_API_KEY": OPENROUTER_API_KEY = value
                            Case "OPENAI_URL": OPENAI_URL = value
                            Case "MISTRAL_URL": MISTRAL_URL = value
                            Case "NEBIUS_URL": NEBIUS_URL = value
                            Case "SCALEWAY_URL": SCALEWAY_URL = value
                            Case "OPENROUTER_URL": OPENROUTER_URL = value
                            Case "OLLAMA_BASE_URL"
                                OLLAMA_BASE_URL = value
                                Debug.Print "[Config]   *** SET OLLAMA_BASE_URL = '" & value & "'"
                            Case "CurrentProvider": CurrentProvider = value
                            Case "CurrentModel": CurrentModel = value
                        End Select
                    Else
                        Debug.Print "[Config] Line " & lineNum & ": " & key & " = (empty, keeping default)"
                    End If
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    
    Debug.Print "[Config] === LoadConfig COMPLETE ==="
    Debug.Print "[Config] Final: Provider='" & CurrentProvider & "', Model='" & CurrentModel & "'"
    Debug.Print "[Config] Final: OLLAMA_BASE_URL='" & OLLAMA_BASE_URL & "'"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Config] *** ERROR: " & Err.Number & " - " & Err.Description
    If fileNum > 0 Then Close #fileNum
    Call InitializeDefaults
End Sub

' Save configuration to file
Public Sub SaveConfig()
    Dim filePath As String
    Dim fileNum As Integer
    
    filePath = GetConfigLocation()
    Debug.Print "[Config] === SaveConfig to: " & filePath
    Debug.Print "[Config] Saving: Provider='" & CurrentProvider & "', Model='" & CurrentModel & "'"
    Debug.Print "[Config] Saving: OLLAMA_BASE_URL='" & OLLAMA_BASE_URL & "'"
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    Print #fileNum, "# LLM Excel Add-in Configuration"
    Print #fileNum, "# " & Now
    Print #fileNum, ""
    Print #fileNum, "# API Keys"
    Print #fileNum, "OPENAI_API_KEY=" & OPENAI_API_KEY
    Print #fileNum, "MISTRAL_API_KEY=" & MISTRAL_API_KEY
    Print #fileNum, "NEBIUS_API_KEY=" & NEBIUS_API_KEY
    Print #fileNum, "SCALEWAY_API_KEY=" & SCALEWAY_API_KEY
    Print #fileNum, "OPENROUTER_API_KEY=" & OPENROUTER_API_KEY
    Print #fileNum, ""
    Print #fileNum, "# Provider URLs"
    Print #fileNum, "OPENAI_URL=" & OPENAI_URL
    Print #fileNum, "MISTRAL_URL=" & MISTRAL_URL
    Print #fileNum, "NEBIUS_URL=" & NEBIUS_URL
    Print #fileNum, "SCALEWAY_URL=" & SCALEWAY_URL
    Print #fileNum, "OPENROUTER_URL=" & OPENROUTER_URL
    Print #fileNum, "OLLAMA_BASE_URL=" & OLLAMA_BASE_URL
    Print #fileNum, ""
    Print #fileNum, "# Current Defaults"
    Print #fileNum, "CurrentProvider=" & CurrentProvider
    Print #fileNum, "CurrentModel=" & CurrentModel
    
    Close #fileNum
    
    Debug.Print "[Config] *** Saved successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Config] *** SaveConfig ERROR: " & Err.Description
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error saving config: " & Err.Description, vbExclamation
End Sub

' Initialize default values
Public Sub InitializeDefaults()
    Debug.Print "[Config] --- InitializeDefaults ---"
    
    ' URLs - ALWAYS set
    OPENAI_URL = "https://api.openai.com/v1"
    MISTRAL_URL = "https://api.mistral.ai/v1"
    NEBIUS_URL = "https://api.studio.nebius.ai/v1"
    SCALEWAY_URL = "https://api.scaleway.ai/v1"
    OPENROUTER_URL = "https://openrouter.ai/api/v1"
    OLLAMA_BASE_URL = "http://localhost:11434"
    
    ' Provider/Model - only if empty
    If CurrentProvider = "" Then CurrentProvider = "ollama"
    If CurrentModel = "" Then CurrentModel = "ministral-3:3b-instruct-2512-q4_K_M"
    
    Debug.Print "[Config] Defaults complete"
End Sub

' Get API key for provider
Public Function GetAPIKey(provider As String) As String
    Select Case LCase(Trim(provider))
        Case "openai": GetAPIKey = OPENAI_API_KEY
        Case "mistral": GetAPIKey = MISTRAL_API_KEY
        Case "nebius": GetAPIKey = NEBIUS_API_KEY
        Case "scaleway": GetAPIKey = SCALEWAY_API_KEY
        Case "openrouter": GetAPIKey = OPENROUTER_API_KEY
        Case "ollama": GetAPIKey = ""
        Case Else: GetAPIKey = ""
    End Select
End Function

' Get base URL for provider
Public Function GetBaseURL(provider As String) As String
    ' Ensure defaults are loaded
    If OLLAMA_BASE_URL = "" Or OPENAI_URL = "" Then
        Call InitializeDefaults
    End If
    
    Select Case LCase(Trim(provider))
        Case "openai": GetBaseURL = OPENAI_URL
        Case "mistral": GetBaseURL = MISTRAL_URL
        Case "nebius": GetBaseURL = NEBIUS_URL
        Case "scaleway": GetBaseURL = SCALEWAY_URL
        Case "openrouter": GetBaseURL = OPENROUTER_URL
        Case "ollama": GetBaseURL = OLLAMA_BASE_URL
        Case Else: GetBaseURL = ""
    End Select
End Function



