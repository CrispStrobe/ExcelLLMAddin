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

' Keys for the extra OpenAI-compatible cloud providers (groq, together, cerebras,
' gemini, cohere, huggingface, requesty). Kept in one cross-platform Dictionary
' (vendored, works on Mac + Windows) instead of a fixed variable per provider, and
' persisted as PKEY_<provider>= lines. Their base URLs are fixed literals in
' GetBaseURL, so no per-provider URL variable is needed.
Public ProviderKeys As Dictionary

Private Sub EnsureProviderKeys()
    If ProviderKeys Is Nothing Then Set ProviderKeys = New Dictionary
End Sub

' Public setter so the menu (modMenu) can store a cloud provider's key.
Public Sub SetProviderKey(ByVal provider As String, ByVal key As String)
    EnsureProviderKeys
    ProviderKeys(LCase(Trim(provider))) = key
End Sub

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
    
    If DEBUG_MODE Then Debug.Print "[Config] === LoadConfig START ==="
    
    ' Initialize defaults FIRST
    Call InitializeDefaults
    If DEBUG_MODE Then Debug.Print "[Config] Defaults set: OLLAMA_BASE_URL='" & OLLAMA_BASE_URL & "'"
    
    filePath = GetConfigLocation()
    If DEBUG_MODE Then Debug.Print "[Config] Config path: " & filePath
    
    ' Check if file exists
    If Dir(filePath) = "" Then
        If DEBUG_MODE Then Debug.Print "[Config] No config file, creating with defaults"
        Call SaveConfig
        Exit Sub
    End If
    
    If DEBUG_MODE Then Debug.Print "[Config] File exists, reading..."
    
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
                        If DEBUG_MODE Then Debug.Print "[Config] Line " & lineNum & ": " & key & " = '" & Left(value, 30) & "'"
                        
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
                                If DEBUG_MODE Then Debug.Print "[Config]   *** SET OLLAMA_BASE_URL = '" & value & "'"
                            Case "CurrentProvider": CurrentProvider = value
                            Case "CurrentModel": CurrentModel = value
                            Case Else
                                ' Extra cloud-provider keys: PKEY_<provider>=...
                                If Left(key, 5) = "PKEY_" Then
                                    EnsureProviderKeys
                                    ProviderKeys(LCase(Mid(key, 6))) = value
                                End If
                        End Select
                    Else
                        If DEBUG_MODE Then Debug.Print "[Config] Line " & lineNum & ": " & key & " = (empty, keeping default)"
                    End If
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    
    If DEBUG_MODE Then
        Debug.Print "[Config] === LoadConfig COMPLETE ==="
        Debug.Print "[Config] Final: Provider='" & CurrentProvider & "', Model='" & CurrentModel & "'"
        Debug.Print "[Config] Final: OLLAMA_BASE_URL='" & OLLAMA_BASE_URL & "'"
    End If
    
    Exit Sub
    
ErrorHandler:
    If DEBUG_MODE Then Debug.Print "[Config] *** ERROR: " & Err.Number & " - " & Err.Description
    If fileNum > 0 Then Close #fileNum
    Call InitializeDefaults
End Sub

' Save configuration to file
Public Sub SaveConfig()
    Dim filePath As String
    Dim fileNum As Integer
    
    filePath = GetConfigLocation()
    If DEBUG_MODE Then
        Debug.Print "[Config] === SaveConfig to: " & filePath
        Debug.Print "[Config] Saving: Provider='" & CurrentProvider & "', Model='" & CurrentModel & "'"
        Debug.Print "[Config] Saving: OLLAMA_BASE_URL='" & OLLAMA_BASE_URL & "'"
    End If
    
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

    ' Extra cloud-provider keys (groq, gemini, ...), one PKEY_ line each.
    EnsureProviderKeys
    Dim pk As Variant
    For Each pk In ProviderKeys.Keys
        Print #fileNum, "PKEY_" & pk & "=" & ProviderKeys(pk)
    Next pk

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
    
    If DEBUG_MODE Then Debug.Print "[Config] *** Saved successfully"
    Exit Sub
    
ErrorHandler:
    If DEBUG_MODE Then Debug.Print "[Config] *** SaveConfig ERROR: " & Err.Description
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error saving config: " & Err.Description, vbExclamation
End Sub

' Initialize default values
Public Sub InitializeDefaults()
    If DEBUG_MODE Then Debug.Print "[Config] --- InitializeDefaults ---"
    
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

    EnsureProviderKeys

    If DEBUG_MODE Then Debug.Print "[Config] Defaults complete"
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
        Case Else
            ' Extra cloud providers keep their key in ProviderKeys.
            EnsureProviderKeys
            If ProviderKeys.Exists(LCase(Trim(provider))) Then
                GetAPIKey = ProviderKeys(LCase(Trim(provider)))
            Else
                GetAPIKey = ""
            End If
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
        ' Extra OpenAI-compatible cloud providers (fixed base URLs).
        Case "groq": GetBaseURL = "https://api.groq.com/openai/v1"
        Case "together": GetBaseURL = "https://api.together.xyz/v1"
        Case "cerebras": GetBaseURL = "https://api.cerebras.ai/v1"
        Case "gemini": GetBaseURL = "https://generativelanguage.googleapis.com/v1beta/openai"
        Case "cohere": GetBaseURL = "https://api.cohere.ai/compatibility/v1"
        Case "huggingface": GetBaseURL = "https://router.huggingface.co/v1"
        Case "requesty": GetBaseURL = "https://router.requesty.ai/v1"
        Case Else: GetBaseURL = ""
    End Select
End Function
