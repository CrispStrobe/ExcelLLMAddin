' modMenu:

Option Explicit

' DEBUG MODE - Set to False for production, True for debugging
Public Const DEBUG_MODE As Boolean = False

' Auto_Open
Public Sub Auto_Open()
    If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "========================================="
        Debug.Print "=== AUTO_OPEN STARTING ==="
        Debug.Print "=== Time: " & Now
        Debug.Print "========================================="
        
        Debug.Print "[AUTO_OPEN] Before LoadConfig:"
        Debug.Print "[AUTO_OPEN]   CurrentProvider = '" & CurrentProvider & "'"
        Debug.Print "[AUTO_OPEN]   CurrentModel = '" & CurrentModel & "'"
        Debug.Print "[AUTO_OPEN]   OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
    End If
    
    Call LoadConfig
    
    If DEBUG_MODE Then
        Debug.Print "[AUTO_OPEN] After LoadConfig:"
        Debug.Print "[AUTO_OPEN]   CurrentProvider = '" & CurrentProvider & "'"
        Debug.Print "[AUTO_OPEN]   CurrentModel = '" & CurrentModel & "'"
        Debug.Print "[AUTO_OPEN]   OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
        Debug.Print "[AUTO_OPEN]   OPENAI_URL = '" & OPENAI_URL & "'"
        
        Debug.Print "[AUTO_OPEN] Calling SetupKeyboardShortcuts..."
    End If
    
    Call SetupKeyboardShortcuts
    
    ' Show welcome message only in DEBUG_MODE
    If DEBUG_MODE Then
        Dim msg As String
        msg = "LLM Excel Add-in loaded!" & vbCrLf & vbCrLf & _
              "Current: " & CurrentProvider & " / " & CurrentModel & vbCrLf & vbCrLf & _
              "Run from Tools > Macro > Macros:" & vbCrLf & _
              "- ShowSettings - Configure providers" & vbCrLf & _
              "- QuickTest - Test connection" & vbCrLf & _
              "- FullDiagnostic - Show diagnostics" & vbCrLf & vbCrLf & _
              "(Keyboard shortcuts may not work on Mac)"
        
        msg = msg & vbCrLf & vbCrLf & "DEBUG MODE: ON" & vbCrLf & _
              "Press Cmd+G in VBA to see Immediate Window"
        
        MsgBox msg, vbInformation, "LLM Add-in"
        
        Debug.Print "=== AUTO_OPEN COMPLETE ==="
        Debug.Print "========================================="
        Debug.Print ""
    End If
End Sub

' Auto_Close
Public Sub Auto_Close()
    If DEBUG_MODE Then Debug.Print "[AUTO_CLOSE] Called"
    Call RemoveKeyboardShortcuts
End Sub

' Show model selector
Public Sub ShowModelSelector()
    Dim models As String
    Dim modelArray() As String
    Dim i As Integer
    Dim selectedModel As String

    models = LIST_MODELS()
    If Left(models, 6) = "Error:" Then
        MsgBox "Cannot list models: " & models, vbExclamation
        Exit Sub
    End If

    modelArray = Split(models, vbCrLf)
    selectedModel = InputBox("Available Models:" & vbCrLf & vbCrLf & _
                            Join(modelArray, ", ") & vbCrLf & vbCrLf & _
                            "Enter model name to use:", _
                            "Model Selector", _
                            CurrentModel)

    If selectedModel <> "" Then
        CurrentModel = selectedModel
        Call SaveConfig
        MsgBox "Model set to: " & CurrentModel, vbInformation
    End If
End Sub

' Setup keyboard shortcuts
Private Sub SetupKeyboardShortcuts()
    On Error Resume Next
    
    If DEBUG_MODE Then Debug.Print "[SHORTCUTS] Setting up keyboard shortcuts..."
    
    Application.OnKey "^+L", "ShowSettings"
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   Warning: Could not set ^+L: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   OK ^+L -> ShowSettings"
    End If
    
    Application.OnKey "^+M", "ShowModelSelector"
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   Warning: Could not set ^+M: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   OK ^+M -> ShowModelSelector"
    End If
    
    Application.OnKey "^+T", "QuickTest"
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   Warning: Could not set ^+T: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   OK ^+T -> QuickTest"
    End If
    
    Application.OnKey "^+D", "FullDiagnostic"
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   Warning: Could not set ^+D: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print "[SHORTCUTS]   OK ^+D -> FullDiagnostic"
    End If
    
    On Error GoTo 0
    If DEBUG_MODE Then Debug.Print "[SHORTCUTS] Setup complete"
End Sub

' Remove keyboard shortcuts
Private Sub RemoveKeyboardShortcuts()
    On Error Resume Next
    Application.OnKey "^+L"
    Application.OnKey "^+M"
    Application.OnKey "^+T"
    Application.OnKey "^+D"
    On Error GoTo 0
End Sub

' Show Settings - USING InputBox (not Application.InputBox)
Public Sub ShowSettings()
    If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "========================================="
        Debug.Print "=== ShowSettings START ==="
        Debug.Print "=== Time: " & Now
        Debug.Print "========================================="
    End If
    
    On Error GoTo ErrorHandler
    
    ' Ensure config is loaded
    If CurrentProvider = "" Or OLLAMA_BASE_URL = "" Then
        If DEBUG_MODE Then Debug.Print "[ShowSettings] Config appears empty, calling LoadConfig..."
        Call LoadConfig
    End If
    
    Dim choice As String
    Dim providers As String
    
    providers = "1=OpenAI, 2=Mistral, 3=Nebius, 4=Scaleway, 5=OpenRouter, 6=Ollama"
    
    ' Use regular InputBox (works on Mac)
    choice = InputBox( _
        "LLM Add-in Settings" & vbCrLf & vbCrLf & _
        "Current Provider: " & CurrentProvider & vbCrLf & _
        "Current Model: " & CurrentModel & vbCrLf & vbCrLf & _
        providers & vbCrLf & vbCrLf & _
        "Choose action:" & vbCrLf & _
        "1-6: Configure provider" & vbCrLf & _
        "7: Set default model" & vbCrLf & _
        "8: Show current config" & vbCrLf & _
        "9: Test connection" & vbCrLf & _
        "D: Full diagnostics" & vbCrLf & _
        "0: Exit", _
        "Settings Menu", _
        "6")
    
    ' Check if user cancelled (empty string)
    If choice = "" Then
        If DEBUG_MODE Then Debug.Print "[ShowSettings] User cancelled"
        Exit Sub
    End If
    
    ' Convert to uppercase and trim
    choice = UCase(Trim(choice))
    If DEBUG_MODE Then Debug.Print "[ShowSettings] User choice: '" & choice & "'"
    
    ' Handle the choice
    Select Case choice
        Case "1"
            Call ConfigureProvider("openai", "OpenAI")
            Call ShowSettings
            
        Case "2"
            Call ConfigureProvider("mistral", "Mistral")
            Call ShowSettings
            
        Case "3"
            Call ConfigureProvider("nebius", "Nebius")
            Call ShowSettings
            
        Case "4"
            Call ConfigureProvider("scaleway", "Scaleway")
            Call ShowSettings
            
        Case "5"
            Call ConfigureProvider("openrouter", "OpenRouter")
            Call ShowSettings
            
        Case "6"
            Call ConfigureOllama
            Call ShowSettings
            
        Case "7"
            Call SetDefaultModel
            Call ShowSettings
            
        Case "8"
            Call ShowCurrentConfig
            Call ShowSettings
            
        Case "9"
            Call QuickTest
            Call ShowSettings
            
        Case "D"
            Call FullDiagnostic
            Call ShowSettings
            
        Case "0"
            ' Exit
            Exit Sub
            
        Case Else
            MsgBox "Invalid choice: " & choice, vbExclamation
            Call ShowSettings
    End Select
    
    Exit Sub
    
ErrorHandler:
    If DEBUG_MODE Then Debug.Print "[ShowSettings] ERROR: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Configure Ollama
Private Sub ConfigureOllama()
    On Error GoTo ErrorHandler
    
    Dim ollamaURL As String
    Dim response As String
    Dim suggestedModel As String
    
    ollamaURL = OLLAMA_BASE_URL
    If ollamaURL = "" Then ollamaURL = "http://localhost:11434"
    
    ' Get Ollama URL using regular InputBox
    response = InputBox( _
        "Enter Ollama server URL:" & vbCrLf & vbCrLf & _
        "Default: http://localhost:11434" & vbCrLf & vbCrLf & _
        "Press OK to use default, or enter custom URL:", _
        "Ollama Configuration", _
        ollamaURL)
    
    If response = "" Then
        ollamaURL = "http://localhost:11434"
    Else
        ollamaURL = response
    End If
    
    OLLAMA_BASE_URL = ollamaURL
    
    ' Suggest a model
    suggestedModel = "ministral-3:3b-instruct-2512-q4_K_M"
    
    response = InputBox( _
        "Enter Ollama model name:" & vbCrLf & vbCrLf & _
        "Suggested: " & suggestedModel & vbCrLf & vbCrLf & _
        "Enter model name or press OK for suggested:", _
        "Ollama Model", _
        suggestedModel)
    
    If response <> "" Then suggestedModel = response
    
    ' Ask if this should be default
    Dim msgResponse As Integer
    msgResponse = MsgBox( _
        "Set Ollama as default provider?" & vbCrLf & vbCrLf & _
        "Model: " & suggestedModel, _
        vbYesNo + vbQuestion, _
        "Set Default?")
    
    If msgResponse = vbYes Then
        CurrentProvider = "ollama"
        CurrentModel = suggestedModel
    End If
    
    Call SaveConfig
    MsgBox "Ollama configured successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    If DEBUG_MODE Then Debug.Print "[ConfigureOllama] ERROR: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Configure provider (OpenAI, Mistral, etc.) - USING InputBox
Private Sub ConfigureProvider(providerKey As String, providerName As String)
    On Error GoTo ErrorHandler
    
    Dim apiKey As String
    Dim apiURL As String
    Dim currentKey As String
    Dim defaultURL As String
    Dim suggestedModel As String
    Dim response As String
    
    ' Get current values
    currentKey = GetAPIKey(providerKey)
    defaultURL = GetBaseURL(providerKey)
    
    ' Set default URL if empty
    Select Case providerKey
        Case "openai"
            If defaultURL = "" Then defaultURL = "https://api.openai.com/v1"
            suggestedModel = "gpt-4o-mini"
        Case "mistral"
            If defaultURL = "" Then defaultURL = "https://api.mistral.ai/v1"
            suggestedModel = "mistral-small-latest"
        Case "nebius"
            If defaultURL = "" Then defaultURL = "https://api.studio.nebius.ai/v1"
            suggestedModel = "meta-llama/Llama-3.3-70B-Instruct"
        Case "scaleway"
            If defaultURL = "" Then defaultURL = "https://api.scaleway.ai/v1"
            suggestedModel = "llama-3.3-70b-instruct"
        Case "openrouter"
            If defaultURL = "" Then defaultURL = "https://openrouter.ai/api/v1"
            suggestedModel = "meta-llama/llama-3.3-70b-instruct"
    End Select
    
    ' Get API key using regular InputBox
    response = InputBox( _
        "Enter " & providerName & " API Key:" & vbCrLf & vbCrLf & _
        "Current: " & IIf(currentKey = "", "(not set)", "***" & Right(currentKey, 4)) & vbCrLf & vbCrLf & _
        "Leave blank to skip, or enter new key:", _
        providerName & " API Key", _
        currentKey)
    
    If response = "" Then Exit Sub
    apiKey = response
    
    ' Get API URL using regular InputBox
    response = InputBox( _
        "Enter " & providerName & " API URL:" & vbCrLf & vbCrLf & _
        "Default: " & defaultURL & vbCrLf & vbCrLf & _
        "Press OK to use default, or enter custom URL:", _
        providerName & " API URL", _
        defaultURL)
    
    If response = "" Then
        apiURL = defaultURL
    Else
        apiURL = response
    End If
    
    ' Save settings
    Select Case providerKey
        Case "openai"
            If apiKey <> "" Then OPENAI_API_KEY = apiKey
            OPENAI_URL = apiURL
        Case "mistral"
            If apiKey <> "" Then MISTRAL_API_KEY = apiKey
            MISTRAL_URL = apiURL
        Case "nebius"
            If apiKey <> "" Then NEBIUS_API_KEY = apiKey
            NEBIUS_URL = apiURL
        Case "scaleway"
            If apiKey <> "" Then SCALEWAY_API_KEY = apiKey
            SCALEWAY_URL = apiURL
        Case "openrouter"
            If apiKey <> "" Then OPENROUTER_API_KEY = apiKey
            OPENROUTER_URL = apiURL
    End Select
    
    ' Ask if this should be default
    Dim msgResponse As Integer
    msgResponse = MsgBox( _
        "Set " & providerName & " as default provider?" & vbCrLf & vbCrLf & _
        "Suggested model: " & suggestedModel, _
        vbYesNo + vbQuestion, _
        "Set Default?")
    
    If msgResponse = vbYes Then
        CurrentProvider = providerKey
        CurrentModel = suggestedModel
    End If
    
    Call SaveConfig
    MsgBox providerName & " configured successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    If DEBUG_MODE Then Debug.Print "[ConfigureProvider] ERROR: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Set default model - USING InputBox
Private Sub SetDefaultModel()
    Dim response As String
    
    response = InputBox( _
        "Current Provider: " & CurrentProvider & vbCrLf & _
        "Current Model: " & CurrentModel & vbCrLf & vbCrLf & _
        "Enter new default model name:", _
        "Set Default Model", _
        CurrentModel)
    
    If response = "" Then Exit Sub
    
    CurrentModel = response
    Call SaveConfig
    
    MsgBox "Default model set to: " & CurrentModel, vbInformation
End Sub

' Show current configuration
Private Sub ShowCurrentConfig()
    Dim msg As String
    
    msg = "Current Configuration:" & vbCrLf & vbCrLf
    msg = msg & "Provider: " & CurrentProvider & vbCrLf
    msg = msg & "Model: " & CurrentModel & vbCrLf & vbCrLf
    
    msg = msg & "Ollama URL: " & OLLAMA_BASE_URL & vbCrLf
    msg = msg & "OpenAI URL: " & OPENAI_URL & vbCrLf & vbCrLf
    
    msg = msg & "API Keys:" & vbCrLf
    msg = msg & "OpenAI: " & IIf(OPENAI_API_KEY = "", "(not set)", "***" & Right(OPENAI_API_KEY, 4)) & vbCrLf
    msg = msg & "Mistral: " & IIf(MISTRAL_API_KEY = "", "(not set)", "***" & Right(MISTRAL_API_KEY, 4)) & vbCrLf
    msg = msg & "Nebius: " & IIf(NEBIUS_API_KEY = "", "(not set)", "***" & Right(NEBIUS_API_KEY, 4)) & vbCrLf
    msg = msg & "Scaleway: " & IIf(SCALEWAY_API_KEY = "", "(not set)", "***" & Right(SCALEWAY_API_KEY, 4)) & vbCrLf
    msg = msg & "OpenRouter: " & IIf(OPENROUTER_API_KEY = "", "(not set)", "***" & Right(OPENROUTER_API_KEY, 4)) & vbCrLf & vbCrLf
    
    msg = msg & "Config file: " & GetConfigLocation() & vbCrLf
    msg = msg & "File exists: " & IIf(Dir(GetConfigLocation()) <> "", "Yes", "No")
    
    MsgBox msg, vbInformation, "Current Configuration"
End Sub

' Quick test
Public Sub QuickTest()
    Dim result As String
    
    If DEBUG_MODE Then Debug.Print "=== QuickTest START ==="
    
    ' Ensure config is loaded
    If CurrentProvider = "" Then
        Call LoadConfig
    End If
    
    If CurrentProvider = "" Then
        MsgBox "No provider configured! Run ShowSettings first.", vbCritical
        Exit Sub
    End If
    
    Application.StatusBar = "Testing connection to " & CurrentProvider & "..."
    result = prompt("Say 'Hello Excel' in a friendly way")
    Application.StatusBar = False
    
    If DEBUG_MODE Then Debug.Print "[QuickTest] Result: " & Left(result, 100)
    
    If Left(result, 6) = "Error:" Then
        MsgBox "Test FAILED:" & vbCrLf & vbCrLf & _
               "Provider: " & CurrentProvider & vbCrLf & _
               "Model: " & CurrentModel & vbCrLf & vbCrLf & _
               result, vbCritical, "Connection Test"
    Else
        MsgBox "Test SUCCESSFUL!" & vbCrLf & vbCrLf & _
               "Provider: " & CurrentProvider & vbCrLf & _
               "Model: " & CurrentModel & vbCrLf & vbCrLf & _
               "Response: " & result, vbInformation, "Connection Test"
    End If
    
    If DEBUG_MODE Then Debug.Print "=== QuickTest END ==="
End Sub

' Full diagnostic
Public Sub FullDiagnostic()
    Dim msg As String
    Dim testResult As String
    
    msg = "=== FULL DIAGNOSTIC ===" & vbCrLf & vbCrLf
    
    ' Config status
    msg = msg & "Config Variables:" & vbCrLf
    msg = msg & "CurrentProvider = '" & CurrentProvider & "'" & vbCrLf
    msg = msg & "CurrentModel = '" & CurrentModel & "'" & vbCrLf
    msg = msg & "OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'" & vbCrLf
    msg = msg & "OPENAI_URL = '" & OPENAI_URL & "'" & vbCrLf & vbCrLf
    
    ' Config file
    msg = msg & "Config File:" & vbCrLf
    msg = msg & "Location: " & GetConfigLocation() & vbCrLf
    msg = msg & "Exists: " & IIf(Dir(GetConfigLocation()) <> "", "Yes", "No") & vbCrLf & vbCrLf
    
    ' Test GetBaseURL
    msg = msg & "GetBaseURL Tests:" & vbCrLf
    msg = msg & "GetBaseURL(""ollama"") = '" & GetBaseURL("ollama") & "'" & vbCrLf
    msg = msg & "GetBaseURL(""openai"") = '" & GetBaseURL("openai") & "'" & vbCrLf & vbCrLf
    
    ' Test PROMPT function
    msg = msg & "Function Tests:" & vbCrLf
    On Error Resume Next
    testResult = prompt("test", "ollama", "ministral-3:3b-instruct-2512-q4_K_M")
    If Err.Number = 0 Then
        msg = msg & "PROMPT function: OK" & vbCrLf
        msg = msg & "Result: " & Left(testResult, 50) & vbCrLf
    Else
        msg = msg & "PROMPT function: ERROR - " & Err.Description & vbCrLf
    End If
    On Error GoTo 0
    
    msg = msg & vbCrLf & "DEBUG_MODE: " & IIf(DEBUG_MODE, "ON", "OFF") & vbCrLf
    msg = msg & vbCrLf & "** Press OK, then run QuickTest **"
    If DEBUG_MODE Then
        msg = msg & vbCrLf & "Check Immediate Window (Cmd+G) for detailed logs"
    End If
    
    MsgBox msg, vbInformation, "Full Diagnostic"
    
    Debug.Print msg
End Sub

' Simple menu (for testing)
Public Sub ShowSettingsSimple()
    Call ShowSettings
End Sub

' Test curl connectivity
Public Sub TestCurlConnection()
    If DEBUG_MODE Then Debug.Print "=== TestCurlConnection START ==="
    
    Dim result As String
    
    MsgBox "This will test if curl can connect to Ollama." & vbCrLf & vbCrLf & _
           "Make sure Ollama is running first!" & vbCrLf & vbCrLf & _
           "Run 'ollama serve' in Terminal if not running.", vbInformation
    
    result = TestCurl()
    
    If DEBUG_MODE Then Debug.Print "[TestCurlConnection] Result: " & result
    
    If Left(result, 6) = "Error:" Then
        MsgBox "Curl test FAILED:" & vbCrLf & vbCrLf & result & vbCrLf & vbCrLf & _
               "Check:" & vbCrLf & _
               "1. Ollama is running (ollama serve)" & vbCrLf & _
               "2. Check Immediate Window (Cmd+G) for details", vbCritical
    Else
        MsgBox "Curl test SUCCESS!" & vbCrLf & vbCrLf & _
               "Response: " & result, vbInformation
    End If
    
    If DEBUG_MODE Then Debug.Print "=== TestCurlConnection END ==="
End Sub
