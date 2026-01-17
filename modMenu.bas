' modMenu:

Option Explicit

' DEBUG MODE
Public Const DEBUG_MODE As Boolean = True

' Auto_Open
Public Sub Auto_Open()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "=== AUTO_OPEN STARTING ==="
    Debug.Print "=== Time: " & Now
    Debug.Print "========================================="
    
    Debug.Print "[AUTO_OPEN] Before LoadConfig:"
    Debug.Print "[AUTO_OPEN]   CurrentProvider = '" & CurrentProvider & "'"
    Debug.Print "[AUTO_OPEN]   CurrentModel = '" & CurrentModel & "'"
    Debug.Print "[AUTO_OPEN]   OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
    
    Debug.Print "[AUTO_OPEN] Calling LoadConfig..."
    Call LoadConfig
    
    Debug.Print "[AUTO_OPEN] After LoadConfig:"
    Debug.Print "[AUTO_OPEN]   CurrentProvider = '" & CurrentProvider & "'"
    Debug.Print "[AUTO_OPEN]   CurrentModel = '" & CurrentModel & "'"
    Debug.Print "[AUTO_OPEN]   OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
    Debug.Print "[AUTO_OPEN]   OPENAI_URL = '" & OPENAI_URL & "'"
    
    Debug.Print "[AUTO_OPEN] Calling SetupKeyboardShortcuts..."
    Call SetupKeyboardShortcuts
    
    Dim msg As String
    msg = "LLM Excel Add-in loaded!" & vbCrLf & vbCrLf & _
          "Current: " & CurrentProvider & " / " & CurrentModel & vbCrLf & vbCrLf & _
          "Run from Tools > Macro > Macros:" & vbCrLf & _
          "• ShowSettings - Configure providers" & vbCrLf & _
          "• QuickTest - Test connection" & vbCrLf & _
          "• FullDiagnostic - Show diagnostics" & vbCrLf & vbCrLf & _
          "(Keyboard shortcuts may not work on Mac)"
    
    If DEBUG_MODE Then
        msg = msg & vbCrLf & vbCrLf & "DEBUG MODE: ON" & vbCrLf & _
              "Press ?+G in VBA to see Immediate Window"
    End If
    
    MsgBox msg, vbInformation, "LLM Add-in"
    Debug.Print "=== AUTO_OPEN COMPLETE ==="
    Debug.Print "========================================="
    Debug.Print ""
End Sub

' Auto_Close
Public Sub Auto_Close()
    Debug.Print "[AUTO_CLOSE] Called"
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
    
    Debug.Print "[SHORTCUTS] Setting up keyboard shortcuts..."
    
    Application.OnKey "^+L", "ShowSettings"
    If Err.Number <> 0 Then
        Debug.Print "[SHORTCUTS]   Warning: Could not set ^+L: " & Err.Description
        Err.Clear
    Else
        Debug.Print "[SHORTCUTS]   ? ^+L -> ShowSettings"
    End If
    
    Application.OnKey "^+M", "ShowModelSelector"
    If Err.Number <> 0 Then
        Debug.Print "[SHORTCUTS]   Warning: Could not set ^+M: " & Err.Description
        Err.Clear
    Else
        Debug.Print "[SHORTCUTS]   ? ^+M -> ShowModelSelector"
    End If
    
    Application.OnKey "^+T", "QuickTest"
    If Err.Number <> 0 Then
        Debug.Print "[SHORTCUTS]   Warning: Could not set ^+T: " & Err.Description
        Err.Clear
    Else
        Debug.Print "[SHORTCUTS]   ? ^+T -> QuickTest"
    End If
    
    Application.OnKey "^+D", "FullDiagnostic"
    If Err.Number <> 0 Then
        Debug.Print "[SHORTCUTS]   Warning: Could not set ^+D: " & Err.Description
        Err.Clear
    Else
        Debug.Print "[SHORTCUTS]   ? ^+D -> FullDiagnostic"
    End If
    
    On Error GoTo 0
    Debug.Print "[SHORTCUTS] Setup complete"
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
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "=== ShowSettings START ==="
    Debug.Print "=== Time: " & Now
    Debug.Print "========================================="
    
    On Error GoTo ErrorHandler
    
    ' Ensure config is loaded
    Debug.Print "[ShowSettings] Checking config..."
    Debug.Print "[ShowSettings]   CurrentProvider = '" & CurrentProvider & "'"
    Debug.Print "[ShowSettings]   OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
    
    If CurrentProvider = "" Or OLLAMA_BASE_URL = "" Then
        Debug.Print "[ShowSettings] Config appears empty, calling LoadConfig..."
        Call LoadConfig
        Debug.Print "[ShowSettings] After LoadConfig:"
        Debug.Print "[ShowSettings]   CurrentProvider = '" & CurrentProvider & "'"
        Debug.Print "[ShowSettings]   OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
    End If
    
    Dim choice As String
    Dim providers As String
    
    providers = "1=OpenAI, 2=Mistral, 3=Nebius, 4=Scaleway, 5=OpenRouter, 6=Ollama"
    
    Debug.Print "[ShowSettings] Showing InputBox..."
    
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
    
    Debug.Print "[ShowSettings] InputBox returned: '" & choice & "'"
    
    ' Check if user cancelled (empty string)
    If choice = "" Then
        Debug.Print "[ShowSettings] User cancelled"
        Debug.Print "=== ShowSettings END (cancelled) ==="
        Debug.Print "========================================="
        Exit Sub
    End If
    
    ' Convert to uppercase and trim
    choice = UCase(Trim(choice))
    Debug.Print "[ShowSettings] User choice: '" & choice & "'"
    
    ' Handle the choice
    Select Case choice
        Case "1"
            Debug.Print "[ShowSettings] -> ConfigureProvider(openai)"
            Call ConfigureProvider("openai", "OpenAI")
            Call ShowSettings
            
        Case "2"
            Debug.Print "[ShowSettings] -> ConfigureProvider(mistral)"
            Call ConfigureProvider("mistral", "Mistral")
            Call ShowSettings
            
        Case "3"
            Debug.Print "[ShowSettings] -> ConfigureProvider(nebius)"
            Call ConfigureProvider("nebius", "Nebius")
            Call ShowSettings
            
        Case "4"
            Debug.Print "[ShowSettings] -> ConfigureProvider(scaleway)"
            Call ConfigureProvider("scaleway", "Scaleway")
            Call ShowSettings
            
        Case "5"
            Debug.Print "[ShowSettings] -> ConfigureProvider(openrouter)"
            Call ConfigureProvider("openrouter", "OpenRouter")
            Call ShowSettings
            
        Case "6"
            Debug.Print "[ShowSettings] -> ConfigureOllama"
            Call ConfigureOllama
            Call ShowSettings
            
        Case "7"
            Debug.Print "[ShowSettings] -> SetDefaultModel"
            Call SetDefaultModel
            Call ShowSettings
            
        Case "8"
            Debug.Print "[ShowSettings] -> ShowCurrentConfig"
            Call ShowCurrentConfig
            Call ShowSettings
            
        Case "9"
            Debug.Print "[ShowSettings] -> QuickTest"
            Call QuickTest
            Call ShowSettings
            
        Case "D"
            Debug.Print "[ShowSettings] -> FullDiagnostic"
            Call FullDiagnostic
            Call ShowSettings

        Case "M"
            Debug.Print "[ShowSettings] -> SelectOllamaModel"
            Call SelectOllamaModel
            Call ShowSettings
            
        Case "0"
            Debug.Print "[ShowSettings] User chose to exit"
            
        Case Else
            Debug.Print "[ShowSettings] Invalid input: " & choice
            MsgBox "Please enter a number from 0 to 9, or D for diagnostics", vbExclamation, "Invalid Input"
            Call ShowSettings
    End Select
    
    Debug.Print "=== ShowSettings END ==="
    Debug.Print "========================================="
    Debug.Print ""
    Exit Sub
    
ErrorHandler:
    Debug.Print "*** ShowSettings ERROR ***"
    Debug.Print "   Error Number: " & Err.Number
    Debug.Print "   Error Description: " & Err.Description
    Debug.Print "========================================="
    MsgBox "Error in Settings: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' Configure Ollama - USING InputBox
Private Sub ConfigureOllama()
    Debug.Print ""
    Debug.Print "=== ConfigureOllama START ==="
    
    On Error GoTo ErrorHandler
    
    ' Ensure defaults
    If OLLAMA_BASE_URL = "" Then
        OLLAMA_BASE_URL = "http://localhost:11434"
    End If
    
    Dim newURL As String
    Dim newModel As String
    Dim response As Integer
    
    Debug.Print "[ConfigureOllama] Current URL: '" & OLLAMA_BASE_URL & "'"
    Debug.Print "[ConfigureOllama] Current Model: '" & CurrentModel & "'"
    
    ' Use regular InputBox (works on Mac)
    newURL = InputBox( _
        "Enter Ollama Base URL:" & vbCrLf & vbCrLf & _
        "Default: http://localhost:11434" & vbCrLf & _
        "Current: " & OLLAMA_BASE_URL, _
        "Ollama Configuration", _
        OLLAMA_BASE_URL)
    
    Debug.Print "[ConfigureOllama] User entered URL: '" & newURL & "'"
    
    If newURL = "" Then
        Debug.Print "[ConfigureOllama] User cancelled URL input"
        Exit Sub
    End If
    
    OLLAMA_BASE_URL = Trim(newURL)
    Debug.Print "[ConfigureOllama] Set OLLAMA_BASE_URL = '" & OLLAMA_BASE_URL & "'"
    
    ' Ask for model name
    newModel = InputBox( _
        "Enter Ollama model name:" & vbCrLf & vbCrLf & _
        "Recommended: ministral-3:3b-instruct-2512-q4_K_M" & vbCrLf & _
        "Fallback: llama-3.2-3b-instruct:latest" & vbCrLf & _
        "Current: " & CurrentModel, _
        "Ollama Model", _
        CurrentModel)
    
    Debug.Print "[ConfigureOllama] User entered model: '" & newModel & "'"
    
    If newModel = "" Then
        Debug.Print "[ConfigureOllama] User cancelled model input"
        Exit Sub
    End If
    
    response = MsgBox( _
        "Set Ollama as default provider?" & vbCrLf & vbCrLf & _
        "Model: " & newModel & vbCrLf & _
        "URL: " & OLLAMA_BASE_URL, _
        vbYesNo + vbQuestion, _
        "Set Default?")
    
    If response = vbYes Then
        CurrentProvider = "ollama"
        CurrentModel = Trim(newModel)
        Debug.Print "[ConfigureOllama] Set provider='ollama', model='" & CurrentModel & "'"
    End If
    
    Debug.Print "[ConfigureOllama] Calling SaveConfig..."
    Call SaveConfig
    
    MsgBox "Ollama configured!" & vbCrLf & _
           "Provider: " & CurrentProvider & vbCrLf & _
           "Model: " & CurrentModel & vbCrLf & _
           "URL: " & OLLAMA_BASE_URL, vbInformation
    
    Debug.Print "=== ConfigureOllama END ===" & vbCrLf
    Exit Sub
    
ErrorHandler:
    Debug.Print "*** ConfigureOllama ERROR: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Configure a provider - USING InputBox
Private Sub ConfigureProvider(providerKey As String, providerName As String)
    Debug.Print "[ConfigureProvider] Called for: " & providerKey
    
    Dim apiKey As String
    Dim apiURL As String
    Dim suggestedModel As String
    Dim defaultURL As String
    Dim response As String
    Dim currentKey As String
    
    On Error GoTo ErrorHandler
    
    ' Get current values
    Select Case providerKey
        Case "openai"
            currentKey = OPENAI_API_KEY
            defaultURL = "https://api.openai.com/v1"
            suggestedModel = "gpt-3.5-turbo"
        Case "mistral"
            currentKey = MISTRAL_API_KEY
            defaultURL = "https://api.mistral.ai/v1"
            suggestedModel = "mistral-small-latest"
        Case "nebius"
            currentKey = NEBIUS_API_KEY
            defaultURL = "https://api.studio.nebius.ai/v1"
            suggestedModel = "meta-llama/Meta-Llama-3.1-8B-Instruct"
        Case "scaleway"
            currentKey = SCALEWAY_API_KEY
            defaultURL = "https://api.scaleway.ai/v1"
            suggestedModel = "llama-3.1-8b-instruct"
        Case "openrouter"
            currentKey = OPENROUTER_API_KEY
            defaultURL = "https://openrouter.ai/api/v1"
            suggestedModel = "meta-llama/llama-3.2-3b-instruct:free"
    End Select
    
    ' Get API Key using regular InputBox
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
    
    Debug.Print "[ConfigureProvider] Setting URL to: " & apiURL
    
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
        Debug.Print "[ConfigureProvider] Set defaults: " & providerKey & " / " & suggestedModel
    End If
    
    Call SaveConfig
    MsgBox providerName & " configured successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    Debug.Print "[ConfigureProvider] ERROR: " & Err.Description
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
    
    Debug.Print ""
    Debug.Print "=== QuickTest START ==="
    
    ' Ensure config is loaded
    If CurrentProvider = "" Then
        Debug.Print "[QuickTest] Calling LoadConfig..."
        Call LoadConfig
    End If
    
    Debug.Print "[QuickTest] Provider='" & CurrentProvider & "', Model='" & CurrentModel & "'"
    
    If CurrentProvider = "" Then
        MsgBox "No provider configured! Run ShowSettings first.", vbCritical
        Exit Sub
    End If
    
    Application.StatusBar = "Testing connection to " & CurrentProvider & "..."
    result = prompt("Say 'Hello Excel' in a friendly way")
    Application.StatusBar = False
    
    Debug.Print "[QuickTest] Result: " & Left(result, 100)
    
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
    
    Debug.Print "=== QuickTest END ==="
    Debug.Print ""
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
    
    msg = msg & vbCrLf & "** Press OK, then run QuickTest **"
    msg = msg & vbCrLf & "Check Immediate Window (?+G) for detailed logs"
    
    MsgBox msg, vbInformation, "Full Diagnostic"
    
    Debug.Print msg
End Sub

' Simple menu (for testing)
Public Sub ShowSettingsSimple()
    Debug.Print "[ShowSettingsSimple] Called"
    Call ShowSettings
End Sub

' Test curl connectivity
Public Sub TestCurlConnection()
    Debug.Print ""
    Debug.Print "=== TestCurlConnection START ==="
    
    Dim result As String
    
    MsgBox "This will test if curl can connect to Ollama." & vbCrLf & vbCrLf & _
           "Make sure Ollama is running first!" & vbCrLf & vbCrLf & _
           "Run 'ollama serve' in Terminal if not running.", vbInformation
    
    result = TestCurl()
    
    Debug.Print "[TestCurlConnection] Result: " & result
    
    If Left(result, 6) = "Error:" Then
        MsgBox "Curl test FAILED:" & vbCrLf & vbCrLf & result & vbCrLf & vbCrLf & _
               "Check:" & vbCrLf & _
               "1. Ollama is running (ollama serve)" & vbCrLf & _
               "2. Check Immediate Window (?+G) for details", vbCritical
    Else
        MsgBox "Curl test SUCCESS!" & vbCrLf & vbCrLf & _
               "Response: " & result, vbInformation
    End If
    
    Debug.Print "=== TestCurlConnection END ==="
End Sub

